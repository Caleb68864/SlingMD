using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// A calendar folder opened for export. The caller owns <see cref="Folder"/> and must release it.
    /// </summary>
    public sealed class OpenedCalendar
    {
        public OpenedCalendar(CalendarFolderInfo info, MAPIFolder folder)
        {
            Info = info;
            Folder = folder;
        }

        public CalendarFolderInfo Info { get; }

        public MAPIFolder Folder { get; }
    }

    /// <summary>
    /// Finds the calendars Outlook can see and reopens the ones chosen for bulk appointment export.
    /// Two sources are combined: every store's default calendar and the sub-calendars beneath it
    /// (additional accounts, auto-mapped shared mailboxes, user-created calendars), and the Calendar
    /// navigation pane, which is the only object-model route to a calendar another user shared
    /// without full mailbox access, and to internet and group calendars.
    /// </summary>
    public class CalendarFolderService
    {
        // Sub-calendars sit under a store's Calendar folder, rarely more than one level deep. The cap
        // only keeps a pathological folder tree from stalling the ribbon click.
        private const int MaxSubCalendarDepth = 5;

        private readonly Application _application;

        public CalendarFolderService(Application application)
        {
            _application = application ?? throw new ArgumentNullException(nameof(application));
        }

        /// <summary>
        /// Lists every calendar Outlook exposes, with each account's default calendar flagged via
        /// <see cref="CalendarFolderInfo.IsAccountDefault"/>. A folder reachable from both sources
        /// is listed once.
        /// </summary>
        public List<CalendarFolderInfo> GetAvailableCalendars()
        {
            List<CalendarFolderInfo> defaults = GetAccountDefaultCalendars();
            List<CalendarFolderInfo> calendars = new List<CalendarFolderInfo>();

            VisitStoreCalendars((label, folder) =>
            {
                AddUnique(calendars, Describe(folder, label), defaults);
                return false;
            });
            VisitNavigationPaneCalendars((label, folder) =>
            {
                AddUnique(calendars, Describe(folder, label), defaults);
                return false;
            });

            // An account default neither walk reached is still offered, so it can be pre-selected.
            foreach (CalendarFolderInfo calendar in defaults)
            {
                AddUnique(calendars, calendar, defaults);
            }

            return calendars;
        }

        /// <summary>
        /// Returns the default calendar of each configured account, falling back to the session's
        /// default calendar when no account yields one. This is what bulk export uses when no
        /// calendars have been chosen.
        /// </summary>
        public List<CalendarFolderInfo> GetAccountDefaultCalendars()
        {
            List<CalendarFolderInfo> defaults = new List<CalendarFolderInfo>();
            NameSpace session = null;
            Accounts accounts = null;
            try
            {
                session = _application.Session;
                accounts = session.Accounts;
                foreach (Account account in accounts)
                {
                    Store store = null;
                    MAPIFolder calendar = null;
                    try
                    {
                        // account.Session.GetDefaultFolder would return the PRIMARY store's calendar
                        // for every account; the account's own calendar lives in its delivery store.
                        store = account.DeliveryStore;
                        calendar = store.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                        AddDefault(defaults, Describe(calendar, FormatCalendarLabel(calendar.Name, store.DisplayName)));
                    }
                    catch (COMException ex)
                    {
                        // Some account types have no delivery store or no calendar folder.
                        Logger.Instance.Debug($"CalendarFolderService: account has no default calendar: {ex.Message}");
                    }
                    finally
                    {
                        if (calendar != null) Marshal.ReleaseComObject(calendar);
                        if (store != null) Marshal.ReleaseComObject(store);
                        if (account != null) Marshal.ReleaseComObject(account);
                    }
                }

                if (defaults.Count == 0)
                {
                    MAPIFolder calendar = null;
                    try
                    {
                        calendar = session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                        AddDefault(defaults, Describe(calendar, calendar.Name));
                    }
                    finally
                    {
                        if (calendar != null) Marshal.ReleaseComObject(calendar);
                    }
                }
            }
            catch (COMException ex)
            {
                Logger.Instance.Warning($"CalendarFolderService: could not read default calendars: {ex.Message}");
            }
            finally
            {
                if (accounts != null) Marshal.ReleaseComObject(accounts);
                if (session != null) Marshal.ReleaseComObject(session);
            }

            return defaults;
        }

        /// <summary>
        /// Opens the requested calendars. Each is opened by EntryID first; any that fail are looked
        /// up in the Calendar navigation pane, because a calendar shared by another user can refuse
        /// to open by ID while the pane still holds the live folder. Calendars found by neither route
        /// are added to <paramref name="unavailable"/>. The caller must release every returned folder.
        /// </summary>
        public List<OpenedCalendar> OpenCalendars(IEnumerable<CalendarFolderInfo> wanted, List<CalendarFolderInfo> unavailable)
        {
            List<OpenedCalendar> opened = new List<OpenedCalendar>();
            List<CalendarFolderInfo> pending = new List<CalendarFolderInfo>();

            NameSpace session = null;
            try
            {
                session = _application.Session;
                foreach (CalendarFolderInfo calendar in wanted ?? Enumerable.Empty<CalendarFolderInfo>())
                {
                    if (calendar == null
                        || string.IsNullOrEmpty(calendar.EntryId)
                        || opened.Any(o => o.Info.IsSameFolder(calendar))
                        || pending.Any(p => p.IsSameFolder(calendar)))
                    {
                        continue;
                    }

                    MAPIFolder folder = null;
                    try
                    {
                        folder = string.IsNullOrEmpty(calendar.StoreId)
                            ? session.GetFolderFromID(calendar.EntryId)
                            : session.GetFolderFromID(calendar.EntryId, calendar.StoreId);
                    }
                    catch (COMException ex)
                    {
                        Logger.Instance.Debug($"CalendarFolderService: could not open '{calendar.DisplayName}' by ID: {ex.Message}");
                    }

                    if (folder != null)
                    {
                        opened.Add(new OpenedCalendar(calendar, folder));
                    }
                    else
                    {
                        pending.Add(calendar);
                    }
                }
            }
            finally
            {
                if (session != null) Marshal.ReleaseComObject(session);
            }

            if (pending.Count > 0)
            {
                VisitNavigationPaneCalendars((label, folder) =>
                {
                    string entryId;
                    try
                    {
                        entryId = folder.EntryID;
                    }
                    catch (COMException)
                    {
                        return false;
                    }

                    CalendarFolderInfo match = pending.FirstOrDefault(
                        p => string.Equals(p.EntryId, entryId, StringComparison.OrdinalIgnoreCase));
                    if (match == null)
                    {
                        return false;
                    }

                    pending.Remove(match);
                    opened.Add(new OpenedCalendar(match, folder));
                    return true;
                });
            }

            unavailable?.AddRange(pending);
            return opened;
        }

        /// <summary>
        /// Calendars to pre-check in the Save Date Range dialog: the saved choice, limited to the
        /// calendars that still exist, or every account default when nothing saved is available.
        /// </summary>
        public static List<CalendarFolderInfo> GetInitialSelection(
            IEnumerable<CalendarFolderInfo> available,
            IEnumerable<CalendarFolderInfo> saved)
        {
            List<CalendarFolderInfo> availableList = (available ?? Enumerable.Empty<CalendarFolderInfo>())
                .Where(c => c != null)
                .ToList();
            List<CalendarFolderInfo> savedList = (saved ?? Enumerable.Empty<CalendarFolderInfo>())
                .Where(c => c != null)
                .ToList();

            List<CalendarFolderInfo> stillAvailable = availableList
                .Where(a => savedList.Any(s => s.IsSameFolder(a)))
                .ToList();
            if (stillAvailable.Count > 0)
            {
                return stillAvailable;
            }

            return availableList.Where(a => a.IsAccountDefault).ToList();
        }

        /// <summary>
        /// Builds the list label for a calendar: its name, qualified by the store or navigation group
        /// it came from so the several folders all named "Calendar" can be told apart. The qualifier
        /// is dropped when the name already contains it (shared calendars are often named after their
        /// owner).
        /// </summary>
        internal static string FormatCalendarLabel(string folderName, string qualifier)
        {
            string name = string.IsNullOrWhiteSpace(folderName) ? "Calendar" : folderName.Trim();
            if (string.IsNullOrWhiteSpace(qualifier))
            {
                return name;
            }

            string trimmedQualifier = qualifier.Trim();
            if (name.IndexOf(trimmedQualifier, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return name;
            }

            return string.Format("{0} ({1})", name, trimmedQualifier);
        }

        private static void AddUnique(List<CalendarFolderInfo> calendars, CalendarFolderInfo calendar, List<CalendarFolderInfo> defaults)
        {
            if (calendar == null || calendars.Any(c => c.IsSameFolder(calendar)))
            {
                return;
            }

            calendar.IsAccountDefault = defaults.Any(d => d.IsSameFolder(calendar));
            calendars.Add(calendar);
        }

        private static void AddDefault(List<CalendarFolderInfo> defaults, CalendarFolderInfo calendar)
        {
            if (calendar == null || defaults.Any(d => d.IsSameFolder(calendar)))
            {
                return;
            }

            calendar.IsAccountDefault = true;
            defaults.Add(calendar);
        }

        private static CalendarFolderInfo Describe(MAPIFolder folder, string label)
        {
            try
            {
                string entryId = folder.EntryID;
                if (string.IsNullOrEmpty(entryId))
                {
                    return null;
                }

                return new CalendarFolderInfo
                {
                    EntryId = entryId,
                    StoreId = folder.StoreID ?? string.Empty,
                    DisplayName = label ?? string.Empty
                };
            }
            catch (COMException ex)
            {
                Logger.Instance.Debug($"CalendarFolderService: could not read IDs for '{label}': {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Walks each store's default calendar and its calendar subfolders. <paramref name="visit"/>
        /// receives the label and folder and returns true when it keeps the folder (and so must
        /// release it); otherwise the folder is released here.
        /// </summary>
        private void VisitStoreCalendars(Func<string, MAPIFolder, bool> visit)
        {
            NameSpace session = null;
            Stores stores = null;
            try
            {
                session = _application.Session;
                stores = session.Stores;
                foreach (Store store in stores)
                {
                    MAPIFolder calendar = null;
                    bool kept = false;
                    try
                    {
                        string storeName = store.DisplayName;
                        calendar = store.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                        kept = visit(FormatCalendarLabel(calendar.Name, storeName), calendar);
                        VisitSubCalendars(calendar, storeName, visit, 1);
                    }
                    catch (COMException ex)
                    {
                        // Public folders, online archives and some PSTs have no calendar.
                        Logger.Instance.Debug($"CalendarFolderService: store has no calendar: {ex.Message}");
                    }
                    finally
                    {
                        if (calendar != null && !kept) Marshal.ReleaseComObject(calendar);
                        if (store != null) Marshal.ReleaseComObject(store);
                    }
                }
            }
            catch (COMException ex)
            {
                Logger.Instance.Warning($"CalendarFolderService: could not enumerate stores: {ex.Message}");
            }
            finally
            {
                if (stores != null) Marshal.ReleaseComObject(stores);
                if (session != null) Marshal.ReleaseComObject(session);
            }
        }

        private void VisitSubCalendars(MAPIFolder parent, string storeName, Func<string, MAPIFolder, bool> visit, int depth)
        {
            if (depth > MaxSubCalendarDepth)
            {
                return;
            }

            Folders subFolders = null;
            try
            {
                subFolders = parent.Folders;
                foreach (MAPIFolder subFolder in subFolders)
                {
                    bool kept = false;
                    try
                    {
                        if (subFolder.DefaultItemType == OlItemType.olAppointmentItem)
                        {
                            kept = visit(FormatCalendarLabel(subFolder.Name, storeName), subFolder);
                            VisitSubCalendars(subFolder, storeName, visit, depth + 1);
                        }
                    }
                    catch (COMException ex)
                    {
                        Logger.Instance.Debug($"CalendarFolderService: skipped a subfolder: {ex.Message}");
                    }
                    finally
                    {
                        if (!kept) Marshal.ReleaseComObject(subFolder);
                    }
                }
            }
            catch (COMException ex)
            {
                Logger.Instance.Debug($"CalendarFolderService: could not list subfolders: {ex.Message}");
            }
            finally
            {
                if (subFolders != null) Marshal.ReleaseComObject(subFolders);
            }
        }

        /// <summary>
        /// Walks the calendars listed in the Calendar navigation pane (My Calendars, Shared Calendars,
        /// Other Calendars and any user-created groups). Ownership follows the same rule as
        /// <see cref="VisitStoreCalendars"/>. Does nothing when Outlook has no explorer window.
        /// </summary>
        private void VisitNavigationPaneCalendars(Func<string, MAPIFolder, bool> visit)
        {
            Explorer explorer = null;
            NavigationPane pane = null;
            NavigationModules modules = null;
            NavigationModule module = null;
            NavigationGroups groups = null;
            try
            {
                explorer = _application.ActiveExplorer();
                if (explorer == null)
                {
                    return;
                }

                pane = explorer.NavigationPane;
                modules = pane.Modules;
                module = modules.GetNavigationModule(OlNavigationModuleType.olModuleCalendar);
                CalendarModule calendarModule = module as CalendarModule;
                if (calendarModule == null)
                {
                    return;
                }

                groups = calendarModule.NavigationGroups;
                foreach (NavigationGroup group in groups)
                {
                    NavigationFolders navigationFolders = null;
                    try
                    {
                        string groupName = group.Name;
                        navigationFolders = group.NavigationFolders;
                        foreach (NavigationFolder navigationFolder in navigationFolders)
                        {
                            MAPIFolder folder = null;
                            bool kept = false;
                            try
                            {
                                string displayName = navigationFolder.DisplayName;
                                folder = navigationFolder.Folder;
                                kept = visit(FormatCalendarLabel(displayName, groupName), folder);
                            }
                            catch (COMException ex)
                            {
                                // A shared calendar that is offline or no longer shared throws here.
                                Logger.Instance.Debug($"CalendarFolderService: skipped a navigation pane calendar: {ex.Message}");
                            }
                            finally
                            {
                                if (folder != null && !kept) Marshal.ReleaseComObject(folder);
                                if (navigationFolder != null) Marshal.ReleaseComObject(navigationFolder);
                            }
                        }
                    }
                    catch (COMException ex)
                    {
                        Logger.Instance.Debug($"CalendarFolderService: skipped a navigation group: {ex.Message}");
                    }
                    finally
                    {
                        if (navigationFolders != null) Marshal.ReleaseComObject(navigationFolders);
                        if (group != null) Marshal.ReleaseComObject(group);
                    }
                }
            }
            catch (COMException ex)
            {
                Logger.Instance.Warning($"CalendarFolderService: could not read the Calendar navigation pane: {ex.Message}");
            }
            finally
            {
                if (groups != null) Marshal.ReleaseComObject(groups);
                if (module != null) Marshal.ReleaseComObject(module);
                if (modules != null) Marshal.ReleaseComObject(modules);
                if (pane != null) Marshal.ReleaseComObject(pane);
                if (explorer != null) Marshal.ReleaseComObject(explorer);
            }
        }
    }
}
