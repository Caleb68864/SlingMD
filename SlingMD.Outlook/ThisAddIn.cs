using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using SlingMD.Outlook.Helpers;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using SlingMD.Outlook.Forms;
using SlingMD.Outlook.Ribbon;

namespace SlingMD.Outlook
{
    public partial class ThisAddIn
    {
        private ObsidianSettings _settings;
        private EmailProcessor _emailProcessor;
        private AppointmentProcessor _appointmentProcessor;
        private ContactProcessor _contactProcessor;
        private FileService _fileService;
        private NotificationService _notificationService;
        private FolderMonitorService _folderMonitorService;
        private AutoSlingService _autoSlingService;
        private FlagMonitorService _flagMonitorService;
        private SlingRibbon _ribbon;
        private Explorer _activeExplorer;
        private bool _startupComplete;

        // Shared across all monitor services (and preserved when they are recreated on settings save)
        // so an email seen by more than one monitor is processed only once.
        private readonly SlingMD.Outlook.Helpers.BoundedHashSet _autoSlingProcessedIds = new SlingMD.Outlook.Helpers.BoundedHashSet();

        // A control created on the UI thread so (a) a WindowsFormsSynchronizationContext is installed
        // on the main thread — making async monitor continuations resume on the UI/STA thread instead
        // of a thread-pool thread — and (b) background notifications can marshal onto the UI thread.
        private Control _uiMarshaler;

        // Serializes top-level user-initiated operations. The ribbon handlers all run on the single
        // Outlook UI thread but are `async void`, so a second click while one is still awaiting would
        // otherwise start an overlapping operation (double-processing the same item, stacked dialogs).
        // A non-atomic bool is safe because there is no true parallelism on this thread.
        private bool _operationInProgress;

        private bool TryBeginOperation()
        {
            if (_operationInProgress)
            {
                MessageBox.Show(
                    "SlingMD is already processing. Please wait for the current operation to finish.",
                    "SlingMD",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return false;
            }
            _operationInProgress = true;
            return true;
        }

        private void EndOperation()
        {
            _operationInProgress = false;
        }

        /// <summary>
        /// Creates the STA marshaling control and installs a WindowsFormsSynchronizationContext so
        /// background-monitor async continuations resume on the Outlook STA thread. Returns true only
        /// when both are in place; the caller must NOT start the monitors when this returns false,
        /// because their post-await COM/UI access would then run off-STA (crash / undefined behavior).
        /// </summary>
        private bool InitializeUiMarshaler()
        {
            try
            {
                _uiMarshaler = new Control();
                // Accessing Handle forces window-handle creation on this (UI/STA) thread, giving a
                // valid BeginInvoke target and causing WinForms to install a
                // WindowsFormsSynchronizationContext on the thread when one isn't already present.
                IntPtr handle = _uiMarshaler.Handle;

                if (System.Threading.SynchronizationContext.Current == null)
                {
                    System.Threading.SynchronizationContext.SetSynchronizationContext(
                        new WindowsFormsSynchronizationContext());
                }

                Forms.ToastForm.SetUiMarshaler(_uiMarshaler);

                bool ready = handle != IntPtr.Zero && System.Threading.SynchronizationContext.Current != null;
                if (!ready)
                {
                    Logger.Instance.Warning("ThisAddIn: UI marshaler initialized without a usable handle/sync context; background monitors will be disabled.");
                }
                return ready;
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ThisAddIn: could not initialize UI marshaler: {ex.Message}");
                return false;
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new SlingRibbon(this);
            return _ribbon;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // The whole startup path is guarded: an exception escaping a VSTO Startup handler can
            // cause Outlook to hard-disable the add-in (LoadBehavior=2) so the user must manually
            // re-enable it. Degrading to "loaded but some features off" is strictly better.
            try
            {
                bool isFirstLaunchAfterInstall;

                // Runs on the Outlook main (STA) thread. Create a marshaling control here so a
                // WindowsFormsSynchronizationContext is installed on this thread — without it, async
                // continuations in the background monitors would resume on a thread-pool thread and
                // touch Outlook COM / create WinForms windows off the STA thread. If this fails we
                // must not start the monitors at all (see below).
                bool marshalerReady = InitializeUiMarshaler();

                _settings = LoadSettings(out isFirstLaunchAfterInstall);
                bool vaultFound = ValidateStartupHealth();
                _emailProcessor = new EmailProcessor(_settings);
                _appointmentProcessor = new AppointmentProcessor(_settings);
                _contactProcessor = new ContactProcessor(_settings);
                _fileService = new FileService(_settings);
                _notificationService = new NotificationService(_settings);

                try
                {
                    _activeExplorer = Application.ActiveExplorer();
                    if (_activeExplorer != null)
                    {
                        _activeExplorer.SelectionChange += Explorer_SelectionChange;
                    }
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ThisAddIn.Startup: could not hook explorer selection change: {ex.Message}");
                }

                _startupComplete = true;

                // The background monitors rely on the marshaler to bounce their async continuations
                // back onto the STA thread before touching COM/UI. If the marshaler isn't ready,
                // starting them would resume off-STA — disable them and run in a degraded (manual
                // Sling still works) mode rather than risk crashing Outlook.
                if (marshalerReady)
                {
                    if (_settings.WatchedFolders != null && _settings.WatchedFolders.Count > 0)
                    {
                        _folderMonitorService = new FolderMonitorService(_settings, _emailProcessor, _notificationService, Application, _autoSlingProcessedIds);
                        _folderMonitorService.StartWatching(_settings.WatchedFolders);
                    }

                    _autoSlingService = new AutoSlingService(_settings, _emailProcessor, _notificationService, _autoSlingProcessedIds);
                    _autoSlingService.Start(Application);

                    if (_settings.EnableFlagToSling)
                    {
                        _flagMonitorService = new FlagMonitorService(_settings, _emailProcessor, _notificationService, _autoSlingProcessedIds);
                        _flagMonitorService.Start(Application);
                    }
                }
                else
                {
                    Logger.Instance.Warning("ThisAddIn.Startup: UI marshaler unavailable — auto-sling, flag, and folder monitors are disabled for this session. Manual Sling is unaffected.");
                }

                _ribbon?.UpdateSlingButtonLabel(GetSelectedItemLabel());

                // Offered once everything is running, because Settings restarts the services.
                if (!vaultFound)
                {
                    OfferVaultSetup();
                }

                if (isFirstLaunchAfterInstall && !_settings.HasShownSupportPrompt)
                {
                    ShowFirstRunSupportPrompt();
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"ThisAddIn.Startup failed; add-in will load in a degraded state: {ex.Message}", ex);
                // Ensure the add-in still considers itself started so callbacks don't NRE.
                _startupComplete = true;
            }
        }

        private void Explorer_SelectionChange()
        {
            _ribbon?.UpdateSlingButtonLabel(GetSelectedItemLabel());
        }

        public string GetSelectedItemLabel()
        {
            if (!_startupComplete)
            {
                return "Sling";
            }

            // This runs on every selection change, so COM objects obtained here must be released
            // or they accumulate for the whole session (a classic cause of Outlook failing to
            // close). Each `.Selection` property access mints a fresh COM object.
            Explorer explorer = null;
            Selection selection = null;
            object selected = null;
            try
            {
                explorer = Application.ActiveExplorer();
                if (explorer == null)
                {
                    return "Sling";
                }

                selection = explorer.Selection;
                if (selection == null || selection.Count == 0)
                {
                    return "Sling";
                }

                if (selection.Count > 1) return "Sling Selected";

                selected = selection[1];
                if (selected is MailItem) return "Sling Email";
                if (selected is AppointmentItem) return "Sling Appointment";
                if (selected is ContactItem) return "Sling Contact";
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Warning($"ThisAddIn.GetSelectedItemLabel: {ex.Message}");
            }
            finally
            {
                if (selected != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(selected);
                if (selection != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                if (explorer != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(explorer);
            }

            return "Sling";
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Signal shutdown FIRST so any in-flight OnItemChange/OnItemAdded continuation sees
            // _shuttingDown and bails out before touching COM objects we're about to release
            // (mirrors ShowSettings). Then stop/unhook and release the handles.
            _flagMonitorService?.SignalShutdown();
            _folderMonitorService?.SignalShutdown();

            // Guarded like Startup: an exception escaping a VSTO Shutdown handler risks Outlook
            // hard-disabling the add-in on next launch, and would skip the teardown steps and
            // settings save below. Unhooking/releasing an Explorer RCW during Outlook teardown is
            // a known COMException source.
            try
            {
                _autoSlingService?.Shutdown();
                _flagMonitorService?.Stop();
                _folderMonitorService?.StopWatching();

                if (_activeExplorer != null)
                {
                    _activeExplorer.SelectionChange -= Explorer_SelectionChange;
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_activeExplorer);
                    _activeExplorer = null;
                }

                _ribbon?.Dispose();

                if (_uiMarshaler != null)
                {
                    _uiMarshaler.Dispose();
                    _uiMarshaler = null;
                }
            }
            catch (System.Exception ex)
            {
                Logger.Instance.Error($"ThisAddIn.Shutdown: teardown failed: {ex.Message}");
            }

            if (_settings != null)
            {
                // Persist any late in-memory changes on the way out. This must never throw during
                // shutdown, and a failure here must not surface as a crash — the settings the user
                // saved via the dialog are already on disk (SettingsForm saves explicitly).
                try
                {
                    _settings.Save();
                }
                catch (System.Exception ex)
                {
                    Logger.Instance.Warning($"ThisAddIn.Shutdown: could not save settings: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Logs whether the configured vault exists. Returns false when it doesn't, so startup can
        /// offer to open Settings once the add-in is running.
        /// </summary>
        private bool ValidateStartupHealth()
        {
            string vaultPath = _settings.GetFullVaultPath();
            string inboxPath = _settings.GetInboxPath();

            if (!System.IO.Directory.Exists(vaultPath))
            {
                Logger.Instance.Warning($"Startup health: vault path does not exist: {vaultPath}");
                return false;
            }

            if (!System.IO.Directory.Exists(inboxPath))
            {
                Logger.Instance.Info($"Startup health: inbox folder does not exist yet: {inboxPath}. Will be created on first export.");
            }

            Logger.Instance.Info($"Startup health: vault path OK: {vaultPath}");
            return true;
        }

        private void OfferVaultSetup()
        {
            string vaultPath = _settings.GetFullVaultPath();
            DialogResult choice = MessageBox.Show(
                $"SlingMD can't find the folder it saves notes to (\"{vaultPath}\"), so emails can't be saved yet.\n\n"
                + "Open SlingMD settings to choose your Obsidian vault, or any folder if you don't use Obsidian?",
                "SlingMD - Set up your vault",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (choice == DialogResult.Yes)
            {
                ShowSettings();
            }
        }

        private ObsidianSettings LoadSettings(out bool isFirstLaunchAfterInstall)
        {
            ObsidianSettings settings = new ObsidianSettings();
            isFirstLaunchAfterInstall = !settings.HasSavedSettings();
            settings.Load();
            return settings;
        }

        private void ShowFirstRunSupportPrompt()
        {
            SupportService.ShowBuyMeACoffeePrompt();
            _settings.HasShownSupportPrompt = true;

            try
            {
                _settings.Save();
            }
            catch (ArgumentException ex)
            {
                ShowFirstRunStateSaveWarning(ex.Message);
            }
            catch (IOException ex)
            {
                ShowFirstRunStateSaveWarning(ex.Message);
            }
            catch (UnauthorizedAccessException ex)
            {
                ShowFirstRunStateSaveWarning(ex.Message);
            }
        }

        private void ShowFirstRunStateSaveWarning(string errorMessage)
        {
            MessageBox.Show(
                "SlingMD showed the support prompt but could not save the first-run state."
                    + Environment.NewLine
                    + Environment.NewLine
                    + errorMessage,
                "SlingMD",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        public async void ProcessSelection()
        {
            // Sling on several items used to export only the first one, silently.
            if (GetActiveSelectionCount() > 1)
            {
                SlingMultipleEmails();
                return;
            }

            if (!TryBeginOperation())
            {
                return;
            }

            Explorer explorer = null;
            Selection selection = null;
            object selected = null;
            try
            {
                explorer = Application.ActiveExplorer();
                selection = explorer?.Selection;
                if (selection == null || selection.Count == 0)
                {
                    MessageBox.Show("Please select an email, appointment, or contact first.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                selected = selection[1];
                MailItem mail = selected as MailItem;
                AppointmentItem appointment = selected as AppointmentItem;
                ContactItem contact = selected as ContactItem;

                if (mail != null)
                {
                    // When the user has opted into per-sling folder routing, ask where this one email
                    // should land. Mirrors SlingMultipleEmails: never mutate the shared live settings —
                    // route through a cloned settings/processor so background and auto slings keep
                    // writing to the real Inbox.
                    EmailProcessor singleProcessor = _emailProcessor;
                    string routedFolder = null;
                    bool routedFolderIsNew = false;
                    if (_settings.PromptForFolderOnSling)
                    {
                        BatchFolderPickerForm.PickerResult pick;
                        string chosenFolder;
                        bool createdNew;
                        using (BatchFolderPickerForm picker = new BatchFolderPickerForm(1, _settings.GetInboxPath()))
                        {
                            picker.ShowDialog();
                            pick = picker.Result;
                            chosenFolder = picker.SelectedFolderPath;
                            createdNew = picker.CreatedNewFolder;
                        }

                        if (pick == BatchFolderPickerForm.PickerResult.Cancel)
                        {
                            return;
                        }

                        if (pick == BatchFolderPickerForm.PickerResult.UseSubfolder
                            && !string.IsNullOrEmpty(chosenFolder))
                        {
                            string subName = System.IO.Path.GetFileName(
                                chosenFolder.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar));

                            // An empty leaf would make InboxFolder collapse back to the Inbox itself.
                            // Falling through to the default processor says that plainly.
                            if (!string.IsNullOrEmpty(subName))
                            {
                                // NOTE: this clone's InboxFolder holds a two-segment relative path
                                // ("Inbox\Sub"), which ValidateFolderName rejects. It is safe only
                                // because a clone is never Save()d — Save() runs against the live
                                // _settings alone. Do not persist this object.
                                ObsidianSettings singleSettings = _settings.Clone();
                                singleSettings.InboxFolder = System.IO.Path.Combine(_settings.InboxFolder, subName);
                                singleProcessor = new EmailProcessor(singleSettings);
                                routedFolder = chosenFolder;
                                routedFolderIsNew = createdNew;
                            }
                        }
                    }

                    bool slung = await singleProcessor.ProcessEmail(mail);
                    if (!slung)
                    {
                        RemoveFolderIfUnused(routedFolder, routedFolderIsNew);
                    }
                }
                else if (appointment != null)
                {
                    await _appointmentProcessor.ProcessAppointment(appointment, bulkMode: false);
                }
                else if (contact != null)
                {
                    ContactProcessingResult contactResult = _contactProcessor.ProcessContact(contact);
                    if (contactResult == ContactProcessingResult.Success && _settings.LaunchObsidian)
                    {
                        _fileService.LaunchObsidian(_settings.VaultName, _settings.GetContactsPath());
                    }
                }
                else
                {
                    MessageBox.Show("Please select an email, appointment, or contact.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error saving item: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (selected != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(selected);
                if (selection != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                if (explorer != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(explorer);
                EndOperation();
            }
        }

        public async void ProcessCurrentAppointment()
        {
            if (!TryBeginOperation())
            {
                return;
            }

            Inspector inspector = null;
            AppointmentItem appointment = null;
            try
            {
                inspector = Application.ActiveInspector();
                if (inspector == null)
                {
                    MessageBox.Show("No item is currently open.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                appointment = inspector.CurrentItem as AppointmentItem;
                if (appointment == null)
                {
                    MessageBox.Show("The open item is not an appointment.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (!appointment.Saved)
                {
                    DialogResult choice = MessageBox.Show(
                        "This appointment has unsaved changes. Save before exporting to Obsidian?",
                        "Unsaved Changes",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question);

                    if (choice == DialogResult.Cancel)
                    {
                        return;
                    }

                    if (choice == DialogResult.Yes)
                    {
                        appointment.Save();
                    }
                }

                await _appointmentProcessor.ProcessAppointment(appointment, bulkMode: false);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error saving appointment: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (appointment != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(appointment);
                if (inspector != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(inspector);
                EndOperation();
            }
        }

        /// <summary>
        /// Returns how many items are selected in the active explorer, or 0 when the selection can't be read.
        /// </summary>
        private int GetActiveSelectionCount()
        {
            Explorer explorer = null;
            Selection selection = null;
            try
            {
                explorer = Application.ActiveExplorer();
                selection = explorer?.Selection;
                return selection?.Count ?? 0;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                // Some views (e.g. a folder with nothing selectable) throw on Selection.
                return 0;
            }
            finally
            {
                if (selection != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                if (explorer != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(explorer);
            }
        }

        public void ProcessSelectedEmail()
        {
            ProcessSelection();
        }

        /// <summary>
        /// Copies the selected email bodies, optionally with YAML metadata and optionally expanded
        /// to every message Outlook exposes in each selected email's conversation.
        /// </summary>
        public void CopySelectedEmailsToClipboard(bool includeFrontmatter, bool includeThreads)
        {
            if (!TryBeginOperation())
            {
                return;
            }

            Explorer explorer = null;
            Selection selection = null;
            Dictionary<string, ClipboardEmail> snapshots = new Dictionary<string, ClipboardEmail>(StringComparer.OrdinalIgnoreCase);
            try
            {
                explorer = Application.ActiveExplorer();
                selection = explorer?.Selection;
                if (selection == null || selection.Count == 0)
                {
                    MessageBox.Show("Please select one or more emails first.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                for (int index = 1; index <= selection.Count; index++)
                {
                    object selectedItem = null;
                    try
                    {
                        selectedItem = selection[index];
                        MailItem mail = selectedItem as MailItem;
                        if (mail == null)
                        {
                            continue;
                        }

                        if (includeThreads)
                        {
                            AddConversationSnapshots(mail, snapshots);
                        }
                        else
                        {
                            AddClipboardSnapshot(mail, snapshots);
                        }
                    }
                    finally
                    {
                        if (selectedItem != null)
                        {
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(selectedItem);
                        }
                    }
                }

                if (snapshots.Count == 0)
                {
                    MessageBox.Show("None of the selected items are emails.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ClipboardEmailService clipboardService = new ClipboardEmailService();
                string content = clipboardService.Format(snapshots.Values, includeFrontmatter);

                if (clipboardService.ExceedsClipboardWarningSize(content))
                {
                    double sizeInMegabytes = System.Text.Encoding.Unicode.GetByteCount(content) / (1024d * 1024d);
                    DialogResult largeContentChoice = MessageBox.Show(
                        $"The generated text is {sizeInMegabytes:F1} MB. Very large clipboard contents may fail or be truncated by the receiving application."
                            + Environment.NewLine + Environment.NewLine
                            + "Choose Yes to save it as a text file, No to try the clipboard anyway, or Cancel to do nothing.",
                        "Large Clipboard Content",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Warning);

                    if (largeContentChoice == DialogResult.Cancel)
                    {
                        return;
                    }

                    if (largeContentChoice == DialogResult.Yes)
                    {
                        SaveClipboardContentToFile(content, snapshots.Count, includeThreads);
                        return;
                    }
                }

                try
                {
                    Clipboard.SetText(content);
                }
                catch (System.Runtime.InteropServices.ExternalException ex)
                {
                    Logger.Instance.Debug($"Windows rejected clipboard content: {ex.Message}");
                    OfferClipboardFileFallback(content, snapshots.Count, includeThreads);
                    return;
                }
                catch (OutOfMemoryException ex)
                {
                    Logger.Instance.Debug($"Clipboard content exceeded available memory: {ex.Message}");
                    OfferClipboardFileFallback(content, snapshots.Count, includeThreads);
                    return;
                }

                string scope = includeThreads ? "thread email" : "selected email";
                string noun = snapshots.Count == 1 ? scope : scope + "s";
                ToastForm.ShowToast($"SlingMD: copied {snapshots.Count} {noun} to the clipboard.");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error copying email contents: {ex.Message}", "SlingMD Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (selection != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                if (explorer != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(explorer);
                EndOperation();
            }
        }

        private static void OfferClipboardFileFallback(string content, int emailCount, bool includesThreads)
        {
            DialogResult choice = MessageBox.Show(
                "Windows could not place all of the generated email content on the clipboard. Save it as a UTF-8 text file instead?",
                "Clipboard Copy Failed",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (choice == DialogResult.Yes)
            {
                SaveClipboardContentToFile(content, emailCount, includesThreads);
            }
        }

        private static bool SaveClipboardContentToFile(string content, int emailCount, bool includesThreads)
        {
            string scope = includesThreads ? "email-threads" : "emails";
            string defaultFileName = $"SlingMD-{scope}-{DateTime.Now:yyyy-MM-dd-HHmmss}.txt";
            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = $"Save {emailCount} Email{(emailCount == 1 ? string.Empty : "s")}";
                dialog.Filter = "Text files (*.txt)|*.txt|Markdown files (*.md)|*.md|All files (*.*)|*.*";
                dialog.DefaultExt = "txt";
                dialog.AddExtension = true;
                dialog.FileName = defaultFileName;
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                dialog.OverwritePrompt = true;

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return false;
                }

                File.WriteAllText(dialog.FileName, content ?? string.Empty, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                ToastForm.ShowToast($"SlingMD: saved {emailCount} email{(emailCount == 1 ? string.Empty : "s")} to {dialog.FileName}");
                return true;
            }
        }

        private static void AddConversationSnapshots(MailItem selectedMail, Dictionary<string, ClipboardEmail> snapshots)
        {
            Conversation conversation = null;
            SimpleItems roots = null;
            try
            {
                conversation = selectedMail.GetConversation();
                if (conversation == null)
                {
                    AddClipboardSnapshot(selectedMail, snapshots);
                    return;
                }

                roots = conversation.GetRootItems();
                AddConversationItems(roots, conversation, snapshots);
            }
            catch (System.Exception ex)
            {
                // Conversation expansion is unavailable for some stores/items. The selected mail is
                // still useful, so fall back to it rather than failing the entire clipboard action.
                Logger.Instance.Debug($"Could not expand Outlook conversation: {ex.Message}");
                AddClipboardSnapshot(selectedMail, snapshots);
            }
            finally
            {
                if (roots != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(roots);
                if (conversation != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(conversation);
            }
        }

        private static void AddConversationItems(SimpleItems items, Conversation conversation, Dictionary<string, ClipboardEmail> snapshots)
        {
            if (items == null)
            {
                return;
            }

            object item = items.GetFirst();
            while (item != null)
            {
                object current = item;
                try
                {
                    MailItem mail = current as MailItem;
                    if (mail != null)
                    {
                        AddClipboardSnapshot(mail, snapshots);
                    }

                    SimpleItems children = null;
                    try
                    {
                        children = conversation.GetChildren(current);
                        AddConversationItems(children, conversation, snapshots);
                    }
                    finally
                    {
                        if (children != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(children);
                    }

                    item = items.GetNext();
                }
                finally
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(current);
                }
            }
        }

        private static void AddClipboardSnapshot(MailItem mail, Dictionary<string, ClipboardEmail> snapshots)
        {
            ClipboardEmail snapshot = CreateClipboardSnapshot(mail);
            string key = !string.IsNullOrWhiteSpace(snapshot.EntryId)
                ? snapshot.EntryId
                : snapshot.ConversationId + "|" + snapshot.SentOn.Ticks + "|" + snapshot.Subject;
            if (!snapshots.ContainsKey(key))
            {
                snapshots.Add(key, snapshot);
            }
        }

        private static ClipboardEmail CreateClipboardSnapshot(MailItem mail)
        {
            ClipboardEmail snapshot = new ClipboardEmail
            {
                EntryId = mail.EntryID ?? string.Empty,
                ConversationId = mail.ConversationID ?? string.Empty,
                Subject = mail.Subject ?? string.Empty,
                SenderName = mail.SenderName ?? string.Empty,
                SenderEmail = mail.SenderEmailAddress ?? string.Empty,
                SentOn = mail.SentOn,
                ReceivedTime = mail.ReceivedTime,
                Body = mail.Body ?? string.Empty
            };

            Recipients recipients = null;
            try
            {
                recipients = mail.Recipients;
                foreach (Recipient recipient in recipients)
                {
                    try
                    {
                        string display = string.IsNullOrWhiteSpace(recipient.Address)
                            ? recipient.Name
                            : $"{recipient.Name} <{recipient.Address}>";
                        if (recipient.Type == (int)OlMailRecipientType.olCC)
                        {
                            snapshot.Cc.Add(display);
                        }
                        else if (recipient.Type == (int)OlMailRecipientType.olTo)
                        {
                            snapshot.To.Add(display);
                        }
                    }
                    finally
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(recipient);
                    }
                }
            }
            finally
            {
                if (recipients != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(recipients);
            }

            return snapshot;
        }

        /// <summary>
        /// Removes a folder the picker created when the sling it was created for wrote nothing —
        /// a duplicate email, for instance, returns early without producing a note. Without this
        /// the vault accumulates empty folders from slings that visibly did nothing.
        /// Only removes a folder this run created, and only while it is still empty.
        /// </summary>
        private static void RemoveFolderIfUnused(string folderPath, bool createdByPicker)
        {
            if (!createdByPicker || string.IsNullOrEmpty(folderPath))
            {
                return;
            }

            try
            {
                if (Directory.Exists(folderPath)
                    && Directory.GetFileSystemEntries(folderPath).Length == 0)
                {
                    Directory.Delete(folderPath);
                }
            }
            catch (System.Exception ex)
            {
                // Best-effort tidy-up. A stray empty folder is a cosmetic problem; failing the
                // sling over it, or surfacing a dialog, would be worse.
                Logger.Instance.Debug($"Could not remove unused sling folder '{folderPath}': {ex.Message}");
            }
        }

        /// <summary>
        /// Slings every selected email at once. Prompts (via <see cref="BatchFolderPickerForm"/>) for
        /// an Inbox subfolder to route the batch into (or the default Inbox), then processes each email
        /// automatically — no per-email contact/task dialogs and a single Obsidian launch at the end.
        /// </summary>
        public async void SlingMultipleEmails()
        {
            if (!TryBeginOperation())
            {
                return;
            }

            Explorer explorer = null;
            Selection selection = null;
            List<MailItem> mails = new List<MailItem>();
            try
            {
                explorer = Application.ActiveExplorer();
                selection = explorer?.Selection;
                if (selection == null || selection.Count == 0)
                {
                    MessageBox.Show("Please select one or more emails first.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                for (int i = 1; i <= selection.Count; i++)
                {
                    object item = selection[i];
                    MailItem selectedMail = item as MailItem;
                    if (selectedMail != null)
                    {
                        mails.Add(selectedMail);
                    }
                    else if (item != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(item);
                    }
                }

                if (mails.Count == 0)
                {
                    MessageBox.Show("None of the selected items are emails.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string inboxPath = _settings.GetInboxPath();
                BatchFolderPickerForm.PickerResult pick;
                string chosenFolder;
                bool chosenFolderIsNew;
                using (BatchFolderPickerForm picker = new BatchFolderPickerForm(mails.Count, inboxPath))
                {
                    picker.ShowDialog();
                    pick = picker.Result;
                    chosenFolder = picker.SelectedFolderPath;
                    chosenFolderIsNew = picker.CreatedNewFolder;
                }

                if (pick == BatchFolderPickerForm.PickerResult.Cancel)
                {
                    return;
                }

                bool useSubfolder = pick == BatchFolderPickerForm.PickerResult.UseSubfolder
                    && !string.IsNullOrEmpty(chosenFolder);

                // Route the batch WITHOUT mutating the shared live settings: when a subfolder is
                // chosen, process against an isolated cloned settings whose InboxFolder points at the
                // subfolder, via a dedicated processor. This keeps concurrent/background slings (which
                // use the live _settings/_emailProcessor) writing to the real Inbox. bulkMode suppresses
                // per-email progress windows, date prompts, and Obsidian launches.
                string vaultRelativeTarget = _settings.InboxFolder;
                EmailProcessor batchProcessor = _emailProcessor;
                if (useSubfolder)
                {
                    string subName = System.IO.Path.GetFileName(
                        chosenFolder.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar));
                    ObsidianSettings batchSettings = _settings.Clone();
                    batchSettings.InboxFolder = System.IO.Path.Combine(_settings.InboxFolder, subName);
                    vaultRelativeTarget = batchSettings.InboxFolder;
                    batchProcessor = new EmailProcessor(batchSettings);
                }

                BatchSlingReport report = await SlingBatchAsync(batchProcessor, mails, "SlingMD - Sling Multiple");
                int slung = report.Saved;

                if (slung == 0)
                {
                    RemoveFolderIfUnused(useSubfolder ? chosenFolder : null, chosenFolderIsNew);
                }

                // Launch Obsidian once, honoring the user's setting, using a vault-relative path
                // (LaunchObsidian expects a path relative to the vault, not an absolute filesystem path).
                if (_settings.LaunchObsidian && slung > 0)
                {
                    _fileService.LaunchObsidian(_settings.VaultName, vaultRelativeTarget);
                }

                string noun = mails.Count == 1 ? "email" : "emails";
                MessageBox.Show(report.BuildMessage(noun), "SlingMD", MessageBoxButtons.OK,
                    report.Failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error slinging emails: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                foreach (MailItem selectedMail in mails)
                {
                    if (selectedMail != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(selectedMail);
                    }
                }
                if (selection != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                if (explorer != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(explorer);
                EndOperation();
            }
        }

        public void SlingAllContacts()
        {
            int saved = 0;
            int skipped = 0;
            int errors = 0;

            MAPIFolder contactsFolder = null;
            NameSpace session = null;
            try
            {
                try
                {
                    session = Application.Session;
                    contactsFolder = session.GetDefaultFolder(OlDefaultFolders.olFolderContacts);
                    _contactProcessor.ProcessAddressBook(contactsFolder, out saved, out skipped, out errors);
                }
                finally
                {
                    if (session != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(session);
                    }
                    if (contactsFolder != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(contactsFolder);
                    }
                }

                System.Collections.Generic.List<string> bulkErrors = _contactProcessor.GetBulkErrors();
                string summary = string.Format(
                    "Saved {0} contacts.\nSkipped: {1} (already exist or not a contact)\nErrors: {2}",
                    saved, skipped, errors);

                if (bulkErrors.Count > 0)
                {
                    summary += "\n\nError details:\n" + string.Join("\n", bulkErrors);
                }

                MessageBox.Show(
                    summary,
                    "Sling All Contacts",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                if (_settings.LaunchObsidian && saved > 0)
                {
                    _fileService.LaunchObsidian(_settings.VaultName, _settings.GetContactsPath());
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    string.Format("Error exporting contacts: {0}", ex.Message),
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public async void SaveTodaysAppointments()
        {
            if (!TryBeginOperation())
            {
                return;
            }

            int saved = 0;
            int skipped = 0;
            int errors = 0;
            int total = 0;

            NameSpace session = null;
            try
            {
                session = Application.Session;
                Accounts accounts = session.Accounts;
                try
                {
                    foreach (Account account in accounts)
                    {
                        MAPIFolder calendar = null;
                        Items items = null;
                        Items restricted = null;
                        NameSpace accountSession = null;
                        try
                        {
                            accountSession = account.Session;
                            calendar = accountSession.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                            items = calendar.Items;
                            items.IncludeRecurrences = true;
                            items.Sort("[Start]");

                            DateTime today = DateTime.Today;
                            DateTime tomorrow = today.AddDays(1);
                            string filter = string.Format(
                                "[Start] >= '{0}' AND [Start] < '{1}'",
                                today.ToString("g"),
                                tomorrow.ToString("g"));
                            restricted = items.Restrict(filter);

                            foreach (object item in restricted)
                            {
                                AppointmentItem appointment = item as AppointmentItem;
                                if (appointment == null) continue;

                                try
                                {
                                    total++;

                                    AppointmentProcessingResult result =
                                        await _appointmentProcessor.ProcessAppointment(
                                            appointment, bulkMode: true);

                                    switch (result)
                                    {
                                        case AppointmentProcessingResult.Success:
                                            saved++;
                                            break;
                                        case AppointmentProcessingResult.Skipped:
                                            skipped++;
                                            break;
                                        case AppointmentProcessingResult.Error:
                                            errors++;
                                            break;
                                    }
                                }
                                finally
                                {
                                    if (appointment != null)
                                    {
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(appointment);
                                    }
                                }
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            errors++;
                        }
                        finally
                        {
                            if (restricted != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(restricted);
                            if (items != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
                            if (calendar != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(calendar);
                            if (accountSession != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(accountSession);
                            if (account != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(account);
                        }
                    }
                }
                finally
                {
                    if (accounts != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(accounts);
                }

                List<string> bulkErrors = _appointmentProcessor.GetBulkErrors();
                int ambiguousCount = _appointmentProcessor.GetAndClearAmbiguousCount();
                string summary = string.Format(
                    "Saved {0}/{1} appointments.\nSkipped: {2} (duplicates/cancelled)\nErrors: {3}",
                    saved, total, skipped, errors);

                if (ambiguousCount > 0)
                {
                    string logStem = System.IO.Path.GetFileNameWithoutExtension(_settings.BulkAmbiguousMatchLogPath);
                    summary += string.Format("\n> ⚠ {0} ambiguous matches — see [[{1}]]", ambiguousCount, logStem);
                }

                if (bulkErrors.Count > 0)
                {
                    summary += "\n\nError details:\n" + string.Join("\n", bulkErrors);
                }

                MessageBox.Show(
                    summary,
                    "Save Today's Appointments",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                if (_settings.LaunchObsidian && saved > 0)
                {
                    _fileService.LaunchObsidian(_settings.VaultName, _settings.GetAppointmentsPath());
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    string.Format("Error saving today's appointments: {0}", ex.Message),
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (session != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(session);
                EndOperation();
            }
        }

        public async void SaveAppointmentRange(DateTime start, DateTime end)
        {
            if (!TryBeginOperation())
            {
                return;
            }

            int saved = 0;
            int skipped = 0;
            int errors = 0;
            int total = 0;

            NameSpace session = null;
            try
            {
                session = Application.Session;
                Accounts accounts = session.Accounts;
                try
                {
                    foreach (Account account in accounts)
                    {
                        MAPIFolder calendar = null;
                        Items items = null;
                        Items restricted = null;
                        NameSpace accountSession = null;
                        try
                        {
                            accountSession = account.Session;
                            calendar = accountSession.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                            items = calendar.Items;
                            items.IncludeRecurrences = true;
                            items.Sort("[Start]");

                            DateTime rangeEnd = end.AddDays(1);
                            string filter = string.Format(
                                "[Start] >= '{0}' AND [Start] < '{1}'",
                                start.ToString("g"),
                                rangeEnd.ToString("g"));
                            restricted = items.Restrict(filter);

                            foreach (object item in restricted)
                            {
                                AppointmentItem appointment = item as AppointmentItem;
                                if (appointment == null) continue;

                                try
                                {
                                    total++;

                                    AppointmentProcessingResult result =
                                        await _appointmentProcessor.ProcessAppointment(
                                            appointment, bulkMode: true);

                                    switch (result)
                                    {
                                        case AppointmentProcessingResult.Success:
                                            saved++;
                                            break;
                                        case AppointmentProcessingResult.Skipped:
                                            skipped++;
                                            break;
                                        case AppointmentProcessingResult.Error:
                                            errors++;
                                            break;
                                    }
                                }
                                finally
                                {
                                    if (appointment != null)
                                    {
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(appointment);
                                    }
                                }
                            }
                        }
                        catch (System.Runtime.InteropServices.COMException)
                        {
                            errors++;
                        }
                        finally
                        {
                            if (restricted != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(restricted);
                            if (items != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(items);
                            if (calendar != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(calendar);
                            if (accountSession != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(accountSession);
                            if (account != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(account);
                        }
                    }
                }
                finally
                {
                    if (accounts != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(accounts);
                }

                List<string> bulkErrors = _appointmentProcessor.GetBulkErrors();
                int ambiguousCount = _appointmentProcessor.GetAndClearAmbiguousCount();
                string summary = string.Format(
                    "Saved {0}/{1} appointments.\nSkipped: {2} (duplicates/cancelled)\nErrors: {3}\nDate range: {4} to {5}",
                    saved, total, skipped, errors,
                    start.ToString("d"), end.ToString("d"));

                if (ambiguousCount > 0)
                {
                    string logStem = System.IO.Path.GetFileNameWithoutExtension(_settings.BulkAmbiguousMatchLogPath);
                    summary += string.Format("\n> ⚠ {0} ambiguous matches — see [[{1}]]", ambiguousCount, logStem);
                }

                if (bulkErrors.Count > 0)
                {
                    summary += "\n\nError details:\n" + string.Join("\n", bulkErrors);
                }

                MessageBox.Show(
                    summary,
                    "Save Appointment Range",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                if (_settings.LaunchObsidian && saved > 0)
                {
                    _fileService.LaunchObsidian(_settings.VaultName, _settings.GetAppointmentsPath());
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(
                    string.Format("Error saving appointments: {0}", ex.Message),
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (session != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(session);
                EndOperation();
            }
        }

        public async void CompleteThread()
        {
            if (!TryBeginOperation())
            {
                return;
            }

            Explorer explorer = null;
            Selection selection = null;
            MailItem mail = null;
            List<MailItem> missingEmails = null;
            try
            {
                explorer = Application.ActiveExplorer();
                selection = explorer?.Selection;
                if (explorer == null || selection == null || selection.Count == 0)
                {
                    MessageBox.Show("Please select an email first.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                mail = selection[1] as MailItem;
                if (mail == null)
                {
                    MessageBox.Show("Please select an email to complete its thread.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                TemplateService templateService = new TemplateService(_fileService);
                ThreadService threadService = new ThreadService(_fileService, templateService, _settings);
                string conversationId = threadService.GetConversationId(mail);

                ThreadCompletionService completionService = new ThreadCompletionService(_fileService, _settings);
                MAPIFolder inbox = null;
                NameSpace session = null;
                try
                {
                    session = Application.Session;
                    inbox = session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
                    missingEmails = completionService.FindMissingEmails(conversationId, inbox, threadService.GetConversationId);
                }
                finally
                {
                    if (inbox != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(inbox);
                    }
                    if (session != null)
                    {
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(session);
                    }
                }

                if (missingEmails.Count == 0)
                {
                    MessageBox.Show("All emails in this thread have already been slung to Obsidian.", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (ThreadCompletionDialog dialog = new ThreadCompletionDialog(missingEmails))
                {
                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        List<MailItem> selected = dialog.SelectedEmails;
                        // Bulk mode, as in SlingMultipleEmails: no task-date dialog, contact prompt,
                        // countdown or Obsidian launch per email.
                        BatchSlingReport report = await SlingBatchAsync(_emailProcessor, selected, "SlingMD - Complete Thread");

                        if (_settings.LaunchObsidian && report.Saved > 0)
                        {
                            _fileService.LaunchObsidian(_settings.VaultName, GetThreadLaunchTarget(threadService, conversationId));
                        }

                        MessageBox.Show(report.BuildMessage("emails from this thread"), "SlingMD", MessageBoxButtons.OK,
                            report.Failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error completing thread: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // FindMissingEmails returns live MailItem COM objects we own; release each so a
                // multi-email thread doesn't leak one handle per email on every Complete-Thread run.
                if (missingEmails != null)
                {
                    foreach (MailItem missing in missingEmails)
                    {
                        if (missing != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(missing);
                    }
                }
                if (mail != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(mail);
                if (selection != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
                if (explorer != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(explorer);
                EndOperation();
            }
        }

        /// <summary>
        /// Slings <paramref name="mails"/> in bulk mode behind a progress window with a Cancel button,
        /// and returns what happened to each. Cancel takes effect between emails.
        /// </summary>
        private async System.Threading.Tasks.Task<BatchSlingReport> SlingBatchAsync(EmailProcessor processor, IList<MailItem> mails, string title)
        {
            BatchSlingReport report = new BatchSlingReport(mails.Count);
            using (BatchProgressForm progress = new BatchProgressForm(title, mails.Count))
            {
                progress.Show();
                for (int i = 0; i < mails.Count; i++)
                {
                    // Yield to the message loop so the window repaints and Cancel can be clicked;
                    // ProcessEmail itself mostly runs synchronously on this thread.
                    await System.Threading.Tasks.Task.Yield();
                    if (progress.CancelRequested)
                    {
                        report.Cancelled = true;
                        break;
                    }

                    string subject = GetSubjectOrEmpty(mails[i]);
                    progress.ShowItem(i + 1, subject);
                    EmailSlingResult result = await processor.ProcessEmailWithResult(
                        mails[i], contactMode: ContactInteractionMode.Automated, bulkMode: true);
                    report.Add(subject, result);
                }
            }
            return report;
        }

        private static string GetSubjectOrEmpty(MailItem mail)
        {
            try
            {
                return mail.Subject ?? string.Empty;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Returns the vault-relative path Obsidian should open after a thread was completed: the
        /// thread's summary note when the emails are grouped in a thread folder, otherwise the inbox.
        /// </summary>
        private string GetThreadLaunchTarget(ThreadService threadService, string conversationId)
        {
            string inboxPath = _settings.GetInboxPath();
            string threadName = threadService.FindExistingThread(conversationId, inboxPath).earliestEmailThreadName;
            if (string.IsNullOrEmpty(threadName))
            {
                return _settings.InboxFolder;
            }

            return System.IO.Path.Combine(_settings.InboxFolder, threadName, "0-" + threadName + ".md");
        }

        public void ShowSettings()
        {
            try
            {
                using (SettingsForm form = new SettingsForm(_settings))
                {
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        // Signal shutdown to in-flight handlers before recreating
                        _autoSlingService?.Shutdown();
                        _flagMonitorService?.SignalShutdown();
                        _flagMonitorService?.Stop();
                        _folderMonitorService?.SignalShutdown();
                        _folderMonitorService?.StopWatching();

                        // Settings are automatically saved by the form
                        // Recreate processors with new settings
                        _emailProcessor = new EmailProcessor(_settings);
                        _appointmentProcessor = new AppointmentProcessor(_settings);
                        _contactProcessor = new ContactProcessor(_settings);
                        _fileService = new FileService(_settings);
                        _notificationService = new NotificationService(_settings);

                        // Restart folder monitoring with updated watched folders
                        _folderMonitorService = null;
                        if (_settings.WatchedFolders != null && _settings.WatchedFolders.Count > 0)
                        {
                            _folderMonitorService = new FolderMonitorService(_settings, _emailProcessor, _notificationService, Application, _autoSlingProcessedIds);
                            _folderMonitorService.StartWatching(_settings.WatchedFolders);
                        }

                        // Restart auto-sling and flag monitor with updated settings. The shared
                        // processed-id set is reused (not recreated), so an item already slung by a
                        // monitor before the settings change isn't re-slung after the restart.
                        _autoSlingService = new AutoSlingService(_settings, _emailProcessor, _notificationService, _autoSlingProcessedIds);
                        _autoSlingService.Start(Application);

                        _flagMonitorService = null;
                        if (_settings.EnableFlagToSling)
                        {
                            _flagMonitorService = new FlagMonitorService(_settings, _emailProcessor, _notificationService, _autoSlingProcessedIds);
                            _flagMonitorService.Start(Application);
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error showing settings: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
