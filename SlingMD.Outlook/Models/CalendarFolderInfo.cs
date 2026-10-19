using System;
using Newtonsoft.Json;

namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// Identifies one Outlook calendar folder for bulk appointment export. Persisted in
    /// <see cref="ObsidianSettings.AppointmentExportCalendars"/> so the last choice made in the
    /// Save Date Range dialog is reused by Save Today. Pure DTO — no Outlook Interop dependencies.
    /// </summary>
    public class CalendarFolderInfo
    {
        /// <summary>Gets or sets the folder EntryID (the stable key used to reopen the folder).</summary>
        public string EntryId { get; set; } = string.Empty;

        /// <summary>Gets or sets the StoreID of the store that holds the folder.</summary>
        public string StoreId { get; set; } = string.Empty;

        /// <summary>Gets or sets the label shown to the user, e.g. "Calendar (jane@firm.com)".</summary>
        public string DisplayName { get; set; } = string.Empty;

        /// <summary>
        /// True when this is the default calendar of a configured Outlook account. Those calendars
        /// are pre-selected when no choice has been saved. Not persisted: recomputed on every lookup.
        /// </summary>
        [JsonIgnore]
        public bool IsAccountDefault { get; set; }

        /// <summary>
        /// True when <paramref name="other"/> refers to the same Outlook folder. EntryIDs are hex
        /// strings whose case is not guaranteed to be stable, so the comparison ignores case.
        /// </summary>
        public bool IsSameFolder(CalendarFolderInfo other)
        {
            return other != null
                && !string.IsNullOrEmpty(EntryId)
                && string.Equals(EntryId, other.EntryId, StringComparison.OrdinalIgnoreCase);
        }

        public override string ToString()
        {
            return string.IsNullOrEmpty(DisplayName) ? EntryId : DisplayName;
        }
    }
}
