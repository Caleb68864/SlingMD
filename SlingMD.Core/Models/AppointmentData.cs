using System;
using System.Collections.Generic;

namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// Outlook's recurrence state of an appointment (<c>OlRecurrenceState</c>) without the interop
    /// type, so the engine can read it.
    /// </summary>
    public enum AppointmentRecurrence
    {
        NotRecurring,
        Master,
        Occurrence,
        Exception
    }

    /// <summary>
    /// Everything the appointment note engine reads from an Outlook <c>AppointmentItem</c>, copied
    /// out as plain values so <see cref="SlingMD.Outlook.Services.AppointmentNoteEngine"/> runs
    /// without COM. The add-in fills it with <c>OutlookAppointmentReader</c>; tests build it directly.
    /// Strings are never null (the reader turns unreadable or null values into empty strings, as the
    /// processor always did).
    /// </summary>
    public class AppointmentData
    {
        public string Subject { get; set; } = string.Empty;
        public string Body { get; set; } = string.Empty;
        public string Location { get; set; } = string.Empty;
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public string OrganizerName { get; set; } = string.Empty;

        /// <summary>The organizer's <c>AddressEntry.Address</c> (may be an Exchange DN, not SMTP).</summary>
        public string OrganizerEmail { get; set; } = string.Empty;

        public string GlobalAppointmentId { get; set; } = string.Empty;
        public AppointmentRecurrence Recurrence { get; set; } = AppointmentRecurrence.NotRecurring;

        /// <summary>Calendar folder name qualified by its store, or empty when Outlook does not expose it.</summary>
        public string CalendarName { get; set; } = string.Empty;

        public List<string> RequiredAttendees { get; set; } = new List<string>();
        public List<string> OptionalAttendees { get; set; } = new List<string>();

        /// <summary>Rooms and equipment (resource recipients).</summary>
        public List<string> ResourceAttendees { get; set; } = new List<string>();
    }
}
