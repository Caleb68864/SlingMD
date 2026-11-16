using System;

namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// Computed due dates and reminder information for a task.
    /// Pure DTO with no Outlook Interop dependencies.
    /// </summary>
    public class TaskDueDates
    {
        /// <summary>
        /// Gets or sets the computed due date.
        /// </summary>
        public DateTime DueDate { get; set; }

        /// <summary>
        /// Gets or sets the reminder date (date portion only).
        /// </summary>
        public DateTime ReminderDate { get; set; }

        /// <summary>
        /// Gets or sets the reminder date and time.
        /// </summary>
        public DateTime ReminderDateTime { get; set; }

        /// <summary>
        /// Gets or sets whether the computed reminder time is in the past.
        /// </summary>
        public bool PastReminderWarning { get; set; }
    }
}
