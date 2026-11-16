namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// Settings for computing task due dates and reminders.
    /// Pure DTO with no Outlook Interop dependencies.
    /// </summary>
    public class TaskDueSettings
    {
        /// <summary>
        /// Gets or sets the default number of days until the task is due.
        /// </summary>
        public int DefaultDueDays { get; set; }

        /// <summary>
        /// Gets or sets whether reminders are relative to due date (true)
        /// or absolute from now (false).
        /// </summary>
        public bool UseRelativeReminder { get; set; }

        /// <summary>
        /// Gets or sets the default number of days for reminder calculation.
        /// In relative mode: days before due date.
        /// In absolute mode: days from now.
        /// </summary>
        public int DefaultReminderDays { get; set; }

        /// <summary>
        /// Gets or sets the hour of day for the reminder (0-23).
        /// </summary>
        public int DefaultReminderHour { get; set; }
    }
}
