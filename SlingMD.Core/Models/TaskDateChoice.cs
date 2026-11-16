namespace SlingMD.Outlook.Models
{
    /// <summary>
    /// What the user picked in the task-dates prompt (<c>TaskOptionsForm</c> in the add-in), as plain
    /// values. Returned by <see cref="SlingMD.Outlook.Services.IEmailPrompts.PromptTaskOptions"/>.
    /// </summary>
    public sealed class TaskDateChoice
    {
        public TaskDateChoice(int dueDays, int reminderDays, int reminderHour, bool useRelativeReminder)
        {
            DueDays = dueDays;
            ReminderDays = reminderDays;
            ReminderHour = reminderHour;
            UseRelativeReminder = useRelativeReminder;
        }

        public int DueDays { get; }

        public int ReminderDays { get; }

        public int ReminderHour { get; }

        public bool UseRelativeReminder { get; }
    }
}
