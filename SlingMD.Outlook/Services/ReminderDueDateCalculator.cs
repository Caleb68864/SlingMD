using System;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Calculates task due dates and reminder dates based on settings.
    /// Pure helper with no Outlook Interop dependencies.
    /// </summary>
    public class ReminderDueDateCalculator
    {
        /// <summary>
        /// Calculates the due date, reminder date, and warning flags based on the provided settings.
        /// </summary>
        /// <param name="now">The current date/time.</param>
        /// <param name="settings">The task due settings containing days and mode configuration.</param>
        /// <returns>A TaskDueDates object with computed dates and warning flags.</returns>
        public TaskDueDates Calculate(DateTime now, TaskDueSettings settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            DateTime dueDate = now.Date.AddDays(settings.DefaultDueDays);

            DateTime reminderDate;
            if (settings.UseRelativeReminder)
            {
                // Relative mode: reminder is X days before due date
                reminderDate = dueDate.AddDays(-settings.DefaultReminderDays);
            }
            else
            {
                // Absolute mode: reminder is X days from now
                reminderDate = now.Date.AddDays(settings.DefaultReminderDays);
            }

            DateTime reminderDateTime = reminderDate.AddHours(settings.DefaultReminderHour);
            bool pastReminderWarning = false;

            // Check if reminder is in the past and apply the same logic as TaskService
            if (reminderDateTime < now)
            {
                pastReminderWarning = true;

                // Match pre-change TaskService behavior (lines 122-132):
                // If reminder lands today but hour has passed: use now + 1 hour
                // Otherwise: use tomorrow at the configured reminder hour
                if (reminderDateTime.Date == now.Date)
                {
                    reminderDateTime = now.AddHours(1);
                }
                else
                {
                    reminderDateTime = now.Date.AddDays(1).AddHours(settings.DefaultReminderHour);
                }
            }

            return new TaskDueDates
            {
                DueDate = dueDate,
                ReminderDate = reminderDate,
                ReminderDateTime = reminderDateTime,
                PastReminderWarning = pastReminderWarning
            };
        }
    }
}
