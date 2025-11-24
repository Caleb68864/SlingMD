using System;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;
using System.Collections.Generic;
using System.Linq;

namespace SlingMD.Outlook.Services
{
    public class TaskService
    {
        private readonly ObsidianSettings _settings;
        private int? _taskDueDays;
        private int? _taskReminderDays;
        private int? _taskReminderHour;
        private bool _useRelativeReminder;
        private bool _createTasks = true;

        public TaskService(ObsidianSettings settings)
        {
            _settings = settings;
        }

        public void InitializeTaskSettings(int? dueDays = null, int? reminderDays = null, int? reminderHour = null, bool? useRelativeReminder = null)
        {
            _taskDueDays = dueDays ?? _settings.DefaultDueDays;
            _taskReminderDays = reminderDays ?? _settings.DefaultReminderDays;
            _taskReminderHour = reminderHour ?? _settings.DefaultReminderHour;
            _useRelativeReminder = useRelativeReminder ?? _settings.UseRelativeReminder;
        }

        public bool ShouldCreateTasks => _createTasks;

        public void DisableTaskCreation()
        {
            _createTasks = false;
        }

        /// <summary>
        /// Generates a single-line Obsidian task with tags and dates.
        /// </summary>
        /// <param name="fileName">The note file name (without extension).</param>
        /// <param name="taskTags">A list of tags to include in the task line (e.g., ["FollowUp", "ActionItem"]).</param>
        /// <returns>The full Obsidian task line, including tags and dates, on a single line.</returns>
        public string GenerateObsidianTask(string fileName, List<string> taskTags = null)
        {
            if (!_createTasks) return string.Empty;

            // Defensive: Ensure task settings are initialized (use defaults if not)
            int dueDays = _taskDueDays ?? _settings.DefaultDueDays;
            int reminderDays = _taskReminderDays ?? _settings.DefaultReminderDays;

            string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
            string dueDate = DateTime.Now.Date.AddDays(dueDays).ToString("yyyy-MM-dd");

            // Calculate reminder date based on setting
            DateTime reminderDateTime;
            if (_useRelativeReminder)
            {
                // Relative: Calculate from due date
                reminderDateTime = DateTime.Now.Date.AddDays(dueDays - reminderDays);
            }
            else
            {
                // Absolute: Calculate from today
                reminderDateTime = DateTime.Now.Date.AddDays(reminderDays);
            }
            string reminderDate = reminderDateTime.ToString("yyyy-MM-dd");

            // Format tags as #tag
            string tagsPart = (taskTags != null && taskTags.Count > 0)
                ? string.Join(" ", taskTags.Select(t => t.StartsWith("#") ? t : "#" + t))
                : "#FollowUp";

            // All on one line
            return $"- [ ] [[{fileName}]] {tagsPart} ➕ {currentDate} 🛫 {reminderDate} 📅 {dueDate}";
        }

        public Task CreateOutlookTask(MailItem mail)
        {
            if (!_createTasks) return Task.CompletedTask;

            Microsoft.Office.Interop.Outlook.Application outlookApp = null;
            Microsoft.Office.Interop.Outlook.TaskItem task = null;

            try
            {
                // Defensive: Ensure task settings are initialized (use defaults if not)
                int dueDays = _taskDueDays ?? _settings.DefaultDueDays;
                int reminderDays = _taskReminderDays ?? _settings.DefaultReminderDays;
                int reminderHour = _taskReminderHour ?? _settings.DefaultReminderHour;

                outlookApp = mail.Application;
                task = (Microsoft.Office.Interop.Outlook.TaskItem)outlookApp.CreateItem(OlItemType.olTaskItem);
                task.Subject = $"Follow up: {mail.Subject ?? "No Subject"}";
                task.Body = $"Follow up on email from {mail.SenderName ?? "Unknown Sender"}\n\nOriginal email:\n{mail.Body ?? string.Empty}";

                // Set due date based on settings
                var dueDate = DateTime.Now.Date.AddDays(dueDays);
                task.DueDate = dueDate;
                task.ReminderSet = true;

                // Calculate reminder time based on setting
                DateTime reminderDate;
                if (_useRelativeReminder)
                {
                    // Relative: Calculate from due date
                    reminderDate = dueDate.AddDays(-reminderDays);
                }
                else
                {
                    // Absolute: Calculate from today
                    reminderDate = DateTime.Now.Date.AddDays(reminderDays);
                }
                var reminderTime = reminderDate.AddHours(reminderHour);

                // If reminder would be in the past, set it to the next possible time
                if (reminderTime < DateTime.Now)
                {
                    if (reminderTime.Date == DateTime.Now.Date)
                    {
                        // If it's today but earlier hour, set to next hour
                        reminderTime = DateTime.Now.AddHours(1);
                    }
                    else
                    {
                        // If it's a past day, set to tomorrow at the specified hour
                        reminderTime = DateTime.Now.Date.AddDays(1).AddHours(reminderHour);
                    }
                }

                task.ReminderTime = reminderTime;
                task.Save();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Failed to create Outlook task: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                // Release COM objects to prevent memory leaks
                if (task != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(task);
                }
                if (outlookApp != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(outlookApp);
                }
            }

            return Task.CompletedTask;
        }
    }
} 