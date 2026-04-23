using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Infrastructure;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class TaskService
    {
        private readonly ObsidianSettings _settings;
        private readonly TemplateService _templateService;
        private readonly IClock _clock;
        private readonly ReminderDueDateCalculator _calculator;
        private int? _taskDueDays;
        private int? _taskReminderDays;
        private int? _taskReminderHour;
        private bool _useRelativeReminder;
        private bool _createTasks = true;

        public TaskService(ObsidianSettings settings, TemplateService templateService = null, IClock clock = null)
        {
            _settings = settings;
            _templateService = templateService ?? new TemplateService(new FileService(settings));
            _clock = clock ?? new SystemClock();
            _calculator = new ReminderDueDateCalculator();
        }

        private TaskDueDates ComputeDueDates()
        {
            int dueDays = _taskDueDays ?? _settings.DefaultDueDays;
            int reminderDays = _taskReminderDays ?? _settings.DefaultReminderDays;
            int reminderHour = _taskReminderHour ?? _settings.DefaultReminderHour;
            return _calculator.Calculate(_clock.Now, new TaskDueSettings
            {
                DefaultDueDays = dueDays,
                UseRelativeReminder = _useRelativeReminder,
                DefaultReminderDays = reminderDays,
                DefaultReminderHour = reminderHour
            });
        }

        public void InitializeTaskSettings(int? dueDays = null, int? reminderDays = null, int? reminderHour = null, bool? useRelativeReminder = null)
        {
            _taskDueDays = dueDays ?? _settings.DefaultDueDays;
            _taskReminderDays = reminderDays ?? _settings.DefaultReminderDays;
            _taskReminderHour = reminderHour ?? _settings.DefaultReminderHour;
            _useRelativeReminder = useRelativeReminder ?? _settings.UseRelativeReminder;
            _createTasks = true;
        }

        public bool ShouldCreateTasks => _createTasks;

        public void DisableTaskCreation()
        {
            _createTasks = false;
        }

        /// <summary>
        /// Generates a single-line Obsidian task with tags and dates.
        /// </summary>
        public string GenerateObsidianTask(string fileName, List<string> taskTags = null)
        {
            if (!_createTasks)
            {
                return string.Empty;
            }

            TaskDueDates dates = ComputeDueDates();
            string currentDate = _clock.Now.ToString("yyyy-MM-dd");
            string dueDate = dates.DueDate.ToString("yyyy-MM-dd");
            string reminderDate = dates.ReminderDate.ToString("yyyy-MM-dd");
            List<string> effectiveTags = GetEffectiveTaskTags(taskTags);
            string tagsPart = string.Join(" ", effectiveTags.Select(tag => tag.StartsWith("#") ? tag : "#" + tag));

            TaskTemplateContext context = new TaskTemplateContext
            {
                NoteLink = string.IsNullOrWhiteSpace(fileName) ? string.Empty : $"[[{fileName}]]",
                NoteName = fileName ?? string.Empty,
                Tags = tagsPart,
                CreatedDate = currentDate,
                ReminderDate = reminderDate,
                DueDate = dueDate
            };

            return _templateService.RenderTaskLine(context);
        }

        public Task CreateOutlookTask(MailItem mail)
        {
            if (!_createTasks)
            {
                return Task.CompletedTask;
            }

            Microsoft.Office.Interop.Outlook.Application outlookApp = null;
            Microsoft.Office.Interop.Outlook.TaskItem task = null;

            try
            {
                TaskDueDates dates = ComputeDueDates();

                outlookApp = mail.Application;
                task = (Microsoft.Office.Interop.Outlook.TaskItem)outlookApp.CreateItem(OlItemType.olTaskItem);
                task.Subject = $"Follow up: {mail.Subject ?? "No Subject"}";
                task.Body = $"Follow up on email from {mail.SenderName ?? "Unknown Sender"}\n\nOriginal email:\n{mail.Body ?? string.Empty}";

                task.DueDate = dates.DueDate;
                task.ReminderSet = true;
                task.ReminderTime = dates.ReminderDateTime;
                task.Save();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Failed to create Outlook task: {ex.Message}", "SlingMD", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
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

        private List<string> GetEffectiveTaskTags(List<string> taskTags)
        {
            if (taskTags != null && taskTags.Count > 0)
            {
                return taskTags;
            }

            if (_settings.DefaultTaskTags != null && _settings.DefaultTaskTags.Count > 0)
            {
                return _settings.DefaultTaskTags;
            }

            return new List<string> { "FollowUp" };
        }
    }
}
