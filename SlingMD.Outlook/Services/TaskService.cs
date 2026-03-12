using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services
{
    public class TaskService
    {
        private readonly ObsidianSettings _settings;
        private readonly TemplateService _templateService;
        private int? _taskDueDays;
        private int? _taskReminderDays;
        private int? _taskReminderHour;
        private bool _useRelativeReminder;
        private bool _createTasks = true;

        public TaskService(ObsidianSettings settings, TemplateService templateService = null)
        {
            _settings = settings;
            _templateService = templateService ?? new TemplateService(new FileService(settings));
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

            int dueDays = _taskDueDays ?? _settings.DefaultDueDays;
            int reminderDays = _taskReminderDays ?? _settings.DefaultReminderDays;

            string currentDate = DateTime.Now.ToString("yyyy-MM-dd");
            string dueDate = DateTime.Now.Date.AddDays(dueDays).ToString("yyyy-MM-dd");

            DateTime reminderDateTime;
            if (_useRelativeReminder)
            {
                reminderDateTime = DateTime.Now.Date.AddDays(dueDays - reminderDays);
            }
            else
            {
                reminderDateTime = DateTime.Now.Date.AddDays(reminderDays);
            }

            string reminderDate = reminderDateTime.ToString("yyyy-MM-dd");
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
                int dueDays = _taskDueDays ?? _settings.DefaultDueDays;
                int reminderDays = _taskReminderDays ?? _settings.DefaultReminderDays;
                int reminderHour = _taskReminderHour ?? _settings.DefaultReminderHour;

                outlookApp = mail.Application;
                task = (Microsoft.Office.Interop.Outlook.TaskItem)outlookApp.CreateItem(OlItemType.olTaskItem);
                task.Subject = $"Follow up: {mail.Subject ?? "No Subject"}";
                task.Body = $"Follow up on email from {mail.SenderName ?? "Unknown Sender"}\n\nOriginal email:\n{mail.Body ?? string.Empty}";

                DateTime dueDate = DateTime.Now.Date.AddDays(dueDays);
                task.DueDate = dueDate;
                task.ReminderSet = true;

                DateTime reminderDate;
                if (_useRelativeReminder)
                {
                    reminderDate = dueDate.AddDays(-reminderDays);
                }
                else
                {
                    reminderDate = DateTime.Now.Date.AddDays(reminderDays);
                }

                DateTime reminderTime = reminderDate.AddHours(reminderHour);
                if (reminderTime < DateTime.Now)
                {
                    if (reminderTime.Date == DateTime.Now.Date)
                    {
                        reminderTime = DateTime.Now.AddHours(1);
                    }
                    else
                    {
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
