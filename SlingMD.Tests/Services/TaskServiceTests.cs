using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;

namespace SlingMD.Tests.Services
{
    public class TaskServiceTests : IDisposable
    {
        private readonly string _testDir;

        public TaskServiceTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "TaskService");
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }

            Directory.CreateDirectory(_testDir);
        }

        [Fact]
        public void GenerateObsidianTask_UsesTagsAndDates_SingleLine()
        {
            ObsidianSettings settings = new ObsidianSettings();
            TaskService service = new TaskService(settings);
            service.InitializeTaskSettings(1, 1, 9, false);

            List<string> tags = new List<string> { "FollowUp", "ActionItem" };
            string result = service.GenerateObsidianTask("TestNote", tags);

            Assert.StartsWith("- [ ] [[TestNote]] #FollowUp #ActionItem", result);
            Assert.Contains("➕", result);
            Assert.Contains("🛫", result);
            Assert.Contains("📅", result);
            Assert.DoesNotContain("\n", result.TrimEnd());
        }

        [Fact]
        public void GenerateObsidianTask_FormatsTagsWithHash()
        {
            ObsidianSettings settings = new ObsidianSettings();
            TaskService service = new TaskService(settings);
            service.InitializeTaskSettings();

            List<string> tags = new List<string> { "foo", "#bar", "baz" };
            string result = service.GenerateObsidianTask("Note", tags);

            Assert.Contains("#foo", result);
            Assert.Contains("#bar", result);
            Assert.Contains("#baz", result);
        }

        [Fact]
        public void GenerateObsidianTask_FallsBackToDefaultTag()
        {
            ObsidianSettings settings = new ObsidianSettings();
            TaskService service = new TaskService(settings);
            service.InitializeTaskSettings();

            string result = service.GenerateObsidianTask("Note", null);
            Assert.Contains("#FollowUp", result);
        }

        [Fact]
        public void GenerateObsidianTask_EmptyTagList_FallsBackToDefault()
        {
            ObsidianSettings settings = new ObsidianSettings();
            TaskService service = new TaskService(settings);
            service.InitializeTaskSettings();

            string result = service.GenerateObsidianTask("Note", new List<string>());
            Assert.Contains("#FollowUp", result);
        }

        [Fact]
        public void GenerateObsidianTask_DisabledTaskCreation_ReturnsEmpty()
        {
            ObsidianSettings settings = new ObsidianSettings();
            TaskService service = new TaskService(settings);
            service.DisableTaskCreation();

            string result = service.GenerateObsidianTask("Note", new List<string> { "foo" });
            Assert.Equal(string.Empty, result);
        }

        [Fact]
        public void GenerateObsidianTask_Dates_AreCorrectlyFormatted()
        {
            ObsidianSettings settings = new ObsidianSettings();
            TaskService service = new TaskService(settings);
            service.InitializeTaskSettings(2, 1, 9, false);

            string result = service.GenerateObsidianTask("Note", new List<string> { "foo" });
            Assert.Matches(@"\d{4}-\d{2}-\d{2}", result);
        }

        [Fact]
        public void InitializeTaskSettings_AfterDisableTaskCreation_ReEnablesTaskGeneration()
        {
            ObsidianSettings settings = new ObsidianSettings();
            TaskService service = new TaskService(settings);

            service.DisableTaskCreation();
            Assert.False(service.ShouldCreateTasks);

            service.InitializeTaskSettings();
            Assert.True(service.ShouldCreateTasks);
        }

        [Fact]
        public void ShouldCreateTasks_AfterCancelThenInitialize_ReturnsTrue()
        {
            ObsidianSettings settings = new ObsidianSettings();
            TaskService service = new TaskService(settings);

            // Simulate a canceled task dialog (which calls DisableTaskCreation)
            service.DisableTaskCreation();

            // Simulate the next export attempt calling InitializeTaskSettings
            service.InitializeTaskSettings(2, 1, 9, false);

            Assert.True(service.ShouldCreateTasks);
            // Verify tasks can actually be generated after re-enable
            string result = service.GenerateObsidianTask("TestNote", new List<string> { "FollowUp" });
            Assert.NotEqual(string.Empty, result);
        }

        [Fact]
        public void GenerateObsidianTask_UsesCustomTemplateWhenPresent()
        {
            string templatesPath = Path.Combine(_testDir, "Templates");
            Directory.CreateDirectory(templatesPath);
            File.WriteAllText(Path.Combine(templatesPath, "TaskTemplate.md"), "TODO {{noteName}} {{tags}} {{dueDate}}");

            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault",
                TemplatesFolder = templatesPath
            };

            TaskService service = new TaskService(settings);
            service.InitializeTaskSettings(1, 0, 9, false);

            string result = service.GenerateObsidianTask("FollowUpNote", new List<string> { "custom" });

            Assert.StartsWith("TODO FollowUpNote #custom", result);
        }

        public void Dispose()
        {
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }
        }
    }
}
