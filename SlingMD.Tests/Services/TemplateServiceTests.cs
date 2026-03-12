using System;
using System.Collections.Generic;
using System.IO;
using Xunit;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;

namespace SlingMD.Tests.Services
{
    public class TemplateServiceTests : IDisposable
    {
        private readonly string _testDir;
        private readonly string _templatesDir;
        private readonly ObsidianSettings _settings;
        private readonly FileService _fileService;
        private readonly TemplateService _templateService;

        public TemplateServiceTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "TemplateService");
            _templatesDir = Path.Combine(_testDir, "Vault", "Templates");

            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, true);
            }

            Directory.CreateDirectory(_templatesDir);

            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault",
                TemplatesFolder = "Templates"
            };
            _fileService = new FileService(_settings);
            _templateService = new TemplateService(_fileService);
        }

        [Fact]
        public void LoadTemplate_PrefersConfiguredVaultTemplateFolder()
        {
            string templatePath = Path.Combine(_templatesDir, "EmailTemplate.md");
            File.WriteAllText(templatePath, "custom email template");

            string template = _templateService.LoadTemplate("EmailTemplate.md");

            Assert.Equal("custom email template", template);
        }

        [Fact]
        public void RenderFilename_SanitizesInvalidCharacters()
        {
            Dictionary<string, string> replacements = new Dictionary<string, string>
            {
                { "Subject", "Re: Status / Update" },
                { "Sender", "Jane Doe" }
            };

            string fileName = _templateService.RenderFilename("{Subject}-{Sender}", replacements, "fallback");

            Assert.DoesNotContain(":", fileName);
            Assert.DoesNotContain("/", fileName);
            Assert.Contains("Jane Doe", fileName);
        }

        [Fact]
        public void RenderEmailContent_UsesDefaultTemplateWhenCustomTemplateMissing()
        {
            EmailTemplateContext context = new EmailTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", "Test Email" },
                    { "from", "[[Tester]]" }
                },
                NoteTitle = "Test Email",
                Subject = "Test Email",
                SenderName = "Tester",
                SenderShortName = "Tester",
                Body = "Hello body",
                TaskBlock = "- [ ] [[TestEmail]] #FollowUp\n\n",
                FileName = "TestEmail.md",
                FileNameWithoutExtension = "TestEmail",
                ThreadId = "thread-1"
            };

            string content = _templateService.RenderEmailContent(context);

            Assert.Contains("---", content);
            Assert.Contains("Hello body", content);
            Assert.Contains("[[TestEmail]]", content);
        }

        [Fact]
        public void RenderTaskLine_UsesConfiguredTemplateFile()
        {
            File.WriteAllText(Path.Combine(_templatesDir, "TaskTemplate.md"), "TASK {{noteName}} {{createdDate}} {{dueDate}}");

            TaskTemplateContext context = new TaskTemplateContext
            {
                NoteName = "ProjectNote",
                NoteLink = "[[ProjectNote]]",
                CreatedDate = "2026-03-12",
                DueDate = "2026-03-13",
                ReminderDate = "2026-03-12",
                Tags = "#followup"
            };

            string content = _templateService.RenderTaskLine(context);

            Assert.Equal("TASK ProjectNote 2026-03-12 2026-03-13", content);
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
