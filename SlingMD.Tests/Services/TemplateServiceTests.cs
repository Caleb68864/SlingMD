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

        /// <summary>
        /// Verifies that string scalar values containing double quotes, backslashes, and embedded
        /// newlines are escaped so that the resulting YAML frontmatter remains parseable.
        /// </summary>
        [Fact]
        public void BuildFrontMatter_EscapesQuotesBackslashesAndNewlines()
        {
            Dictionary<string, object> metadata = new Dictionary<string, object>
            {
                { "title", "He said \"hello\"" },
                { "path", @"C:\Users\Test\note.md" },
                { "body", "Line one\nLine two" }
            };

            string frontmatter = _templateService.BuildFrontMatter(metadata);

            // Double quotes inside a YAML double-quoted scalar must be escaped as \"
            Assert.Contains("title: \"He said \\\"hello\\\"\"", frontmatter);
            // Backslashes must be escaped as \\
            Assert.Contains("path: \"C:\\\\Users\\\\Test\\\\note.md\"", frontmatter);
            // Newlines must be escaped as \n
            Assert.Contains("body: \"Line one\\nLine two\"", frontmatter);
        }

        /// <summary>
        /// Verifies that list entries containing special characters are also escaped, and that
        /// the list shape (YAML block sequence) is preserved.
        /// </summary>
        [Fact]
        public void BuildFrontMatter_EscapesListValuesWithoutChangingListShape()
        {
            Dictionary<string, object> metadata = new Dictionary<string, object>
            {
                { "to", new System.Collections.Generic.List<string> { "Normal Person", "Has \"Quotes\"", @"Back\Slash" } }
            };

            string frontmatter = _templateService.BuildFrontMatter(metadata);

            // List block sequence prefix must still be present
            Assert.Contains("  - \"Normal Person\"", frontmatter);
            Assert.Contains("  - \"Has \\\"Quotes\\\"\"", frontmatter);
            Assert.Contains("  - \"Back\\\\Slash\"", frontmatter);
        }

        /// <summary>
        /// Unit-tests the escaping helper directly to cover edge cases.
        /// </summary>
        [Fact]
        public void EscapeYamlDoubleQuotedScalar_HandlesAllSpecialCharacters()
        {
            Assert.Equal("\\\\", TemplateService.EscapeYamlDoubleQuotedScalar("\\"));
            Assert.Equal("\\\"", TemplateService.EscapeYamlDoubleQuotedScalar("\""));
            Assert.Equal("\\n", TemplateService.EscapeYamlDoubleQuotedScalar("\n"));
            Assert.Equal("\\n", TemplateService.EscapeYamlDoubleQuotedScalar("\r\n"));
            Assert.Equal("plain", TemplateService.EscapeYamlDoubleQuotedScalar("plain"));
            Assert.Equal(string.Empty, TemplateService.EscapeYamlDoubleQuotedScalar(null));
        }

        [Fact]
        public void RenderContactContent_WithAllFields_ContainsContactDetails()
        {
            ContactTemplateContext context = new ContactTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", "Jane Doe" }
                },
                ContactName = "Jane Doe",
                ContactShortName = "Jane",
                Created = "2026-03-13",
                FileName = "Jane Doe.md",
                FileNameWithoutExtension = "Jane Doe",
                Phone = "555-1234",
                Email = "jane@example.com",
                Company = "Acme Corp",
                JobTitle = "Engineer",
                Address = "123 Main St",
                Birthday = "1990-01-01",
                Notes = "Some notes here",
                IncludeDetails = true
            };

            string content = _templateService.RenderContactContent(context);

            Assert.Contains("## Contact Details", content);
            Assert.Contains("555-1234", content);
            Assert.Contains("jane@example.com", content);
            Assert.Contains("Acme Corp", content);
            Assert.Contains("Engineer", content);
            Assert.Contains("123 Main St", content);
            Assert.Contains("1990-01-01", content);
        }

        [Fact]
        public void RenderContactContent_WithIncludeDetailsFalse_OmitsContactDetails()
        {
            // IncludeDetails branching only applies when no ContactTemplate file is found.
            // Point at a name that doesn't exist anywhere on the candidate path so the
            // built-in default (driven by IncludeDetails) is used.
            _settings.ContactTemplateFile = "NonExistentContactTemplate-" + System.Guid.NewGuid().ToString("N") + ".md";

            ContactTemplateContext context = new ContactTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", "Jane Doe" }
                },
                ContactName = "Jane Doe",
                ContactShortName = "Jane",
                Created = "2026-03-13",
                FileName = "Jane Doe.md",
                FileNameWithoutExtension = "Jane Doe",
                Phone = "555-1234",
                Email = "jane@example.com",
                Company = "Acme Corp",
                JobTitle = "Engineer",
                Address = "123 Main St",
                Birthday = "1990-01-01",
                Notes = "Some notes here",
                IncludeDetails = false
            };

            string content = _templateService.RenderContactContent(context);

            Assert.DoesNotContain("## Contact Details", content);
        }

        [Fact]
        public void RenderContactContent_WithEmptyFields_DoesNotContainNull()
        {
            ContactTemplateContext context = new ContactTemplateContext
            {
                Metadata = new Dictionary<string, object>
                {
                    { "title", "Jane Doe" }
                },
                ContactName = "Jane Doe",
                ContactShortName = "Jane",
                Created = "2026-03-13",
                FileName = "Jane Doe.md",
                FileNameWithoutExtension = "Jane Doe",
                Phone = string.Empty,
                Email = string.Empty,
                Company = string.Empty,
                JobTitle = string.Empty,
                Address = string.Empty,
                Birthday = string.Empty,
                Notes = string.Empty,
                IncludeDetails = true
            };

            string content = _templateService.RenderContactContent(context);

            Assert.DoesNotContain("null", content, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("## Contact Details", content);
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
