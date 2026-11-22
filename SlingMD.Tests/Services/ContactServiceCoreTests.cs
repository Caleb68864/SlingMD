using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using SlingMD.Outlook.Services.Formatting;
using SlingMD.Tests.Infrastructure;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// <see cref="ContactService"/> in SlingMD.Core: the ambiguous-contact prompt behind
    /// <see cref="IContactMatchPrompt"/> (the add-in's <c>DialogContactMatchPrompt</c> shows the
    /// form), and the contact-note context built from values read from an Outlook contact.
    /// Linked into SlingMD.Core.Tests.
    /// </summary>
    public class ContactServiceCoreTests : IDisposable
    {
        private readonly string _testDir;
        private readonly ObsidianSettings _settings;

        public ContactServiceCoreTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ContactCore", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(_testDir, "Vault", "Contacts"));
            _settings = new ObsidianSettings
            {
                VaultBasePath = _testDir,
                VaultName = "Vault",
                ContactsFolder = "Contacts",
                EnableContactSaving = true,
                EnableContactFuzzyMatching = true,
                SearchEntireVaultForContacts = false,
                BulkAmbiguousMatchLogPath = "Logs/ambiguous.md"
            };
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(_testDir, true);
            }
            catch (IOException)
            {
            }
        }

        private sealed class FakePrompt : IContactMatchPrompt
        {
            private readonly Func<bool> _answer;
            public int Calls;
            public List<string> Candidates = new List<string>();
            public ContactInteractionMode Mode;
            public bool IncludeDetails;

            public FakePrompt(Func<bool> answer)
            {
                _answer = answer;
            }

            public bool ChooseExisting(IReadOnlyList<ContactIndexEntry> candidates, ContactInteractionMode mode, bool includeDetails)
            {
                Calls++;
                Candidates = candidates.Select(c => Path.GetFileNameWithoutExtension(c.FilePath)).OrderBy(n => n).ToList();
                Mode = mode;
                IncludeDetails = includeDetails;
                return _answer();
            }
        }

        private ContactService NewService(IContactMatchPrompt prompt)
        {
            FileService files = new FileService(_settings);
            return new ContactService(files, new TemplateService(files), new FakeClock(new DateTime(2026, 9, 23, 9, 0, 0)), prompt);
        }

        private void WriteAmbiguousPair()
        {
            string contacts = Path.Combine(_testDir, "Vault", "Contacts");
            File.WriteAllText(Path.Combine(contacts, "Bob Smith Jr.md"), "---\ntitle: Bob Smith Jr.\n---\n");
            File.WriteAllText(Path.Combine(contacts, "Bob Smith Sr.md"), "---\ntitle: Bob Smith Sr.\n---\n");
        }

        private string LogPath => Path.Combine(_testDir, "Vault", "Logs", "ambiguous.md");

        [Fact]
        public void Interactive_Ambiguous_UserPicksExisting_Handled()
        {
            WriteAmbiguousPair();
            _settings.ContactNoteIncludeDetails = true;
            FakePrompt prompt = new FakePrompt(() => true);
            ContactService service = NewService(prompt);

            bool handled = service.TryHandleFuzzyMatch("Bob James Smith", null, ContactInteractionMode.Interactive, "Subject");

            Assert.True(handled);
            Assert.Equal(1, prompt.Calls);
            Assert.Equal(new List<string> { "Bob Smith Jr", "Bob Smith Sr" }, prompt.Candidates);
            Assert.Equal(ContactInteractionMode.Interactive, prompt.Mode);
            Assert.True(prompt.IncludeDetails);
            Assert.Equal(0, service.GetAndClearAmbiguousCount());
            Assert.False(File.Exists(LogPath));
        }

        [Fact]
        public void Interactive_Ambiguous_UserCreatesNew_NotHandled()
        {
            WriteAmbiguousPair();
            ContactService service = NewService(new FakePrompt(() => false));

            Assert.False(service.TryHandleFuzzyMatch("Bob James Smith", null, ContactInteractionMode.Interactive, "Subject"));
            Assert.Equal(0, service.GetAndClearAmbiguousCount());
        }

        [Fact]
        public void Interactive_Ambiguous_PromptRefuses_LoggedAsAmbiguous()
        {
            // The form's constructor throws InvalidOperationException in Automated mode; that path
            // logs the candidates and counts the contact as ambiguous, as before.
            WriteAmbiguousPair();
            ContactService service = NewService(new FakePrompt(() => { throw new InvalidOperationException("automated"); }));

            Assert.True(service.TryHandleFuzzyMatch("Bob James Smith", null, ContactInteractionMode.Interactive, "Subject"));
            Assert.Equal(1, service.GetAndClearAmbiguousCount());
            Assert.Contains("[[Bob Smith Jr]]", File.ReadAllText(LogPath));
        }

        [Fact]
        public void Interactive_Ambiguous_NoPrompt_LoggedAsAmbiguous()
        {
            WriteAmbiguousPair();
            ContactService service = NewService(null);

            Assert.True(service.TryHandleFuzzyMatch("Bob James Smith", null, ContactInteractionMode.Interactive, "Subject"));
            Assert.Equal(1, service.GetAndClearAmbiguousCount());
        }

        [Fact]
        public void Automated_Ambiguous_NeverPrompts()
        {
            WriteAmbiguousPair();
            FakePrompt prompt = new FakePrompt(() => true);
            ContactService service = NewService(prompt);

            Assert.True(service.TryHandleFuzzyMatch("Bob James Smith", null, ContactInteractionMode.Automated, "Subject"));
            Assert.Equal(0, prompt.Calls);
            Assert.Equal(1, service.GetAndClearAmbiguousCount());
        }

        [Fact]
        public void BuildContactTemplateContext_AllFields()
        {
            ContactTemplateContext context = NewService(null).BuildContactTemplateContext(
                "Jane Q. Doe", "555-0100", "jane@example.com", "Contoso", "CFO", "1 Main St", "1980-05-04", "Met at the conference");

            Assert.Equal(new[] { "title", "type", "created", "tags", "company", "email" }, context.Metadata.Keys.ToArray());
            Assert.Equal("Jane Q. Doe", context.Metadata["title"]);
            Assert.Equal("contact", context.Metadata["type"]);
            Assert.Equal("2026-09-23", context.Metadata["created"]);
            Assert.Equal(new List<string> { "contact" }, context.Metadata["tags"]);
            Assert.Equal("Contoso", context.Metadata["company"]);
            Assert.Equal("jane@example.com", context.Metadata["email"]);
            Assert.Equal("Jane Q. Doe", context.ContactName);
            Assert.Equal("JaneD", context.ContactShortName);
            Assert.Equal("2026-09-23", context.Created);
            Assert.Equal("Jane Q. Doe.md", context.FileName);
            Assert.Equal("Jane Q. Doe", context.FileNameWithoutExtension);
            Assert.Equal("555-0100", context.Phone);
            Assert.Equal("CFO", context.JobTitle);
            Assert.Equal("1 Main St", context.Address);
            Assert.Equal("1980-05-04", context.Birthday);
            Assert.Equal("Met at the conference", context.Notes);
            Assert.True(context.IncludeDetails);
            Assert.Equal("Jane", context.FirstName);
            Assert.Equal("Doe", context.LastName);
        }

        [Fact]
        public void BuildContactTemplateContext_NoCompanyOrEmail_LeavesThemOutOfTheProperties()
        {
            _settings.ContactDateFormat = "dd/MM/yyyy";

            ContactTemplateContext context = NewService(null).BuildContactTemplateContext(
                "Unknown Contact", "", "", "", "", "", "", "");

            Assert.Equal(new[] { "title", "type", "created", "tags" }, context.Metadata.Keys.ToArray());
            Assert.Equal("23/09/2026", context.Created);
        }
    }
}
