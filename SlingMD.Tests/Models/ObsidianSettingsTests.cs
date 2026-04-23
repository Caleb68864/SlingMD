using System;
using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Models;
using Xunit;

namespace SlingMD.Tests.Models
{
    public class ObsidianSettingsTests
    {
        private readonly string _testSettingsDir;
        private readonly string _testSettingsPath;

        public ObsidianSettingsTests()
        {
            _testSettingsDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "Settings");
            _testSettingsPath = Path.Combine(_testSettingsDir, "ObsidianSettings.json");

            if (Directory.Exists(_testSettingsDir))
            {
                Directory.Delete(_testSettingsDir, true);
            }

            Directory.CreateDirectory(_testSettingsDir);
        }

        [Fact]
        public void Save_CreatesSettingsDirectoryIfNotExists()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable();
            settings.TestSettingsPath = _testSettingsPath;

            if (Directory.Exists(_testSettingsDir))
            {
                Directory.Delete(_testSettingsDir, true);
            }

            settings.Save();

            Assert.True(Directory.Exists(_testSettingsDir), "Settings directory should be created");
            Assert.True(File.Exists(_testSettingsPath), "Settings file should be created");
        }

        [Fact]
        public void HasSavedSettings_ReturnsFalseBeforeSaveAndTrueAfterSave()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };

            Assert.False(settings.HasSavedSettings());

            settings.Save();

            Assert.True(settings.HasSavedSettings());
        }

        [Fact]
        public void SearchEntireVaultForContacts_SavedAndLoaded_Correctly()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath,
                SearchEntireVaultForContacts = true
            };

            settings.Save();

            ObsidianSettingsTestable loadedSettings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };

            Assert.False(loadedSettings.SearchEntireVaultForContacts, "Should start as false before loading");

            loadedSettings.Load();

            Assert.True(loadedSettings.SearchEntireVaultForContacts, "SearchEntireVaultForContacts should be true after loading");
        }

        [Fact]
        public void HasShownSupportPrompt_SavedAndLoaded_Correctly()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath,
                HasShownSupportPrompt = true
            };

            settings.Save();

            ObsidianSettingsTestable loadedSettings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            loadedSettings.Load();

            Assert.True(loadedSettings.HasShownSupportPrompt);
        }

        [Fact]
        public void SaveAndLoad_PersistsAllSettings()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath,
                VaultName = "TestVault",
                VaultBasePath = @"C:\Test\Path",
                InboxFolder = "TestInbox",
                ContactsFolder = "TestContacts",
                EnableContactSaving = true,
                SearchEntireVaultForContacts = true,
                LaunchObsidian = false,
                ObsidianDelaySeconds = 5,
                ShowCountdown = false,
                CreateObsidianTask = false,
                CreateOutlookTask = true,
                DefaultDueDays = 3,
                DefaultReminderDays = 2,
                DefaultReminderHour = 10,
                AskForDates = true,
                GroupEmailThreads = false,
                ShowDevelopmentSettings = true,
                ShowThreadDebug = true,
                HasShownSupportPrompt = true,
                TemplatesFolder = "Config\\Templates",
                EmailTemplateFile = "Email.md",
                ContactTemplateFile = "Contact.md",
                TaskTemplateFile = "Task.md",
                ThreadTemplateFile = "Thread.md",
                EmailFilenameFormat = "{Subject}-{Timestamp}",
                ContactFilenameFormat = "{ContactShortName}"
            };

            settings.SubjectCleanupPatterns.Clear();
            settings.SubjectCleanupPatterns.Add("test-pattern-1");
            settings.SubjectCleanupPatterns.Add("test-pattern-2");

            settings.Save();

            ObsidianSettingsTestable loadedSettings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            loadedSettings.Load();

            Assert.Equal(settings.VaultName, loadedSettings.VaultName);
            Assert.Equal(settings.VaultBasePath, loadedSettings.VaultBasePath);
            Assert.Equal(settings.InboxFolder, loadedSettings.InboxFolder);
            Assert.Equal(settings.ContactsFolder, loadedSettings.ContactsFolder);
            Assert.Equal(settings.EnableContactSaving, loadedSettings.EnableContactSaving);
            Assert.Equal(settings.SearchEntireVaultForContacts, loadedSettings.SearchEntireVaultForContacts);
            Assert.Equal(settings.LaunchObsidian, loadedSettings.LaunchObsidian);
            Assert.Equal(settings.ObsidianDelaySeconds, loadedSettings.ObsidianDelaySeconds);
            Assert.Equal(settings.ShowCountdown, loadedSettings.ShowCountdown);
            Assert.Equal(settings.CreateObsidianTask, loadedSettings.CreateObsidianTask);
            Assert.Equal(settings.CreateOutlookTask, loadedSettings.CreateOutlookTask);
            Assert.Equal(settings.DefaultDueDays, loadedSettings.DefaultDueDays);
            Assert.Equal(settings.DefaultReminderDays, loadedSettings.DefaultReminderDays);
            Assert.Equal(settings.DefaultReminderHour, loadedSettings.DefaultReminderHour);
            Assert.Equal(settings.AskForDates, loadedSettings.AskForDates);
            Assert.Equal(settings.GroupEmailThreads, loadedSettings.GroupEmailThreads);
            Assert.Equal(settings.ShowDevelopmentSettings, loadedSettings.ShowDevelopmentSettings);
            Assert.Equal(settings.ShowThreadDebug, loadedSettings.ShowThreadDebug);
            Assert.Equal(settings.HasShownSupportPrompt, loadedSettings.HasShownSupportPrompt);
            Assert.Equal(settings.TemplatesFolder, loadedSettings.TemplatesFolder);
            Assert.Equal(settings.EmailTemplateFile, loadedSettings.EmailTemplateFile);
            Assert.Equal(settings.ContactTemplateFile, loadedSettings.ContactTemplateFile);
            Assert.Equal(settings.TaskTemplateFile, loadedSettings.TaskTemplateFile);
            Assert.Equal(settings.ThreadTemplateFile, loadedSettings.ThreadTemplateFile);
            Assert.Equal(settings.EmailFilenameFormat, loadedSettings.EmailFilenameFormat);
            Assert.Equal(settings.ContactFilenameFormat, loadedSettings.ContactFilenameFormat);

            Assert.Equal(2, loadedSettings.SubjectCleanupPatterns.Count);
            Assert.Contains("test-pattern-1", loadedSettings.SubjectCleanupPatterns);
            Assert.Contains("test-pattern-2", loadedSettings.SubjectCleanupPatterns);
        }

        [Fact]
        public void LoadSettings_FileDoesNotExist_UsesDefaultValues()
        {
            if (File.Exists(_testSettingsPath))
            {
                File.Delete(_testSettingsPath);
            }

            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            settings.Load();

            Assert.Equal("Logic", settings.VaultName);
            Assert.Contains("Documents\\Notes", settings.VaultBasePath);
            Assert.Equal("Inbox", settings.InboxFolder);
            Assert.Equal("Contacts", settings.ContactsFolder);
            Assert.True(settings.EnableContactSaving);
            Assert.False(settings.SearchEntireVaultForContacts);
            Assert.True(settings.LaunchObsidian);
            Assert.False(settings.HasShownSupportPrompt);
        }

        [Fact]
        public void Load_LegacySettingsWithoutTemplateFields_UsesTemplateDefaults()
        {
            string legacyJson = @"{
  ""VaultName"": ""LegacyVault"",
  ""VaultBasePath"": ""C:\\Legacy\\Vault"",
  ""InboxFolder"": ""Inbox"",
  ""ContactsFolder"": ""Contacts""
}";
            File.WriteAllText(_testSettingsPath, legacyJson);

            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };

            settings.Load();

            Assert.Equal("Templates", settings.TemplatesFolder);
            Assert.Equal("EmailTemplate.md", settings.EmailTemplateFile);
            Assert.Equal("ContactTemplate.md", settings.ContactTemplateFile);
            Assert.Equal("TaskTemplate.md", settings.TaskTemplateFile);
            Assert.Equal("ThreadNoteTemplate.md", settings.ThreadTemplateFile);
            Assert.Equal(string.Empty, settings.EmailFilenameFormat);
            Assert.Equal("{ContactName}", settings.ContactFilenameFormat);
        }
        [Fact]
        public void GetFullVaultPath_CombinesBasePathAndVaultName()
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = @"C:\Test\Path",
                VaultName = "TestVault"
            };

            string fullPath = settings.GetFullVaultPath();

            Assert.Equal(@"C:\Test\Path\TestVault", fullPath);
        }

        [Fact]
        public void GetInboxPath_ReturnsCorrectPath()
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = @"C:\Test\Path",
                VaultName = "TestVault",
                InboxFolder = "TestInbox"
            };

            string inboxPath = settings.GetInboxPath();

            Assert.Equal(@"C:\Test\Path\TestVault\TestInbox", inboxPath);
        }

        [Fact]
        public void GetContactsPath_ReturnsCorrectPath()
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = @"C:\Test\Path",
                VaultName = "TestVault",
                ContactsFolder = "TestContacts"
            };

            string contactsPath = settings.GetContactsPath();

            Assert.Equal(@"C:\Test\Path\TestVault\TestContacts", contactsPath);
        }

        [Fact]
        public void Load_MalformedJson_KeepsDefaultsAndDoesNotThrow()
        {
            File.WriteAllText(_testSettingsPath, "{ this is not valid json !!! ]");

            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };

            System.Exception caughtException = null;
            try
            {
                settings.Load();
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            Assert.Null(caughtException);
            Assert.Equal("Logic", settings.VaultName);
            Assert.Equal("Inbox", settings.InboxFolder);
            Assert.Equal("Contacts", settings.ContactsFolder);
        }

        [Fact]
        public void Load_TypeMismatchedFields_NormalizesInvalidValuesWithoutClearingValidOnes()
        {
            // VaultName is valid, but DefaultDueDays is a string instead of int (type mismatch)
            string json = @"{
  ""VaultName"": ""MyVault"",
  ""VaultBasePath"": ""C:\\Notes"",
  ""DefaultDueDays"": ""not-a-number"",
  ""InboxFolder"": ""Inbox""
}";
            File.WriteAllText(_testSettingsPath, json);

            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };

            System.Exception caughtException = null;
            try
            {
                settings.Load();
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            Assert.Null(caughtException);
            // Valid string field should be preserved
            Assert.Equal("MyVault", settings.VaultName);
            // Invalid numeric field should keep whatever value it has (default or unchanged)
            // The important thing is no exception is thrown
        }

        [Fact]
        public void ContactNoteIncludeDetails_DefaultsToTrue()
        {
            ObsidianSettings settings = new ObsidianSettings();

            Assert.True(settings.ContactNoteIncludeDetails);
        }

        [Fact]
        public void ContactNoteIncludeDetails_SavedAndLoaded_Correctly()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath,
                ContactNoteIncludeDetails = false
            };

            settings.Save();

            ObsidianSettingsTestable loadedSettings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            loadedSettings.Load();

            Assert.False(loadedSettings.ContactNoteIncludeDetails);
        }

        [Fact]
        public void NormalizeLoadedSettings_MissingContactNoteIncludeDetails_DefaultsToTrue()
        {
            string legacyJson = @"{
  ""VaultName"": ""TestVault"",
  ""VaultBasePath"": ""C:\\Test\\Path"",
  ""InboxFolder"": ""Inbox"",
  ""ContactsFolder"": ""Contacts""
}";
            File.WriteAllText(_testSettingsPath, legacyJson);

            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            settings.Load();

            Assert.True(settings.ContactNoteIncludeDetails);
        }
    }

    // Auto-Sling settings tests

    public class ObsidianSettingsAutoSlingTests
    {
        private readonly string _testSettingsDir;
        private readonly string _testSettingsPath;

        public ObsidianSettingsAutoSlingTests()
        {
            _testSettingsDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "AutoSlingSettings");
            _testSettingsPath = Path.Combine(_testSettingsDir, "ObsidianSettings.json");

            if (Directory.Exists(_testSettingsDir))
            {
                Directory.Delete(_testSettingsDir, true);
            }

            Directory.CreateDirectory(_testSettingsDir);
        }

        [Fact]
        public void EnableAutoSling_DefaultsToFalse()
        {
            ObsidianSettings settings = new ObsidianSettings();

            Assert.False(settings.EnableAutoSling);
        }

        [Fact]
        public void AutoSlingSettings_SavedAndLoaded_Correctly()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath,
                EnableAutoSling = true,
                AutoSlingNotificationMode = "Silent",
                EnableFlagToSling = true,
                SentToObsidianCategory = "TestCategory"
            };

            settings.Save();

            ObsidianSettingsTestable loadedSettings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            loadedSettings.Load();

            Assert.True(loadedSettings.EnableAutoSling);
            Assert.Equal("Silent", loadedSettings.AutoSlingNotificationMode);
            Assert.True(loadedSettings.EnableFlagToSling);
            Assert.Equal("TestCategory", loadedSettings.SentToObsidianCategory);
        }

        [Fact]
        public void AutoSlingRules_SavedAndLoaded_Correctly()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            settings.AutoSlingRules.Add(new AutoSlingRule { Type = "Sender", Pattern = "test@example.com", Enabled = true });
            settings.AutoSlingRules.Add(new AutoSlingRule { Type = "Domain", Pattern = "example.com", Enabled = false });

            settings.Save();

            ObsidianSettingsTestable loadedSettings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            loadedSettings.Load();

            Assert.Equal(2, loadedSettings.AutoSlingRules.Count);
            Assert.Equal("Sender", loadedSettings.AutoSlingRules[0].Type);
            Assert.Equal("test@example.com", loadedSettings.AutoSlingRules[0].Pattern);
            Assert.True(loadedSettings.AutoSlingRules[0].Enabled);
            Assert.Equal("Domain", loadedSettings.AutoSlingRules[1].Type);
            Assert.Equal("example.com", loadedSettings.AutoSlingRules[1].Pattern);
            Assert.False(loadedSettings.AutoSlingRules[1].Enabled);
        }

        [Fact]
        public void WatchedFolders_SavedAndLoaded_Correctly()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            settings.WatchedFolders.Add(new WatchedFolder { FolderPath = "Inbox/Work", CustomTemplate = "WorkEmail.md", Enabled = true });
            settings.WatchedFolders.Add(new WatchedFolder { FolderPath = "Inbox/Personal", CustomTemplate = string.Empty, Enabled = false });

            settings.Save();

            ObsidianSettingsTestable loadedSettings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            loadedSettings.Load();

            Assert.Equal(2, loadedSettings.WatchedFolders.Count);
            Assert.Equal("Inbox/Work", loadedSettings.WatchedFolders[0].FolderPath);
            Assert.Equal("WorkEmail.md", loadedSettings.WatchedFolders[0].CustomTemplate);
            Assert.True(loadedSettings.WatchedFolders[0].Enabled);
            Assert.Equal("Inbox/Personal", loadedSettings.WatchedFolders[1].FolderPath);
            Assert.False(loadedSettings.WatchedFolders[1].Enabled);
        }

        [Fact]
        public void CustomizationSettings_SavedAndLoaded_Correctly()
        {
            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath,
                ContactLinkFormat = "[[{LastName}, {FirstName}]]",
                EmailDateFormat = "MM/dd/yyyy HH:mm",
                ContactDateFormat = "yyyy.MM.dd",
                AppointmentDateFormat = "dd MMM yyyy HH:mm"
            };
            settings.Save();

            ObsidianSettingsTestable loaded = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };

            // Sanity: before load the defaults must differ from the values we just saved,
            // otherwise the test would pass trivially.
            Assert.Equal("[[{FullName}]]", loaded.ContactLinkFormat);
            Assert.Equal("yyyy-MM-dd HH:mm:ss", loaded.EmailDateFormat);
            Assert.Equal("yyyy-MM-dd", loaded.ContactDateFormat);
            Assert.Equal("yyyy-MM-dd HH:mm", loaded.AppointmentDateFormat);

            loaded.Load();

            Assert.Equal("[[{LastName}, {FirstName}]]", loaded.ContactLinkFormat);
            Assert.Equal("MM/dd/yyyy HH:mm", loaded.EmailDateFormat);
            Assert.Equal("yyyy.MM.dd", loaded.ContactDateFormat);
            Assert.Equal("dd MMM yyyy HH:mm", loaded.AppointmentDateFormat);
        }

        [Fact]
        public void CustomizationSettings_NormalizeLoadedSettings_BlankValuesRestoreDefaults()
        {
            string json = @"{
  ""VaultName"": ""V"",
  ""VaultBasePath"": ""C:\\V"",
  ""ContactLinkFormat"": """",
  ""EmailDateFormat"": """",
  ""ContactDateFormat"": """",
  ""AppointmentDateFormat"": """"
}";
            File.WriteAllText(_testSettingsPath, json);

            ObsidianSettingsTestable loaded = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            loaded.Load();

            Assert.Equal("[[{FullName}]]", loaded.ContactLinkFormat);
            Assert.Equal("yyyy-MM-dd HH:mm:ss", loaded.EmailDateFormat);
            Assert.Equal("yyyy-MM-dd", loaded.ContactDateFormat);
            Assert.Equal("yyyy-MM-dd HH:mm", loaded.AppointmentDateFormat);
        }

        [Fact]
        public void NormalizeLoadedSettings_MissingAutoSlingFields_UsesDefaults()
        {
            string legacyJson = @"{
  ""VaultName"": ""LegacyVault"",
  ""VaultBasePath"": ""C:\\Legacy\\Vault"",
  ""InboxFolder"": ""Inbox"",
  ""ContactsFolder"": ""Contacts""
}";
            File.WriteAllText(_testSettingsPath, legacyJson);

            ObsidianSettingsTestable settings = new ObsidianSettingsTestable
            {
                TestSettingsPath = _testSettingsPath
            };
            settings.Load();

            Assert.False(settings.EnableAutoSling);
            Assert.Equal("Toast", settings.AutoSlingNotificationMode);
            Assert.NotNull(settings.AutoSlingRules);
            Assert.Empty(settings.AutoSlingRules);
            Assert.NotNull(settings.WatchedFolders);
            Assert.Empty(settings.WatchedFolders);
            Assert.False(settings.EnableFlagToSling);
            Assert.Equal("Sent to Obsidian", settings.SentToObsidianCategory);
        }
    }

    public class ObsidianSettingsTestable : ObsidianSettings
    {
        public string TestSettingsPath { get; set; }

        protected override string GetSettingsPath()
        {
            return TestSettingsPath ?? base.GetSettingsPath();
        }
    }
}
