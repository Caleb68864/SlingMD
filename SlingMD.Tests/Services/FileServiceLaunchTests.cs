using System;
using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// <see cref="FileService.LaunchObsidian"/> with the URI opener replaced: the URI it opens, and the
    /// <see cref="IUserNotifier"/> warning when Obsidian can't be launched. Linked into SlingMD.Core.Tests.
    /// </summary>
    public class FileServiceLaunchTests
    {
        private sealed class RecordingNotifier : IUserNotifier
        {
            public readonly List<string> Warnings = new List<string>();

            public void ShowWarning(string message, string title)
            {
                Warnings.Add(title + "|" + message);
            }

            public void ShowInformation(string message, string title)
            {
                throw new InvalidOperationException("not expected");
            }
        }

        private static ObsidianSettings Settings()
        {
            return new ObsidianSettings
            {
                VaultBasePath = Path.Combine(Path.GetTempPath(), "SlingMDTests", "Launch"),
                VaultName = "Vault"
            };
        }

        [Fact]
        public void LaunchObsidian_OpensTheVaultRelativeUri_AndShowsNothing()
        {
            ObsidianSettings settings = Settings();
            RecordingNotifier notifier = new RecordingNotifier();
            List<string> opened = new List<string>();
            FileService service = new FileService(settings, notifier, uri => opened.Add(uri));

            service.LaunchObsidian("Vault", Path.Combine(settings.GetFullVaultPath(), "Inbox", "Budget review.md"));

            Assert.Equal(new List<string> { "obsidian://open?vault=Vault&file=Inbox%2FBudget%20review" }, opened);
            Assert.Empty(notifier.Warnings);
        }

        public static IEnumerable<object[]> LaunchFailures()
        {
            yield return new object[] { new System.ComponentModel.Win32Exception(2, "The system cannot find the file specified") };
            yield return new object[] { new InvalidOperationException("No file name was specified") };
            yield return new object[] { new FileNotFoundException("obsidian not found") };
        }

        [Theory]
        [MemberData(nameof(LaunchFailures))]
        public void LaunchObsidian_LaunchFails_WarnsThroughTheNotifier(Exception failure)
        {
            RecordingNotifier notifier = new RecordingNotifier();
            FileService service = new FileService(Settings(), notifier, uri => { throw failure; });

            service.LaunchObsidian("Vault", "Inbox/Note");

            Assert.Equal(
                new List<string> { "SlingMD|Failed to launch Obsidian: " + failure.Message + "\n\nIs Obsidian installed and registered for obsidian:// links?" },
                notifier.Warnings);
        }

        [Fact]
        public void LaunchObsidian_LaunchFailsWithoutNotifier_DoesNotThrow()
        {
            FileService service = new FileService(Settings(), null, uri => { throw new System.ComponentModel.Win32Exception(2); });

            service.LaunchObsidian("Vault", "Inbox/Note");
        }

        [Fact]
        public void LaunchObsidian_UnexpectedFailure_Propagates()
        {
            // Only the launch failures are absorbed; anything else is a bug and surfaces.
            RecordingNotifier notifier = new RecordingNotifier();
            FileService service = new FileService(Settings(), notifier, uri => { throw new NullReferenceException(); });

            Assert.Throws<NullReferenceException>(() => service.LaunchObsidian("Vault", "Inbox/Note"));
            Assert.Empty(notifier.Warnings);
        }

        [Fact]
        public void LaunchObsidian_EmptyVaultOrPath_OpensNothing()
        {
            List<string> opened = new List<string>();
            FileService service = new FileService(Settings(), new RecordingNotifier(), uri => opened.Add(uri));

            service.LaunchObsidian("", "Inbox/Note");
            service.LaunchObsidian("Vault", "");

            Assert.Empty(opened);
        }
    }
}
