using System;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Unit tests for <see cref="AttachmentService"/> link generation, verifying that
    /// the correct path form is used for each <see cref="AttachmentStorageMode"/>.
    /// </summary>
    public class AttachmentServiceTests
    {
        private AttachmentService BuildService(AttachmentStorageMode mode, bool useWikilinks = true)
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = @"C:\Vault",
                VaultName = "MyVault",
                AttachmentStorageMode = mode,
                UseObsidianWikilinks = useWikilinks
            };
            FileService fileService = new FileService(settings);
            return new AttachmentService(settings, fileService);
        }

        // -----------------------------------------------------------------------------------------
        // BuildAttachmentLinkTarget (internal helper) tests
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// For SameAsNote mode the attachment lives in the same directory as the note.
        /// The link target must be the bare filename with no path prefix.
        /// </summary>
        [Fact]
        public void BuildAttachmentLinkTarget_SameFolder_UsesBareFilename()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);

            string noteFile = @"C:\Vault\MyVault\Inbox\2026-03-12_my-email.md";
            string attachFile = @"C:\Vault\MyVault\Inbox\image.png";

            string target = service.BuildAttachmentLinkTarget(attachFile, noteFile);

            Assert.Equal("image.png", target);
        }

        /// <summary>
        /// For SubfolderPerNote mode the attachment lives in a subdirectory of the note directory
        /// named after the note.  The link target must contain the subfolder segment.
        /// </summary>
        [Fact]
        public void BuildAttachmentLinkTarget_SubfolderPerNote_UsesRelativeSubfolderPath()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SubfolderPerNote);

            string noteFile = @"C:\Vault\MyVault\Inbox\2026-03-12_my-email.md";
            string attachFile = @"C:\Vault\MyVault\Inbox\2026-03-12_my-email\image.png";

            string target = service.BuildAttachmentLinkTarget(attachFile, noteFile);

            // Should be relative from note dir: "2026-03-12_my-email/image.png"
            Assert.Equal("2026-03-12_my-email/image.png", target);
        }

        /// <summary>
        /// For Centralized mode the attachment lives in a vault-level attachments folder.
        /// The link target must navigate upward from the note directory then into the attachments path.
        /// </summary>
        [Fact]
        public void BuildAttachmentLinkTarget_CentralizedStorage_UsesRelativeVaultPath()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.Centralized);

            string noteFile = @"C:\Vault\MyVault\Inbox\2026-03-12_my-email.md";
            string attachFile = @"C:\Vault\MyVault\Attachments\2026-03\document.pdf";

            string target = service.BuildAttachmentLinkTarget(attachFile, noteFile);

            // From Inbox/ we need to go up one level then into Attachments/2026-03/
            Assert.Equal("../Attachments/2026-03/document.pdf", target);
        }

        // -----------------------------------------------------------------------------------------
        // GenerateWikilink with path overload tests
        // -----------------------------------------------------------------------------------------

        /// <summary>
        /// GenerateAttachmentLink_SameFolder: the generated wikilink uses only the bare filename.
        /// </summary>
        [Fact]
        public void GenerateAttachmentLink_SameFolder_UsesBareFilename()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);

            string noteFile = @"C:\Vault\MyVault\Inbox\email.md";
            string attachFile = @"C:\Vault\MyVault\Inbox\photo.png";

            string link = service.GenerateWikilink(attachFile, noteFile, isImage: true);

            Assert.Equal("![[photo.png]]", link);
        }

        /// <summary>
        /// GenerateAttachmentLink_SubfolderPerNote: the generated wikilink includes the subfolder prefix.
        /// </summary>
        [Fact]
        public void GenerateAttachmentLink_SubfolderPerNote_UsesRelativeSubfolderPath()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SubfolderPerNote);

            string noteFile = @"C:\Vault\MyVault\Inbox\email.md";
            string attachFile = @"C:\Vault\MyVault\Inbox\email\photo.png";

            string link = service.GenerateWikilink(attachFile, noteFile, isImage: true);

            Assert.Equal("![[email/photo.png]]", link);
        }

        /// <summary>
        /// GenerateAttachmentLink_CentralizedStorage: the generated wikilink uses a vault-relative path.
        /// </summary>
        [Fact]
        public void GenerateAttachmentLink_CentralizedStorage_UsesRelativeVaultPath()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.Centralized);

            string noteFile = @"C:\Vault\MyVault\Inbox\email.md";
            string attachFile = @"C:\Vault\MyVault\Attachments\2026-03\report.pdf";

            string link = service.GenerateWikilink(attachFile, noteFile, isImage: false);

            Assert.Equal("[[../Attachments/2026-03/report.pdf]]", link);
        }

        /// <summary>
        /// The single-parameter overload (bare filename) must continue to work unchanged
        /// for callers that only have the filename available.
        /// </summary>
        [Fact]
        public void GenerateWikilink_LegacyOverload_StillWorks()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);

            string link = service.GenerateWikilink("image.png", isImage: true);

            Assert.Equal("![[image.png]]", link);
        }
    }
}
