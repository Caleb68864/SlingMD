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

        // -----------------------------------------------------------------------------------------
        // ClassifyFromSignals (internal helper) tests
        // -----------------------------------------------------------------------------------------

        [Fact]
        public void ClassifyFromSignals_EmptySignals_ClassifiedAsReal()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);

            bool isInline = service.ClassifyFromSignals(new AttachmentSignals());

            Assert.False(isInline);
        }

        [Fact]
        public void ClassifyFromSignals_IsHidden_ClassifiedAsInline()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);

            bool isInline = service.ClassifyFromSignals(new AttachmentSignals { IsHidden = true });

            Assert.True(isInline);
        }

        [Fact]
        public void ClassifyFromSignals_HasMhtmlRef_ClassifiedAsInline()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);

            bool isInline = service.ClassifyFromSignals(new AttachmentSignals { HasMhtmlRef = true });

            Assert.True(isInline);
        }

        [Fact]
        public void ClassifyFromSignals_IsEmbeddedItem_ClassifiedAsInline()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);

            bool isInline = service.ClassifyFromSignals(new AttachmentSignals { IsEmbeddedItem = true });

            Assert.True(isInline);
        }

        [Fact]
        public void ClassifyFromSignals_NonEmptyContentId_ClassifiedAsInline()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);

            bool isInline = service.ClassifyFromSignals(new AttachmentSignals { ContentId = "abc123" });

            Assert.True(isInline);
        }

        // -----------------------------------------------------------------------------------------
        // ShouldSave (internal helper) tests
        // -----------------------------------------------------------------------------------------

        [Fact]
        public void ShouldSave_RealAttachmentWithSaveRealAttachmentsEnabled_ReturnsTrue()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);
            ObsidianSettings settings = new ObsidianSettings
            {
                SaveRealAttachments = true,
                SaveInlineImages = false,
                SaveAllAttachments = false
            };

            bool shouldSave = service.ShouldSave(isInline: false, settings: settings);

            Assert.True(shouldSave);
        }

        [Fact]
        public void ShouldSave_InlineAttachmentWithSaveInlineImagesEnabled_ReturnsTrue()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);
            ObsidianSettings settings = new ObsidianSettings
            {
                SaveRealAttachments = false,
                SaveInlineImages = true,
                SaveAllAttachments = false
            };

            bool shouldSave = service.ShouldSave(isInline: true, settings: settings);

            Assert.True(shouldSave);
        }

        [Fact]
        public void ShouldSave_SaveAllAttachmentsEnabled_OverridesOtherFlags()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);
            ObsidianSettings settings = new ObsidianSettings
            {
                SaveRealAttachments = false,
                SaveInlineImages = false,
                SaveAllAttachments = true
            };

            bool shouldSave = service.ShouldSave(isInline: true, settings: settings);

            Assert.True(shouldSave);
        }

        [Fact]
        public void ShouldSave_AllFlagsFalse_ReturnsFalse()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);
            ObsidianSettings settings = new ObsidianSettings
            {
                SaveRealAttachments = false,
                SaveInlineImages = false,
                SaveAllAttachments = false
            };

            bool shouldSaveReal = service.ShouldSave(isInline: false, settings: settings);
            bool shouldSaveInline = service.ShouldSave(isInline: true, settings: settings);

            Assert.False(shouldSaveReal);
            Assert.False(shouldSaveInline);
        }

        [Fact]
        public void ResolveInlineImageLinks_ReplacesCidWithSavedImage()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SubfolderPerNote);
            AttachmentInfo info = new AttachmentInfo();
            info.SavedAttachments.Add(new SavedAttachment
            {
                FullPath = @"C:\Vault\MyVault\Inbox\email\logo.png",
                ContentId = "<logo@example>",
                IsInline = true
            });

            string result = service.ResolveInlineImageLinks(
                "Before ![Logo](cid:logo@example) after", info,
                @"C:\Vault\MyVault\Inbox\email.md");

            Assert.Equal("Before ![[email/logo.png]] after", result);
        }

        [Fact]
        public void ResolveInlineImageLinks_MarkdownMode_PreservesAltText()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote, useWikilinks: false);
            AttachmentInfo info = new AttachmentInfo();
            info.SavedAttachments.Add(new SavedAttachment
            {
                FullPath = @"C:\Vault\MyVault\Inbox\logo.png",
                ContentId = "logo@example",
                IsInline = true
            });

            string result = service.ResolveInlineImageLinks(
                "![Company logo](cid:logo@example)", info,
                @"C:\Vault\MyVault\Inbox\email.md");

            Assert.Equal("![Company logo](logo.png)", result);
        }

        // -----------------------------------------------------------------------------------------
        // Duplicate reuse / empty folder cleanup
        // -----------------------------------------------------------------------------------------

        private static AttachmentService BuildServiceOnDisk(string root, bool reuseDuplicates)
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = root,
                VaultName = "Vault",
                InboxFolder = "Inbox",
                AttachmentStorageMode = AttachmentStorageMode.SubfolderPerNote,
                ReuseDuplicateAttachments = reuseDuplicates
            };
            return new AttachmentService(settings, new FileService(settings));
        }

        private static string CreateTempRoot()
        {
            string root = Path.Combine(Path.GetTempPath(), "SlingMDTests", "Attachment_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(Path.Combine(root, "Vault", "Inbox"));
            return root;
        }

        private static string WriteFile(string root, string relativePath, string content)
        {
            string path = Path.Combine(root, "Vault", relativePath);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, content);
            return path;
        }

        [Fact]
        public void ReuseDuplicateIfAny_Disabled_KeepsOwnCopy()
        {
            string root = CreateTempRoot();
            try
            {
                AttachmentService service = BuildServiceOnDisk(root, reuseDuplicates: false);
                WriteFile(root, Path.Combine("Inbox", "first", "logo.png"), "same-bytes");
                string second = WriteFile(root, Path.Combine("Inbox", "second", "logo.png"), "same-bytes");

                Assert.Equal(second, service.ReuseDuplicateIfAny(second));
                Assert.True(File.Exists(second));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        [Fact]
        public void ReuseDuplicateIfAny_SecondAndThirdCopies_LandOnOneSharedFile()
        {
            string root = CreateTempRoot();
            try
            {
                AttachmentService service = BuildServiceOnDisk(root, reuseDuplicates: true);
                string first = WriteFile(root, Path.Combine("Inbox", "first", "logo.png"), "same-bytes");
                string second = WriteFile(root, Path.Combine("Inbox", "second", "logo.png"), "same-bytes");

                string sharedPath = service.ReuseDuplicateIfAny(second);

                Assert.Equal(Path.Combine(root, "Vault", "Attachments", "Shared", "logo.png"), sharedPath);
                Assert.False(File.Exists(second));
                Assert.True(File.Exists(first));

                string third = WriteFile(root, Path.Combine("Inbox", "third", "image001.png"), "same-bytes");
                Assert.Equal(sharedPath, service.ReuseDuplicateIfAny(third));
                Assert.False(File.Exists(third));
                Assert.Single(Directory.GetFiles(Path.Combine(root, "Vault", "Attachments", "Shared")));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        [Fact]
        public void ReuseDuplicateIfAny_SameSizeDifferentContent_KeepsOwnCopy()
        {
            string root = CreateTempRoot();
            try
            {
                AttachmentService service = BuildServiceOnDisk(root, reuseDuplicates: true);
                WriteFile(root, Path.Combine("Inbox", "first", "logo.png"), "aaaa");
                string second = WriteFile(root, Path.Combine("Inbox", "second", "logo.png"), "bbbb");

                Assert.Equal(second, service.ReuseDuplicateIfAny(second));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        [Fact]
        public void BuildAttachmentLinkTarget_SharedAttachment_UsesVaultRootPath()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SubfolderPerNote);

            string target = service.BuildAttachmentLinkTarget(
                @"C:\Vault\MyVault\Attachments\Shared\logo.png",
                @"C:\Vault\MyVault\Inbox\Some thread\email.md");

            Assert.Equal("Attachments/Shared/logo.png", target);
        }

        [Fact]
        public void RemoveEmptyNoteFolder_DeletesOnlyEmptyFolders()
        {
            string root = CreateTempRoot();
            try
            {
                AttachmentService service = BuildServiceOnDisk(root, reuseDuplicates: false);
                string emptyFolder = Path.Combine(root, "Vault", "Inbox", "empty-note");
                Directory.CreateDirectory(emptyFolder);
                string usedFile = WriteFile(root, Path.Combine("Inbox", "used-note", "doc.pdf"), "pdf");

                service.RemoveEmptyNoteFolder(emptyFolder);
                service.RemoveEmptyNoteFolder(Path.GetDirectoryName(usedFile));

                Assert.False(Directory.Exists(emptyFolder));
                Assert.True(File.Exists(usedFile));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        // -----------------------------------------------------------------------------------------
        // Grouping the note with its attachments
        // -----------------------------------------------------------------------------------------

        private static AttachmentService BuildGroupingService(string root, bool group)
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = root,
                VaultName = "Vault",
                InboxFolder = "Inbox",
                AttachmentStorageMode = AttachmentStorageMode.SubfolderPerNote,
                GroupNoteWithAttachments = group
            };
            return new AttachmentService(settings, new FileService(settings));
        }

        [Fact]
        public void GroupNoteWithAttachments_FolderExists_MovesNoteIntoIt()
        {
            string root = CreateTempRoot();
            try
            {
                AttachmentService service = BuildGroupingService(root, group: true);
                string note = WriteFile(root, Path.Combine("Inbox", "email.md"), "note");
                string attachment = WriteFile(root, Path.Combine("Inbox", "email", "doc.pdf"), "pdf");

                string grouped = service.GroupNoteWithAttachments(note);

                Assert.Equal(Path.Combine(root, "Vault", "Inbox", "email", "email.md"), grouped);
                Assert.True(File.Exists(grouped));
                Assert.False(File.Exists(note));
                Assert.Equal("doc.pdf", service.BuildAttachmentLinkTarget(attachment, grouped));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        [Fact]
        public void GroupNoteWithAttachments_NoAttachmentFolder_LeavesNoteInPlace()
        {
            string root = CreateTempRoot();
            try
            {
                AttachmentService service = BuildGroupingService(root, group: true);
                string note = WriteFile(root, Path.Combine("Inbox", "email.md"), "note");

                Assert.Equal(note, service.GroupNoteWithAttachments(note));
                Assert.True(File.Exists(note));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        [Fact]
        public void GroupNoteWithAttachments_Disabled_LeavesNoteInPlace()
        {
            string root = CreateTempRoot();
            try
            {
                AttachmentService service = BuildGroupingService(root, group: false);
                string note = WriteFile(root, Path.Combine("Inbox", "email.md"), "note");
                WriteFile(root, Path.Combine("Inbox", "email", "doc.pdf"), "pdf");

                Assert.Equal(note, service.GroupNoteWithAttachments(note));
            }
            finally
            {
                Directory.Delete(root, true);
            }
        }

        [Fact]
        public void GenerateWikilink_MarkdownLinks_EscapeSpacesInTarget()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote, useWikilinks: false);

            string link = service.GenerateWikilink(@"C:\Vault\MyVault\Inbox\Q3 Report #2.pdf", @"C:\Vault\MyVault\Inbox\email.md", isImage: false);

            Assert.Equal("[Q3 Report #2.pdf](Q3%20Report%20%232.pdf)", link);
        }

        [Fact]
        public void BuildAttachmentSection_ListsFilesAsBulletsAndSkipsInlineImagesShownInBody()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);
            string noteFile = @"C:\Vault\MyVault\Inbox\email.md";
            AttachmentInfo info = new AttachmentInfo();
            info.SavedAttachments.Add(new SavedAttachment { FullPath = @"C:\Vault\MyVault\Inbox\report.pdf" });
            info.SavedAttachments.Add(new SavedAttachment { FullPath = @"C:\Vault\MyVault\Inbox\logo.png", IsInline = true });
            info.SavedAttachments.Add(new SavedAttachment { FullPath = @"C:\Vault\MyVault\Inbox\unused.png", IsInline = true });

            string section = service.BuildAttachmentSection(info, noteFile, "Hello\n\n![[logo.png]]\n");

            Assert.Equal("\n\n## Attachments\n\n- [[report.pdf]]\n- [[unused.png]]\n", section);
        }

        [Fact]
        public void BuildAttachmentSection_OnlyInlineImagesShown_ReturnsEmpty()
        {
            AttachmentService service = BuildService(AttachmentStorageMode.SameAsNote);
            AttachmentInfo info = new AttachmentInfo();
            info.SavedAttachments.Add(new SavedAttachment { FullPath = @"C:\Vault\MyVault\Inbox\logo.png", IsInline = true });

            Assert.Equal(string.Empty, service.BuildAttachmentSection(info, @"C:\Vault\MyVault\Inbox\email.md", "![[logo.png]]"));
        }
    }
}
