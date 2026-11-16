using System;
using System.IO;
using System.Text;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class DottedFolderLinkRepairServiceTests : IDisposable
    {
        private readonly string _vault;

        public DottedFolderLinkRepairServiceTests()
        {
            _vault = Path.Combine(Path.GetTempPath(), "SlingMDDottedLinks_" + Guid.NewGuid().ToString("N"));
            // As Windows created them before FolderNameSanitizer: the trailing dot is gone on disk.
            Directory.CreateDirectory(Path.Combine(_vault, "Inbox", "Re_ Budget-Robert Smith Jr", "Re_ Budget-Robert Smith Jr"));
            File.WriteAllText(Path.Combine(_vault, "Inbox", "Re_ Budget-Robert Smith Jr", "Re_ Budget-Robert Smith Jr", "file.pdf"), "pdf");
            Directory.CreateDirectory(Path.Combine(_vault, ".obsidian"));
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(_vault, true);
            }
            catch (IOException)
            {
            }
        }

        private string WriteNote(string relativePath, string content, bool bom = false)
        {
            string path = Path.Combine(_vault, relativePath);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, content, new UTF8Encoding(bom));
            return path;
        }

        [Fact]
        public void RepairVault_RewritesNoteRelativeAndVaultRelativeLinks()
        {
            string threadNote = WriteNote(Path.Combine("Inbox", "Re_ Budget-Robert Smith Jr", "Re_ Budget-Robert Smith Jr..md"),
                "Body\n![[Re_ Budget-Robert Smith Jr./file.pdf]]\n- [file.pdf](Re_%20Budget-Robert%20Smith%20Jr./file.pdf)\n");
            string elsewhere = WriteNote(Path.Combine("Notes", "Mine.md"),
                "See [[Inbox/Re_ Budget-Robert Smith Jr./Re_ Budget-Robert Smith Jr./file.pdf|the pdf]].\n");

            DottedFolderLinkRepairService service = new DottedFolderLinkRepairService(_vault);
            DottedFolderLinkRepairReport report = service.RepairVault();

            Assert.Equal("Body\n![[Re_ Budget-Robert Smith Jr/file.pdf]]\n- [file.pdf](Re_%20Budget-Robert%20Smith%20Jr/file.pdf)\n",
                File.ReadAllText(threadNote));
            Assert.Equal("See [[Inbox/Re_ Budget-Robert Smith Jr/Re_ Budget-Robert Smith Jr/file.pdf|the pdf]].\n",
                File.ReadAllText(elsewhere));
            Assert.Equal(2, report.NotesScanned);
            Assert.Equal(2, report.NotesChanged);
            Assert.Equal(3, report.LinksRewritten);
            Assert.Empty(report.FailedNotes);
        }

        [Fact]
        public void RepairVault_LeavesBrokenLinksWithNoRepairAndUnrelatedNotesUntouched()
        {
            string note = WriteNote("Other.md", "[[Missing Jr./file.pdf]] [[Inbox/Re_ Budget-Robert Smith Jr/Re_ Budget-Robert Smith Jr/file.pdf]]\n");
            DateTime before = File.GetLastWriteTimeUtc(note);

            DottedFolderLinkRepairReport report = new DottedFolderLinkRepairService(_vault).RepairVault();

            Assert.Equal("[[Missing Jr./file.pdf]] [[Inbox/Re_ Budget-Robert Smith Jr/Re_ Budget-Robert Smith Jr/file.pdf]]\n",
                File.ReadAllText(note));
            Assert.Equal(before, File.GetLastWriteTimeUtc(note));
            Assert.Equal(0, report.NotesChanged);
        }

        [Fact]
        public void RepairVault_KeepsByteOrderMarkAndSkipsDotFolders()
        {
            string withBom = WriteNote("Bom.md", "[[Inbox/Re_ Budget-Robert Smith Jr./Re_ Budget-Robert Smith Jr./file.pdf]]", bom: true);
            string config = WriteNote(Path.Combine(".obsidian", "Snippet.md"), "[[Inbox/Re_ Budget-Robert Smith Jr./Re_ Budget-Robert Smith Jr./file.pdf]]");

            new DottedFolderLinkRepairService(_vault).RepairVault();

            byte[] bytes = File.ReadAllBytes(withBom);
            Assert.Equal(new byte[] { 0xEF, 0xBB, 0xBF }, new[] { bytes[0], bytes[1], bytes[2] });
            Assert.Equal("[[Inbox/Re_ Budget-Robert Smith Jr/Re_ Budget-Robert Smith Jr/file.pdf]]",
                Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3));
            Assert.Equal("[[Inbox/Re_ Budget-Robert Smith Jr./Re_ Budget-Robert Smith Jr./file.pdf]]", File.ReadAllText(config));
            Assert.Empty(Directory.GetFiles(_vault, "*.tmp.*", SearchOption.AllDirectories));
        }

        [Fact]
        public void Resolves_ComparesFolderNamesExactly()
        {
            // Windows opens "Jr.\file.pdf" as "Jr\file.pdf"; the check must not be fooled by that.
            DottedFolderLinkRepairService service = new DottedFolderLinkRepairService(_vault);
            string inbox = Path.Combine(_vault, "Inbox");

            Assert.True(service.Resolves(inbox, "Re_ Budget-Robert Smith Jr/Re_ Budget-Robert Smith Jr/file.pdf"));
            Assert.False(service.Resolves(inbox, "Re_ Budget-Robert Smith Jr./Re_ Budget-Robert Smith Jr/file.pdf"));
            Assert.True(service.Resolves(inbox, "/Inbox/Re_ Budget-Robert Smith Jr"));
            Assert.True(service.Resolves(Path.Combine(_vault, "Notes"), "../Inbox/Re_ Budget-Robert Smith Jr/Re_ Budget-Robert Smith Jr/file.pdf"));
            Assert.False(service.Resolves(inbox, "Re_ Budget-Robert Smith Jr/nothing.pdf"));
        }

        [Fact]
        public void Resolves_NoteLinkWithoutExtension_FindsMarkdownFile()
        {
            WriteNote(Path.Combine("Inbox", "Re_ Budget-Robert Smith Jr", "0-Thread.md"), "x");
            DottedFolderLinkRepairService service = new DottedFolderLinkRepairService(_vault);

            Assert.True(service.Resolves(_vault, "Inbox/Re_ Budget-Robert Smith Jr/0-Thread"));
        }

        [Fact]
        public void RepairVault_MissingVault_ReportsNothing()
        {
            DottedFolderLinkRepairReport report =
                new DottedFolderLinkRepairService(Path.Combine(_vault, "nope")).RepairVault();
            Assert.Equal(0, report.NotesScanned);
        }
    }
}
