using System;
using System.IO;
using SlingMD.Outlook.Forms;
using Xunit;

namespace SlingMD.Tests.Forms
{
    public class SettingsFormVaultStatusTests : IDisposable
    {
        private readonly string _testDir;

        public SettingsFormVaultStatusTests()
        {
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "VaultStatus_" + Guid.NewGuid().ToString("N").Substring(0, 8));
            Directory.CreateDirectory(_testDir);
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(_testDir, true);
            }
            catch (System.Exception)
            {
                // Best-effort cleanup
            }
        }

        [Fact]
        public void DescribeVault_Blank_AsksToBrowse()
        {
            bool found;
            string status = SettingsForm.DescribeVault(string.Empty, "Notes", out found);

            Assert.False(found);
            Assert.Contains("Browse", status);
        }

        [Fact]
        public void DescribeVault_FolderWithObsidianConfig_IsFound()
        {
            Directory.CreateDirectory(Path.Combine(_testDir, "Notes", ".obsidian"));

            bool found;
            string status = SettingsForm.DescribeVault(_testDir, "Notes", out found);

            Assert.True(found);
            Assert.StartsWith("✓", status);
        }

        [Fact]
        public void DescribeVault_BasePathIsTheVault_ExplainsHowToFix()
        {
            string vault = Path.Combine(_testDir, "Notes");
            Directory.CreateDirectory(Path.Combine(vault, ".obsidian"));

            bool found;
            string status = SettingsForm.DescribeVault(vault, "Notes", out found);

            Assert.False(found);
            Assert.Contains("is itself the vault", status);
        }

        [Fact]
        public void DescribeVault_FolderWithoutObsidianConfig_Warns()
        {
            Directory.CreateDirectory(Path.Combine(_testDir, "Notes"));

            bool found;
            string status = SettingsForm.DescribeVault(_testDir, "Notes", out found);

            Assert.False(found);
            Assert.Contains("has no .obsidian folder", status);
        }

        [Fact]
        public void DescribeVault_MissingFolder_SaysNotFound()
        {
            bool found;
            string status = SettingsForm.DescribeVault(_testDir, "Missing", out found);

            Assert.False(found);
            Assert.Contains("Folder not found", status);
        }
    }
}
