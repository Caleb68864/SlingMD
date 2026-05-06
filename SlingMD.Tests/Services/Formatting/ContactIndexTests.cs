using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class ContactIndexTests
    {
        // ── helpers ──────────────────────────────────────────────────────────────

        private static string CreateTempDir() =>
            Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), Path.GetRandomFileName())).FullName;

        private static void WriteNote(string dir, string filename, string frontmatter = null)
        {
            string content = frontmatter == null
                ? $"# {Path.GetFileNameWithoutExtension(filename)}\n"
                : frontmatter;
            File.WriteAllText(Path.Combine(dir, filename), content);
        }

        private static string BlockFrontmatter(params string[] aliases)
        {
            string aliasLines = string.Join("\n", System.Array.ConvertAll(aliases, a => $"  - {a}"));
            return $"---\naliases:\n{aliasLines}\n---\n# body\n";
        }

        // ── structural ───────────────────────────────────────────────────────────

        [Fact]
        public void ContactIndexEntry_ExposesRequiredProperties()
        {
            List<string> aliases = new List<string> { "Lisa" };
            List<string> normAliases = new List<string> { "lisa" };

            ContactIndexEntry entry = new ContactIndexEntry(
                "/path/file.md",
                "Lisa Angle",
                aliases.AsReadOnly(),
                "lisa angle",
                normAliases.AsReadOnly());

            Assert.Equal("/path/file.md", entry.FilePath);
            Assert.Equal("Lisa Angle", entry.DisplayName);
            Assert.Single(entry.Aliases);
            Assert.Equal("lisa angle", entry.NormalizedDisplayName);
            Assert.Single(entry.NormalizedAliases);
        }

        // ── tier 1: contacts folder ───────────────────────────────────────────

        [Fact]
        public void BuildContactsFolderTier_LookupByDisplayName_ReturnsEntry()
        {
            string dir = CreateTempDir();
            try
            {
                WriteNote(dir, "Lisa Angle.md");

                ContactIndex index = new ContactIndex();
                index.BuildContactsFolderTier(dir);

                IReadOnlyList<ContactIndexEntry> results = index.Lookup("lisa angle");
                Assert.Single(results);
                Assert.Equal("Lisa Angle", results[0].DisplayName);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        [Fact]
        public void BuildContactsFolderTier_LookupByAlias_ReturnsEntry()
        {
            string dir = CreateTempDir();
            try
            {
                WriteNote(dir, "Lisa Angle.md", BlockFrontmatter("L. Angle", "LA"));

                ContactIndex index = new ContactIndex();
                index.BuildContactsFolderTier(dir);

                // alias "L. Angle" normalizes to "l angle" (dot stripped, initial removed by normalizer)
                // Try the raw alias key: lookup by "la"
                IReadOnlyList<ContactIndexEntry> results = index.Lookup("la");
                Assert.Single(results);
                Assert.Equal("Lisa Angle", results[0].DisplayName);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        [Fact]
        public void BuildContactsFolderTier_Over5000Files_ReturnsEmptyLookup()
        {
            string dir = CreateTempDir();
            try
            {
                for (int i = 0; i <= 5000; i++)
                    File.WriteAllText(Path.Combine(dir, $"contact_{i:D5}.md"), "# body\n");

                ContactIndex index = new ContactIndex();
                index.BuildContactsFolderTier(dir);

                IReadOnlyList<ContactIndexEntry> results = index.Lookup("contact 00001");
                Assert.Empty(results);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        // ── tier 2: vault ─────────────────────────────────────────────────────

        [Fact]
        public void BuildVaultTier_LookupByDisplayName_ReturnsEntry()
        {
            string dir = CreateTempDir();
            try
            {
                string sub = Directory.CreateDirectory(Path.Combine(dir, "contacts")).FullName;
                WriteNote(sub, "John Smith.md");

                ContactIndex index = new ContactIndex();
                index.BuildVaultTier(dir);

                IReadOnlyList<ContactIndexEntry> results = index.Lookup("john smith");
                Assert.Single(results);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        [Fact]
        public void BuildVaultTier_Over5000Files_ReturnsEmptyLookup()
        {
            string dir = CreateTempDir();
            try
            {
                for (int i = 0; i <= 5000; i++)
                    File.WriteAllText(Path.Combine(dir, $"note_{i:D5}.md"), "# body\n");

                ContactIndex index = new ContactIndex();
                index.BuildVaultTier(dir);

                IReadOnlyList<ContactIndexEntry> results = index.Lookup("note 00001");
                Assert.Empty(results);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        // ── frontmatter: max 30 lines ─────────────────────────────────────────

        [Fact]
        public void BuildContactsFolderTier_GarbagePastLine30_DoesNotThrow()
        {
            string dir = CreateTempDir();
            try
            {
                // Build a file with valid frontmatter in first 5 lines then garbage after line 30
                System.Text.StringBuilder sb = new System.Text.StringBuilder();
                sb.AppendLine("---");
                sb.AppendLine("aliases:");
                sb.AppendLine("  - GarbageAlias");
                // Fill lines 4..30 inside frontmatter block (never close with ---)
                for (int i = 4; i <= 30; i++)
                    sb.AppendLine($"field_{i}: value");
                // Past line 30: malformed garbage
                sb.AppendLine("{{{{{{ NOT YAML {{{{{{");
                sb.AppendLine("---");   // closing dash would be ignored; reader stopped at line 30
                sb.AppendLine("# body content");

                File.WriteAllText(Path.Combine(dir, "GarbageNote.md"), sb.ToString());

                ContactIndex index = new ContactIndex();
                // Should not throw even with garbage past line 30
                System.Exception caught = Record.Exception(() => index.BuildContactsFolderTier(dir));
                Assert.Null(caught);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        // ── multi-match: both display name and alias across different notes ───

        [Fact]
        public void Lookup_KeyMatchesBothDisplayNameAndAliasOnDifferentNotes_ReturnsBothEntries()
        {
            string dir = CreateTempDir();
            try
            {
                // "Alex" is the display name of one note
                WriteNote(dir, "Alex.md");
                // "Alex" is an alias of another note
                WriteNote(dir, "Alexander Smith.md", BlockFrontmatter("Alex"));

                ContactIndex index = new ContactIndex();
                index.BuildContactsFolderTier(dir);

                IReadOnlyList<ContactIndexEntry> results = index.Lookup("alex");
                Assert.Equal(2, results.Count);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        // ── edge cases ────────────────────────────────────────────────────────

        [Fact]
        public void Lookup_EmptyKey_ReturnsEmpty()
        {
            ContactIndex index = new ContactIndex();
            Assert.Empty(index.Lookup(string.Empty));
            Assert.Empty(index.Lookup(null));
        }

        [Fact]
        public void Lookup_UnknownKey_ReturnsEmpty()
        {
            string dir = CreateTempDir();
            try
            {
                WriteNote(dir, "John Smith.md");
                ContactIndex index = new ContactIndex();
                index.BuildContactsFolderTier(dir);

                Assert.Empty(index.Lookup("nobody"));
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        [Fact]
        public void BuildContactsFolderTier_MissingDirectory_DoesNotThrow()
        {
            ContactIndex index = new ContactIndex();
            System.Exception ex = Record.Exception(
                () => index.BuildContactsFolderTier(@"C:\does\not\exist\xyz"));
            Assert.Null(ex);
        }

        [Fact]
        public void BuildVaultTier_MissingDirectory_DoesNotThrow()
        {
            ContactIndex index = new ContactIndex();
            System.Exception ex = Record.Exception(
                () => index.BuildVaultTier(@"C:\does\not\exist\xyz"));
            Assert.Null(ex);
        }
    }
}
