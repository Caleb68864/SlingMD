using System.IO;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class ContactMatcherTests
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
        public void MatchTier_HasExpectedValues()
        {
            // [STRUCTURAL] MatchTier enum declares Exact, HighConfidence, Ambiguous, None
            Assert.Equal(0, (int)MatchTier.Exact);
            Assert.Equal(1, (int)MatchTier.HighConfidence);
            Assert.Equal(2, (int)MatchTier.Ambiguous);
            Assert.Equal(3, (int)MatchTier.None);
        }

        [Fact]
        public void MatchResult_ExposesTierAndCandidates()
        {
            // [STRUCTURAL] MatchResult exposes Tier and Candidates
            MatchResult result = new MatchResult(MatchTier.None, System.Array.Empty<ContactIndexEntry>());
            Assert.Equal(MatchTier.None, result.Tier);
            Assert.Empty(result.Candidates);
        }

        [Fact]
        public void ContactMatcher_ExposesMatchMethod()
        {
            // [STRUCTURAL] ContactMatcher exposes MatchResult Match(string displayName, string email)
            ContactMatcher matcher = new ContactMatcher();
            MatchResult result = matcher.Match("Any Name", "any@example.com");
            Assert.NotNull(result);
        }

        // ── exact match: normalized display name ──────────────────────────────

        [Fact]
        public void Match_NormalizedDisplayNameEqualsIndexedNormalizedName_ReturnsExact()
        {
            // [BEHAVIORAL] normalized displayName == indexed normalized name → Tier.Exact, one candidate
            string dir = CreateTempDir();
            try
            {
                WriteNote(dir, "Lisa Angle.md");

                ContactMatcher matcher = new ContactMatcher(contactsFolderPath: dir);
                MatchResult result = matcher.Match("Lisa Angle", "");

                Assert.Equal(MatchTier.Exact, result.Tier);
                Assert.Single(result.Candidates);
                Assert.Equal("Lisa Angle", result.Candidates[0].DisplayName);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        [Fact]
        public void Match_NormalizedDisplayNameWithNoise_ReturnsExact()
        {
            // normalize("Bob M Smith (Acme)") == "bob smith" == normalize("Bob Smith") → Exact
            string dir = CreateTempDir();
            try
            {
                WriteNote(dir, "Bob Smith.md");

                ContactMatcher matcher = new ContactMatcher(contactsFolderPath: dir);
                MatchResult result = matcher.Match("Bob M Smith (Acme)", "");

                Assert.Equal(MatchTier.Exact, result.Tier);
                Assert.Single(result.Candidates);
                Assert.Equal("Bob Smith", result.Candidates[0].DisplayName);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        // ── exact match: alias ────────────────────────────────────────────────

        [Fact]
        public void Match_NormalizedDisplayNameEqualsIndexedNormalizedAlias_ReturnsExact()
        {
            // [BEHAVIORAL] normalized displayName == indexed normalized alias → Tier.Exact, one candidate
            string dir = CreateTempDir();
            try
            {
                WriteNote(dir, "Lisa Angle.md", BlockFrontmatter("LA", "L. Angle"));

                ContactMatcher matcher = new ContactMatcher(contactsFolderPath: dir);
                // "LA" normalizes to "la"; match on that alias
                MatchResult result = matcher.Match("LA", "");

                Assert.Equal(MatchTier.Exact, result.Tier);
                Assert.Single(result.Candidates);
                Assert.Equal("Lisa Angle", result.Candidates[0].DisplayName);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        // ── high confidence: first+last match, display strings differ ─────────

        [Fact]
        public void Match_FirstLastMatchSingleEntry_ReturnsHighConfidence()
        {
            // [BEHAVIORAL] normalized first+last match single indexed entry but display strings differ
            // "Bob Smith Jr." normalize → "bob smith jr" (not in index)
            // NormalizeFirstLast → ("bob","smith"), compositeKey "bob smith" → finds "Bob Smith"
            string dir = CreateTempDir();
            try
            {
                WriteNote(dir, "Bob Smith.md");

                ContactMatcher matcher = new ContactMatcher(contactsFolderPath: dir);
                MatchResult result = matcher.Match("Bob Smith Jr.", "");

                Assert.Equal(MatchTier.HighConfidence, result.Tier);
                Assert.Single(result.Candidates);
                Assert.Equal("Bob Smith", result.Candidates[0].DisplayName);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        [Fact]
        public void Match_FirstLastMatchSingleEntryViaMiddleName_ReturnsHighConfidence()
        {
            // "Bob James Smith" normalize → "bob james smith" (not in index as "Bob Smith")
            // NormalizeFirstLast → ("bob","smith"), firstLastIndex["bob smith"] → finds "Bob Smith"
            string dir = CreateTempDir();
            try
            {
                WriteNote(dir, "Bob Smith.md");

                ContactMatcher matcher = new ContactMatcher(contactsFolderPath: dir);
                MatchResult result = matcher.Match("Bob James Smith", "");

                Assert.Equal(MatchTier.HighConfidence, result.Tier);
                Assert.Single(result.Candidates);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        // ── ambiguous ─────────────────────────────────────────────────────────

        [Fact]
        public void Match_TwoEntriesSameNormalizedFirstLast_ReturnsAmbiguous()
        {
            // [BEHAVIORAL] two or more indexed entries share same normalized first+last → Ambiguous
            // "Bob Smith Jr." and "Bob Smith Sr." both have NormalizeFirstLast → ("bob","smith")
            // Incoming "Bob James Smith" → NormalizeFirstLast → ("bob","smith") → two hits → Ambiguous
            string dir = CreateTempDir();
            try
            {
                WriteNote(dir, "Bob Smith Jr..md");
                WriteNote(dir, "Bob Smith Sr..md");

                ContactMatcher matcher = new ContactMatcher(contactsFolderPath: dir);
                MatchResult result = matcher.Match("Bob James Smith", "");

                Assert.Equal(MatchTier.Ambiguous, result.Tier);
                Assert.Equal(2, result.Candidates.Count);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        [Fact]
        public void Match_TwoEntriesSameExactNormalizedName_ReturnsAmbiguous()
        {
            // Two different files that normalize to the same full key → Ambiguous via exact path
            string dir = CreateTempDir();
            try
            {
                // Both "John Smith.md" and an alias "John Smith" on another file
                WriteNote(dir, "John Smith.md");
                WriteNote(dir, "Jonathan Smith.md", BlockFrontmatter("John Smith"));

                ContactMatcher matcher = new ContactMatcher(contactsFolderPath: dir);
                MatchResult result = matcher.Match("John Smith", "");

                Assert.Equal(MatchTier.Ambiguous, result.Tier);
                Assert.Equal(2, result.Candidates.Count);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        // ── none ──────────────────────────────────────────────────────────────

        [Fact]
        public void Match_NoCandidatesMatch_ReturnsNone()
        {
            // [BEHAVIORAL] no candidates match → Tier.None, empty list
            string dir = CreateTempDir();
            try
            {
                WriteNote(dir, "Alice Jones.md");

                ContactMatcher matcher = new ContactMatcher(contactsFolderPath: dir);
                MatchResult result = matcher.Match("Nobody Here", "");

                Assert.Equal(MatchTier.None, result.Tier);
                Assert.Empty(result.Candidates);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        [Fact]
        public void Match_EmptyDisplayName_ReturnsNone()
        {
            ContactMatcher matcher = new ContactMatcher();
            MatchResult result = matcher.Match("", "any@example.com");
            Assert.Equal(MatchTier.None, result.Tier);
            Assert.Empty(result.Candidates);
        }

        [Fact]
        public void Match_NullDisplayName_ReturnsNone()
        {
            ContactMatcher matcher = new ContactMatcher();
            MatchResult result = matcher.Match(null, "any@example.com");
            Assert.Equal(MatchTier.None, result.Tier);
            Assert.Empty(result.Candidates);
        }

        // ── vault tier ────────────────────────────────────────────────────────

        [Fact]
        public void Match_VaultPath_FindsEntryInSubdirectory()
        {
            string dir = CreateTempDir();
            try
            {
                string sub = Directory.CreateDirectory(Path.Combine(dir, "contacts")).FullName;
                WriteNote(sub, "Jane Doe.md");

                ContactMatcher matcher = new ContactMatcher(vaultPath: dir);
                MatchResult result = matcher.Match("Jane Doe", "");

                Assert.Equal(MatchTier.Exact, result.Tier);
                Assert.Single(result.Candidates);
            }
            finally
            {
                Directory.Delete(dir, recursive: true);
            }
        }

        // ── edge: missing directories ─────────────────────────────────────────

        [Fact]
        public void Constructor_MissingContactsFolder_DoesNotThrow()
        {
            System.Exception ex = Record.Exception(
                () => new ContactMatcher(contactsFolderPath: @"C:\does\not\exist\xyz"));
            Assert.Null(ex);
        }

        [Fact]
        public void Constructor_MissingVaultPath_DoesNotThrow()
        {
            System.Exception ex = Record.Exception(
                () => new ContactMatcher(vaultPath: @"C:\does\not\exist\xyz"));
            Assert.Null(ex);
        }
    }
}
