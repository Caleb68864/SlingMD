using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class TemplatePathResolverTests
    {
        private readonly TemplatePathResolver _resolver = new TemplatePathResolver();

        private static ObsidianSettings DefaultSettings(string vault = "C:\\Vault", string templatesFolder = "Templates")
        {
            return new ObsidianSettings
            {
                VaultBasePath = vault,
                VaultName = "MyVault",
                TemplatesFolder = templatesFolder
            };
        }

        [Fact]
        public void Resolve_VaultTemplatesFirst()
        {
            ObsidianSettings settings = DefaultSettings();
            List<string> result = _resolver.Resolve("EmailTemplate.md", settings, new[] { "C:\\Install" });
            Assert.NotEmpty(result);
            Assert.Equal(Path.Combine(settings.GetTemplatesPath(), "EmailTemplate.md"), result[0]);
        }

        [Fact]
        public void Resolve_RelativeTemplatesFolder_AddsBaseDirsWithFolderName()
        {
            ObsidianSettings settings = DefaultSettings(templatesFolder: "MyTemplates");
            List<string> result = _resolver.Resolve("X.md", settings, new[] { "C:\\Install", "C:\\App" });

            Assert.Contains(Path.Combine("C:\\Install", "MyTemplates", "X.md"), result);
            Assert.Contains(Path.Combine("C:\\App", "MyTemplates", "X.md"), result);
        }

        [Fact]
        public void Resolve_AlwaysIncludesLiteralTemplatesFallback()
        {
            ObsidianSettings settings = DefaultSettings(templatesFolder: "CustomFolder");
            List<string> result = _resolver.Resolve("X.md", settings, new[] { "C:\\Install" });

            Assert.Contains(Path.Combine("C:\\Install", "Templates", "X.md"), result);
        }

        [Fact]
        public void Resolve_AbsoluteTemplatesFolder_AppearsOnce_AndLiteralTemplatesFallbackPresent()
        {
            // When TemplatesFolder is rooted, GetTemplatesPath() returns it directly (so the
            // absolute path appears exactly once). The resolver must NOT additionally combine
            // base dirs with the rooted folder. The literal "Templates" fallback still fires.
            ObsidianSettings settings = DefaultSettings(templatesFolder: "C:\\AbsoluteTemplates");
            List<string> result = _resolver.Resolve("X.md", settings, new[] { "C:\\Install", "C:\\App" });

            string absoluteCandidate = Path.Combine("C:\\AbsoluteTemplates", "X.md");
            int hits = 0;
            foreach (string p in result)
            {
                if (string.Equals(p, absoluteCandidate, System.StringComparison.OrdinalIgnoreCase))
                {
                    hits++;
                }
            }
            Assert.Equal(1, hits);
            Assert.Contains(Path.Combine("C:\\Install", "Templates", "X.md"), result);
            Assert.Contains(Path.Combine("C:\\App", "Templates", "X.md"), result);
        }

        [Fact]
        public void Resolve_DuplicateBaseDirs_DedupedCaseInsensitive()
        {
            ObsidianSettings settings = DefaultSettings();
            List<string> result = _resolver.Resolve("X.md", settings, new[] { "C:\\Install", "c:\\install", "C:\\Install" });

            int templatesPathHits = 0;
            string expected = Path.Combine("C:\\Install", "Templates", "X.md");
            foreach (string p in result)
            {
                if (string.Equals(p, expected, System.StringComparison.OrdinalIgnoreCase))
                {
                    templatesPathHits++;
                }
            }
            Assert.Equal(1, templatesPathHits);
        }

        [Fact]
        public void Resolve_NullBaseDirs_StillReturnsVaultTemplatesPath()
        {
            ObsidianSettings settings = DefaultSettings();
            List<string> result = _resolver.Resolve("X.md", settings, null);
            Assert.Single(result);
            Assert.Equal(Path.Combine(settings.GetTemplatesPath(), "X.md"), result[0]);
        }

        [Fact]
        public void Resolve_NullSettings_OnlyReturnsBaseDirFallbacks()
        {
            List<string> result = _resolver.Resolve("X.md", null, new[] { "C:\\Install" });
            Assert.Single(result);
            Assert.Equal(Path.Combine("C:\\Install", "Templates", "X.md"), result[0]);
        }

        [Fact]
        public void Resolve_BlankBaseDirEntry_Ignored()
        {
            ObsidianSettings settings = DefaultSettings();
            List<string> result = _resolver.Resolve("X.md", settings, new[] { "", null, "C:\\Install" });

            Assert.Contains(Path.Combine("C:\\Install", "Templates", "X.md"), result);
            // No path containing an empty base dir should appear.
            foreach (string p in result)
            {
                Assert.False(p.StartsWith(Path.Combine("", "Templates")));
            }
        }

        [Fact]
        public void Resolve_OrderingMatchesSearchPriority()
        {
            ObsidianSettings settings = DefaultSettings(templatesFolder: "T");
            List<string> result = _resolver.Resolve("X.md", settings, new[] { "C:\\A", "C:\\B" });

            // Vault path first, then "C:\\A\\T", "C:\\B\\T", "C:\\A\\Templates", "C:\\B\\Templates".
            Assert.Equal(5, result.Count);
            Assert.Equal(Path.Combine(settings.GetTemplatesPath(), "X.md"), result[0]);
            Assert.Equal(Path.Combine("C:\\A", "T", "X.md"), result[1]);
            Assert.Equal(Path.Combine("C:\\B", "T", "X.md"), result[2]);
            Assert.Equal(Path.Combine("C:\\A", "Templates", "X.md"), result[3]);
            Assert.Equal(Path.Combine("C:\\B", "Templates", "X.md"), result[4]);
        }
    }
}
