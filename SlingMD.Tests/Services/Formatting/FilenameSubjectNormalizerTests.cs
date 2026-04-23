using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class FilenameSubjectNormalizerTests
    {
        private readonly FilenameSubjectNormalizer _norm = new FilenameSubjectNormalizer();

        [Fact]
        public void Normalize_NullInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _norm.Normalize(null));
        }

        [Fact]
        public void Normalize_EmptyInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _norm.Normalize(string.Empty));
        }

        [Fact]
        public void Normalize_ColonSpace_BecomesUnderscore()
        {
            Assert.Equal("Re_foo", _norm.Normalize("Re: foo"));
        }

        [Fact]
        public void Normalize_ColonNoSpace_BecomesUnderscore()
        {
            Assert.Equal("Re_foo", _norm.Normalize("Re:foo"));
        }

        [Fact]
        public void Normalize_RepeatedReUnderscore_CollapsedToSingle()
        {
            Assert.Equal("Re_topic", _norm.Normalize("Re_Re_Re_topic"));
        }

        [Fact]
        public void Normalize_MixedCaseRePrefixes_CollapsedToReUnderscore()
        {
            Assert.Equal("Re_topic", _norm.Normalize("Re_RE_Re_topic"));
        }

        [Fact]
        public void Normalize_RepeatedFwUnderscore_CollapsedToSingle()
        {
            Assert.Equal("Fw_topic", _norm.Normalize("Fw_Fw_Fw_topic"));
        }

        [Fact]
        public void Normalize_MixedCaseFwPrefixes_CollapsedToFwUnderscore()
        {
            Assert.Equal("Fw_topic", _norm.Normalize("Fw_FW_Fw_topic"));
        }

        [Fact]
        public void Normalize_DropsTrailingSpaceAfterReUnderscore()
        {
            Assert.Equal("Re_topic", _norm.Normalize("Re_   topic"));
        }

        [Fact]
        public void Normalize_DropsTrailingSpaceAfterFwUnderscore()
        {
            Assert.Equal("Fw_topic", _norm.Normalize("Fw_   topic"));
        }

        [Fact]
        public void Normalize_PlainSubject_Unchanged()
        {
            Assert.Equal("Quarterly review", _norm.Normalize("Quarterly review"));
        }

        [Fact]
        public void Normalize_ReColonSpace_MultiplePrefixes_BecomesSingleReUnderscore()
        {
            // "Re: Re: Re: foo" → after colon-space pass: "Re_Re_Re_foo" → after collapse: "Re_foo"
            Assert.Equal("Re_foo", _norm.Normalize("Re: Re: Re: foo"));
        }

        [Fact]
        public void DefaultConstructor_UsesBuiltInRules_WhenSettingsAbsent()
        {
            FilenameSubjectNormalizer n = new FilenameSubjectNormalizer();
            // BuiltInDefaults must produce the canonical "Re_foo" output.
            Assert.Equal("Re_foo", n.Normalize("Re: Re: Re: foo"));
        }

        [Fact]
        public void Normalize_NullSettings_FallsBackToBuiltInDefaults()
        {
            FilenameSubjectNormalizer n = new FilenameSubjectNormalizer(null);
            Assert.Equal("Re_foo", n.Normalize("Re: foo"));
        }

        [Fact]
        public void Normalize_EmptySettingsList_FallsBackToBuiltInDefaults()
        {
            ObsidianSettings s = new ObsidianSettings { FilenameSubjectPatterns = new List<FilenameSubjectRule>() };
            FilenameSubjectNormalizer n = new FilenameSubjectNormalizer(s);
            Assert.Equal("Re_foo", n.Normalize("Re: foo"));
        }

        [Fact]
        public void Normalize_CustomSettingsRules_OverrideDefaults()
        {
            ObsidianSettings s = new ObsidianSettings
            {
                FilenameSubjectPatterns = new List<FilenameSubjectRule>
                {
                    new FilenameSubjectRule { Pattern = @"foo", Replacement = "bar" }
                }
            };
            FilenameSubjectNormalizer n = new FilenameSubjectNormalizer(s);
            // Custom rule applied; defaults NOT applied (no Re_ collapse).
            Assert.Equal("Re: bar", n.Normalize("Re: foo"));
        }

        [Fact]
        public void Normalize_InvalidUserPattern_IsSkippedSilently()
        {
            ObsidianSettings s = new ObsidianSettings
            {
                FilenameSubjectPatterns = new List<FilenameSubjectRule>
                {
                    new FilenameSubjectRule { Pattern = "[unbalanced", Replacement = "x" },
                    new FilenameSubjectRule { Pattern = @"foo", Replacement = "bar" }
                }
            };
            FilenameSubjectNormalizer n = new FilenameSubjectNormalizer(s);
            Assert.Equal("Re: bar", n.Normalize("Re: foo"));
        }

        [Fact]
        public void BuiltInDefaults_MatchObsidianSettingsDefaults()
        {
            // Pin: the ObsidianSettings default factory must produce the same rule set as
            // FilenameSubjectNormalizer.BuiltInDefaults (single source of truth).
            List<FilenameSubjectRule> settingsDefaults = ObsidianSettings.CreateDefaultFilenameSubjectPatterns();
            Assert.Equal(FilenameSubjectNormalizer.BuiltInDefaults.Count, settingsDefaults.Count);
            for (int i = 0; i < settingsDefaults.Count; i++)
            {
                Assert.Equal(FilenameSubjectNormalizer.BuiltInDefaults[i].Pattern, settingsDefaults[i].Pattern);
                Assert.Equal(FilenameSubjectNormalizer.BuiltInDefaults[i].Replacement, settingsDefaults[i].Replacement);
            }
        }
    }
}
