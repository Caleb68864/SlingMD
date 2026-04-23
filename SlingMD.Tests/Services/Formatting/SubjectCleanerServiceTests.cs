using System;
using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class SubjectCleanerServiceTests
    {
        private static ObsidianSettings DefaultSettings()
        {
            return new ObsidianSettings();
        }

        [Fact]
        public void Clean_StripsLeadingRePrefix()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            Assert.Equal("Status update", svc.Clean("Re: Status update"));
        }

        [Fact]
        public void Clean_StripsRepeatedRePrefixes()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            Assert.Equal("Status update", svc.Clean("Re: Re: RE: Status update"));
        }

        [Fact]
        public void Clean_StripsFwdPrefix()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            Assert.Equal("Heads up", svc.Clean("Fwd: Heads up"));
        }

        [Fact]
        public void Clean_PreservesPreReleaseInsideWords()
        {
            // The bug fix that motivated SS-02: "pre-release" must not be corrupted.
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            string cleaned = svc.Clean("pre-release notes");
            Assert.Contains("pre-release", cleaned);
        }

        [Fact]
        public void Clean_StripsExternalTag()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            Assert.Equal("Quarterly review", svc.Clean("[EXTERNAL] Quarterly review"));
        }

        [Fact]
        public void Clean_NormalizesMultipleSpacesToSingle_WhenNoCustomPatterns()
        {
            // Use an empty pattern set so only the trailing whitespace-normalization step runs.
            ObsidianSettings settings = DefaultSettings();
            settings.SubjectCleanupPatterns = new List<string>();
            SubjectCleanerService svc = new SubjectCleanerService(settings);
            Assert.Equal("a b c", svc.Clean("a   b\tc"));
        }

        [Fact]
        public void Clean_NullInput_ReturnsEmpty()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            Assert.Equal(string.Empty, svc.Clean(null));
        }

        [Fact]
        public void Clean_EmptyInput_ReturnsEmpty()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            Assert.Equal(string.Empty, svc.Clean(string.Empty));
        }

        [Fact]
        public void Clean_InvalidRegexPattern_IsSkippedSilently()
        {
            ObsidianSettings settings = DefaultSettings();
            settings.SubjectCleanupPatterns = new List<string> { "[unbalanced", "^Re:\\s*" };
            SubjectCleanerService svc = new SubjectCleanerService(settings);
            // Invalid pattern is skipped; valid Re: stripper still runs.
            Assert.Equal("Status update", svc.Clean("Re: Status update"));
        }

        [Fact]
        public void Clean_NullPatternsList_ReturnsTrimmedSubject()
        {
            ObsidianSettings settings = DefaultSettings();
            settings.SubjectCleanupPatterns = null;
            SubjectCleanerService svc = new SubjectCleanerService(settings);
            Assert.Equal("Hello", svc.Clean("  Hello  "));
        }

        [Fact]
        public void Constructor_NullSettings_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => new SubjectCleanerService(null));
        }

        [Fact]
        public void NormalizeForGrouping_StripsLeadingRePrefix()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            Assert.Equal("Status update", svc.NormalizeForGrouping("Re: Status update"));
        }

        [Fact]
        public void NormalizeForGrouping_StripsLeadingExternalTag()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            Assert.Equal("Status update", svc.NormalizeForGrouping("[EXTERNAL] Status update"));
        }

        [Fact]
        public void NormalizeForGrouping_StripsExternalThenRe()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            Assert.Equal("Status update", svc.NormalizeForGrouping("[EXTERNAL] Re: Status update"));
        }

        [Fact]
        public void NormalizeForGrouping_PreservesInWordRe()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            string normalized = svc.NormalizeForGrouping("pre-release notes");
            Assert.Contains("pre-release", normalized);
        }

        [Fact]
        public void NormalizeForGrouping_NullInput_ReturnsEmpty()
        {
            SubjectCleanerService svc = new SubjectCleanerService(DefaultSettings());
            Assert.Equal(string.Empty, svc.NormalizeForGrouping(null));
        }
    }
}
