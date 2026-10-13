using System;
using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class BodyCleanerServiceTests
    {
        private static BodyCleanerService ServiceWith(params string[] patterns)
        {
            ObsidianSettings settings = new ObsidianSettings { BodyCleanupPatterns = new List<string>(patterns) };
            return new BodyCleanerService(settings);
        }

        [Fact]
        public void BodyCleanupPatterns_DefaultsToEmpty()
        {
            Assert.Empty(new ObsidianSettings().BodyCleanupPatterns);
        }

        [Fact]
        public void Clean_NoPatterns_ReturnsBodyUnchangedAndNotFlagged()
        {
            string body = "Hello\r\n\r\n\r\nWorld  ";
            BodyCleanerService svc = new BodyCleanerService(new ObsidianSettings());

            string cleaned = svc.Clean(body, out bool removed);

            Assert.Same(body, cleaned);
            Assert.False(removed);
        }

        [Fact]
        public void Clean_PatternWithNoMatch_ReturnsBodyUnchangedAndNotFlagged()
        {
            string body = "Line one\r\nLine two";
            BodyCleanerService svc = ServiceWith(@"^CAUTION:.*$");

            string cleaned = svc.Clean(body, out bool removed);

            Assert.Equal(body, cleaned);
            Assert.False(removed);
        }

        [Fact]
        public void Clean_LineAnchoredPattern_RemovesOnlyThatLine()
        {
            string body = "CAUTION: This email originated from outside the organization.\n\nPlease see the attached order.\n\nThanks,\nPat";
            BodyCleanerService svc = ServiceWith(@"^CAUTION:.*$");

            string cleaned = svc.Clean(body, out bool removed);

            Assert.True(removed);
            Assert.Equal("Please see the attached order.\n\nThanks,\nPat", cleaned);
        }

        [Fact]
        public void Clean_IsCaseInsensitive()
        {
            BodyCleanerService svc = ServiceWith(@"\[external\]\s*");

            string cleaned = svc.Clean("[EXTERNAL] Hearing moved to Friday.", out bool removed);

            Assert.True(removed);
            Assert.Equal("Hearing moved to Friday.", cleaned);
        }

        [Fact]
        public void Clean_SinglelineInlineFlag_RemovesMultiLineDisclaimer()
        {
            string body = "Draft attached.\n\nCONFIDENTIALITY NOTICE: This message\nmay contain privileged material.\nEND NOTICE\n\nRegards";
            BodyCleanerService svc = ServiceWith(@"(?s)CONFIDENTIALITY NOTICE:.*?END NOTICE");

            string cleaned = svc.Clean(body, out bool removed);

            Assert.True(removed);
            Assert.Equal("Draft attached.\n\nRegards", cleaned);
        }

        [Fact]
        public void Clean_WithoutSinglelineFlag_DotDoesNotCrossLines()
        {
            string body = "Keep this\nCONFIDENTIAL footer line\nKeep this too";
            BodyCleanerService svc = ServiceWith(@"CONFIDENTIAL.*");

            string cleaned = svc.Clean(body, out bool removed);

            Assert.True(removed);
            Assert.Equal("Keep this\n\nKeep this too", cleaned);
        }

        [Fact]
        public void Clean_CrlfBody_LineAnchorsMatchAndNoStrayCarriageReturnRemains()
        {
            string body = "Real content\r\n[Sent from my phone]\r\nMore content";
            BodyCleanerService svc = ServiceWith(@"^\[Sent from my phone\]$");

            string cleaned = svc.Clean(body, out bool removed);

            Assert.True(removed);
            Assert.DoesNotContain("\r", cleaned);
            Assert.Equal("Real content\n\nMore content", cleaned);
        }

        [Fact]
        public void Clean_CollapsesBlankLinesLeftByRemoval_PreservesParagraphs()
        {
            string body = "First paragraph.\n\nBANNER\n\nSecond paragraph.\n- item one\n- item two";
            BodyCleanerService svc = ServiceWith(@"^BANNER$");

            string cleaned = svc.Clean(body, out bool removed);

            Assert.True(removed);
            Assert.Equal("First paragraph.\n\nSecond paragraph.\n- item one\n- item two", cleaned);
        }

        [Fact]
        public void Clean_PatternsApplyInOrder()
        {
            BodyCleanerService svc = ServiceWith(@"foo bar", @"bar");

            string cleaned = svc.Clean("foo bar baz", out bool removed);

            // Reversed, "bar" would run first and leave "foo  baz" for "foo bar" to miss.
            Assert.True(removed);
            Assert.Equal("baz", cleaned);
        }

        [Fact]
        public void Clean_ZeroLengthMatchOnly_IsNotFlagged()
        {
            string body = "No leading whitespace here";
            BodyCleanerService svc = ServiceWith(@"^\s*");

            string cleaned = svc.Clean(body, out bool removed);

            Assert.False(removed);
            Assert.Equal(body, cleaned);
        }

        [Fact]
        public void Clean_InvalidRegexPattern_IsSkippedAndValidPatternStillRuns()
        {
            BodyCleanerService svc = ServiceWith("[unbalanced", @"^Sent from my iPhone$");

            string cleaned = svc.Clean("See you then.\nSent from my iPhone", out bool removed);

            Assert.True(removed);
            Assert.Equal("See you then.", cleaned);
        }

        [Fact]
        public void Clean_BlankAndNullPatternEntries_AreSkipped()
        {
            ObsidianSettings settings = new ObsidianSettings { BodyCleanupPatterns = new List<string> { null, "   ", string.Empty } };
            BodyCleanerService svc = new BodyCleanerService(settings);

            string cleaned = svc.Clean("Body text", out bool removed);

            Assert.False(removed);
            Assert.Equal("Body text", cleaned);
        }

        [Fact]
        public void Clean_NullPatternsList_ReturnsBodyUnchanged()
        {
            ObsidianSettings settings = new ObsidianSettings { BodyCleanupPatterns = null };
            BodyCleanerService svc = new BodyCleanerService(settings);

            string cleaned = svc.Clean("Body text", out bool removed);

            Assert.False(removed);
            Assert.Equal("Body text", cleaned);
        }

        [Fact]
        public void Clean_NullBody_ReturnsEmptyAndNotFlagged()
        {
            BodyCleanerService svc = ServiceWith(@"anything");

            string cleaned = svc.Clean(null, out bool removed);

            Assert.False(removed);
            Assert.Equal(string.Empty, cleaned);
        }

        [Fact]
        public void Constructor_NullSettings_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => new BodyCleanerService(null));
        }
    }
}
