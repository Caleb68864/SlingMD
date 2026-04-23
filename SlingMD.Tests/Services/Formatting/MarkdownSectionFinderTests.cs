using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class MarkdownSectionFinderTests
    {
        private readonly MarkdownSectionFinder _finder = new MarkdownSectionFinder();

        [Fact]
        public void FindSectionStart_HeadingPresent_ReturnsIndex()
        {
            string content = "Intro line\n## Notes\nSome note.\n";
            int idx = _finder.FindSectionStart(content, "## Notes");
            Assert.Equal(content.IndexOf("## Notes"), idx);
        }

        [Fact]
        public void FindSectionStart_HeadingAbsent_ReturnsNegativeOne()
        {
            string content = "Intro line\nBody only\n";
            Assert.Equal(-1, _finder.FindSectionStart(content, "## Notes"));
        }

        [Fact]
        public void FindSectionStart_EmptyContent_ReturnsNegativeOne()
        {
            Assert.Equal(-1, _finder.FindSectionStart(string.Empty, "## Notes"));
        }

        [Fact]
        public void FindSectionStart_NullContent_ReturnsNegativeOne()
        {
            Assert.Equal(-1, _finder.FindSectionStart(null, "## Notes"));
        }

        [Fact]
        public void FindSectionStart_EmptyHeading_ReturnsNegativeOne()
        {
            Assert.Equal(-1, _finder.FindSectionStart("## Notes\nx", string.Empty));
        }

        [Fact]
        public void FindSectionStart_HeadingMustBeOwnLine_NotEmbeddedInText()
        {
            string content = "See ## Notes below\n## Notes\nreal section.\n";
            int idx = _finder.FindSectionStart(content, "## Notes");
            // The first occurrence is embedded in text; the real match is the one
            // on its own line.
            Assert.NotEqual(content.IndexOf("## Notes"), idx);
            Assert.True(idx > content.IndexOf("## Notes"));
        }

        [Fact]
        public void FindSectionStart_HonorsStartIndex_SkipsEarlierMatches()
        {
            string content = "## Notes\nfirst\n## Notes\nsecond\n";
            int firstIdx = _finder.FindSectionStart(content, "## Notes");
            int secondIdx = _finder.FindSectionStart(content, "## Notes", firstIdx + 1);
            Assert.NotEqual(firstIdx, secondIdx);
            Assert.True(secondIdx > firstIdx);
        }

        [Fact]
        public void FindSectionStart_StartIndexBeyondContent_ReturnsNegativeOne()
        {
            Assert.Equal(-1, _finder.FindSectionStart("## Notes\n", "## Notes", 999));
        }

        [Fact]
        public void FindSectionStart_NegativeStartIndex_ReturnsNegativeOne()
        {
            Assert.Equal(-1, _finder.FindSectionStart("## Notes\n", "## Notes", -1));
        }

        [Fact]
        public void FindSectionStart_HeadingWithRegexMetacharacters_IsEscaped()
        {
            // Brackets in a heading shouldn't be interpreted as regex char classes.
            string content = "body\n## [Archived]\ncontent\n";
            int idx = _finder.FindSectionStart(content, "## [Archived]");
            Assert.Equal(content.IndexOf("## [Archived]"), idx);
        }

        [Fact]
        public void FindSectionStart_HeadingWithTrailingSpaces_Matches()
        {
            // The anchor "\\s*$" allows trailing whitespace after the heading text.
            string content = "body\n## Notes   \ncontent\n";
            int idx = _finder.FindSectionStart(content, "## Notes");
            Assert.Equal(content.IndexOf("## Notes"), idx);
        }
    }
}
