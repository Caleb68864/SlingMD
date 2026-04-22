using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class FrontmatterReaderTests
    {
        private readonly FrontmatterReader _reader = new FrontmatterReader();

        private const string SampleFrontmatter =
            "---\n" +
            "title: \"Quarterly review - 2026-04-21\"\n" +
            "type: \"email\"\n" +
            "from: \"[[John Smith]]\"\n" +
            "fromEmail: \"john@example.com\"\n" +
            "to:\n" +
            "  - \"[[Jane Doe]]\"\n" +
            "  - \"[[Bob Brown]]\"\n" +
            "threadId: \"AF6350AAEA21DDC93136\"\n" +
            "date: \"2026-04-21 09:30:00\"\n" +
            "---\n" +
            "\n" +
            "Body text here.\n";

        [Fact]
        public void ExtractThreadId_ReturnsValue()
        {
            Assert.Equal("AF6350AAEA21DDC93136", _reader.ExtractThreadId(SampleFrontmatter));
        }

        [Fact]
        public void ExtractRawDate_ReturnsFullTimestamp()
        {
            Assert.Equal("2026-04-21 09:30:00", _reader.ExtractRawDate(SampleFrontmatter));
        }

        [Fact]
        public void ExtractRawDate_AcceptsLegacyMinutePrecision()
        {
            string legacy = "date: 2026-04-21 09:30\n";
            Assert.Equal("2026-04-21 09:30", _reader.ExtractRawDate(legacy));
        }

        [Fact]
        public void ExtractTitle_ReturnsValue()
        {
            Assert.Equal("Quarterly review - 2026-04-21", _reader.ExtractTitle(SampleFrontmatter));
        }

        [Fact]
        public void ExtractFromName_ReturnsWikilinkTarget()
        {
            Assert.Equal("John Smith", _reader.ExtractFromName(SampleFrontmatter));
        }

        [Fact]
        public void ExtractFirstToName_ReturnsFirstWikilinkInToList()
        {
            Assert.Equal("Jane Doe", _reader.ExtractFirstToName(SampleFrontmatter));
        }

        [Fact]
        public void ExtractMissingField_ReturnsNull()
        {
            string noThreadId = "title: \"x\"\nfrom: \"[[A]]\"\n";
            Assert.Null(_reader.ExtractThreadId(noThreadId));
        }

        [Fact]
        public void ExtractFromEmptyContent_ReturnsNull()
        {
            Assert.Null(_reader.ExtractThreadId(string.Empty));
            Assert.Null(_reader.ExtractTitle(null));
        }

        [Fact]
        public void ExtractFromName_HandlesEmailAddressBeforeWikilink()
        {
            // Real-world from-line has "fromEmail|[[Name]]" pattern.
            string content = "from: \"john@example.com|[[John Smith]]\"\n";
            Assert.Equal("John Smith", _reader.ExtractFromName(content));
        }

        [Fact]
        public void ExtractFirstToName_ReturnsOnlyFirst_WhenMultipleListed()
        {
            // The reader's contract is "first" — should not return all entries.
            string content = "to:\n  - \"[[A]]\"\n  - \"[[B]]\"\n  - \"[[C]]\"\n";
            Assert.Equal("A", _reader.ExtractFirstToName(content));
        }
    }
}
