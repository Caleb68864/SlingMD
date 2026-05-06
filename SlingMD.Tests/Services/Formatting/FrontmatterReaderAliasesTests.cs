using System.Collections.Generic;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class FrontmatterReaderAliasesTests
    {
        private readonly FrontmatterReader _reader = new FrontmatterReader();

        [Fact]
        public void ExtractAliases_BlockStyle_ReturnsList()
        {
            string content =
                "---\n" +
                "aliases:\n" +
                "  - Lisa\n" +
                "  - L. Angle\n" +
                "---\n";

            IReadOnlyList<string> result = _reader.ExtractAliases(content);

            Assert.Equal(2, result.Count);
            Assert.Equal("Lisa", result[0]);
            Assert.Equal("L. Angle", result[1]);
        }

        [Fact]
        public void ExtractAliases_InlineArray_ReturnsList()
        {
            string content =
                "---\n" +
                "aliases: [Lisa, L. Angle]\n" +
                "---\n";

            IReadOnlyList<string> result = _reader.ExtractAliases(content);

            Assert.Equal(2, result.Count);
            Assert.Equal("Lisa", result[0]);
            Assert.Equal("L. Angle", result[1]);
        }

        [Fact]
        public void ExtractAliases_NoAliasesKey_ReturnsEmpty()
        {
            string content =
                "---\n" +
                "title: \"My Note\"\n" +
                "---\n";

            IReadOnlyList<string> result = _reader.ExtractAliases(content);

            Assert.Empty(result);
        }

        [Fact]
        public void ExtractAliases_NoLeadingTripleDash_ReturnsEmpty()
        {
            string content =
                "title: \"My Note\"\n" +
                "aliases:\n" +
                "  - Lisa\n";

            IReadOnlyList<string> result = _reader.ExtractAliases(content);

            Assert.Empty(result);
        }

        [Fact]
        public void ExtractAliases_MalformedYaml_ReturnsEmptyAndDoesNotThrow()
        {
            string content =
                "---\n" +
                "aliases: :::invalid::: {{{ yaml\n" +
                "---\n";

            IReadOnlyList<string> result = _reader.ExtractAliases(content);

            Assert.Empty(result);
        }

        [Fact]
        public void ExtractAliases_NullContent_ReturnsEmpty()
        {
            IReadOnlyList<string> result = _reader.ExtractAliases(null);

            Assert.Empty(result);
        }

        [Fact]
        public void ExtractAliases_EmptyContent_ReturnsEmpty()
        {
            IReadOnlyList<string> result = _reader.ExtractAliases(string.Empty);

            Assert.Empty(result);
        }
    }
}
