using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class ContactNameNormalizerTests
    {
        private readonly ContactNameNormalizer _normalizer = new ContactNameNormalizer();

        [Theory]
        [InlineData("Bob M Smith", "Bob Smith")]
        [InlineData("Bob M. Smith", "Bob Smith")]
        [InlineData("Alice J. Doe", "Alice Doe")]
        public void Normalize_MiddleInitial_IsStripped(string withInitial, string withoutInitial)
        {
            Assert.Equal(_normalizer.Normalize(withoutInitial), _normalizer.Normalize(withInitial));
        }

        [Theory]
        [InlineData("Bob Smith (Acme Corp)", "Bob Smith")]
        [InlineData("Bob M Smith (Acme)", "Bob Smith")]
        [InlineData("Jane Doe (Contractor)", "Jane Doe")]
        public void Normalize_ParenthesizedCompany_IsStripped(string withCompany, string withoutCompany)
        {
            Assert.Equal(_normalizer.Normalize(withoutCompany), _normalizer.Normalize(withCompany));
        }

        [Theory]
        [InlineData("Dr. Robert Smith")]
        [InlineData("Mr. John Doe")]
        [InlineData("Mrs. Jane Doe")]
        [InlineData("Prof. Alice Brown")]
        public void Normalize_Honorific_IsStripped(string nameWithHonorific)
        {
            string stripped = nameWithHonorific.Substring(nameWithHonorific.IndexOf(' ') + 1);
            Assert.Equal(_normalizer.Normalize(stripped), _normalizer.Normalize(nameWithHonorific));
        }

        [Fact]
        public void Normalize_TrailingCommaSuffix_MatchesInlineSuffix()
        {
            // "Dr. Robert Smith, Jr." should normalize same as "Robert Smith Jr"
            Assert.Equal(_normalizer.Normalize("Robert Smith Jr"), _normalizer.Normalize("Dr. Robert Smith, Jr."));
        }

        [Theory]
        [InlineData("BOB   smith ", "Bob Smith")]
        [InlineData("  ALICE   DOE  ", "Alice Doe")]
        public void Normalize_WhitespaceAndCase_AreCollapsed(string messy, string clean)
        {
            Assert.Equal(_normalizer.Normalize(clean), _normalizer.Normalize(messy));
        }

        [Theory]
        [InlineData("BOB SMITH", "bob smith")]
        [InlineData("Alice Doe", "ALICE DOE")]
        public void Normalize_CaseInsensitive(string a, string b)
        {
            Assert.Equal(_normalizer.Normalize(a), _normalizer.Normalize(b));
        }

        [Theory]
        [InlineData("")]
        [InlineData("   ")]
        public void Normalize_EmptyOrWhitespace_ReturnsEmpty(string input)
        {
            Assert.Equal(string.Empty, _normalizer.Normalize(input));
        }

        [Fact]
        public void Normalize_Initials_AreNotExpanded()
        {
            // B. Smith should NOT equal Bob Smith — initials are kept as-is
            Assert.NotEqual(_normalizer.Normalize("Bob Smith"), _normalizer.Normalize("B. Smith"));
        }

        [Theory]
        [InlineData("André Smith", "andré smith")]
        [InlineData("Søren Aaberg", "søren aaberg")]
        public void Normalize_Unicode_HandledSafely(string input, string expected)
        {
            Assert.Equal(expected, _normalizer.Normalize(input));
        }

        [Fact]
        public void NormalizeFirstLast_ExtractsFirstAndLast()
        {
            (string first, string last) = _normalizer.NormalizeFirstLast("Dr. Robert Smith, Jr.");
            Assert.Equal("robert", first);
            Assert.Equal("smith", last);
        }

        [Fact]
        public void NormalizeFirstLast_EmptyInput_ReturnsEmptyTuple()
        {
            (string first, string last) = _normalizer.NormalizeFirstLast("   ");
            Assert.Equal(string.Empty, first);
            Assert.Equal(string.Empty, last);
        }
    }
}
