using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class ContactNameParserTests
    {
        private readonly ContactNameParser _parser = new ContactNameParser();

        [Fact]
        public void Parse_FirstLast_PopulatesNamesAndFullName()
        {
            ContactName r = _parser.Parse("John Smith");
            Assert.Equal("John", r.FirstName);
            Assert.Equal("Smith", r.LastName);
            Assert.Equal("John Smith", r.FullName);
            Assert.Equal("John Smith", r.DisplayName);
            Assert.Equal("John", r.ShortName);
            Assert.Equal(string.Empty, r.MiddleName);
            Assert.Equal(string.Empty, r.Suffix);
        }

        [Fact]
        public void Parse_LastCommaFirst_PopulatesPartsAndRebuildsFullName()
        {
            ContactName r = _parser.Parse("Smith, John");
            Assert.Equal("John", r.FirstName);
            Assert.Equal("Smith", r.LastName);
            Assert.Equal("John Smith", r.FullName);
            Assert.Equal("Smith, John", r.DisplayName);
        }

        [Fact]
        public void Parse_FirstMiddleLast_AssignsMiddleName()
        {
            ContactName r = _parser.Parse("John Quincy Adams");
            Assert.Equal("John", r.FirstName);
            Assert.Equal("Adams", r.LastName);
            Assert.Equal("Quincy", r.MiddleName);
            Assert.Equal("John Adams", r.FullName);
        }

        [Fact]
        public void Parse_FirstMiddleLastSuffix_DetectsSuffix()
        {
            ContactName r = _parser.Parse("John Quincy Adams Jr.");
            Assert.Equal("John", r.FirstName);
            Assert.Equal("Adams", r.LastName);
            Assert.Equal("Quincy", r.MiddleName);
            Assert.Equal("Jr.", r.Suffix);
        }

        [Fact]
        public void Parse_LastCommaFirstSuffix_DetectsSuffix()
        {
            ContactName r = _parser.Parse("Adams, John Q. Jr.");
            Assert.Equal("John", r.FirstName);
            Assert.Equal("Adams", r.LastName);
            Assert.Equal("Q.", r.MiddleName);
            Assert.Equal("Jr.", r.Suffix);
        }

        [Fact]
        public void Parse_SingleName_UsesAsFirstNameOnly()
        {
            ContactName r = _parser.Parse("Madonna");
            Assert.Equal("Madonna", r.FirstName);
            Assert.Equal("Madonna", r.FullName);
            Assert.Equal(string.Empty, r.LastName);
            Assert.Equal("Madonna", r.ShortName);
        }

        [Fact]
        public void Parse_FirstLastWithSuffix_DetectsSuffix()
        {
            ContactName r = _parser.Parse("John Smith III");
            Assert.Equal("John", r.FirstName);
            Assert.Equal("Smith", r.LastName);
            Assert.Equal("III", r.Suffix);
        }

        [Fact]
        public void Parse_EmptyDisplayName_FallsBackToEmailLocalPart()
        {
            ContactName r = _parser.Parse(string.Empty, "jane.doe@example.com");
            Assert.Equal("jane.doe", r.FirstName);
            Assert.Equal("jane.doe", r.FullName);
            Assert.Equal("jane.doe", r.DisplayName);
            Assert.Equal("jane.doe@example.com", r.Email);
        }

        [Fact]
        public void Parse_NullDisplayNameAndNullEmail_ReturnsEmptyStrings()
        {
            ContactName r = _parser.Parse(null, null);
            Assert.Equal(string.Empty, r.FirstName);
            Assert.Equal(string.Empty, r.LastName);
            Assert.Equal(string.Empty, r.FullName);
            Assert.Equal(string.Empty, r.Email);
        }

        [Fact]
        public void Parse_StoresEmailWhenProvided()
        {
            ContactName r = _parser.Parse("John Smith", "john@example.com");
            Assert.Equal("john@example.com", r.Email);
        }

        [Fact]
        public void Parse_TrimsWhitespace()
        {
            ContactName r = _parser.Parse("   John   Smith   ");
            Assert.Equal("John", r.FirstName);
            Assert.Equal("Smith", r.LastName);
        }

        [Fact]
        public void Parse_EmailWithoutAtSign_TreatedAsLocalPart()
        {
            ContactName r = _parser.Parse(null, "weirdaddress");
            Assert.Equal("weirdaddress", r.FirstName);
        }
    }
}
