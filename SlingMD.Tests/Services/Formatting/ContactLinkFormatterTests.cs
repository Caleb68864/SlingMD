using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class ContactLinkFormatterTests
    {
        private readonly ContactLinkFormatter _formatter = new ContactLinkFormatter();

        private static ContactName JohnSmith()
        {
            return new ContactName
            {
                FirstName = "John",
                LastName = "Smith",
                FullName = "John Smith",
                DisplayName = "John Smith",
                ShortName = "John",
                Email = "john@example.com"
            };
        }

        [Fact]
        public void Format_WikilinkDefault_RendersFullNameInBrackets()
        {
            Assert.Equal("[[John Smith]]", _formatter.Format(JohnSmith(), "[[{FullName}]]"));
        }

        [Fact]
        public void Format_AtMentionFirstLast_RendersAtJohnSmith()
        {
            Assert.Equal("@JohnSmith", _formatter.Format(JohnSmith(), "@{FirstName}{LastName}"));
        }

        [Fact]
        public void Format_Initials_UpperCasesSingleChars()
        {
            Assert.Equal("@JS", _formatter.Format(JohnSmith(), "@{FirstInitial}{LastInitial}"));
        }

        [Fact]
        public void Format_TokenIsCaseInsensitive()
        {
            Assert.Equal("[[John Smith]]", _formatter.Format(JohnSmith(), "[[{fullname}]]"));
        }

        [Fact]
        public void Format_UnknownToken_RendersAsEmpty()
        {
            // {Address} isn't a known token — it renders empty, but FullName is still present
            Assert.Equal("[[John Smith / ]]", _formatter.Format(JohnSmith(), "[[{FullName} / {Address}]]"));
        }

        [Fact]
        public void Format_NullName_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _formatter.Format(null, "[[{FullName}]]"));
        }

        [Fact]
        public void Format_EmptyFormat_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _formatter.Format(JohnSmith(), string.Empty));
        }

        [Fact]
        public void Format_OnlyUnknownTokens_FallsBackToDisplayName()
        {
            ContactName name = new ContactName { DisplayName = "Fallback Display", FullName = "Should Not Be Used" };
            Assert.Equal("Fallback Display", _formatter.Format(name, "{Unknown1}{Unknown2}"));
        }

        [Fact]
        public void Format_OnlyUnknownTokensAndNoDisplayName_FallsBackToFullName()
        {
            ContactName name = new ContactName { FullName = "Just FullName" };
            Assert.Equal("Just FullName", _formatter.Format(name, "{Unknown}"));
        }

        [Fact]
        public void Format_EmptyTokenValueDoesNotCountAsRendered_FallsBackOnly()
        {
            // FirstName empty, LastName empty: counts as no rendered tokens, fallback to DisplayName.
            ContactName name = new ContactName { DisplayName = "Display Only" };
            Assert.Equal("Display Only", _formatter.Format(name, "{FirstName}{LastName}"));
        }

        [Fact]
        public void FormatList_JoinsWithSeparator()
        {
            List<ContactName> names = new List<ContactName>
            {
                JohnSmith(),
                new ContactName { FullName = "Jane Doe" }
            };
            Assert.Equal("[[John Smith]], [[Jane Doe]]", _formatter.FormatList(names, "[[{FullName}]]", ", "));
        }

        [Fact]
        public void FormatList_SkipsNullEntries()
        {
            List<ContactName> names = new List<ContactName> { JohnSmith(), null, new ContactName { FullName = "Jane Doe" } };
            Assert.Equal("[[John Smith]], [[Jane Doe]]", _formatter.FormatList(names, "[[{FullName}]]", ", "));
        }

        [Fact]
        public void FormatList_NullList_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _formatter.FormatList(null, "[[{FullName}]]", ", "));
        }

        [Fact]
        public void GetUnknownTokens_ReturnsOnlyUnknown()
        {
            List<string> unknown = _formatter.GetUnknownTokens("[[{FullName}]] {Mystery} {Email} {Address}");
            Assert.Contains("Mystery", unknown);
            Assert.Contains("Address", unknown);
            Assert.DoesNotContain("FullName", unknown);
            Assert.DoesNotContain("Email", unknown);
        }

        [Fact]
        public void GetUnknownTokens_EmptyFormat_ReturnsEmpty()
        {
            Assert.Empty(_formatter.GetUnknownTokens(string.Empty));
        }
    }
}
