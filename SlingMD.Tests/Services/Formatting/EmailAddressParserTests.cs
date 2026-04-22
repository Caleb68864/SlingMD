using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class EmailAddressParserTests
    {
        private readonly EmailAddressParser _p = new EmailAddressParser();

        [Fact]
        public void Domain_Normal_ReturnsAfterAt()
        {
            Assert.Equal("example.com", _p.Domain("jane@example.com"));
        }

        [Fact]
        public void Domain_Subdomain_ReturnsFullTail()
        {
            Assert.Equal("mail.corp.example.com", _p.Domain("jane@mail.corp.example.com"));
        }

        [Fact]
        public void Domain_Null_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _p.Domain(null));
        }

        [Fact]
        public void Domain_EmptyString_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _p.Domain(string.Empty));
        }

        [Fact]
        public void Domain_NoAtSign_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _p.Domain("justaword"));
        }

        [Fact]
        public void Domain_TrailingAtSign_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _p.Domain("jane@"));
        }

        [Fact]
        public void Domain_MultipleAtSigns_ReturnsTailFromFirst()
        {
            // MIME allows weird addresses; the first '@' wins. Domain is everything after.
            Assert.Equal("sub@example.com", _p.Domain("jane@sub@example.com"));
        }

        [Fact]
        public void LocalPart_Normal_ReturnsBeforeAt()
        {
            Assert.Equal("jane", _p.LocalPart("jane@example.com"));
        }

        [Fact]
        public void LocalPart_DottedLocal_IsPreserved()
        {
            Assert.Equal("jane.doe", _p.LocalPart("jane.doe@example.com"));
        }

        [Fact]
        public void LocalPart_Null_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _p.LocalPart(null));
        }

        [Fact]
        public void LocalPart_Whitespace_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _p.LocalPart("   "));
        }

        [Fact]
        public void LocalPart_NoAtSign_ReturnsWholeInput()
        {
            Assert.Equal("justaword", _p.LocalPart("justaword"));
        }

        [Fact]
        public void LocalPart_AtSignAtStart_ReturnsWholeInput()
        {
            // Leading '@' means no local-part — contract is to return the whole input.
            Assert.Equal("@example.com", _p.LocalPart("@example.com"));
        }
    }
}
