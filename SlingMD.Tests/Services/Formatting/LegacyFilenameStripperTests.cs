using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class LegacyFilenameStripperTests
    {
        private readonly LegacyFilenameStripper _s = new LegacyFilenameStripper();

        [Fact]
        public void Strip_NullInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _s.Strip(null));
        }

        [Fact]
        public void Strip_EmptyInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _s.Strip(string.Empty));
        }

        [Fact]
        public void Strip_PlainName_Unchanged()
        {
            Assert.Equal("Quarterly review", _s.Strip("Quarterly review"));
        }

        [Fact]
        public void Strip_TrailingEidSuffix_Removed()
        {
            Assert.Equal("foo", _s.Strip("foo-eidABC123def456"));
        }

        [Fact]
        public void Strip_TrailingThreeDigitSuffix_Removed()
        {
            Assert.Equal("foo", _s.Strip("foo-001"));
        }

        [Fact]
        public void Strip_TwoDigitSuffix_NotRemoved()
        {
            // Only 3-digit suffixes count.
            Assert.Equal("foo-99", _s.Strip("foo-99"));
        }

        [Fact]
        public void Strip_FourDigitSuffix_NotRemoved()
        {
            Assert.Equal("foo-1234", _s.Strip("foo-1234"));
        }

        [Fact]
        public void Strip_LeadingDateUnderscoreHHmmss_Removed()
        {
            Assert.Equal("Subject", _s.Strip("2026-04-21_143009_Subject"));
        }

        [Fact]
        public void Strip_LeadingDateDashHHmm_Removed()
        {
            Assert.Equal("Subject", _s.Strip("2026-04-21-1430_Subject"));
        }

        [Fact]
        public void Strip_LeadingDateUnderscoreHHmm_Removed()
        {
            Assert.Equal("Subject", _s.Strip("2026-04-21_1430_Subject"));
        }

        [Fact]
        public void Strip_AllThreeDecorations_StackedAndCleaned()
        {
            // Date prefix + plain body + 3-digit suffix
            Assert.Equal("Subject", _s.Strip("2026-04-21_143009_Subject-001"));
        }

        [Fact]
        public void Strip_DatePrefixPlusEidSuffix_BothRemoved()
        {
            Assert.Equal("Subject", _s.Strip("2026-04-21_143009_Subject-eidXYZ"));
        }

        [Fact]
        public void Strip_NoDatePrefix_LeavesNameUntouched()
        {
            // 2026-04-21 in the middle of the name should not be stripped — only leading.
            Assert.Equal("Notes from 2026-04-21 meeting", _s.Strip("Notes from 2026-04-21 meeting"));
        }
    }
}
