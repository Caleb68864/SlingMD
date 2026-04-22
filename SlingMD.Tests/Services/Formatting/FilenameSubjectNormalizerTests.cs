using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class FilenameSubjectNormalizerTests
    {
        private readonly FilenameSubjectNormalizer _norm = new FilenameSubjectNormalizer();

        [Fact]
        public void Normalize_NullInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _norm.Normalize(null));
        }

        [Fact]
        public void Normalize_EmptyInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _norm.Normalize(string.Empty));
        }

        [Fact]
        public void Normalize_ColonSpace_BecomesUnderscore()
        {
            Assert.Equal("Re_foo", _norm.Normalize("Re: foo"));
        }

        [Fact]
        public void Normalize_ColonNoSpace_BecomesUnderscore()
        {
            Assert.Equal("Re_foo", _norm.Normalize("Re:foo"));
        }

        [Fact]
        public void Normalize_RepeatedReUnderscore_CollapsedToSingle()
        {
            Assert.Equal("Re_topic", _norm.Normalize("Re_Re_Re_topic"));
        }

        [Fact]
        public void Normalize_MixedCaseRePrefixes_CollapsedToReUnderscore()
        {
            Assert.Equal("Re_topic", _norm.Normalize("Re_RE_Re_topic"));
        }

        [Fact]
        public void Normalize_RepeatedFwUnderscore_CollapsedToSingle()
        {
            Assert.Equal("Fw_topic", _norm.Normalize("Fw_Fw_Fw_topic"));
        }

        [Fact]
        public void Normalize_MixedCaseFwPrefixes_CollapsedToFwUnderscore()
        {
            Assert.Equal("Fw_topic", _norm.Normalize("Fw_FW_Fw_topic"));
        }

        [Fact]
        public void Normalize_DropsTrailingSpaceAfterReUnderscore()
        {
            Assert.Equal("Re_topic", _norm.Normalize("Re_   topic"));
        }

        [Fact]
        public void Normalize_DropsTrailingSpaceAfterFwUnderscore()
        {
            Assert.Equal("Fw_topic", _norm.Normalize("Fw_   topic"));
        }

        [Fact]
        public void Normalize_PlainSubject_Unchanged()
        {
            Assert.Equal("Quarterly review", _norm.Normalize("Quarterly review"));
        }

        [Fact]
        public void Normalize_ReColonSpace_MultiplePrefixes_BecomesSingleReUnderscore()
        {
            // "Re: Re: Re: foo" → after colon-space pass: "Re_Re_Re_foo" → after collapse: "Re_foo"
            Assert.Equal("Re_foo", _norm.Normalize("Re: Re: Re: foo"));
        }
    }
}
