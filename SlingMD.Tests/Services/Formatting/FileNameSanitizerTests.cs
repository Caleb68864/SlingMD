using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class FileNameSanitizerTests
    {
        private readonly FileNameSanitizer _sanitizer = new FileNameSanitizer();

        [Fact]
        public void Sanitize_NullInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _sanitizer.Sanitize(null));
        }

        [Fact]
        public void Sanitize_EmptyInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _sanitizer.Sanitize(string.Empty));
        }

        [Fact]
        public void Sanitize_PreservesPlainName()
        {
            Assert.Equal("Quarterly review", _sanitizer.Sanitize("Quarterly review"));
        }

        [Fact]
        public void Sanitize_ReplacesInvalidPathChars_WithUnderscore()
        {
            // Backslash and forward slash are invalid on Windows.
            Assert.Equal("a_b_c", _sanitizer.Sanitize("a/b\\c"));
        }

        [Fact]
        public void Sanitize_DoubleQuotesBecomeUnderscoresViaInvalidCharStrip()
        {
            // " is in Path.GetInvalidFileNameChars on Windows, so it's replaced with _
            // before the post-strip pass runs, leaving "hello _world".
            Assert.Equal("hello _world", _sanitizer.Sanitize("hello \"world\""));
        }

        [Fact]
        public void Sanitize_StripsSingleQuotes()
        {
            Assert.Equal("its broken", _sanitizer.Sanitize("it's broken"));
        }

        [Fact]
        public void Sanitize_StripsBackticks()
        {
            Assert.Equal("hello world", _sanitizer.Sanitize("hello `world`"));
        }

        [Fact]
        public void Sanitize_ReplacesColonWithUnderscore()
        {
            Assert.Equal("a_b", _sanitizer.Sanitize("a:b"));
        }

        [Fact]
        public void Sanitize_StripsSemicolon()
        {
            Assert.Equal("ab", _sanitizer.Sanitize("a;b"));
        }

        [Fact]
        public void Sanitize_StripsLeadingReUnderscorePrefix()
        {
            Assert.Equal("Status update", _sanitizer.Sanitize("Re_Status update"));
        }

        [Fact]
        public void Sanitize_StripsLeadingFwdUnderscorePrefix()
        {
            Assert.Equal("Heads up", _sanitizer.Sanitize("Fwd_Heads up"));
        }

        [Fact]
        public void Sanitize_StripsLeadingPrefix_CaseInsensitive()
        {
            Assert.Equal("Status update", _sanitizer.Sanitize("RE_Status update"));
        }

        [Fact]
        public void Sanitize_CollapsesRepeatedSeparators()
        {
            Assert.Equal("a-b-c", _sanitizer.Sanitize("a___b---c"));
        }

        [Fact]
        public void Sanitize_TrimsLeadingAndTrailingSeparators()
        {
            Assert.Equal("middle", _sanitizer.Sanitize("--middle__"));
        }

        [Fact]
        public void Sanitize_HandlesColonSpace_StripsPrefix_LeavesSpace()
        {
            // ":" is invalid on Windows → underscore. "Re: foo" → "Re_ foo" → prefix strip → " foo".
            // Leading whitespace is preserved (Trim only strips '-' and '_'); the orchestrator
            // does its own Trim() afterwards.
            Assert.Equal(" foo", _sanitizer.Sanitize("Re: foo"));
        }
    }
}
