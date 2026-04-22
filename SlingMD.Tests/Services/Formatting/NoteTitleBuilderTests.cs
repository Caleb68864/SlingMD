using System.Collections.Generic;
using Xunit;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Tests.Services.Formatting
{
    public class NoteTitleBuilderTests
    {
        private readonly NoteTitleBuilder _builder;

        public NoteTitleBuilderTests()
        {
            _builder = new NoteTitleBuilder();
        }

        [Fact]
        public void Build_WithBasicTokens_SubstitutesCorrectly()
        {
            // Arrange
            string format = "{Subject} - {Date}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Hello" },
                { "Date", "2026-04-21" }
            };

            // Act
            string result = _builder.Build(format, tokens, 50);

            // Assert
            Assert.Equal("Hello - 2026-04-21", result);
        }

        [Fact]
        public void Build_ExceedsMaxLength_TruncatesWithEllipsis()
        {
            // Arrange
            string format = "{Subject}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "This is a very long subject line that definitely exceeds the maximum length allowed" }
            };
            int maxLen = 20;

            // Act
            string result = _builder.Build(format, tokens, maxLen);

            // Assert
            Assert.True(result.Length <= maxLen, $"Expected length <= {maxLen}, but got {result.Length}");
            Assert.EndsWith(NoteTitleBuilder.Ellipsis, result);
        }

        [Fact]
        public void Build_ExactlyAtMaxLength_NoTruncation()
        {
            // Arrange
            string format = "{Subject}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Short" }
            };
            int maxLen = 5;

            // Act
            string result = _builder.Build(format, tokens, maxLen);

            // Assert
            Assert.Equal("Short", result);
            Assert.Equal(5, result.Length);
        }

        [Fact]
        public void Build_WithMissingToken_RendersAsEmpty()
        {
            // Arrange
            string format = "{Subject} - {MissingToken}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Hello" }
            };

            // Act
            string result = _builder.Build(format, tokens, 50);

            // Assert
            Assert.Equal("Hello -", result);
        }

        [Fact]
        public void Build_WithNullTokenValue_RendersAsEmpty()
        {
            // Arrange
            string format = "{Subject} by {Author}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Hello" },
                { "Author", null }
            };

            // Act
            string result = _builder.Build(format, tokens, 50);

            // Assert
            Assert.Equal("Hello by", result);
        }

        [Fact]
        public void Build_WithEmptyFormat_ReturnsEmpty()
        {
            // Arrange
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Hello" }
            };

            // Act
            string result = _builder.Build(string.Empty, tokens, 50);

            // Assert
            Assert.Equal(string.Empty, result);
        }

        [Fact]
        public void Build_WithNullFormat_ReturnsEmpty()
        {
            // Arrange
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Hello" }
            };

            // Act
            string result = _builder.Build(null, tokens, 50);

            // Assert
            Assert.Equal(string.Empty, result);
        }

        [Fact]
        public void Build_WithNullTokens_LeavesUnmatchedTokensEmpty()
        {
            // Arrange
            string format = "{Subject} - {Date}";

            // Act
            string result = _builder.Build(format, null, 50);

            // Assert
            Assert.Equal("-", result);
        }

        [Fact]
        public void Build_WithZeroMaxLength_ReturnsEmpty()
        {
            // Arrange
            string format = "{Subject}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Hello" }
            };

            // Act
            string result = _builder.Build(format, tokens, 0);

            // Assert
            Assert.Equal(string.Empty, result);
        }

        [Fact]
        public void Build_WithNegativeMaxLength_ReturnsEmpty()
        {
            // Arrange
            string format = "{Subject}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Hello" }
            };

            // Act
            string result = _builder.Build(format, tokens, -5);

            // Assert
            Assert.Equal(string.Empty, result);
        }

        [Fact]
        public void Build_VerySmallMaxLength_TruncatesEllipsis()
        {
            // Arrange
            string format = "{Subject}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Hello World" }
            };
            int maxLen = 2;

            // Act
            string result = _builder.Build(format, tokens, maxLen);

            // Assert
            Assert.True(result.Length <= maxLen);
        }

        [Fact]
        public void Build_MultipleTokens_SubstitutesAll()
        {
            // Arrange
            string format = "{Sender} sent '{Subject}' on {Date}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Sender", "John" },
                { "Subject", "Report" },
                { "Date", "2026-04-21" }
            };

            // Act
            string result = _builder.Build(format, tokens, 100);

            // Assert
            Assert.Equal("John sent 'Report' on 2026-04-21", result);
        }

        [Fact]
        public void Build_PreservesNonTokenBraces()
        {
            // Arrange
            // A brace pattern that doesn't look like a token (contains non-alphanumeric)
            string format = "{Subject} {not-a-token}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Hello" }
            };

            // Act
            string result = _builder.Build(format, tokens, 50);

            // Assert
            Assert.Equal("Hello {not-a-token}", result);
        }

        [Fact]
        public void Build_CaseSensitiveTokens()
        {
            // Arrange
            string format = "{Subject} vs {subject}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Upper" },
                { "subject", "lower" }
            };

            // Act
            string result = _builder.Build(format, tokens, 50);

            // Assert
            Assert.Equal("Upper vs lower", result);
        }

        [Fact]
        public void Build_TrimsWhitespace()
        {
            // Arrange
            string format = "  {Subject}  ";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "Subject", "Hello" }
            };

            // Act
            string result = _builder.Build(format, tokens, 50);

            // Assert
            Assert.Equal("Hello", result);
        }

        [Fact]
        public void Build_TruncationMaintainsMaxLength()
        {
            // Arrange
            string format = "{LongText}";
            Dictionary<string, string> tokens = new Dictionary<string, string>
            {
                { "LongText", "ABCDEFGHIJKLMNOPQRSTUVWXYZ" }
            };

            for (int maxLen = 5; maxLen <= 30; maxLen++)
            {
                // Act
                string result = _builder.Build(format, tokens, maxLen);

                // Assert
                Assert.True(result.Length <= maxLen, $"For maxLen={maxLen}, got length {result.Length}");
            }
        }

        [Fact]
        public void Ellipsis_IsThreeDots()
        {
            // Assert that the ellipsis constant matches the one used in EmailProcessor
            Assert.Equal("...", NoteTitleBuilder.Ellipsis);
        }
    }
}
