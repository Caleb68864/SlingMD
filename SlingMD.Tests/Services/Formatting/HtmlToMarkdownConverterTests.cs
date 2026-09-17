using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class HtmlToMarkdownConverterTests
    {
        [Fact]
        public void Convert_NullOrEmptyHtml_ReturnsPlainTextFallback()
        {
            Assert.Equal("Plain text", HtmlToMarkdownConverter.Convert(null, "Plain text"));
            Assert.Equal("Plain text", HtmlToMarkdownConverter.Convert(string.Empty, "Plain text"));
            Assert.Equal(string.Empty, HtmlToMarkdownConverter.Convert("   ", null));
        }

        [Fact]
        public void Convert_UsesBodyContentAndRemovesNonVisibleElements()
        {
            string html = "<html><head><style>.x { display:none; }</style></head>"
                + "<body><p>Hello <span>world</span>.</p><script>alert('bad');</script></body></html>";

            Assert.Equal("Hello world.", HtmlToMarkdownConverter.Convert(html, "fallback"));
        }

        [Fact]
        public void Convert_ConvertsHeadingsParagraphsAndLineBreaks()
        {
            string html = "<h2>Weekly update</h2><p>First line<br>second line</p><div>Next block</div>";

            Assert.Equal("## Weekly update\n\nFirst line\nsecond line\n\nNext block", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_ConvertsLinksWithQuotedAndUnquotedAttributes()
        {
            string html = "<p><a href=\"https://example.com?a=1&amp;b=2\">Read <b>this</b></a> "
                + "<a href='https://example.org'>Other</a> <a href=https://example.net>Third</a></p>";

            Assert.Equal(
                "[Read this](https://example.com?a=1&b=2) [Other](https://example.org) [Third](https://example.net)",
                HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_ConvertsImagesAndSkipsTrackingPixels()
        {
            string html = "<p>Before</p>"
                + "<img src=\"https://cdn.example/image.png\" alt=\"Logo\">"
                + "<img src=\"https://cdn.example/p.gif\" width=\"1\" height=\"1\">"
                + "<img src=\"https://cdn.example/other.png\" style=\"width: 1px\">"
                + "<p>After</p>";

            Assert.Equal("Before\n\n![Logo](https://cdn.example/image.png)\n\nAfter", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_ConvertsListsBlockquotesAndEmphasis()
        {
            string html = "<ul><li><strong>One</strong></li><li><em>Two</em></li></ul>"
                + "<blockquote>Quoted <strong>text</strong><br>next line</blockquote><hr>End";

            Assert.Equal("- **One**\n- *Two*\n\n> Quoted **text**\n> next line\n\n---\n\nEnd", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_DecodesEntitiesAndNonBreakingSpaces()
        {
            string html = "<p>Tom &amp; Jerry&nbsp;&mdash;  hello&nbsp;there</p>";

            Assert.Equal("Tom & Jerry\u00A0— hello there".Replace('\u00A0', ' '), HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_MalformedMarkup_RetainsReadableTextAndDoesNotThrow()
        {
            string html = "<div>Start <a href=\"https://example.com\">link<div>middle"
                + "<img src=\"https://cdn.example/photo.jpg\" alt=\"photo\">end";

            string result = HtmlToMarkdownConverter.Convert(html, "fallback");

            Assert.Contains("Start", result);
            Assert.Contains("link", result);
            Assert.Contains("middle", result);
            Assert.Contains("photo", result);
            Assert.DoesNotContain("<div>", result);
        }

        [Fact]
        public void Convert_EmptyVisibleHtml_FallsBackToPlainText()
        {
            string html = "<!-- comment --><head><title>Hidden</title></head><body>"
                + "<img src=\"https://cdn.example/open?pixel=1\" width=\"1\"></body>";

            Assert.Equal("Plain fallback", HtmlToMarkdownConverter.Convert(html, "Plain fallback"));
        }

        [Fact]
        public void Convert_LeavesPlainTextFallbackNullAsEmpty()
        {
            Assert.Equal(string.Empty, HtmlToMarkdownConverter.Convert("<p> </p>", null));
        }
    }
}
