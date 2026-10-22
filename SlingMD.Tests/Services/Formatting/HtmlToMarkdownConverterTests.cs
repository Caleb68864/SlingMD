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

        [Fact]
        public void Convert_DataTable_BecomesPipeTable()
        {
            string html = "<table><tr><th>Name</th><th>Value</th></tr><tr><td>Rate</td><td>5%</td></tr></table>";

            Assert.Equal("| Name | Value |\n| --- | --- |\n| Rate | 5% |", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_PipeTableCell_EscapesPipes()
        {
            string html = "<table><tr><td>A</td><td>B</td></tr><tr><td>x | y</td><td>2</td></tr></table>";

            Assert.Equal("| A | B |\n| --- | --- |\n| x \\| y | 2 |", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_LayoutTable_KeepsCellsApart()
        {
            string html = "<table><tr><td>John Smith</td><td>Branch Manager</td></tr><tr><td>555-1234</td></tr></table>";

            Assert.Equal("John Smith Branch Manager\n555-1234", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_LayoutTableWithParagraphs_KeepsParagraphs()
        {
            string html = "<table><tr><td><p>Hi there</p><p>Thanks</p></td></tr></table>";

            Assert.Equal("Hi there\n\nThanks", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_NestedTables_ConvertsInnerDataTable()
        {
            string html = "<table><tr><td><p>Outer</p><table><tr><td>A</td><td>B</td></tr><tr><td>1</td><td>2</td></tr></table></td></tr></table>";

            Assert.Equal("Outer\n\n| A | B |\n| --- | --- |\n| 1 | 2 |", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_SpacerCells_AreIgnored()
        {
            string html = "<table><tr><td>&nbsp;</td><td>Logo text</td></tr><tr><td>&nbsp;</td><td>&nbsp;</td></tr></table>";

            Assert.Equal("Logo text", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_SafeLink_UnwrapsToOriginalAddress()
        {
            string html = "<a href=\"https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fexample.com%2Freport%3Fid%3D1&amp;data=05%7C01&amp;reserved=0\">Report</a>";

            Assert.Equal("[Report](https://example.com/report?id=1)", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_SafeLinkShownAsText_BecomesBareAddress()
        {
            string wrapped = "https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fexample.com%2Fa&amp;data=1";
            string html = "<a href=\"" + wrapped + "\">" + wrapped + "</a>";

            Assert.Equal("https://example.com/a", HtmlToMarkdownConverter.Convert(html, null));
        }

        [Fact]
        public void Convert_LinkTextSameAsAddress_BecomesBareAddress()
        {
            Assert.Equal("https://example.com", HtmlToMarkdownConverter.Convert("<a href=\"https://example.com\">https://example.com</a>", null));
        }

        [Fact]
        public void UnwrapSafeLink_OrdinaryUrl_Unchanged()
        {
            Assert.Equal("https://example.com/?url=x", HtmlToMarkdownConverter.UnwrapSafeLink("https://example.com/?url=x"));
        }
    }
}
