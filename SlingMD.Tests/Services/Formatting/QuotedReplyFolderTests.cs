using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class QuotedReplyFolderTests
    {
        [Fact]
        public void ConvertFolded_OutlookDesktopReply_FoldsHistory()
        {
            string html = "<html><body><p>Sounds good, thanks.</p>"
                + "<div><div style=\"border:none;border-top:solid #E1E1E1 1.0pt;padding:3.0pt 0in 0in 0in\">"
                + "<p><b>From:</b> Sara<br><b>Sent:</b> Monday</p></div></div><p>Can you send the report?</p></body></html>";

            Assert.Equal(
                "Sounds good, thanks.\n\n> [!quote]- Earlier messages\n> **From:** Sara\n> **Sent:** Monday\n>\n> Can you send the report?",
                QuotedReplyFolder.ConvertFolded(html, null));
        }

        [Fact]
        public void ConvertFolded_OutlookWebReply_DropsSeparatorLine()
        {
            string html = "<body><div>Yes.</div><div id=\"appendonsend\"></div><hr style=\"display:inline-block\">"
                + "<div id=\"divRplyFwdMsg\"><b>From:</b> Sara</div><div>Original</div></body>";

            Assert.Equal(
                "Yes.\n\n> [!quote]- Earlier messages\n> **From:** Sara\n>\n> Original",
                QuotedReplyFolder.ConvertFolded(html, null));
        }

        [Fact]
        public void ConvertFolded_NoHistory_SameAsPlainConversion()
        {
            Assert.Equal("Just a note", QuotedReplyFolder.ConvertFolded("<p>Just a note</p>", null));
        }

        [Fact]
        public void ConvertFolded_NothingAboveHistory_KeepsEverything()
        {
            string html = "<body><div id=\"divRplyFwdMsg\"><b>From:</b> Sara</div><div>Original</div></body>";

            Assert.Equal("**From:** Sara\n\nOriginal", QuotedReplyFolder.ConvertFolded(html, null));
        }

        [Fact]
        public void ConvertFolded_PlainTextOriginalMessage_FoldsHistory()
        {
            string text = "Thanks!\r\n\r\n-----Original Message-----\r\nFrom: Sara\r\nSent: Monday\r\n\r\nHi";

            Assert.Equal(
                "Thanks!\n\n> [!quote]- Earlier messages\n> -----Original Message-----\n> From: Sara\n> Sent: Monday\n>\n> Hi",
                QuotedReplyFolder.ConvertFolded(null, text));
        }

        [Fact]
        public void ConvertFolded_PlainTextOnWrote_FoldsHistory()
        {
            string text = "Ok\n\nOn Mon, Sep 1, 2026 at 9:00 AM Sara <s@x.com> wrote:\n> hi";

            Assert.Equal(
                "Ok\n\n> [!quote]- Earlier messages\n> On Mon, Sep 1, 2026 at 9:00 AM Sara <s@x.com> wrote:\n> > hi",
                QuotedReplyFolder.ConvertFolded(string.Empty, text));
        }
    }
}
