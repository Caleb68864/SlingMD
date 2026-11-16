using System;
using System.Collections.Generic;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class DottedFolderLinkRewriterTests
    {
        // What exists in the fake vault, by decoded link path. "Robert Smith Jr" is the folder Windows
        // created for "Robert Smith Jr."; nothing exists under a dotted name.
        private static readonly HashSet<string> Existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Re_ Budget-Robert Smith Jr/file.pdf",
            "Re_ Budget-Robert Smith Jr/image one.png",
            "Re_ Budget-Robert Smith Jr/Inc/report.pdf",
            "Re_ Budget-Robert Smith Jr/note",
            "Re_ Budget-Robert Smith Jr/a#b.pdf",
            "Inbox/Thread Jr/Note Jr/file.pdf",
            "Thread/file.pdf",
            "Acme Inc/file.pdf",
        };

        private static bool Resolves(string path)
        {
            return Existing.Contains(path);
        }

        private static string Rewrite(string content)
        {
            return DottedFolderLinkRewriter.Rewrite(content, Resolves);
        }

        [Fact]
        public void Wikilink_DottedFolder_IsRewritten()
        {
            Assert.Equal("See [[Re_ Budget-Robert Smith Jr/file.pdf]].",
                Rewrite("See [[Re_ Budget-Robert Smith Jr./file.pdf]]."));
        }

        [Fact]
        public void Embed_DottedFolder_IsRewritten()
        {
            Assert.Equal("![[Re_ Budget-Robert Smith Jr/image one.png]]",
                Rewrite("![[Re_ Budget-Robert Smith Jr./image one.png]]"));
        }

        [Fact]
        public void Wikilink_AliasAndHeading_AreKept()
        {
            Assert.Equal("[[Re_ Budget-Robert Smith Jr/file.pdf|the file. Jr./x]] [[Re_ Budget-Robert Smith Jr/note#Part Jr./2]]",
                Rewrite("[[Re_ Budget-Robert Smith Jr./file.pdf|the file. Jr./x]] [[Re_ Budget-Robert Smith Jr./note#Part Jr./2]]"));
        }

        [Fact]
        public void Wikilink_EscapedPipeInTable_KeepsBackslash()
        {
            Assert.Equal("| [[Re_ Budget-Robert Smith Jr/file.pdf\\|file]] |",
                Rewrite("| [[Re_ Budget-Robert Smith Jr./file.pdf\\|file]] |"));
        }

        [Fact]
        public void Wikilink_ToNoteWithoutExtension_IsRewritten()
        {
            Assert.Equal("[[Re_ Budget-Robert Smith Jr/note]]", Rewrite("[[Re_ Budget-Robert Smith Jr./note]]"));
        }

        [Fact]
        public void MarkdownLink_UrlEncoded_IsRewrittenInSameEncoding()
        {
            Assert.Equal("[file.pdf](Re_%20Budget-Robert%20Smith%20Jr/file.pdf)",
                Rewrite("[file.pdf](Re_%20Budget-Robert%20Smith%20Jr./file.pdf)"));
        }

        [Fact]
        public void MarkdownImage_UrlEncoded_IsRewritten()
        {
            Assert.Equal("![image one.png](Re_%20Budget-Robert%20Smith%20Jr/image%20one.png)",
                Rewrite("![image one.png](Re_%20Budget-Robert%20Smith%20Jr./image%20one.png)"));
        }

        [Fact]
        public void MarkdownLink_EncodedHash_IsRewritten()
        {
            Assert.Equal("[a#b.pdf](Re_%20Budget-Robert%20Smith%20Jr/a%23b.pdf)",
                Rewrite("[a#b.pdf](Re_%20Budget-Robert%20Smith%20Jr./a%23b.pdf)"));
        }

        [Fact]
        public void MarkdownLink_EncodedTrailingDotOrSpace_IsStripped()
        {
            Assert.Equal("[f](Re_%20Budget-Robert%20Smith%20Jr/file.pdf) [g](Thread/file.pdf)",
                Rewrite("[f](Re_%20Budget-Robert%20Smith%20Jr%2E/file.pdf) [g](Thread%20/file.pdf)"));
        }

        [Fact]
        public void MarkdownLink_Unescaped_AsTheAttachmentListWritesIt_IsRewritten()
        {
            Assert.Equal("- [file.pdf](Re_ Budget-Robert Smith Jr/file.pdf)",
                Rewrite("- [file.pdf](Re_ Budget-Robert Smith Jr./file.pdf)"));
        }

        [Fact]
        public void MarkdownLink_AngleBrackets_TitleAndFragment_AreKept()
        {
            Assert.Equal(
                "[a](<Re_ Budget-Robert Smith Jr/file.pdf>) [b](Re_%20Budget-Robert%20Smith%20Jr/file.pdf \"Jr./x\") [c](Re_%20Budget-Robert%20Smith%20Jr/note#Jr./y)",
                Rewrite(
                "[a](<Re_ Budget-Robert Smith Jr./file.pdf>) [b](Re_%20Budget-Robert%20Smith%20Jr./file.pdf \"Jr./x\") [c](Re_%20Budget-Robert%20Smith%20Jr./note#Jr./y)"));
        }

        [Fact]
        public void EveryDottedFolderSegment_IsRepaired()
        {
            Assert.Equal("[[Inbox/Thread Jr/Note Jr/file.pdf]] [x](Re_%20Budget-Robert%20Smith%20Jr/Inc/report.pdf)",
                Rewrite("[[Inbox/Thread Jr./Note Jr./file.pdf]] [x](Re_%20Budget-Robert%20Smith%20Jr./Inc./report.pdf)"));
        }

        [Fact]
        public void TrailingSpaceAndDots_AreAllStripped()
        {
            Assert.Equal("[[Thread/file.pdf]] [[Thread/file.pdf]]", Rewrite("[[Thread. /file.pdf]] [[Thread../file.pdf]]"));
        }

        [Fact]
        public void FileNameDot_IsNeverTouched()
        {
            const string content = "[[Thread/file.]] [[Acme Inc./]] [x](Thread/file.)";
            Assert.Equal(content, Rewrite(content));
        }

        [Fact]
        public void LinkThatAlreadyResolves_IsLeftAlone()
        {
            // A vault synced from a Mac or Linux may really hold a folder named "Acme Inc.".
            HashSet<string> both = new HashSet<string>(StringComparer.Ordinal) { "Acme Inc./file.pdf", "Acme Inc/file.pdf" };
            const string content = "[[Acme Inc./file.pdf]] [x](Acme%20Inc./file.pdf)";
            Assert.Equal(content, DottedFolderLinkRewriter.Rewrite(content, both.Contains));
        }

        [Fact]
        public void RepairedTargetThatDoesNotExist_IsLeftAlone()
        {
            const string content = "[[Missing Jr./file.pdf]] [x](Missing%20Jr./file.pdf) [[Re_ Budget-Robert Smith Jr./gone.pdf]]";
            Assert.Equal(content, Rewrite(content));
        }

        [Fact]
        public void ResolvesIsAskedWithDecodedPaths()
        {
            List<string> asked = new List<string>();
            DottedFolderLinkRewriter.Rewrite("[x](A%20B./c%20d.pdf)", p => { asked.Add(p); return false; });
            Assert.Equal(new[] { "A B./c d.pdf", "A B/c d.pdf" }, asked);
        }

        [Fact]
        public void TextOutsideLinks_IsUntouched()
        {
            string content = "---\ntitle: \"Re_ Budget-Robert Smith Jr./file.pdf\"\n---\n"
                + "Plain Re_ Budget-Robert Smith Jr./file.pdf and [Re_ Budget-Robert Smith Jr./file.pdf] text.\n"
                + "[[Re_ Budget-Robert Smith Jr./file.pdf]]\n"
                + "Signed, Robert Smith Jr./file.pdf\n";
            string expected = "---\ntitle: \"Re_ Budget-Robert Smith Jr./file.pdf\"\n---\n"
                + "Plain Re_ Budget-Robert Smith Jr./file.pdf and [Re_ Budget-Robert Smith Jr./file.pdf] text.\n"
                + "[[Re_ Budget-Robert Smith Jr/file.pdf]]\n"
                + "Signed, Robert Smith Jr./file.pdf\n";
            Assert.Equal(expected, Rewrite(content));
        }

        [Fact]
        public void CrLfLineEndings_ArePreserved()
        {
            Assert.Equal("a\r\n[[Thread/file.pdf]]\r\nb\r\n", Rewrite("a\r\n[[Thread./file.pdf]]\r\nb\r\n"));
        }

        [Fact]
        public void CodeSpansAndFencedBlocks_AreUntouched()
        {
            string content = "`[[Thread./file.pdf]]`\n```\n[[Thread./file.pdf]]\n[x](Thread./file.pdf)\n```\n"
                + "~~~md\n[[Thread./file.pdf]]\n~~~\n[[Thread./file.pdf]]\n";
            string expected = "`[[Thread./file.pdf]]`\n```\n[[Thread./file.pdf]]\n[x](Thread./file.pdf)\n```\n"
                + "~~~md\n[[Thread./file.pdf]]\n~~~\n[[Thread/file.pdf]]\n";
            Assert.Equal(expected, Rewrite(content));
        }

        [Fact]
        public void ExternalUrls_AreUntouched()
        {
            const string content = "[x](https://example.com./file.pdf) [y](file:///C:/Thread./file.pdf) [[obsidian://open?a./b]]";
            Assert.Equal(content, DottedFolderLinkRewriter.Rewrite(content, p => true));
        }

        [Fact]
        public void DotSegmentsAndOnlyDots_AreUntouched()
        {
            // "../" climbs a folder, and "..." could never have been created as a folder.
            const string content = "[[../Thread/file.pdf]] [[./Thread/file.pdf]] [[.../file.pdf]] [x](../file.pdf)";
            Assert.Equal(content, DottedFolderLinkRewriter.Rewrite(content, p => !p.Contains("...")));
        }

        [Fact]
        public void ParentRelativeDottedFolder_IsRewritten()
        {
            HashSet<string> exists = new HashSet<string> { "../Thread/file.pdf" };
            Assert.Equal("[[../Thread/file.pdf]]", DottedFolderLinkRewriter.Rewrite("[[../Thread./file.pdf]]", exists.Contains));
        }

        [Fact]
        public void ReservedDeviceName_IsUntouched()
        {
            const string content = "[[CON./file.pdf]]";
            Assert.Equal(content, DottedFolderLinkRewriter.Rewrite(content, p => p != content));
        }

        [Fact]
        public void LinkToFolderWithoutDots_IsUntouchedAndNotQueried()
        {
            bool asked = false;
            const string content = "[[Thread/file.pdf]] [x](Thread/file.pdf) ![[a.png]]";
            Assert.Equal(content, DottedFolderLinkRewriter.Rewrite(content, p => { asked = true; return true; }));
            Assert.False(asked);
        }

        [Fact]
        public void Count_ReportsRewrittenTargets()
        {
            int count;
            string result = DottedFolderLinkRewriter.Rewrite(
                "[[Thread./file.pdf]] [x](Thread./file.pdf) [[Missing./file.pdf]]", Resolves, out count);
            Assert.Equal(2, count);
            Assert.Equal("[[Thread/file.pdf]] [x](Thread/file.pdf) [[Missing./file.pdf]]", result);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        [InlineData("no links here.")]
        public void EmptyOrLinkless_IsReturnedAsIs(string content)
        {
            Assert.Equal(content, Rewrite(content));
        }

        [Fact]
        public void NullResolver_ChangesNothing()
        {
            const string content = "[[Thread./file.pdf]]";
            Assert.Equal(content, DottedFolderLinkRewriter.Rewrite(content, null));
        }
    }
}
