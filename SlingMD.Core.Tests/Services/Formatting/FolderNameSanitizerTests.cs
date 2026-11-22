using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Core.Tests.Services.Formatting
{
    /// <summary>
    /// <see cref="FolderNameSanitizer"/>: the step every folder-building site applies on top of
    /// <see cref="FileNameSanitizer"/>, whose output keeps a trailing dot for contact names.
    /// </summary>
    public class FolderNameSanitizerTests
    {
        private readonly FileNameSanitizer _fileNameSanitizer = new FileNameSanitizer();

        [Fact]
        public void FileNameSanitizer_StillKeepsTrailingDot_ForContactNames()
        {
            // The contact-name behaviour the folder step must not change.
            Assert.Equal("Robert Smith Jr.", _fileNameSanitizer.Sanitize("Robert Smith Jr."));
        }

        [Theory]
        [InlineData("Robert Smith Jr.", "Robert Smith Jr")]
        [InlineData("Budget review-Bob-Robert Smith Jr.", "Budget review-Bob-Robert Smith Jr")]
        [InlineData("Weekly sync ", "Weekly sync")]
        [InlineData("Ends in dots...", "Ends in dots")]
        [InlineData("Mixed . . ", "Mixed")]
        public void Sanitize_DropsTrailingDotsAndSpaces(string input, string expected)
        {
            Assert.Equal(expected, FolderNameSanitizer.Sanitize(input));
        }

        [Fact]
        public void FolderBuiltFromSanitizedContactName_HasNoTrailingDot()
        {
            // How a thread folder or attachment subfolder name is produced: the file-name sanitiser
            // first, then the folder step. Windows would drop the trailing dot from the folder on disk,
            // so the name SlingMD links to must not have one either.
            string folder = FolderNameSanitizer.Sanitize(
                "Re_ Budget-Bob-" + _fileNameSanitizer.Sanitize("Robert Smith Jr."));

            Assert.False(folder.EndsWith("."), folder);
            Assert.False(folder.EndsWith(" "), folder);
            Assert.Equal("Re_ Budget-Bob-Robert Smith Jr", folder);
        }

        [Theory]
        [InlineData("Normal folder")]
        [InlineData("2026-11-16 Thread-Alice-Bob")]
        [InlineData("v1.2 notes")]
        [InlineData(".hidden-style")]
        [InlineData(" leading space")]
        public void Sanitize_LeavesSafeNamesUnchanged(string input)
        {
            Assert.Equal(input, FolderNameSanitizer.Sanitize(input));
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void Sanitize_NullOrEmpty_ReturnsEmpty(string input)
        {
            Assert.Equal(string.Empty, FolderNameSanitizer.Sanitize(input));
        }

        [Theory]
        [InlineData(".")]
        [InlineData("..")]
        [InlineData("...")]
        [InlineData(" . ")]
        public void Sanitize_NameOfOnlyDotsAndSpaces_BecomesPlaceholder(string input)
        {
            // "..." as a path segment would resolve to the parent folder.
            Assert.Equal(FolderNameSanitizer.EmptyNameReplacement, FolderNameSanitizer.Sanitize(input));
        }

        [Theory]
        [InlineData("CON.", "_CON")]
        [InlineData("nul ", "_nul")]
        [InlineData("COM1..", "_COM1")]
        [InlineData("Con.Ed bill", "_Con.Ed bill")]
        [InlineData("AUX .x", "_AUX .x")]
        [InlineData("LPT0", "_LPT0")]
        [InlineData("COM\u00B2.", "_COM\u00B2")]
        public void Sanitize_NameThatTrimsToReservedDevice_IsPrefixed(string input, string expected)
        {
            Assert.Equal(expected, FolderNameSanitizer.Sanitize(input));
        }
    }
}
