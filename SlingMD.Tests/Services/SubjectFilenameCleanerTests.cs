using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class SubjectFilenameCleanerTests
    {
        private static SubjectFilenameCleaner NewCleaner()
        {
            ObsidianSettings settings = new ObsidianSettings();
            FileService fileService = new FileService(settings);
            return new SubjectFilenameCleaner(settings, fileService);
        }

        [Fact]
        public void Clean_NullOrEmpty_ReturnsEmpty()
        {
            SubjectFilenameCleaner cleaner = NewCleaner();
            Assert.Equal(string.Empty, cleaner.Clean(null));
            Assert.Equal(string.Empty, cleaner.Clean(string.Empty));
        }

        [Fact]
        public void Clean_StripsRePrefix()
        {
            string result = NewCleaner().Clean("Re: project update");
            Assert.DoesNotContain("Re:", result);
            Assert.Contains("project", result);
        }

        [Fact]
        public void Clean_StripsExternalTag()
        {
            string result = NewCleaner().Clean("[EXTERNAL] Quarterly review");
            Assert.DoesNotContain("[EXTERNAL]", result);
            Assert.Contains("Quarterly review", result);
        }

        [Fact]
        public void Clean_PreservesPreReleaseWordBoundary()
        {
            // Legacy bug: the "Re" prefix strip used to match inside "pre-release".
            string result = NewCleaner().Clean("pre-release notes");
            Assert.Contains("pre-release notes", result);
        }

        [Fact]
        public void Clean_RemovesColonForFilename()
        {
            // FilenameSubjectNormalizer step should replace ": " with "_".
            string result = NewCleaner().Clean("Status: approved");
            Assert.DoesNotContain(":", result);
        }

        [Fact]
        public void Clean_CollapsesRepeatedRe()
        {
            string result = NewCleaner().Clean("Re: Re: Re: ping");
            // After SubjectCleanerService strips the leading Re runs, the result should be "ping".
            Assert.Equal("ping", result);
        }
    }
}
