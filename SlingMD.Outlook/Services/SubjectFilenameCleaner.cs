using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Services
{
    /// <summary>
    /// Composes the subject-cleaning pipeline used for filename generation:
    /// SubjectCleanerService.Clean → FilenameSubjectNormalizer.Normalize → FileService.CleanFileName.
    /// Extracted to eliminate the duplicated private CleanSubject helper in email and
    /// appointment processors.
    /// </summary>
    public class SubjectFilenameCleaner
    {
        private readonly SubjectCleanerService _subjectCleaner;
        private readonly FilenameSubjectNormalizer _filenameNormalizer;
        private readonly FileService _fileService;

        public SubjectFilenameCleaner(ObsidianSettings settings, FileService fileService)
        {
            _subjectCleaner = new SubjectCleanerService(settings ?? new ObsidianSettings());
            _filenameNormalizer = new FilenameSubjectNormalizer(settings);
            _fileService = fileService;
        }

        /// <summary>
        /// Runs the full cleanup pipeline and returns a filename-safe representation of the subject.
        /// Returns an empty string for null/empty input.
        /// </summary>
        public string Clean(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return string.Empty;
            }

            string cleaned = _subjectCleaner.Clean(subject);
            cleaned = _filenameNormalizer.Normalize(cleaned);
            return _fileService.CleanFileName(cleaned.Trim());
        }
    }
}
