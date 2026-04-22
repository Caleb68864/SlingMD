using System;
using System.IO;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Resolves filename collisions by appending an incrementing "_N" suffix until a free
    /// name is found or <see cref="MaxAttempts"/> is exhausted. Pure helper — file existence
    /// is checked through an injected predicate so tests can supply a fake.
    /// </summary>
    public class UniqueFilenameResolver
    {
        /// <summary>
        /// Maximum number of "_N" suffixes to try before giving up. Matches the prior inline
        /// behavior that capped collision retries at 10.
        /// </summary>
        public const int MaxAttempts = 10;

        /// <summary>
        /// Returns a full path under <paramref name="targetFolder"/> that does not collide,
        /// or <c>null</c> if a free name could not be found within <see cref="MaxAttempts"/>.
        /// </summary>
        /// <param name="targetFolder">Folder the file will be saved into.</param>
        /// <param name="candidateFilename">Initial filename including extension (e.g. "report.pdf").</param>
        /// <param name="exists">Predicate that returns true when the supplied path is taken. Production callers pass <see cref="File.Exists"/>.</param>
        /// <param name="maxAttempts">Optional override of <see cref="MaxAttempts"/> for tests.</param>
        public string Resolve(string targetFolder, string candidateFilename, Func<string, bool> exists, int maxAttempts = MaxAttempts)
        {
            if (string.IsNullOrWhiteSpace(targetFolder) || string.IsNullOrWhiteSpace(candidateFilename) || exists == null)
            {
                return null;
            }

            string fullPath = Path.Combine(targetFolder, candidateFilename);
            if (!exists(fullPath))
            {
                return fullPath;
            }

            string nameWithoutExt = Path.GetFileNameWithoutExtension(candidateFilename);
            string extension = Path.GetExtension(candidateFilename);

            for (int counter = 1; counter <= maxAttempts; counter++)
            {
                string suffixed = $"{nameWithoutExt}_{counter}{extension}";
                string suffixedPath = Path.Combine(targetFolder, suffixed);
                if (!exists(suffixedPath))
                {
                    return suffixedPath;
                }
            }

            return null;
        }
    }
}
