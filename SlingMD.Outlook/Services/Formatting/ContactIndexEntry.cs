using System.Collections.Generic;

namespace SlingMD.Outlook.Services.Formatting
{
    internal class ContactIndexEntry
    {
        public string FilePath { get; }
        public string DisplayName { get; }
        public IReadOnlyList<string> Aliases { get; }
        public string NormalizedDisplayName { get; }
        public IReadOnlyList<string> NormalizedAliases { get; }

        public ContactIndexEntry(
            string filePath,
            string displayName,
            IReadOnlyList<string> aliases,
            string normalizedDisplayName,
            IReadOnlyList<string> normalizedAliases)
        {
            FilePath = filePath;
            DisplayName = displayName;
            Aliases = aliases;
            NormalizedDisplayName = normalizedDisplayName;
            NormalizedAliases = normalizedAliases;
        }
    }
}
