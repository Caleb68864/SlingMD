using System;
using System.Collections.Generic;
using System.IO;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Maps email addresses to the contact notes that list them in their <c>email</c> frontmatter, so a
    /// sender is recognised even when Outlook shows a different display name ("Smith, John" vs
    /// "John Smith"). Addresses compare case-insensitively; the first note listing an address wins.
    /// </summary>
    internal class ContactEmailIndex
    {
        private const int MaxFiles = 5000;

        private readonly Dictionary<string, string> _stemsByEmail = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly FrontmatterReader _frontmatter = new FrontmatterReader();

        /// <param name="contactsFolderPath">The contacts folder (searched at the top level).</param>
        /// <param name="vaultPath">The whole vault, when contacts may live anywhere in it; otherwise null.</param>
        public ContactEmailIndex(string contactsFolderPath, string vaultPath = null)
        {
            IndexDirectory(contactsFolderPath, SearchOption.TopDirectoryOnly);
            IndexDirectory(vaultPath, SearchOption.AllDirectories);
        }

        /// <summary>Returns the file name (without .md) of the contact note listing <paramref name="email"/>, or null.</summary>
        public string FindStem(string email)
        {
            if (string.IsNullOrWhiteSpace(email))
                return null;

            string stem;
            return _stemsByEmail.TryGetValue(email.Trim(), out stem) ? stem : null;
        }

        /// <summary>Records a note written after the index was built.</summary>
        public void Add(string email, string stem)
        {
            if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(stem) || email.IndexOf('@') <= 0)
                return;

            string key = email.Trim();
            if (!_stemsByEmail.ContainsKey(key))
                _stemsByEmail[key] = stem;
        }

        private void IndexDirectory(string directory, SearchOption searchOption)
        {
            if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
                return;

            string[] files = Directory.GetFiles(directory, "*.md", searchOption);
            if (files.Length > MaxFiles)
                return;

            foreach (string filePath in files)
            {
                string stem = Path.GetFileNameWithoutExtension(filePath);
                foreach (string email in _frontmatter.ExtractEmails(ContactMatcher.ReadFrontmatter(filePath)))
                    Add(email, stem);
            }
        }
    }
}
