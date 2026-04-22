using System;
using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Pure helper that builds the ordered list of candidate paths SlingMD searches when
    /// loading a template file. Search order:
    /// <list type="number">
    ///   <item>The user's vault Templates folder (<see cref="ObsidianSettings.GetTemplatesPath"/>).</item>
    ///   <item>If <see cref="ObsidianSettings.TemplatesFolder"/> is relative: each base directory joined with the relative folder.</item>
    ///   <item>Each base directory joined with the literal "Templates" folder name.</item>
    /// </list>
    /// Duplicates are filtered case-insensitively and original ordering is preserved.
    /// No disk I/O is performed.
    /// </summary>
    public class TemplatePathResolver
    {
        /// <summary>
        /// Resolves <paramref name="templateName"/> against <paramref name="settings"/> and the
        /// supplied <paramref name="baseDirectories"/>, returning ordered candidate file paths.
        /// </summary>
        /// <param name="templateName">The template filename (e.g. "EmailTemplate.md"). Must not be rooted.</param>
        /// <param name="settings">The active settings (used for vault Templates path + TemplatesFolder name).</param>
        /// <param name="baseDirectories">Base directories to consider as fallbacks (typically AppDomain.CurrentDomain.BaseDirectory, the executing assembly directory, and CWD). Duplicates are tolerated.</param>
        public List<string> Resolve(string templateName, ObsidianSettings settings, IEnumerable<string> baseDirectories)
        {
            HashSet<string> directories = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<string> ordered = new List<string>();

            if (settings != null)
            {
                AddDirectory(directories, ordered, settings.GetTemplatesPath());
            }

            string templatesFolder = settings?.TemplatesFolder;
            bool templatesFolderIsRelative = !string.IsNullOrEmpty(templatesFolder) && !Path.IsPathRooted(templatesFolder);

            if (baseDirectories != null)
            {
                if (templatesFolderIsRelative)
                {
                    foreach (string baseDir in baseDirectories)
                    {
                        if (!string.IsNullOrWhiteSpace(baseDir))
                        {
                            AddDirectory(directories, ordered, Path.Combine(baseDir, templatesFolder));
                        }
                    }
                }

                foreach (string baseDir in baseDirectories)
                {
                    if (!string.IsNullOrWhiteSpace(baseDir))
                    {
                        AddDirectory(directories, ordered, Path.Combine(baseDir, "Templates"));
                    }
                }
            }

            List<string> candidates = new List<string>(ordered.Count);
            foreach (string directory in ordered)
            {
                candidates.Add(Path.Combine(directory, templateName));
            }
            return candidates;
        }

        private static void AddDirectory(HashSet<string> seen, List<string> ordered, string path)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }
            if (seen.Add(path))
            {
                ordered.Add(path);
            }
        }
    }
}
