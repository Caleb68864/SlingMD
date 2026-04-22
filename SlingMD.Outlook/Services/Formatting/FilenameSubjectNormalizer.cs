using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using SlingMD.Outlook.Models;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Normalizes a subject line into a filename-safe stem by applying an ordered list of
    /// regex find/replace rules — by default the same set SlingMD has always shipped with
    /// (collapse repeated "Re_"/"Fw_" runs, convert "colon + optional space" to "_").
    ///
    /// Rules come from <see cref="ObsidianSettings.FilenameSubjectPatterns"/> when settings is
    /// supplied, falling back to <see cref="BuiltInDefaults"/> when the settings list is null
    /// or empty. The built-in defaults are the canonical source of "what works" — preserved
    /// here so a future settings change can never silently regress baseline behavior.
    ///
    /// Pure helper — no Outlook or filesystem deps.
    /// </summary>
    public class FilenameSubjectNormalizer
    {
        /// <summary>
        /// The shipping defaults — restore these by clearing <see cref="ObsidianSettings.FilenameSubjectPatterns"/>.
        /// </summary>
        public static IReadOnlyList<FilenameSubjectRule> BuiltInDefaults { get; } = new List<FilenameSubjectRule>
        {
            new FilenameSubjectRule { Pattern = @":\s*",                         Replacement = "_"    },
            new FilenameSubjectRule { Pattern = @"(?:Re_\s*)+(?:RE_\s*)+",       Replacement = "Re_"  },
            new FilenameSubjectRule { Pattern = @"(?:RE_\s*)+(?:Re_\s*)+",       Replacement = "Re_"  },
            new FilenameSubjectRule { Pattern = @"(?:Re_\s*){2,}",               Replacement = "Re_"  },
            new FilenameSubjectRule { Pattern = @"(?:RE_\s*){2,}",               Replacement = "Re_"  },
            new FilenameSubjectRule { Pattern = @"(?:Fw_\s*)+(?:FW_\s*)+",       Replacement = "Fw_"  },
            new FilenameSubjectRule { Pattern = @"(?:FW_\s*)+(?:Fw_\s*)+",       Replacement = "Fw_"  },
            new FilenameSubjectRule { Pattern = @"(?:Fw_\s*){2,}",               Replacement = "Fw_"  },
            new FilenameSubjectRule { Pattern = @"(?:FW_\s*){2,}",               Replacement = "Fw_"  },
            new FilenameSubjectRule { Pattern = @"Re_\s+",                       Replacement = "Re_"  },
            new FilenameSubjectRule { Pattern = @"Fw_\s+",                       Replacement = "Fw_"  }
        };

        private readonly ObsidianSettings _settings;

        /// <summary>Constructs a normalizer that always uses the built-in defaults.</summary>
        public FilenameSubjectNormalizer() : this(null) { }

        /// <summary>
        /// Constructs a normalizer that reads rules from <paramref name="settings"/>.
        /// Passing null is acceptable — the built-in defaults will be used.
        /// </summary>
        public FilenameSubjectNormalizer(ObsidianSettings settings)
        {
            _settings = settings;
        }

        /// <summary>
        /// Applies the configured (or default) ordered rule list to <paramref name="subject"/>.
        /// </summary>
        public string Normalize(string subject)
        {
            if (string.IsNullOrEmpty(subject))
            {
                return string.Empty;
            }

            IReadOnlyList<FilenameSubjectRule> rules = ResolveRules();
            string cleaned = subject;

            foreach (FilenameSubjectRule rule in rules)
            {
                if (rule == null || string.IsNullOrEmpty(rule.Pattern))
                {
                    continue;
                }

                try
                {
                    cleaned = Regex.Replace(cleaned, rule.Pattern, rule.Replacement ?? string.Empty, RegexOptions.IgnoreCase);
                }
                catch (ArgumentException)
                {
                    // Skip an invalid user-supplied pattern silently rather than blow up the whole sling.
                }
            }

            return cleaned;
        }

        private IReadOnlyList<FilenameSubjectRule> ResolveRules()
        {
            List<FilenameSubjectRule> configured = _settings?.FilenameSubjectPatterns;
            if (configured != null && configured.Count > 0)
            {
                return configured;
            }
            return BuiltInDefaults;
        }
    }
}
