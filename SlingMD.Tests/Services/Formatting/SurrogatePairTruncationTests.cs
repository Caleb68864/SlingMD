using System;
using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    /// <summary>
    /// Shortening a subject or name must not cut an emoji (or any other character outside the Basic
    /// Multilingual Plane) in half. Half a surrogate pair in a file name becomes U+FFFD when the note
    /// text is written as UTF-8, so links and titles in the note no longer match the file on disk.
    /// Linked into SlingMD.Core.Tests.
    /// </summary>
    public class SurrogatePairTruncationTests
    {
        private const string Rocket = "🚀"; // U+1F680, two UTF-16 code units

        private static void AssertNoLoneSurrogate(string value)
        {
            for (int i = 0; i < value.Length; i++)
            {
                if (char.IsHighSurrogate(value[i]))
                {
                    Assert.True(i + 1 < value.Length && char.IsLowSurrogate(value[i + 1]), $"lone high surrogate at {i} in '{value}'");
                    i++;
                }
                else
                {
                    Assert.False(char.IsLowSurrogate(value[i]), $"lone low surrogate at {i} in '{value}'");
                }
            }
        }

        [Fact]
        public void PathLengthBudget_CutThroughAnEmoji_DropsTheWholeEmoji()
        {
            // 9 letters then the emoji: a 10-unit limit falls between its two halves.
            string name = "Abcdefghi" + Rocket + "jklmnop";

            string fitted = new PathLengthBudget().Fit(name, 10);

            AssertNoLoneSurrogate(fitted);
            Assert.Equal("Abcdefghi", fitted);
        }

        [Fact]
        public void PathLengthBudget_LimitAfterTheEmoji_KeepsIt()
        {
            string fitted = new PathLengthBudget().Fit("Abcdefghi" + Rocket + "jklmnop", 11);

            Assert.Equal("Abcdefghi" + Rocket, fitted);
        }

        [Fact]
        public void NoteTitleBuilder_TruncatesBeforeAnEmoji_NotThroughIt()
        {
            // maxLen 15 leaves 12 characters before "...": the 12th unit is the emoji's first half.
            Dictionary<string, string> tokens = new Dictionary<string, string> { { "Subject", "Launch day " + Rocket + Rocket + " recap" } };

            string title = new NoteTitleBuilder().Build("{Subject}", tokens, 15);

            AssertNoLoneSurrogate(title);
            Assert.Equal("Launch day ...", title);
        }

        [Fact]
        public void ContactLinkFormatter_InitialsOfSupplementaryCharacters_KeepThePair()
        {
            // U+1D49C MATHEMATICAL SCRIPT CAPITAL A, as some display names use for styling.
            string scriptA = "\uD835\uDC9C";
            ContactName name = new ContactNameParser().Parse(scriptA + "lice " + scriptA + "dams", null);

            string link = new ContactLinkFormatter().Format(name, "[[{FirstInitial}{LastInitial}]]");

            AssertNoLoneSurrogate(link);
            Assert.Equal("[[" + scriptA + scriptA + "]]", link);
        }

        [Fact]
        public void ThreadFolderName_LongTopicWithEmojiAtTheCut_StaysWhole()
        {
            ObsidianSettings settings = new ObsidianSettings
            {
                VaultBasePath = Path.Combine(Path.GetTempPath(), "SlingMDTests", "Surrogates"),
                VaultName = "Vault"
            };
            FileService files = new FileService(settings);
            ThreadService threads = new ThreadService(files, new TemplateService(files), settings);
            // 46 letters, then the emoji across positions 46-47: the 47-unit cut splits it.
            string topic = new string('a', 46) + Rocket + " and the rest of a long subject line";
            EmailData email = new EmailData { ConversationTopic = topic, Subject = topic, ReceivedTime = new DateTime(2026, 9, 23) };

            string folder = threads.GetThreadFolderName(email, "", "SaraA", "CalebB");
            string note = threads.GetThreadNoteName(email, "", "SaraA", "CalebB");

            AssertNoLoneSurrogate(folder);
            AssertNoLoneSurrogate(note);
            Assert.Equal(new string('a', 46) + "...-SaraA-CalebB", folder);
        }
    }
}
