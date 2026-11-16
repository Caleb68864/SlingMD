using System;
using System.IO;
using Xunit;
using SlingMD.Outlook.Services;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Pins the duplicate detection EmailProcessor used before ProcessedEmailCache moved to
    /// SlingMD.Core: ids come from the front matter only, the cache lives 5 minutes, and a different
    /// inbox path forces a rebuild.
    /// </summary>
    public class ProcessedEmailCacheTests : IDisposable
    {
        private static readonly DateTime T0 = new DateTime(2026, 11, 16, 9, 0, 0);
        private readonly string _inbox;

        public ProcessedEmailCacheTests()
        {
            _inbox = Path.Combine(Path.GetTempPath(), "SlingMDTests", "ProcessedEmailCache", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path.Combine(_inbox, "Thread"));
        }

        public void Dispose()
        {
            try
            {
                Directory.Delete(_inbox, true);
            }
            catch (IOException)
            {
            }
        }

        private void WriteNote(string relativePath, string internetMessageId, string entryId, string bodyText = "")
        {
            File.WriteAllText(Path.Combine(_inbox, relativePath),
                "---\ntitle: \"x\"\ninternetMessageId: \"" + internetMessageId + "\"\nentryId: \"" + entryId + "\"\n---\n\n" + bodyText);
        }

        [Fact]
        public void IsDuplicate_IdInFrontMatterOfNestedNote_True()
        {
            WriteNote(Path.Combine("Thread", "a.md"), "<m1@x>", "E1");
            ProcessedEmailCache cache = new ProcessedEmailCache();

            Assert.True(cache.IsDuplicate(_inbox, "<m1@x>", null, T0));
            Assert.True(cache.IsDuplicate(_inbox, null, "E1", T0));
            Assert.False(cache.IsDuplicate(_inbox, "<m2@x>", "E2", T0));
        }

        [Fact]
        public void IsDuplicate_IdOnlyInBody_False()
        {
            WriteNote("a.md", "<m1@x>", "E1", "entryId: \"E9\"");
            ProcessedEmailCache cache = new ProcessedEmailCache();

            Assert.False(cache.IsDuplicate(_inbox, null, "E9", T0));
        }

        [Fact]
        public void IsDuplicate_MissingInbox_FalseAndNoThrow()
        {
            ProcessedEmailCache cache = new ProcessedEmailCache();

            Assert.False(cache.IsDuplicate(Path.Combine(_inbox, "nope"), "<m1@x>", "E1", T0));
        }

        [Fact]
        public void IsDuplicate_NoteAddedLater_SeenOnlyAfterFiveMinutes()
        {
            WriteNote("a.md", "<m1@x>", "E1");
            ProcessedEmailCache cache = new ProcessedEmailCache();
            Assert.True(cache.IsDuplicate(_inbox, "<m1@x>", null, T0));

            WriteNote("b.md", "<m2@x>", "E2");

            Assert.False(cache.IsDuplicate(_inbox, "<m2@x>", null, T0.AddMinutes(4)));
            Assert.True(cache.IsDuplicate(_inbox, "<m2@x>", null, T0.AddMinutes(5)));
        }

        [Fact]
        public void Add_MakesIdADuplicateWithoutRebuild()
        {
            WriteNote("a.md", "<m1@x>", "E1");
            ProcessedEmailCache cache = new ProcessedEmailCache();
            Assert.False(cache.IsDuplicate(_inbox, "<m2@x>", null, T0));

            cache.Add("<m2@x>", "E2");

            Assert.True(cache.IsDuplicate(_inbox, null, "E2", T0.AddMinutes(1)));
        }

        [Fact]
        public void IsDuplicate_DifferentInboxPath_RebuildsForThatPath()
        {
            WriteNote("a.md", "<m1@x>", "E1");
            WriteNote(Path.Combine("Thread", "b.md"), "<m2@x>", "E2");
            ProcessedEmailCache cache = new ProcessedEmailCache();
            Assert.True(cache.IsDuplicate(_inbox, "<m1@x>", null, T0));

            // Scoped to the subfolder now: a.md is outside it.
            Assert.False(cache.IsDuplicate(Path.Combine(_inbox, "Thread"), "<m1@x>", null, T0));
            Assert.True(cache.IsDuplicate(Path.Combine(_inbox, "Thread"), "<m2@x>", null, T0));
        }
    }
}
