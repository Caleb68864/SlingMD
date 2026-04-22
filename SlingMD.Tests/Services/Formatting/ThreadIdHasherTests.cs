using System;
using System.Text.RegularExpressions;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class ThreadIdHasherTests
    {
        private static ThreadIdHasher Hasher()
        {
            return new ThreadIdHasher(new SubjectCleanerService(new ObsidianSettings()));
        }

        [Fact]
        public void Hash_ReturnsTwentyCharacterHexString()
        {
            string hash = Hasher().Hash("Quarterly review");
            Assert.Equal(20, hash.Length);
            Assert.Matches(new Regex("^[0-9A-F]{20}$"), hash);
        }

        [Fact]
        public void Hash_WithRePrefix_ProducesSameHashAsWithout()
        {
            ThreadIdHasher hasher = Hasher();
            string a = hasher.Hash("Re: Quarterly review");
            string b = hasher.Hash("Quarterly review");
            Assert.Equal(a, b);
        }

        [Fact]
        public void Hash_WithFwdPrefix_ProducesSameHashAsWithout()
        {
            ThreadIdHasher hasher = Hasher();
            string a = hasher.Hash("Fwd: Quarterly review");
            string b = hasher.Hash("Quarterly review");
            Assert.Equal(a, b);
        }

        [Fact]
        public void Hash_DifferentSubjects_ProduceDifferentHashes()
        {
            ThreadIdHasher hasher = Hasher();
            Assert.NotEqual(hasher.Hash("Subject A"), hasher.Hash("Subject B"));
        }

        [Fact]
        public void Hash_NullInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, Hasher().Hash(null));
        }

        [Fact]
        public void Hash_EmptyInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, Hasher().Hash(string.Empty));
        }

        [Fact]
        public void Hash_DeterministicAcrossCalls()
        {
            ThreadIdHasher hasher = Hasher();
            Assert.Equal(hasher.Hash("Same Subject"), hasher.Hash("Same Subject"));
        }

        [Fact]
        public void Constructor_NullCleaner_ThrowsArgumentNullException()
        {
            Assert.Throws<ArgumentNullException>(() => new ThreadIdHasher(null));
        }

        [Fact]
        public void Hash_PreservesInWordPrefix()
        {
            // "pre-release notes" must NOT collide with "release notes" — leading "pre-" is part of the word.
            ThreadIdHasher hasher = Hasher();
            Assert.NotEqual(hasher.Hash("pre-release notes"), hasher.Hash("release notes"));
        }
    }
}
