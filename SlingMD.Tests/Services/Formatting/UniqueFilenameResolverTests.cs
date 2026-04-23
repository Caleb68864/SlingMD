using System.Collections.Generic;
using System.IO;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class UniqueFilenameResolverTests
    {
        private readonly UniqueFilenameResolver _resolver = new UniqueFilenameResolver();

        [Fact]
        public void Resolve_NoCollision_ReturnsCandidatePath()
        {
            string path = _resolver.Resolve("C:\\out", "report.pdf", _ => false);
            Assert.Equal(Path.Combine("C:\\out", "report.pdf"), path);
        }

        [Fact]
        public void Resolve_FirstCollision_AppendsCounterOne()
        {
            HashSet<string> taken = new HashSet<string> { Path.Combine("C:\\out", "report.pdf") };
            string path = _resolver.Resolve("C:\\out", "report.pdf", taken.Contains);
            Assert.Equal(Path.Combine("C:\\out", "report_1.pdf"), path);
        }

        [Fact]
        public void Resolve_MultipleCollisions_FindsFirstFree()
        {
            HashSet<string> taken = new HashSet<string>
            {
                Path.Combine("C:\\out", "report.pdf"),
                Path.Combine("C:\\out", "report_1.pdf"),
                Path.Combine("C:\\out", "report_2.pdf")
            };
            string path = _resolver.Resolve("C:\\out", "report.pdf", taken.Contains);
            Assert.Equal(Path.Combine("C:\\out", "report_3.pdf"), path);
        }

        [Fact]
        public void Resolve_AllAttemptsExhausted_ReturnsNull()
        {
            // Predicate says EVERY path is taken; resolver should give up.
            string path = _resolver.Resolve("C:\\out", "report.pdf", _ => true);
            Assert.Null(path);
        }

        [Fact]
        public void Resolve_HonorsMaxAttemptsOverride()
        {
            HashSet<string> taken = new HashSet<string>
            {
                Path.Combine("C:\\out", "report.pdf"),
                Path.Combine("C:\\out", "report_1.pdf"),
                Path.Combine("C:\\out", "report_2.pdf"),
                Path.Combine("C:\\out", "report_3.pdf")
            };
            // Cap at 2 attempts → "_3" never tried, returns null.
            string path = _resolver.Resolve("C:\\out", "report.pdf", taken.Contains, maxAttempts: 2);
            Assert.Null(path);
        }

        [Fact]
        public void Resolve_NoExtension_StillAppendsCounter()
        {
            HashSet<string> taken = new HashSet<string> { Path.Combine("C:\\out", "README") };
            string path = _resolver.Resolve("C:\\out", "README", taken.Contains);
            Assert.Equal(Path.Combine("C:\\out", "README_1"), path);
        }

        [Fact]
        public void Resolve_NullPredicate_ReturnsNull()
        {
            Assert.Null(_resolver.Resolve("C:\\out", "x.txt", null));
        }

        [Fact]
        public void Resolve_NullCandidate_ReturnsNull()
        {
            Assert.Null(_resolver.Resolve("C:\\out", null, _ => false));
        }

        [Fact]
        public void Resolve_NullTargetFolder_ReturnsNull()
        {
            Assert.Null(_resolver.Resolve(null, "x.txt", _ => false));
        }

        [Fact]
        public void Resolve_DotInName_KeepsExtensionCorrectly()
        {
            // "my.report.v2.pdf" → name "my.report.v2", ext ".pdf"
            HashSet<string> taken = new HashSet<string> { Path.Combine("C:\\out", "my.report.v2.pdf") };
            string path = _resolver.Resolve("C:\\out", "my.report.v2.pdf", taken.Contains);
            Assert.Equal(Path.Combine("C:\\out", "my.report.v2_1.pdf"), path);
        }
    }
}
