using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class BatchSlingReportTests
    {
        private static EmailSlingResult Saved(string path = "note.md") => new EmailSlingResult { Outcome = EmailSlingOutcome.Saved, NotePath = path };
        private static EmailSlingResult Duplicate() => new EmailSlingResult { Outcome = EmailSlingOutcome.Duplicate };
        private static EmailSlingResult Failed(string error) => new EmailSlingResult { Outcome = EmailSlingOutcome.Failed, ErrorMessage = error };

        [Fact]
        public void BuildMessage_AllSaved_ReportsCountOnly()
        {
            BatchSlingReport report = new BatchSlingReport(2);
            report.Add("A", Saved());
            report.Add("B", Saved());

            Assert.Equal("Slung 2 of 2 emails.", report.BuildMessage("emails"));
        }

        [Fact]
        public void BuildMessage_DuplicatesCancelAndFailures_ReportsEach()
        {
            BatchSlingReport report = new BatchSlingReport(5);
            report.Add("A", Saved());
            report.Add("B", Duplicate());
            report.Add("C", Failed("Path too long"));
            report.Cancelled = true;

            Assert.Equal(
                "Slung 1 of 5 emails. 1 was already in your vault. Cancelled; 2 not processed.\n\n1 failed:\n• C: Path too long\n\nDetails are in the SlingMD log.",
                report.BuildMessage("emails"));
        }

        [Fact]
        public void BuildMessage_ManyFailures_ListsFirstTen()
        {
            BatchSlingReport report = new BatchSlingReport(12);
            for (int i = 0; i < 12; i++)
            {
                report.Add("S" + i, Failed("e"));
            }

            string message = report.BuildMessage("emails");

            Assert.Contains("• S9: e", message);
            Assert.DoesNotContain("• S10", message);
            Assert.Contains("…and 2 more.", message);
        }

        [Fact]
        public void Add_MissingSubjectAndResult_CountsAsFailure()
        {
            BatchSlingReport report = new BatchSlingReport(1);
            report.Add(null, null);

            Assert.Equal(1, report.Failed);
            Assert.Contains("• (no subject): unknown error", report.BuildMessage("emails"));
        }

        [Fact]
        public void Add_Saved_TracksLastSavedNotePath()
        {
            BatchSlingReport report = new BatchSlingReport(2);
            report.Add("A", Saved("a.md"));
            report.Add("B", Saved("b.md"));

            Assert.Equal("b.md", report.LastSavedNotePath);
        }
    }
}
