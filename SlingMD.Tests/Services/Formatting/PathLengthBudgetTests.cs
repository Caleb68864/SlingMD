using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class PathLengthBudgetTests
    {
        private readonly PathLengthBudget _budget = new PathLengthBudget();

        [Fact]
        public void Fit_NullInput_ReturnsEmpty()
        {
            Assert.Equal(string.Empty, _budget.Fit(null, 20));
        }

        [Fact]
        public void Fit_NameThatFits_Unchanged()
        {
            Assert.Equal("Budget review-SaraA", _budget.Fit("Budget review-SaraA", 19));
        }

        [Fact]
        public void Fit_LongName_CutsAtWordBreak()
        {
            Assert.Equal("Quarterly budget review for", _budget.Fit("Quarterly budget review for the branch", 30));
        }

        [Fact]
        public void Fit_NoWordBreakInLastThird_CutsAtLimit()
        {
            Assert.Equal("Quarterlybudgetrevi", _budget.Fit("Quarterlybudgetreview for", 19));
        }

        [Fact]
        public void Fit_TrimsTrailingDotsSpacesAndSeparators()
        {
            Assert.Equal("Budget", _budget.Fit("Budget..._- tail", 11));
        }

        [Fact]
        public void Fit_BudgetBelowMinimum_KeepsMinimumLength()
        {
            string result = _budget.Fit("Quarterlybudgetreview", -50);

            Assert.Equal(PathLengthBudget.MinNameLength, result.Length);
        }

        [Fact]
        public void FitInDirectory_KeepsFullPathWithinMaxFilePath()
        {
            string directory = @"C:\" + new string('d', 150);
            string name = new string('n', 200);

            string result = _budget.FitInDirectory(directory, name, ".md");

            Assert.Equal(PathLengthBudget.MaxFilePath, directory.Length + 1 + result.Length + ".md".Length);
        }
    }
}
