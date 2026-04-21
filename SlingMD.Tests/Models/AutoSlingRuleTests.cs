using SlingMD.Outlook.Models;
using Xunit;

namespace SlingMD.Tests.Models
{
    public class AutoSlingRuleTests
    {
        [Fact]
        public void AutoSlingRule_DefaultType_IsSender()
        {
            AutoSlingRule rule = new AutoSlingRule();

            Assert.Equal("Sender", rule.Type);
        }

        [Fact]
        public void AutoSlingRule_DefaultEnabled_IsTrue()
        {
            AutoSlingRule rule = new AutoSlingRule();

            Assert.True(rule.Enabled);
        }

        [Fact]
        public void AutoSlingRule_DefaultPattern_IsEmptyString()
        {
            AutoSlingRule rule = new AutoSlingRule();

            Assert.Equal(string.Empty, rule.Pattern);
        }
    }
}
