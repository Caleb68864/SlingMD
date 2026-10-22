using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class TagNormalizerTests
    {
        [Fact]
        public void Normalize_CleansTagsForObsidian()
        {
            Assert.Equal(
                new[] { "FollowUp", "Follow-Up", "work/projects", "xy", "Café" },
                TagNormalizer.Normalize(new[] { "#FollowUp", " Follow Up ", "work/projects", "x,y!", "", "##", "followup", "Café" }));
        }

        [Fact]
        public void Normalize_Null_ReturnsEmpty()
        {
            Assert.Empty(TagNormalizer.Normalize(null));
        }
    }
}
