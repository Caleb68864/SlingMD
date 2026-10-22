using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class MeetingLinkFinderTests
    {
        [Fact]
        public void Find_TeamsInvitation_ReturnsJoinLink()
        {
            string body = "Microsoft Teams meeting\r\nJoin on your computer\r\nClick here to join the meeting"
                + "<https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=%7b%22Tid%22%3a%22x%22%7d>\r\nMeeting ID: 123";

            Assert.Equal(
                "https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=%7b%22Tid%22%3a%22x%22%7d",
                MeetingLinkFinder.Find(body));
        }

        [Fact]
        public void Find_ZoomInsideSafeLink_ReturnsUnwrappedLink()
        {
            string body = "Join: <https://nam12.safelinks.protection.outlook.com/?url=https%3A%2F%2Fus02web.zoom.us%2Fj%2F123456789%3Fpwd%3Dabc&data=05>";

            Assert.Equal("https://us02web.zoom.us/j/123456789?pwd=abc", MeetingLinkFinder.Find(body));
        }

        [Fact]
        public void Find_SkipsOtherLinks()
        {
            Assert.Equal("https://meet.google.com/abc-defg-hij", MeetingLinkFinder.Find("See https://example.com/agenda and https://meet.google.com/abc-defg-hij."));
            Assert.Null(MeetingLinkFinder.Find("No link here, just https://example.com"));
            Assert.Null(MeetingLinkFinder.Find(null));
        }
    }
}
