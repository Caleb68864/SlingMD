using System;
using System.Text.RegularExpressions;

namespace SlingMD.Outlook.Services.Formatting
{
    /// <summary>
    /// Finds the online-meeting join link (Microsoft Teams, Zoom, Google Meet, Webex) in an
    /// appointment's body, unwrapping Safe Links first. Pure helper.
    /// </summary>
    internal static class MeetingLinkFinder
    {
        private static readonly Regex UrlRegex = new Regex(@"https?://[^\s<>""']+", RegexOptions.IgnoreCase);

        private static readonly Regex JoinLinkRegex = new Regex(
            @"^https?://(?:"
            + @"teams\.microsoft\.com/l/meetup-join/"
            + @"|teams\.live\.com/meet/"
            + @"|(?:[\w-]+\.)?zoom\.us/(?:j|my|w)/"
            + @"|meet\.google\.com/[a-z]"
            + @"|[\w-]+\.webex\.com/(?:meet|join|[\w-]+/j\.php)"
            + @")",
            RegexOptions.IgnoreCase);

        /// <summary>Returns the first join link in <paramref name="body"/>, or null when there is none.</summary>
        public static string Find(string body)
        {
            if (string.IsNullOrEmpty(body))
            {
                return null;
            }

            foreach (Match match in UrlRegex.Matches(body))
            {
                string url = HtmlToMarkdownConverter.UnwrapSafeLink(match.Value.TrimEnd('.', ',', ';', ')', ']', '>'));
                if (JoinLinkRegex.IsMatch(url))
                {
                    return url;
                }
            }

            return null;
        }
    }
}
