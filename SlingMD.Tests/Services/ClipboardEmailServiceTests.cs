using System;
using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class ClipboardEmailServiceTests
    {
        [Fact]
        public void Format_SingleBodyOnly_ReturnsBodyWithoutWrapper()
        {
            ClipboardEmailService service = new ClipboardEmailService();

            string result = service.Format(new[] { Email("Latest body", 2) }, includeFrontmatter: false);

            Assert.Equal("Latest body", result);
        }

        [Fact]
        public void Format_WithFrontmatter_IncludesUsefulEmailMetadata()
        {
            ClipboardEmailService service = new ClipboardEmailService();
            ClipboardEmail email = Email("Message", 1);
            email.Subject = "A \"quoted\" subject";
            email.SenderName = "Ada Lovelace";
            email.SenderEmail = "ada@example.com";
            email.To.Add("Charles Babbage <charles@example.com>");

            string result = service.Format(new[] { email }, includeFrontmatter: true);

            Assert.Contains("title: \"A \\\"quoted\\\" subject\"", result);
            Assert.Contains("fromEmail: \"ada@example.com\"", result);
            Assert.Contains("to:\n  - \"Charles Babbage <charles@example.com>\"", result.Replace("\r\n", "\n"));
            Assert.EndsWith("Message", result);
        }

        [Fact]
        public void Format_MultipleEmails_OrdersChronologicallyAndSeparatesMessages()
        {
            ClipboardEmailService service = new ClipboardEmailService();
            List<ClipboardEmail> emails = new List<ClipboardEmail>
            {
                Email("Second", 2),
                Email("First", 1)
            };

            string result = service.Format(emails, includeFrontmatter: false);

            Assert.True(result.IndexOf("First", StringComparison.Ordinal) < result.IndexOf("Second", StringComparison.Ordinal));
            Assert.Contains("Email 2 of 2", result);
        }

        [Fact]
        public void ExceedsClipboardWarningSize_UsesUtf16PayloadSize()
        {
            ClipboardEmailService service = new ClipboardEmailService();
            string atLimit = new string('a', ClipboardEmailService.ClipboardWarningBytes / 2);

            Assert.False(service.ExceedsClipboardWarningSize(atLimit));
            Assert.True(service.ExceedsClipboardWarningSize(atLimit + "a"));
        }

        private static ClipboardEmail Email(string body, int hour)
        {
            return new ClipboardEmail
            {
                Body = body,
                EntryId = hour.ToString(),
                ConversationId = "thread",
                SentOn = new DateTime(2026, 1, 2, hour, 0, 0),
                ReceivedTime = new DateTime(2026, 1, 2, hour, 1, 0)
            };
        }
    }
}
