using System;
using System.Globalization;
using System.Threading;
using SlingMD.Outlook.Services.Formatting;
using Xunit;

namespace SlingMD.Tests.Services.Formatting
{
    public class DateFormatterTests
    {
        private readonly DateFormatter _formatter = new DateFormatter();
        private static readonly DateTime SampleDate = new DateTime(2026, 4, 21, 14, 5, 9);

        [Fact]
        public void Format_DateOnly_RendersIsoDate()
        {
            Assert.Equal("2026-04-21", _formatter.Format(SampleDate, "yyyy-MM-dd"));
        }

        [Fact]
        public void Format_DateAndTime_RendersFullTimestamp()
        {
            Assert.Equal("2026-04-21 14:05:09", _formatter.Format(SampleDate, "yyyy-MM-dd HH:mm:ss"));
        }

        [Fact]
        public void Format_UsesInvariantCulture_NotCurrent()
        {
            CultureInfo previous = Thread.CurrentThread.CurrentCulture;
            try
            {
                // de-DE would render the month name differently; invariant culture should win.
                Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");
                Assert.Equal("April", _formatter.Format(SampleDate, "MMMM"));
            }
            finally
            {
                Thread.CurrentThread.CurrentCulture = previous;
            }
        }

        [Fact]
        public void Format_InvalidFormatString_ThrowsFormatException()
        {
            // A single backslash with nothing to escape is invalid.
            Assert.Throws<FormatException>(() => _formatter.Format(SampleDate, "\\"));
        }

        [Fact]
        public void FormatOrDefault_ValidFormat_ReturnsFormattedDate()
        {
            Assert.Equal("2026-04-21", _formatter.FormatOrDefault(SampleDate, "yyyy-MM-dd", "fallback"));
        }

        [Fact]
        public void FormatOrDefault_NullFormat_ReturnsFallback()
        {
            Assert.Equal("FB", _formatter.FormatOrDefault(SampleDate, null, "FB"));
        }

        [Fact]
        public void FormatOrDefault_EmptyFormat_ReturnsFallback()
        {
            Assert.Equal("FB", _formatter.FormatOrDefault(SampleDate, string.Empty, "FB"));
        }

        [Fact]
        public void FormatOrDefault_InvalidFormat_ReturnsFallback()
        {
            // Single backslash is invalid as a format specifier.
            Assert.Equal("FB", _formatter.FormatOrDefault(SampleDate, "\\", "FB"));
        }
    }
}
