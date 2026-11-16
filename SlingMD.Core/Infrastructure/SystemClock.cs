using System;

namespace SlingMD.Outlook.Infrastructure
{
    /// <summary>
    /// Default implementation of IClock that returns the system time.
    /// </summary>
    public class SystemClock : IClock
    {
        /// <summary>
        /// Gets the current local date and time from the system clock.
        /// </summary>
        public DateTime Now
        {
            get { return DateTime.Now; }
        }
    }
}
