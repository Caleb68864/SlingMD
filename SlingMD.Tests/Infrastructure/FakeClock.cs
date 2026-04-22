using System;
using SlingMD.Outlook.Infrastructure;

namespace SlingMD.Tests.Infrastructure
{
    /// <summary>
    /// Test double for IClock that returns a fixed DateTime.
    /// </summary>
    internal class FakeClock : IClock
    {
        public DateTime Now { get; set; }

        public FakeClock(DateTime now)
        {
            Now = now;
        }
    }
}
