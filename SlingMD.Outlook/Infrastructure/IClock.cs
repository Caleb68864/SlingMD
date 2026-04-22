using System;

namespace SlingMD.Outlook.Infrastructure
{
    /// <summary>
    /// Abstraction for time operations to support testability.
    /// </summary>
    public interface IClock
    {
        /// <summary>
        /// Gets the current local date and time.
        /// </summary>
        DateTime Now { get; }
    }
}
