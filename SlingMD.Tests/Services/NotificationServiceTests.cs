using System;
using System.Collections.Generic;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Pins what NotificationService did with ToastForm before it moved to SlingMD.Core behind
    /// <see cref="IToastNotifier"/>: Toast mode shows the message (errors with the details appended),
    /// Silent mode shows nothing, and a toast that throws never reaches the caller.
    /// </summary>
    public class NotificationServiceTests
    {
        private sealed class RecordingToastNotifier : IToastNotifier
        {
            public readonly List<string> Calls = new List<string>();

            public void ShowToast(string message, bool isError) => Calls.Add($"{(isError ? "Error" : "Info")}:{message}");
        }

        private sealed class ThrowingToastNotifier : IToastNotifier
        {
            public void ShowToast(string message, bool isError) => throw new InvalidOperationException("no UI");
        }

        private static ObsidianSettings Settings(string mode)
        {
            return new ObsidianSettings { AutoSlingNotificationMode = mode };
        }

        [Fact]
        public void Notify_ToastMode_ShowsInfoToast()
        {
            RecordingToastNotifier toasts = new RecordingToastNotifier();
            new NotificationService(Settings("Toast"), toasts).Notify("Slung 3 emails");

            Assert.Equal(new[] { "Info:Slung 3 emails" }, toasts.Calls);
        }

        [Fact]
        public void NotifyError_ToastMode_AppendsDetails()
        {
            RecordingToastNotifier toasts = new RecordingToastNotifier();
            new NotificationService(Settings("Toast"), toasts).NotifyError("Auto-sling failed", "disk full");

            Assert.Equal(new[] { "Error:Auto-sling failed\n\nDetails: disk full" }, toasts.Calls);
        }

        [Theory]
        [InlineData(null)]
        [InlineData("")]
        public void NotifyError_NoDetails_ShowsMessageOnly(string error)
        {
            RecordingToastNotifier toasts = new RecordingToastNotifier();
            new NotificationService(Settings("Toast"), toasts).NotifyError("Auto-sling failed", error);

            Assert.Equal(new[] { "Error:Auto-sling failed" }, toasts.Calls);
        }

        [Fact]
        public void SilentMode_ShowsNothing()
        {
            RecordingToastNotifier toasts = new RecordingToastNotifier();
            NotificationService service = new NotificationService(Settings("Silent"), toasts);

            service.Notify("hello");
            service.NotifyError("bad", "details");

            Assert.Empty(toasts.Calls);
        }

        [Fact]
        public void ToastThatThrows_DoesNotReachCaller()
        {
            NotificationService service = new NotificationService(Settings("Toast"), new ThrowingToastNotifier());

            service.Notify("hello");
            service.NotifyError("bad", "details");
        }

        [Fact]
        public void Constructor_NullNotifier_Throws()
        {
            Assert.Throws<ArgumentNullException>(() => new NotificationService(Settings("Toast"), null));
        }
    }
}
