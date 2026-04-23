using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    public class FolderMonitorServiceTests
    {
        private static ObsidianSettings CreateValidSettings()
        {
            return new ObsidianSettings
            {
                VaultBasePath = Path.GetTempPath(),
                VaultName = "TestVault",
                IncludeDailyNoteLink = false
            };
        }

        [Fact]
        public void FolderMonitorService_Constructor_DoesNotThrow()
        {
            ObsidianSettings settings = CreateValidSettings();
            EmailProcessor emailProcessor = new EmailProcessor(settings);
            NotificationService notificationService = new NotificationService(settings);

            System.Exception caughtException = null;
            try
            {
                FolderMonitorService service = new FolderMonitorService(settings, emailProcessor, notificationService, null);
                Assert.NotNull(service);
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            Assert.Null(caughtException);
        }

        [Fact]
        public void FolderMonitorService_StopWatching_WhenNotStarted_DoesNotThrow()
        {
            ObsidianSettings settings = CreateValidSettings();
            EmailProcessor emailProcessor = new EmailProcessor(settings);
            NotificationService notificationService = new NotificationService(settings);
            FolderMonitorService service = new FolderMonitorService(settings, emailProcessor, notificationService, null);

            System.Exception caughtException = null;
            try
            {
                service.StopWatching();
            }
            catch (System.Exception ex)
            {
                caughtException = ex;
            }

            Assert.Null(caughtException);
        }
    }
}
