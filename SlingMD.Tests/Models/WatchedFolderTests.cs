using SlingMD.Outlook.Models;
using Xunit;

namespace SlingMD.Tests.Models
{
    public class WatchedFolderTests
    {
        [Fact]
        public void WatchedFolder_DefaultEnabled_IsTrue()
        {
            WatchedFolder folder = new WatchedFolder();

            Assert.True(folder.Enabled);
        }

        [Fact]
        public void WatchedFolder_DefaultCustomTemplate_IsEmptyString()
        {
            WatchedFolder folder = new WatchedFolder();

            Assert.Equal(string.Empty, folder.CustomTemplate);
        }

        [Fact]
        public void WatchedFolder_SetFolderPath_RetainsValue()
        {
            WatchedFolder folder = new WatchedFolder();
            folder.FolderPath = @"\\TestAccount\Inbox\Clients";

            Assert.Equal(@"\\TestAccount\Inbox\Clients", folder.FolderPath);
        }
    }
}
