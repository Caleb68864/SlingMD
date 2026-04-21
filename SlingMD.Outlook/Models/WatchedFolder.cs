namespace SlingMD.Outlook.Models
{
    public class WatchedFolder
    {
        public string FolderPath { get; set; } = string.Empty;
        public string CustomTemplate { get; set; } = string.Empty;
        public bool Enabled { get; set; } = true;
    }
}
