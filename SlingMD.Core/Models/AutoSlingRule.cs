namespace SlingMD.Outlook.Models
{
    public class AutoSlingRule
    {
        public string Type { get; set; } = "Sender";
        public string Pattern { get; set; } = string.Empty;
        public bool Enabled { get; set; } = true;
    }
}
