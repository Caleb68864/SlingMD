using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using SlingMD.Outlook.Models;
using SlingMD.Outlook.Services;
using Xunit;

namespace SlingMD.Tests.Services
{
    /// <summary>
    /// Dates SlingMD writes into names and notes must not depend on Windows' regional settings. Under
    /// th-TH the default calendar is Thai Buddhist (2026 is 2569), and under fi-FI the time separator
    /// is '.', so a bare <c>ToString("yyyy-MM-dd")</c> or <c>"HH:mm:ss"</c> gave titles, file names
    /// and task dates that disagreed with the ISO date properties. Linked into SlingMD.Core.Tests.
    /// </summary>
    public class CultureInvariantDatesTests : IDisposable
    {
        private static readonly DateTime Received = new DateTime(2026, 9, 23, 18, 18, 52);
        private readonly CultureInfo _originalCulture;
        private readonly string _testDir;

        public CultureInvariantDatesTests()
        {
            _originalCulture = CultureInfo.CurrentCulture;
            _testDir = Path.Combine(Path.GetTempPath(), "SlingMDTests", "CultureDates", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_testDir);
        }

        public void Dispose()
        {
            CultureInfo.CurrentCulture = _originalCulture;
            try
            {
                Directory.Delete(_testDir, true);
            }
            catch (IOException)
            {
            }
        }

        private static void UseThai()
        {
            CultureInfo.CurrentCulture = ThaiBuddhist();
            Assert.Equal("2569", Received.ToString("yyyy"));
        }

        /// <summary>
        /// th-TH with its Buddhist calendar set explicitly, so the test does not depend on the
        /// platform's culture data (ICU on Linux, NLS on Windows).
        /// </summary>
        internal static CultureInfo ThaiBuddhist()
        {
            CultureInfo culture = (CultureInfo)new CultureInfo("th-TH").Clone();
            culture.DateTimeFormat.Calendar = new ThaiBuddhistCalendar();
            return culture;
        }

        private ObsidianSettings NewSettings()
        {
            return new ObsidianSettings { VaultBasePath = _testDir, VaultName = "Vault", IncludeDailyNoteLink = false };
        }

        private EmailNoteEngine EmailEngine(ObsidianSettings settings)
        {
            return new EmailNoteEngine(settings, new TemplateService(new FileService(settings)), (name, email) => $"[[{name}]]");
        }

        [Fact]
        public void EmailNoteTitle_ThaiCulture_GregorianDate()
        {
            UseThai();

            Assert.Equal("Budget review - 2026-09-23", EmailEngine(NewSettings()).BuildNoteTitle("Budget review", "SaraA", Received));
        }

        [Fact]
        public void EmailFileNameDateToken_ThaiCulture_GregorianDate()
        {
            UseThai();
            ObsidianSettings settings = NewSettings();
            settings.EmailFilenameFormat = "{Date} {Subject}";

            Assert.Equal("2026-09-23 Budget", EmailEngine(settings).BuildBaseFileName(Received, "Budget", "SaraA", "t", false));
        }

        [Fact]
        public void EmailTemplateDateToken_ThaiCulture_GregorianDate()
        {
            UseThai();
            Directory.CreateDirectory(Path.Combine(_testDir, "Vault", "Templates"));
            File.WriteAllText(Path.Combine(_testDir, "Vault", "Templates", "EmailTemplate.md"), "{{date}}");
            ObsidianSettings settings = NewSettings();
            settings.TemplatesFolder = "Templates";
            EmailData email = new EmailData { Subject = "Budget", SenderName = "Sara", ReceivedTime = Received };

            string note = EmailEngine(settings).RenderNote(email, new Dictionary<string, object>(), "T", "Budget", "SaraA", "", "T", "id", "", "");

            Assert.Equal("2026-09-23", note);
        }

        [Fact]
        public void AppointmentNoteTitle_ThaiCulture_GregorianDate()
        {
            UseThai();
            ObsidianSettings settings = NewSettings();
            AppointmentNoteEngine engine = new AppointmentNoteEngine(settings, new TemplateService(new FileService(settings)), null);

            Assert.Equal("2026-09-23 - Budget sync", engine.BuildNoteTitle("Budget sync", "SaraA", Received));
        }

        [Fact]
        public void ClipboardFrontMatter_FinnishCulture_ColonTimeSeparator()
        {
            // fi-FI writes times as 18.18; set the separator explicitly rather than rely on culture data.
            CultureInfo finnish = (CultureInfo)new CultureInfo("fi-FI").Clone();
            finnish.DateTimeFormat.TimeSeparator = ".";
            CultureInfo.CurrentCulture = finnish;
            Assert.Equal("18.18", Received.ToString("HH:mm"));

            string text = new ClipboardEmailService().Format(
                new[] { new ClipboardEmail { Subject = "S", SentOn = Received, ReceivedTime = Received } }, true);

            Assert.Contains("received: \"2026-09-23 18:18:52\"", text);
        }
    }
}
