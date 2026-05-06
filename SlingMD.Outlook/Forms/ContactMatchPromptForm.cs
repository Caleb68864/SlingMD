using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using SlingMD.Outlook.Services;
using SlingMD.Outlook.Services.Formatting;

namespace SlingMD.Outlook.Forms
{
    internal enum MatchPromptDecision
    {
        Match,
        CreateNew
    }

    internal sealed class MatchPromptResult
    {
        public MatchPromptDecision Decision { get; }
        public ContactIndexEntry ChosenEntry { get; }
        public bool SaveAsAlias { get; }

        public MatchPromptResult(MatchPromptDecision decision, ContactIndexEntry chosenEntry, bool saveAsAlias)
        {
            Decision = decision;
            ChosenEntry = chosenEntry;
            SaveAsAlias = saveAsAlias;
        }
    }

    internal partial class ContactMatchPromptForm : Form
    {
        private readonly IReadOnlyList<ContactIndexEntry> _candidates;
        private readonly bool _defaultSaveAsAlias;

        public MatchPromptResult Result { get; private set; }

        public ContactMatchPromptForm(
            IReadOnlyList<ContactIndexEntry> candidates,
            ContactInteractionMode mode,
            bool defaultSaveAsAlias)
        {
            if (mode == ContactInteractionMode.Automated)
                throw new InvalidOperationException(
                    "ContactMatchPromptForm cannot be shown in Automated interaction mode.");

            _candidates = candidates;
            _defaultSaveAsAlias = defaultSaveAsAlias;

            InitializeComponent();
            PopulateCandidateList();

            Result = new MatchPromptResult(MatchPromptDecision.CreateNew, null, defaultSaveAsAlias);
        }

        private void PopulateCandidateList()
        {
            lstCandidates.Items.Clear();
            foreach (ContactIndexEntry entry in _candidates)
            {
                string label = BuildCandidateLabel(entry);
                lstCandidates.Items.Add(label);
            }

            if (lstCandidates.Items.Count > 0)
                lstCandidates.SelectedIndex = 0;
        }

        private static string BuildCandidateLabel(ContactIndexEntry entry)
        {
            string filename = Path.GetFileNameWithoutExtension(entry.FilePath);
            string company = TryReadFrontmatterField(entry.FilePath, "company");
            string email = TryReadFrontmatterField(entry.FilePath, "email");

            if (string.IsNullOrEmpty(company) && string.IsNullOrEmpty(email))
                return filename;

            List<string> extras = new List<string>();
            if (!string.IsNullOrEmpty(company))
                extras.Add(company);
            if (!string.IsNullOrEmpty(email))
                extras.Add(email);

            return string.Format("{0}  ({1})", filename, string.Join(", ", extras));
        }

        private static string TryReadFrontmatterField(string filePath, string fieldName)
        {
            try
            {
                if (!File.Exists(filePath))
                    return null;

                string prefix = fieldName + ":";
                bool inFrontmatter = false;
                int dashLineCount = 0;

                foreach (string line in File.ReadLines(filePath))
                {
                    string trimmed = line.Trim();

                    if (trimmed == "---")
                    {
                        dashLineCount++;
                        if (dashLineCount == 1)
                        {
                            inFrontmatter = true;
                            continue;
                        }
                        break;
                    }

                    if (!inFrontmatter)
                        break;

                    if (trimmed.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                    {
                        string value = trimmed.Substring(prefix.Length).Trim().Trim('"');
                        return string.IsNullOrEmpty(value) ? null : value;
                    }
                }
            }
            catch (System.Exception)
            {
                // Best-effort; silently ignore file read errors
            }

            return null;
        }

        private void btnMatch_Click(object sender, EventArgs e)
        {
            int idx = lstCandidates.SelectedIndex;
            ContactIndexEntry chosen = (idx >= 0 && idx < _candidates.Count) ? _candidates[idx] : null;
            Result = new MatchPromptResult(MatchPromptDecision.Match, chosen, chkSaveAsAlias.Checked);
            DialogResult = DialogResult.OK;
            Close();
        }

        private void btnCreateNew_Click(object sender, EventArgs e)
        {
            Result = new MatchPromptResult(MatchPromptDecision.CreateNew, null, chkSaveAsAlias.Checked);
            DialogResult = DialogResult.Cancel;
            Close();
        }
    }
}
