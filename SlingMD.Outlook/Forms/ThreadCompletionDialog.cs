using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using Microsoft.Office.Interop.Outlook;

namespace SlingMD.Outlook.Forms
{
    /// <summary>
    /// Dialog that shows un-slung emails in a conversation thread as a checked list,
    /// allowing the user to select which ones to sling to Obsidian.
    /// </summary>
    public partial class ThreadCompletionDialog : Form
    {
        private CheckedListBox clbEmails;
        private Button btnSelectAll;
        private Button btnOk;
        private Button btnCancel;
        private List<MailItem> _emails;

        /// <summary>
        /// Gets the list of emails checked by the user when they clicked "Sling Selected".
        /// </summary>
        public List<MailItem> SelectedEmails
        {
            get
            {
                List<MailItem> selected = new List<MailItem>();
                foreach (int index in clbEmails.CheckedIndices)
                {
                    selected.Add(_emails[index]);
                }
                return selected;
            }
        }

        public ThreadCompletionDialog(List<MailItem> missingEmails)
        {
            _emails = missingEmails;
            InitializeComponents();
            PopulateList();
        }

        private void InitializeComponents()
        {
            this.SuspendLayout();

            // Label
            Label lblInfo = new Label();
            lblInfo.Dock = DockStyle.Top;
            lblInfo.Text = "The following emails from this thread have not been slung to Obsidian:";
            lblInfo.Padding = new Padding(5);
            lblInfo.AutoSize = true;

            // Checked list box
            clbEmails = new CheckedListBox();
            clbEmails.Dock = DockStyle.Fill;
            clbEmails.CheckOnClick = true;
            clbEmails.FormattingEnabled = true;
            clbEmails.Name = "clbEmails";
            clbEmails.Size = new Size(560, 280);

            // Button panel
            Panel btnPanel = new Panel();
            btnPanel.Dock = DockStyle.Bottom;
            btnPanel.Height = 50;

            // Select All button
            btnSelectAll = new Button();
            btnSelectAll.Text = "Select All";
            btnSelectAll.Size = new Size(100, 30);
            btnSelectAll.Location = new Point(10, 10);
            btnSelectAll.Click += BtnSelectAll_Click;

            // OK (Sling Selected) button
            btnOk = new Button();
            btnOk.Text = "Sling Selected";
            btnOk.DialogResult = DialogResult.OK;
            btnOk.Size = new Size(120, 30);
            btnOk.Location = new Point(320, 10);
            btnOk.Click += BtnOk_Click;

            // Cancel button
            btnCancel = new Button();
            btnCancel.Text = "Cancel";
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Size = new Size(90, 30);
            btnCancel.Location = new Point(450, 10);

            btnPanel.Controls.Add(btnSelectAll);
            btnPanel.Controls.Add(btnOk);
            btnPanel.Controls.Add(btnCancel);

            this.Controls.Add(clbEmails);
            this.Controls.Add(lblInfo);
            this.Controls.Add(btnPanel);

            // Form settings
            this.ClientSize = new Size(580, 380);
            this.Text = "Complete Thread";
            this.AcceptButton = btnOk;
            this.CancelButton = btnCancel;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;

            this.ResumeLayout(false);
        }

        private void PopulateList()
        {
            clbEmails.Items.Clear();
            foreach (MailItem mail in _emails)
            {
                try
                {
                    string subject = mail.Subject ?? "(No Subject)";
                    string sender = mail.SenderName ?? "(Unknown Sender)";
                    string dateStr = mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm");
                    string displayText = string.Format("{0} - {1} ({2})", subject, sender, dateStr);
                    clbEmails.Items.Add(displayText, true);
                }
                catch (System.Exception)
                {
                    clbEmails.Items.Add("(Unable to read email details)", true);
                }
            }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            // SelectedEmails is computed on-demand from checked indices; nothing extra needed here
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            bool anyUnchecked = false;
            for (int i = 0; i < clbEmails.Items.Count; i++)
            {
                if (!clbEmails.GetItemChecked(i))
                {
                    anyUnchecked = true;
                    break;
                }
            }

            for (int i = 0; i < clbEmails.Items.Count; i++)
            {
                clbEmails.SetItemChecked(i, anyUnchecked);
            }
        }
    }
}
