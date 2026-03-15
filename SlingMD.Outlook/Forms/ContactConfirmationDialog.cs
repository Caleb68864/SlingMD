using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    public partial class ContactConfirmationDialog : Form
    {
        public List<string> SelectedContacts { get; private set; } = new List<string>();
        private List<string> _contacts;

        public ContactConfirmationDialog(List<string> contacts)
        {
            InitializeComponent();
            _contacts = contacts;
            PopulateContactList();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            
            // Create checklist box
            var chkListContacts = new CheckedListBox();
            chkListContacts.Dock = DockStyle.Fill;
            chkListContacts.CheckOnClick = true;
            chkListContacts.FormattingEnabled = true;
            chkListContacts.Name = "chkListContacts";
            chkListContacts.Size = new System.Drawing.Size(380, 250);
            
            // Create label
            var lblInfo = new Label();
            lblInfo.Dock = DockStyle.Top;
            lblInfo.Text = "Select contacts to create notes for:";
            lblInfo.Padding = new Padding(5);
            lblInfo.AutoSize = true;
            
            // Create button panel
            var btnPanel = new Panel();
            btnPanel.Dock = DockStyle.Bottom;
            btnPanel.Height = 50;
            
            // Create buttons
            var btnOk = new Button();
            btnOk.Text = "Create Selected";
            btnOk.DialogResult = DialogResult.OK;
            btnOk.Size = new System.Drawing.Size(120, 30);
            btnOk.Location = new System.Drawing.Point(120, 10);
            btnOk.Click += BtnOk_Click;
            
            var btnCancel = new Button();
            btnCancel.Text = "Skip Creating Contacts";
            btnCancel.DialogResult = DialogResult.Cancel;
            btnCancel.Size = new System.Drawing.Size(160, 30);
            btnCancel.Location = new System.Drawing.Point(250, 10);
            
            var btnSelectAll = new Button();
            btnSelectAll.Text = "Select All";
            btnSelectAll.Size = new System.Drawing.Size(100, 30);
            btnSelectAll.Location = new System.Drawing.Point(10, 10);
            btnSelectAll.Click += BtnSelectAll_Click;
            
            // Add controls
            btnPanel.Controls.Add(btnOk);
            btnPanel.Controls.Add(btnCancel);
            btnPanel.Controls.Add(btnSelectAll);
            
            this.Controls.Add(chkListContacts);
            this.Controls.Add(lblInfo);
            this.Controls.Add(btnPanel);
            
            // Form settings
            this.ClientSize = new System.Drawing.Size(400, 350);
            this.Text = "Create Contact Notes";
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

        private void PopulateContactList()
        {
            var chkListContacts = Controls.OfType<CheckedListBox>().FirstOrDefault();
            if (chkListContacts != null)
            {
                chkListContacts.Items.Clear();
                foreach (var contact in _contacts)
                {
                    chkListContacts.Items.Add(contact, true);
                }
            }
        }

        private void BtnOk_Click(object sender, EventArgs e)
        {
            var chkListContacts = Controls.OfType<CheckedListBox>().FirstOrDefault();
            if (chkListContacts != null)
            {
                SelectedContacts.Clear();
                for (int i = 0; i < chkListContacts.Items.Count; i++)
                {
                    if (chkListContacts.GetItemChecked(i))
                    {
                        SelectedContacts.Add(chkListContacts.Items[i].ToString());
                    }
                }
            }
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            var chkListContacts = Controls.OfType<CheckedListBox>().FirstOrDefault();
            if (chkListContacts != null)
            {
                bool anyUnchecked = false;
                for (int i = 0; i < chkListContacts.Items.Count; i++)
                {
                    if (!chkListContacts.GetItemChecked(i))
                    {
                        anyUnchecked = true;
                        break;
                    }
                }

                // Toggle all items
                for (int i = 0; i < chkListContacts.Items.Count; i++)
                {
                    chkListContacts.SetItemChecked(i, anyUnchecked);
                }
            }
        }
    }
} 