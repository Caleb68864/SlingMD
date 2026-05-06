namespace SlingMD.Outlook.Forms
{
    partial class ContactMatchPromptForm
    {
        private System.ComponentModel.IContainer components = null;

        private System.Windows.Forms.Label lblPrompt;
        private System.Windows.Forms.ListBox lstCandidates;
        private System.Windows.Forms.CheckBox chkSaveAsAlias;
        private System.Windows.Forms.Button btnMatch;
        private System.Windows.Forms.Button btnCreateNew;
        private System.Windows.Forms.ToolTip toolTip;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
                components.Dispose();
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.lblPrompt = new System.Windows.Forms.Label();
            this.lstCandidates = new System.Windows.Forms.ListBox();
            this.chkSaveAsAlias = new System.Windows.Forms.CheckBox();
            this.btnMatch = new System.Windows.Forms.Button();
            this.btnCreateNew = new System.Windows.Forms.Button();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();

            // lblPrompt
            this.lblPrompt.AutoSize = false;
            this.lblPrompt.Location = new System.Drawing.Point(12, 12);
            this.lblPrompt.Size = new System.Drawing.Size(560, 32);
            this.lblPrompt.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular);
            this.lblPrompt.Text = "Possible existing contacts found. Select a match or create a new contact note:";

            // lstCandidates
            this.lstCandidates.Location = new System.Drawing.Point(12, 52);
            this.lstCandidates.Size = new System.Drawing.Size(560, 130);
            this.lstCandidates.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.lstCandidates.IntegralHeight = false;
            this.lstCandidates.SelectionMode = System.Windows.Forms.SelectionMode.One;
            this.lstCandidates.TabIndex = 0;

            // chkSaveAsAlias
            this.chkSaveAsAlias.AutoSize = true;
            this.chkSaveAsAlias.Location = new System.Drawing.Point(12, 194);
            this.chkSaveAsAlias.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.chkSaveAsAlias.Text = "Save sender name as alias on the matched contact";
            this.chkSaveAsAlias.Checked = _defaultSaveAsAlias;
            this.chkSaveAsAlias.TabIndex = 1;

            // btnMatch
            this.btnMatch.Location = new System.Drawing.Point(12, 226);
            this.btnMatch.Size = new System.Drawing.Size(160, 30);
            this.btnMatch.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnMatch.Text = "Match this contact";
            this.btnMatch.TabIndex = 2;
            this.btnMatch.Click += new System.EventHandler(this.btnMatch_Click);

            // btnCreateNew
            this.btnCreateNew.Location = new System.Drawing.Point(412, 226);
            this.btnCreateNew.Size = new System.Drawing.Size(160, 30);
            this.btnCreateNew.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.btnCreateNew.Text = "Create new contact";
            this.btnCreateNew.TabIndex = 3;
            this.btnCreateNew.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCreateNew.Click += new System.EventHandler(this.btnCreateNew_Click);
            this.toolTip.SetToolTip(this.btnCreateNew,
                "If you’re not sure, choosing this creates a new contact note. You can merge later by adding aliases manually.");

            // ContactMatchPromptForm
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCreateNew;
            this.ClientSize = new System.Drawing.Size(584, 271);
            this.Controls.Add(this.lblPrompt);
            this.Controls.Add(this.lstCandidates);
            this.Controls.Add(this.chkSaveAsAlias);
            this.Controls.Add(this.btnMatch);
            this.Controls.Add(this.btnCreateNew);
            this.Font = new System.Drawing.Font("Segoe UI", 9F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ContactMatchPromptForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Possible Contact Match Found";

            this.ResumeLayout(false);
            this.PerformLayout();
        }
    }
}
