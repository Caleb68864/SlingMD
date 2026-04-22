using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SlingMD.Outlook.Forms
{
    /// <summary>
    /// Searchable settings reference. Populated entirely from <see cref="SettingsHelp"/> —
    /// no content lives in this form. Shown modelessly from the Settings form's Help button.
    /// </summary>
    public class HelpForm : Form
    {
        private TextBox _searchBox;
        private TreeView _tree;
        private RichTextBox _detail;
        private SplitContainer _split;
        private Label _matchCount;

        public HelpForm()
        {
            InitializeComponent();
            PopulateTree(SettingsHelp.All());
            SelectFirstLeaf();
        }

        /// <summary>
        /// Opens the help form focused on a specific entry id.
        /// </summary>
        public void ShowEntry(string entryId)
        {
            SelectEntry(entryId);
            if (!Visible)
            {
                Show();
            }
            else
            {
                Activate();
            }
        }

        private void InitializeComponent()
        {
            this.Text = "SlingMD Settings Help";
            this.Size = new Size(900, 600);
            this.StartPosition = FormStartPosition.CenterParent;
            this.MinimumSize = new Size(600, 400);

            TableLayoutPanel root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(8)
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            // Search bar
            TableLayoutPanel searchBar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 6)
            };
            searchBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            searchBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            searchBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            Label searchLabel = new Label
            {
                Text = "Search:",
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 6, 6, 0)
            };
            _searchBox = new TextBox
            {
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Dock = DockStyle.Fill
            };
            _searchBox.TextChanged += SearchBox_TextChanged;
            _searchBox.KeyDown += SearchBox_KeyDown;
            _matchCount = new Label
            {
                AutoSize = true,
                Anchor = AnchorStyles.Right,
                TextAlign = ContentAlignment.MiddleRight,
                Margin = new Padding(6, 6, 0, 0),
                ForeColor = SystemColors.GrayText
            };
            searchBar.Controls.Add(searchLabel, 0, 0);
            searchBar.Controls.Add(_searchBox, 1, 0);
            searchBar.Controls.Add(_matchCount, 2, 0);
            root.Controls.Add(searchBar, 0, 0);

            // Tree + detail split
            _split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                FixedPanel = FixedPanel.Panel1
            };

            _tree = new TreeView
            {
                Dock = DockStyle.Fill,
                HideSelection = false
            };
            _tree.AfterSelect += Tree_AfterSelect;
            _split.Panel1.Controls.Add(_tree);

            _detail = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = SystemColors.Window,
                Font = new Font("Segoe UI", 9.5F)
            };
            _split.Panel2.Controls.Add(_detail);

            root.Controls.Add(_split, 0, 1);

            // Footer
            FlowLayoutPanel footer = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                FlowDirection = FlowDirection.RightToLeft,
                Margin = new Padding(0, 6, 0, 0)
            };
            Button btnClose = new Button { Text = "Close" };
            btnClose.Click += (s, e) => Close();
            footer.Controls.Add(btnClose);
            root.Controls.Add(footer, 0, 2);

            this.Controls.Add(root);
            this.AcceptButton = btnClose;
            this.CancelButton = btnClose;

            // Split distance after form is laid out
            this.Load += (s, e) => { _split.SplitterDistance = (int)(_split.Width * 0.35); };
        }

        private void PopulateTree(IEnumerable<HelpEntry> entries)
        {
            _tree.BeginUpdate();
            _tree.Nodes.Clear();

            // Group by Tab, preserving tab ordering from the registry
            List<HelpEntry> list = entries.ToList();
            foreach (string tab in list.Select(e => e.Tab).Distinct())
            {
                TreeNode tabNode = new TreeNode(tab) { Tag = null };
                foreach (HelpEntry entry in list.Where(e => e.Tab == tab))
                {
                    TreeNode leaf = new TreeNode(entry.Title) { Tag = entry };
                    tabNode.Nodes.Add(leaf);
                }
                _tree.Nodes.Add(tabNode);
                tabNode.Expand();
            }

            _tree.EndUpdate();

            int total = list.Count;
            _matchCount.Text = total + " setting" + (total == 1 ? string.Empty : "s");
        }

        private void SelectFirstLeaf()
        {
            foreach (TreeNode tab in _tree.Nodes)
            {
                if (tab.Nodes.Count > 0)
                {
                    _tree.SelectedNode = tab.Nodes[0];
                    return;
                }
            }
            _detail.Clear();
        }

        private void SelectEntry(string entryId)
        {
            if (string.IsNullOrEmpty(entryId)) return;
            foreach (TreeNode tab in _tree.Nodes)
            {
                foreach (TreeNode leaf in tab.Nodes)
                {
                    HelpEntry entry = leaf.Tag as HelpEntry;
                    if (entry != null && entry.Id == entryId)
                    {
                        _tree.SelectedNode = leaf;
                        leaf.EnsureVisible();
                        return;
                    }
                }
            }
        }

        private void SearchBox_TextChanged(object sender, EventArgs e)
        {
            string query = _searchBox.Text;
            IEnumerable<HelpEntry> results = SettingsHelp.Search(query);
            PopulateTree(results);
            if (!string.IsNullOrWhiteSpace(query))
            {
                // Expand all and select the first match
                foreach (TreeNode tab in _tree.Nodes) tab.ExpandAll();
                SelectFirstLeaf();
            }
        }

        private void SearchBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                _searchBox.Clear();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Down)
            {
                _tree.Focus();
                e.Handled = true;
            }
        }

        private void Tree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            HelpEntry entry = e.Node?.Tag as HelpEntry;
            RenderDetail(entry);
        }

        private void RenderDetail(HelpEntry entry)
        {
            _detail.Clear();
            if (entry == null)
            {
                return;
            }

            // Title
            _detail.SelectionFont = new Font("Segoe UI", 14F, FontStyle.Bold);
            _detail.AppendText(entry.Title + Environment.NewLine);
            _detail.SelectionFont = new Font("Segoe UI", 9F, FontStyle.Italic);
            _detail.SelectionColor = SystemColors.GrayText;
            _detail.AppendText(entry.Tab + " tab" + Environment.NewLine + Environment.NewLine);

            _detail.SelectionFont = new Font("Segoe UI", 10F, FontStyle.Regular);
            _detail.SelectionColor = SystemColors.ControlText;

            if (!string.IsNullOrWhiteSpace(entry.Summary))
            {
                _detail.AppendText(entry.Summary + Environment.NewLine + Environment.NewLine);
            }
            if (!string.IsNullOrWhiteSpace(entry.Description))
            {
                _detail.AppendText(entry.Description + Environment.NewLine + Environment.NewLine);
            }

            if (!string.IsNullOrWhiteSpace(entry.Default))
            {
                _detail.SelectionFont = new Font("Segoe UI", 10F, FontStyle.Bold);
                _detail.AppendText("Default: ");
                _detail.SelectionFont = new Font("Consolas", 10F);
                _detail.AppendText(entry.Default + Environment.NewLine + Environment.NewLine);
                _detail.SelectionFont = new Font("Segoe UI", 10F, FontStyle.Regular);
            }

            if (entry.Tokens != null && entry.Tokens.Count > 0)
            {
                _detail.SelectionFont = new Font("Segoe UI", 10F, FontStyle.Bold);
                _detail.AppendText("Tokens" + Environment.NewLine);
                int pad = entry.Tokens.Keys.Max(k => k.Length);
                _detail.SelectionFont = new Font("Consolas", 10F);
                foreach (KeyValuePair<string, string> token in entry.Tokens)
                {
                    _detail.AppendText("  " + token.Key.PadRight(pad) + "   " + token.Value + Environment.NewLine);
                }
                _detail.AppendText(Environment.NewLine);
            }

            if (entry.Examples != null && entry.Examples.Count > 0)
            {
                _detail.SelectionFont = new Font("Segoe UI", 10F, FontStyle.Bold);
                _detail.AppendText("Examples" + Environment.NewLine);
                int pad = entry.Examples.Max(ex => (ex.Input ?? string.Empty).Length);
                _detail.SelectionFont = new Font("Consolas", 10F);
                foreach (HelpExample ex in entry.Examples)
                {
                    string input = ex.Input ?? string.Empty;
                    _detail.AppendText("  " + input.PadRight(pad) + "  →  " + (ex.Output ?? string.Empty) + Environment.NewLine);
                }
            }

            _detail.SelectionStart = 0;
            _detail.ScrollToCaret();
        }
    }
}
