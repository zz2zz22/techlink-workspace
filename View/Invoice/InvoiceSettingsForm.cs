using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;
using techlink_workspace.Repositories.PICRepo;

namespace techlink_workspace.View.Invoice
{
    /// <summary>
    /// PIC Management settings.
    /// Left: DataGridView of all PIC entries.
    /// Right: Edit panel (PIC Code, PIC Name, RichTextBox of customer codes – one per line).
    /// Saved as semicolon-separated into PIC_CustomerName.
    /// </summary>
    public class InvoiceSettingsForm : Form
    {
        private readonly PICRepository _repo = new PICRepository();
        private readonly string _byUser;

        private DataGridView dgv;
        private TextBox txtCode, txtName;
        private RichTextBox rtxCustomers;
        private Button btnNew, btnSave, btnDelete, btnClose;
        private Label lblStatus;
        private bool _isNew;

        public InvoiceSettingsForm(string byUser)
        {
            _byUser = byUser;
            Build();
            LoadGrid();
        }

        private void Build()
        {
            Text = "PIC Management Settings";
            ClientSize = new Size(860, 540);
            MinimumSize = new Size(700, 400);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9f);

            // ── Title bar ─────────────────────────────────────────────────
            var hdr = new Label
            {
                Dock = DockStyle.Top,
                Height = 32,
                Text = "  Person-in-Charge (PIC) ↔ Customer Code Mapping",
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                BackColor = Color.FromArgb(30, 40, 60),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // ── Status bar ────────────────────────────────────────────────
            lblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 22,
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                BackColor = Color.LightYellow,
                ForeColor = Color.DarkSlateGray,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(6, 0, 0, 0),
                BorderStyle = BorderStyle.FixedSingle
            };

            // ── Split ─────────────────────────────────────────────────────
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 340
            };

            // Left: grid
            dgv = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BackgroundColor = Color.White,
                RowHeadersVisible = false
            };
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 40, 60);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8.5f, FontStyle.Bold);
            dgv.EnableHeadersVisualStyles = false;
            dgv.SelectionChanged += DgvSelectionChanged;
            split.Panel1.Controls.Add(dgv);

            // Right: edit panel
            var editPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(12, 8, 12, 0),
                BackColor = Color.White
            };

            int y = 0;
            editPanel.Controls.Add(Lbl("PIC Employee Code:", 0, y)); y += 20;
            txtCode = new TextBox { Location = new Point(0, y), Width = 300, MaxLength = 50 };
            editPanel.Controls.Add(txtCode); y += 30;

            editPanel.Controls.Add(Lbl("PIC Full Name:", 0, y)); y += 20;
            txtName = new TextBox { Location = new Point(0, y), Width = 300, MaxLength = 50 };
            editPanel.Controls.Add(txtName); y += 30;

            editPanel.Controls.Add(Lbl("Customer Codes (one per line):", 0, y)); y += 20;
            var hint = new Label
            {
                Text = "Enter one customer code per line. Will be saved as ';' separated.",
                Location = new Point(0, y),
                AutoSize = true,
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 7.5f)
            };
            editPanel.Controls.Add(hint); y += 18;
            rtxCustomers = new RichTextBox
            {
                Location = new Point(0, y),
                Size = new Size(300, 160),
                ScrollBars = RichTextBoxScrollBars.Vertical,
                Font = new Font("Consolas", 9f)
            };
            editPanel.Controls.Add(rtxCustomers); y += 168;

            btnSave = Btn("💾 Save PIC", Color.SeaGreen, 0, y);
            btnNew = Btn("+ New", Color.SteelBlue, 110, y);
            btnDelete = Btn("✕ Delete", Color.IndianRed, 190, y);
            editPanel.Controls.AddRange(new Control[] { btnSave, btnNew, btnDelete });

            btnSave.Click += BtnSave_Click;
            btnNew.Click += (s, e) => ClearForm(isNew: true);
            btnDelete.Click += BtnDelete_Click;

            split.Panel2.Controls.Add(editPanel);

            // Close button at bottom right
            var botPanel = new Panel
            { Dock = DockStyle.Bottom, Height = 46, BackColor = Color.WhiteSmoke };
            btnClose = new Button
            {
                Text = "Close",
                Size = new Size(90, 30),
                Location = new Point(12, 8),
                BackColor = Color.FromArgb(200, 200, 200),
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.OK
            };
            btnClose.FlatAppearance.BorderSize = 0;
            botPanel.Controls.Add(btnClose);

            Controls.Add(split);
            Controls.Add(botPanel);
            Controls.Add(lblStatus);
            Controls.Add(hdr);
        }

        private void LoadGrid()
        {
            dgv.DataSource = null;
            var list = _repo.GetAll();
            var dt = new System.Data.DataTable();
            dt.Columns.Add("PIC_Code", typeof(string));
            dt.Columns.Add("PIC_Name", typeof(string));
            dt.Columns.Add("# Customers", typeof(int));
            foreach (var p in list)
                dt.Rows.Add(p.PIC_Code, p.PIC_Name,
                    string.IsNullOrWhiteSpace(p.PIC_CustomerName)
                        ? 0
                        : p.PIC_CustomerName.Split(';').Count(x => !string.IsNullOrWhiteSpace(x)));
            dgv.DataSource = dt;
            SetStatus($"{list.Count} PIC(s) loaded.");
        }

        private void DgvSelectionChanged(object sender, EventArgs e)
        {
            if (dgv.SelectedRows.Count == 0) return;
            string code = dgv.SelectedRows[0].Cells["PIC_Code"].Value?.ToString();
            if (string.IsNullOrEmpty(code)) return;
            var pic = _repo.GetByCode(code);
            if (pic == null) return;
            _isNew = false;
            txtCode.Text = pic.PIC_Code;
            txtCode.ReadOnly = true;
            txtName.Text = pic.PIC_Name ?? "";
            rtxCustomers.Text = string.IsNullOrWhiteSpace(pic.PIC_CustomerName)
                ? ""
                : string.Join("\n", pic.PIC_CustomerName.Split(';')
                                        .Select(x => x.Trim())
                                        .Where(x => !string.IsNullOrEmpty(x)));
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCode.Text))
            { SetStatus("PIC Code is required.", isError: true); return; }

            // Build semicolon-separated list from RichTextBox lines
            var codes = rtxCustomers.Lines
                .Select(l => l.Trim())
                .Where(l => !string.IsNullOrEmpty(l))
                .ToList();

            var m = new PICModel
            {
                PIC_Code = txtCode.Text.Trim(),
                PIC_Name = txtName.Text.Trim(),
                PIC_CustomerName = string.Join(";", codes)
            };

            try
            {
                _repo.Upsert(m, _byUser);
                LoadGrid();
                SetStatus("Saved successfully.");
            }
            catch (Exception ex)
            {
                CTMessageBox.Show("Save error:\r\n" + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtCode.Text)) return;
            if (CTMessageBox.Show($"Delete PIC '{txtCode.Text}'?", "Confirm",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            try
            {
                _repo.Delete(txtCode.Text.Trim());
                ClearForm(isNew: true);
                LoadGrid();
                SetStatus("Deleted.");
            }
            catch (Exception ex)
            {
                CTMessageBox.Show("Delete error:\r\n" + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearForm(bool isNew)
        {
            _isNew = isNew;
            txtCode.Text = "";
            txtCode.ReadOnly = false;
            txtName.Text = "";
            rtxCustomers.Text = "";
        }

        private void SetStatus(string msg, bool isError = false)
        { lblStatus.Text = msg; lblStatus.ForeColor = isError ? Color.Red : Color.DarkGreen; }

        private static Label Lbl(string t, int x, int y) =>
            new Label
            {
                Text = t,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = Color.FromArgb(40, 40, 80),
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
            };

        private static Button Btn(string t, Color bg, int x, int y)
        {
            var b = new Button
            {
                Text = t,
                Size = new Size(90, 28),
                Location = new Point(x, y),
                BackColor = bg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 8.5f)
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }
    }
}