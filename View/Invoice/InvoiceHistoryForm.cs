using Newtonsoft.Json;
using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;
using techlink_workspace.Repositories.InvoiceRepo;

namespace techlink_workspace.View.Invoice
{

    public class InvoiceHistoryForm : Form
    {
        private readonly InvoiceRepository _repo;
        private readonly string _byUser, _invoiceId, _dispName;

        private DataGridView dgvLog, dgvPreview;
        private Button btnRollback, btnClose;
        private Label lblTitle, lblPreviewHdr;

        public InvoiceHistoryForm(InvoiceRepository repo, string byUser,
                                   string invoiceId, string dispName)
        {
            _repo = repo; _byUser = byUser;
            _invoiceId = invoiceId; _dispName = dispName ?? invoiceId ?? "(all)";
            Build(); Load2();
        }

        private void Build()
        {
            Text = "Version History – Invoice"; ClientSize = new Size(1100, 600);
            StartPosition = FormStartPosition.CenterParent; BackColor = Color.White;
            Font = new Font("Segoe UI", 9f);

            lblTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 30,
                BackColor = Color.SlateBlue,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };

            var bot = new Panel { Dock = DockStyle.Bottom, Height = 46, BackColor = Color.WhiteSmoke };
            btnRollback = new Button
            {
                Text = "↩ Rollback to selected",
                Size = new Size(200, 30),
                Location = new Point(10, 8),
                BackColor = Color.OrangeRed,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Enabled = false
            };
            btnRollback.FlatAppearance.BorderSize = 0;
            btnRollback.Click += BtnRollback_Click;
            btnClose = new Button
            {
                Text = "Close",
                Size = new Size(80, 30),
                Location = new Point(218, 8),
                BackColor = Color.FromArgb(200, 200, 200),
                FlatStyle = FlatStyle.Flat,
                DialogResult = DialogResult.Cancel
            };
            btnClose.FlatAppearance.BorderSize = 0;
            bot.Controls.AddRange(new Control[] { btnRollback, btnClose });

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 280
            };

            dgvLog = BuildGrid(); dgvLog.SelectionChanged += DgvLog_SelectionChanged;
            split.Panel1.Controls.Add(dgvLog);

            lblPreviewHdr = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                BackColor = Color.FromArgb(245, 245, 250),
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Italic),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Text = "← Select a log entry to preview"
            };
            dgvPreview = BuildGrid();
            split.Panel2.Controls.Add(dgvPreview);
            split.Panel2.Controls.Add(lblPreviewHdr);

            Controls.Add(split); Controls.Add(bot); Controls.Add(lblTitle);
        }

        private DataGridView BuildGrid()
        {
            var g = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                BackgroundColor = Color.White,
                RowHeadersVisible = false
            };
            g.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 153, 0);
            g.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            g.ColumnHeadersHeight = 32; g.EnableHeadersVisualStyles = false;
            return g;
        }

        private void Load2()
        {
            lblTitle.Text = $"  Version History — {_dispName}";
            var dt = _repo.GetHistory(_invoiceId);
            dgvLog.DataSource = dt;
            if (dgvLog.Columns.Contains("Log_Id")) dgvLog.Columns["Log_Id"].Visible = false;
            if (dgvLog.Columns.Contains("Log_OldData")) dgvLog.Columns["Log_OldData"].Visible = false;
            if (dgvLog.Columns.Contains("Log_UpdateData")) dgvLog.Columns["Log_UpdateData"].Visible = false;
            dgvLog.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                var func = dgvLog.Rows[e.RowIndex].Cells["Log_Function"]?.Value?.ToString() ?? "";
                Color bg = func == "INSERT" ? Color.FromArgb(220, 255, 220)
                         : func == "UPDATE" ? Color.FromArgb(255, 245, 210)
                         : func == "DELETE" ? Color.FromArgb(255, 220, 220)
                         : func == "ROLLBACK" ? Color.FromArgb(220, 220, 255)
                         : Color.White;
                if (e.CellStyle.BackColor != bg) e.CellStyle.BackColor = bg;
            };
        }

        private void DgvLog_SelectionChanged(object sender, EventArgs e)
        {
            btnRollback.Enabled = false; dgvPreview.DataSource = null;
            if (dgvLog.SelectedRows.Count == 0) return;
            var dt = (DataTable)dgvLog.DataSource;
            var dr = dt.Rows[dgvLog.SelectedRows[0].Index];
            string func = dr["Log_Function"]?.ToString() ?? "";
            string oldJson = dr["Log_OldData"]?.ToString();
            string updJson = dr["Log_UpdateData"]?.ToString();
            string json = (func == "INSERT" || func == "ROLLBACK") ? updJson : oldJson;
            lblPreviewHdr.Text = $"Preview [{func} by {dr["Log_EmployeeId"]} at {dr["Log_WriteDate"]}]";
            if (!string.IsNullOrWhiteSpace(json))
            {
                try
                {
                    var m = JsonConvert.DeserializeObject<InvoiceModel>(json);
                    if (m != null)
                    {
                        var prev = new DataTable();
                        prev.Columns.Add("Field"); prev.Columns.Add("Value");
                        foreach (var p in typeof(InvoiceModel).GetProperties())
                            prev.Rows.Add(p.Name, p.GetValue(m)?.ToString() ?? "");
                        dgvPreview.DataSource = prev;
                        if (dgvPreview.Columns.Contains("Field"))
                            dgvPreview.Columns["Field"].DefaultCellStyle.Font =
                                new Font("Segoe UI", 8, FontStyle.Bold);
                    }
                }
                catch
                {
                    var prev = new DataTable(); prev.Columns.Add("Info");
                    prev.Rows.Add(json); dgvPreview.DataSource = prev;
                }
            }
            btnRollback.Enabled = !string.IsNullOrWhiteSpace(oldJson)
                                 && (func == "UPDATE" || func == "DELETE");
        }

        private void BtnRollback_Click(object sender, EventArgs e)
        {
            var dt = (DataTable)dgvLog.DataSource;
            var dr = dt.Rows[dgvLog.SelectedRows[0].Index];
            if (CTMessageBox.Show($"Rollback [{_dispName}] to state before '{dr["Log_Function"]}' on {dr["Log_WriteDate"]}?",
                "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            try
            {
                _repo.Rollback(dr["Log_Id"].ToString(), _byUser);
                CTMessageBox.Show("Rollback successful.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK; Close();
            }
            catch (Exception ex)
            {
                CTMessageBox.Show("Rollback failed:\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}