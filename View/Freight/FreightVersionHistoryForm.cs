using System;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;
using techlink_workspace.Repositories.FreightRepo;

namespace techlink_workspace.View.Freight
{
    /// <summary>
    /// Shows the full change log for one Forwarder row (or all rows).
    /// Selecting a log entry previews the snapshot and allows rollback.
    /// </summary>
    public partial class FreightVersionHistoryForm : Form
    {
        private readonly ForwarderQuotationRepository _repo;
        private readonly string _employeeCode;
        private readonly string _forwarderId;   // the DB row we're tracking
        private readonly string _fwdDisplayName;

        // ── Controls ─────────────────────────────────────────────────────
        private DataGridView dgvLog;
        private DataGridView dgvPreview;
        private Button btnRollback, btnClose;
        private Label lblTitle, lblPreviewTitle;
        private SplitContainer split;

        public FreightVersionHistoryForm(
            ForwarderQuotationRepository repo,
            string employeeCode,
            string forwarderId,
            string fwdDisplayName = null)
        {
            _repo = repo;
            _employeeCode = employeeCode;
            _forwarderId = forwarderId;
            _fwdDisplayName = fwdDisplayName ?? forwarderId ?? "(all records)";
            Build();
            LoadHistory();
        }

        // ════════════════════════════════════════════════════════════════
        // Build UI
        // ════════════════════════════════════════════════════════════════
        private void Build()
        {
            Text = "Version History – Freight Quotation";
            ClientSize = new Size(1100, 600);
            MinimumSize = new Size(800, 400);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9f);

            // Title bar
            lblTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 30,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.SlateBlue,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(10, 0, 0, 0)
            };

            // Bottom button panel
            var pnlBtn = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                BackColor = Color.WhiteSmoke
            };
            pnlBtn.Paint += (s, e) =>
            {
                using (var pen = new System.Drawing.Pen(Color.FromArgb(210, 210, 215)))
                    e.Graphics.DrawLine(pen, 0, 0, pnlBtn.Width, 0);
            };

            btnRollback = new Button
            {
                Text = "↩  Rollback to selected version",
                Size = new Size(220, 30),
                Location = new Point(10, 8),
                BackColor = Color.OrangeRed,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Enabled = false
            };
            btnRollback.FlatAppearance.BorderSize = 0;
            btnRollback.Click += BtnRollback_Click;

            btnClose = new Button
            {
                Text = "Close",
                Size = new Size(80, 30),
                Location = new Point(238, 8),
                BackColor = Color.FromArgb(200, 200, 200),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnClose.FlatAppearance.BorderSize = 0;
            pnlBtn.Controls.AddRange(new Control[] { btnRollback, btnClose });

            // SplitContainer: top = log grid, bottom = snapshot preview
            split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Horizontal,
                SplitterDistance = 280,
                Panel1MinSize = 120,
                Panel2MinSize = 80
            };

            // ── Log grid (top pane) ───────────────────────────────────────
            dgvLog = BuildGrid();
            dgvLog.SelectionChanged += DgvLog_SelectionChanged;
            split.Panel1.Controls.Add(dgvLog);

            // ── Snapshot preview (bottom pane) ────────────────────────────
            lblPreviewTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold | FontStyle.Italic),
                ForeColor = Color.DimGray,
                BackColor = Color.FromArgb(245, 245, 250),
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(8, 0, 0, 0),
                Text = "← Select a log entry above to preview the snapshot"
            };
            dgvPreview = BuildGrid();
            dgvPreview.SelectionMode = DataGridViewSelectionMode.CellSelect;
            split.Panel2.Controls.Add(dgvPreview);
            split.Panel2.Controls.Add(lblPreviewTitle);

            Controls.Add(split);
            Controls.Add(pnlBtn);
            Controls.Add(lblTitle);
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
                RowHeadersVisible = false,
                AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                { BackColor = Color.FromArgb(245, 248, 255) }
            };
            g.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 153, 0);
            g.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            g.ColumnHeadersHeight = 32;
            g.EnableHeadersVisualStyles = false;
            g.DefaultCellStyle.Font = new Font("Segoe UI", 8f);
            return g;
        }

        // ════════════════════════════════════════════════════════════════
        // Load
        // ════════════════════════════════════════════════════════════════
        private void LoadHistory()
        {
            lblTitle.Text = string.IsNullOrEmpty(_forwarderId)
                ? "  Version History — all records"
                : $"  Version History — {_fwdDisplayName}  (ID: {_forwarderId})";

            var dt = _repo.GetVersionHistory(_forwarderId);

            // Hide raw JSON columns from the log view — show them only in preview
            dgvLog.DataSource = dt;
            if (dgvLog.Columns.Contains("Log_Id"))
                dgvLog.Columns["Log_Id"].Visible = false;
            if (dgvLog.Columns.Contains("Log_OldData"))
                dgvLog.Columns["Log_OldData"].Visible = false;
            if (dgvLog.Columns.Contains("Log_UpdateData"))
                dgvLog.Columns["Log_UpdateData"].Visible = false;

            // Colour code by function
            dgvLog.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                var func = dgvLog.Rows[e.RowIndex].Cells["Log_Function"]?.Value?.ToString() ?? "";
                Color bg;

                switch (func)
                {
                    case "INSERT":
                        bg = Color.FromArgb(220, 255, 220);
                        break;

                    case "UPDATE":
                        bg = Color.FromArgb(255, 245, 210);
                        break;

                    case "DELETE":
                        bg = Color.FromArgb(255, 220, 220);
                        break;

                    case "ROLLBACK":
                        bg = Color.FromArgb(220, 220, 255);
                        break;

                    default:
                        bg = Color.White;
                        break;
                }
                if (e.CellStyle.BackColor != bg)
                    e.CellStyle.BackColor = bg;
            };
        }

        // ════════════════════════════════════════════════════════════════
        // Selection → preview snapshot
        // ════════════════════════════════════════════════════════════════
        private void DgvLog_SelectionChanged(object sender, EventArgs e)
        {
            btnRollback.Enabled = false;
            dgvPreview.DataSource = null;
            lblPreviewTitle.Text = "← Select a log entry above to preview the snapshot";

            if (dgvLog.SelectedRows.Count == 0) return;
            var row = dgvLog.SelectedRows[0];
            var dt = (DataTable)dgvLog.DataSource;
            var dr = dt.Rows[row.Index];

            string func = dr["Log_Function"]?.ToString() ?? "";
            string logId = dr["Log_Id"]?.ToString() ?? "";
            string oldJson = dr["Log_OldData"]?.ToString();
            string updJson = dr["Log_UpdateData"]?.ToString();

            // Decide which JSON to preview:
            // For UPDATE/DELETE show OldData (what it looked like BEFORE the change).
            // For INSERT show UpdateData (the new record).
            // For ROLLBACK show UpdateData (the description string).
            string jsonToShow = (func == "INSERT" || func == "ROLLBACK") ? updJson : oldJson;

            lblPreviewTitle.Text =
                $"Snapshot preview  [{func} by {dr["Log_EmployeeId"]}  at {dr["Log_WriteDate"]}]" +
                (func == "UPDATE" || func == "DELETE" ? "  — state BEFORE this action" :
                 func == "INSERT" ? "  — record as inserted" : "");

            if (!string.IsNullOrWhiteSpace(jsonToShow))
            {
                try
                {
                    var m = JsonConvert.DeserializeObject<ForwarderQuotationModel>(jsonToShow);
                    if (m != null)
                    {
                        var preview = new System.Data.DataTable();
                        preview.Columns.Add("Field");
                        preview.Columns.Add("Value");
                        foreach (var prop in typeof(ForwarderQuotationModel).GetProperties())
                            preview.Rows.Add(prop.Name, prop.GetValue(m)?.ToString() ?? "");
                        dgvPreview.DataSource = preview;
                        dgvPreview.Columns["Field"].DefaultCellStyle.Font =
                            new Font("Segoe UI", 8, FontStyle.Bold);
                        dgvPreview.Columns["Field"].DefaultCellStyle.BackColor =
                            Color.FromArgb(240, 240, 250);
                    }
                }
                catch
                {
                    // If JSON is a plain description (e.g. ROLLBACK note), show as-is
                    var preview = new System.Data.DataTable();
                    preview.Columns.Add("Info");
                    preview.Rows.Add(jsonToShow);
                    dgvPreview.DataSource = preview;
                }
            }
            else
            {
                lblPreviewTitle.Text += "  (no snapshot data available)";
            }

            // Only allow rollback when there is an OldData snapshot to restore
            btnRollback.Enabled = !string.IsNullOrWhiteSpace(oldJson)
                                  && (func == "UPDATE" || func == "DELETE");
        }

        // ════════════════════════════════════════════════════════════════
        // Rollback
        // ════════════════════════════════════════════════════════════════
        private void BtnRollback_Click(object sender, EventArgs e)
        {
            if (dgvLog.SelectedRows.Count == 0) return;

            var dt = (DataTable)dgvLog.DataSource;
            var dr = dt.Rows[dgvLog.SelectedRows[0].Index];
            string logId = dr["Log_Id"]?.ToString();
            string func = dr["Log_Function"]?.ToString();
            string writeDate = dr["Log_WriteDate"]?.ToString();
            string varId = dr["Log_Variable"]?.ToString();

            var confirm = CTMessageBox.Show(
                $"Restore record [{_fwdDisplayName}] to the state\n" +
                $"BEFORE the '{func}' action on {writeDate}?\n\n" +
                "A new log entry will be created for this rollback.",
                "Confirm Rollback", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            try
            {
                _repo.Rollback(logId, _employeeCode);
                CTMessageBox.Show("Rollback successful.", "Done",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                CTMessageBox.Show("Rollback failed:\r\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}