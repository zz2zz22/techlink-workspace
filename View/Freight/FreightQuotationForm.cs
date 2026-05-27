using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;
using techlink_workspace.Repositories.FreightRepo;

namespace techlink_workspace.View.Freight
{
    public partial class FreightQuotationForm : Form
    {
        // ── State ────────────────────────────────────────────────────────
        private readonly UserModel _user;
        private readonly ForwarderQuotationRepository _repo = new ForwarderQuotationRepository();
        private DataTable _fullData;

        private bool IsAdmin => _user.User_permissionLevel == 1 || _user.User_permissionLevel == 2;
        private bool IsType1 => _user.User_permissionLevel == 1;

        // ── Toolbar row 1 ────────────────────────────────────────────────
        private Button btnUpload, btnEdit, btnHistory, btnExportCSV, btnExportExcel, btnRefresh;

        // ── Toolbar row 2 – search/filter ────────────────────────────────
        private TextBox txtKeyword;
        private ComboBox cmbSearchCol, cmbGroupCol, cmbGroupVal, cmbFwd, cmbContainer;
        private DateTimePicker dtpFrom, dtpTo;
        private CheckBox chkDateFilter;
        private Button btnSearch, btnClearSearch;

        // ── Grid & status ────────────────────────────────────────────────
        private DataGridView dgv;
        private Label lblStatus;

        // ════════════════════════════════════════════════════════════════
        public FreightQuotationForm(UserModel user)
        {
            _user = user;
            BuildUI();
            LoadGrid();
        }

        // ════════════════════════════════════════════════════════════════
        // UI
        // ════════════════════════════════════════════════════════════════
        private void BuildUI()
        {
            Text = "Freight Quotation";
            ClientSize = new Size(1440, 740);
            MinimumSize = new Size(1100, 540);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 8.5f);

            // ── Row 1 ─────────────────────────────────────────────────────
            var row1 = new Panel { Dock = DockStyle.Top, Height = 44, BackColor = Color.White };
            row1.Paint += PaintBorder;
            int x = 8;
            btnUpload = TB(row1, "⬆ Upload Excel", Color.FromArgb(0, 51, 153), ref x);
            btnEdit = TB(row1, "✎ Edit", Color.DarkOrange, ref x);
            x += 10;
            btnHistory = TB(row1, "⟳ History", Color.SlateBlue, ref x);
            x += 10;
            btnExportCSV = TB(row1, "⬇ CSV", Color.DarkGreen, ref x);
            btnExportExcel = TB(row1, "⬇ Excel", Color.DarkGreen, ref x);
            x += 10;
            btnRefresh = TB(row1, "↺ Refresh", Color.FromArgb(90, 90, 90), ref x);

            // ── Row 2 – filters ──────────────────────────────────────────
            var row2 = new Panel { Dock = DockStyle.Top, Height = 44, BackColor = Color.FromArgb(248, 248, 252) };
            row2.Paint += PaintBorder;
            int x2 = 8;

            row2.Controls.Add(Lbl("FWD:", x2, 14)); x2 += 36;
            cmbFwd = Cmb(row2, 110, ref x2);

            row2.Controls.Add(Lbl("Container:", x2, 14)); x2 += 68;
            cmbContainer = Cmb(row2, 90, ref x2);

            row2.Controls.Add(Lbl("Group:", x2, 14)); x2 += 48;
            cmbGroupCol = Cmb(row2, 120, ref x2);
            cmbGroupVal = Cmb(row2, 130, ref x2);

            row2.Controls.Add(Lbl("Search in:", x2, 14)); x2 += 64;
            cmbSearchCol = Cmb(row2, 130, ref x2);
            txtKeyword = new TextBox
            {
                Location = new Point(x2, 11),
                Width = 150,
                Font = new Font("Segoe UI", 8.5f)
            };
            row2.Controls.Add(txtKeyword); x2 += 156;

            chkDateFilter = new CheckBox
            {
                Text = "Date:",
                Location = new Point(x2, 13),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f)
            };
            row2.Controls.Add(chkDateFilter); x2 += 52;
            dtpFrom = DTP(row2, ref x2);
            row2.Controls.Add(Lbl("–", x2, 14)); x2 += 14;
            dtpTo = DTP(row2, ref x2); x2 += 6;

            btnSearch = TB(row2, "🔍 Search", Color.SteelBlue, ref x2, h: 26, top: 9);
            btnClearSearch = TB(row2, "✕ Clear", Color.DimGray, ref x2, h: 26, top: 9);

            // ── Status & grid ────────────────────────────────────────────
            lblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 22,
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                ForeColor = Color.DarkSlateGray,
                BackColor = Color.LightYellow,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(6, 0, 0, 0),
                BorderStyle = BorderStyle.FixedSingle
            };

            dgv = new DataGridView
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
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 153, 0);
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            dgv.ColumnHeadersHeight = 36;
            dgv.EnableHeadersVisualStyles = false;
            dgv.DefaultCellStyle.Font = new Font("Segoe UI", 8f);
            dgv.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0 && IsAdmin) OpenEditDialog(); };

            Controls.Add(dgv);
            Controls.Add(lblStatus);
            Controls.Add(row2);
            Controls.Add(row1);

            // ── Permissions ──────────────────────────────────────────────
            btnUpload.Enabled = IsAdmin;
            btnEdit.Enabled = IsAdmin;
            btnHistory.Enabled = IsAdmin;
            btnExportCSV.Enabled = IsType1;
            btnExportExcel.Enabled = IsType1;

            // ── Events ───────────────────────────────────────────────────
            btnUpload.Click += (s, e) => UploadExcel();
            btnEdit.Click += (s, e) => OpenEditDialog();
            btnHistory.Click += (s, e) => OpenHistory();
            btnExportCSV.Click += (s, e) => ExportCSV();
            btnExportExcel.Click += (s, e) => ExportExcel();
            btnRefresh.Click += (s, e) => LoadGrid();
            btnSearch.Click += (s, e) => ApplyFilters();
            btnClearSearch.Click += (s, e) => ClearFilters();
            txtKeyword.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) ApplyFilters(); };
            cmbGroupCol.SelectedIndexChanged += CmbGroupCol_Changed;
            cmbFwd.SelectedIndexChanged += (s, e) => ApplyFilters();
            cmbContainer.SelectedIndexChanged += (s, e) => ApplyFilters();
            cmbGroupVal.SelectedIndexChanged += (s, e) => ApplyFilters();
        }

        // ════════════════════════════════════════════════════════════════
        // Data load
        // ════════════════════════════════════════════════════════════════
        private void LoadGrid()
        {
            try
            {
                dgv.DataSource = null;
                _fullData = _repo.GetAll();
                PopulateFilterCombos();
                BindGrid(_fullData);
                SetStatus($"{_fullData.Rows.Count} records loaded.", Color.DarkSlateGray);
            }
            catch (Exception ex)
            {
                SetStatus("Load error: " + ex.Message, Color.Red);
                CTMessageBox.Show("Failed to load data:\r\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Bind grid — NO column renaming to avoid key mismatch ─────────
        private void BindGrid(DataTable dt)
        {
            var view = dt.Copy();
            dgv.DataSource = view;

            // Hide internal ID column
            if (dgv.Columns.Contains("Forwarder_ID"))
                dgv.Columns["Forwarder_ID"].Visible = false;

            // Friendly headers — set HeaderText only, do NOT use column name as key after this
            var headers = new Dictionary<string, string>
            {
                {"Forwarder_name",         "FWD"},
                {"Forwarder_portDelivery", "Port/Delivery"},
                {"Forwarder_term",         "Term"},
                {"Forwarder_container",    "Container"},
                {"Forwarder_commodity",    "Commodity"},
                {"Forwarder_hsCode",       "HS Code"},
                {"Forwarder_carrier",      "Carrier"},
                {"Forwarder_total",        "Total USD"},
                {"Forwarder_of",           "OF"},
                {"Forwarder_localPol",     "Local POL"},
                {"Forwarder_destCharge",   "Dest Chg"},
                {"Forwarder_delivery",     "Delivery"},
                {"Forwarder_otherCharge",  "Other Chg"},
                {"Forwarder_remark",       "Remark"},
                {"Forwarder_volumn",       "Vol/Month"},
                {"Forwarder_validDate",    "Valid Date"},
                {"create_date",            "Created"},
                {"create_by",              "Created By"},
                {"update_date",            "Updated"},
                {"update_by",              "Updated By"},
            };
            foreach (var kv in headers)
                if (dgv.Columns.Contains(kv.Key))           // use DB col name as key
                    dgv.Columns[kv.Key].HeaderText = kv.Value;

            // Highlight — always reference by DB column name, not HeaderText
            foreach (string col in new[] { "Forwarder_total", "Forwarder_of" })
                if (dgv.Columns.Contains(col))
                {
                    dgv.Columns[col].DefaultCellStyle.BackColor = Color.LightYellow;
                    dgv.Columns[col].DefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
                }

            if (dgv.Columns.Contains("Forwarder_name"))
            {
                // Move FWD to first visible position before freezing
                dgv.Columns["Forwarder_name"].DisplayIndex = dgv.Columns["Forwarder_ID"].Visible
                    ? 1 : 0;
                dgv.Columns["Forwarder_name"].Frozen = true;
                dgv.Columns["Forwarder_name"].DefaultCellStyle.BackColor = Color.LightBlue;
                dgv.Columns["Forwarder_name"].DefaultCellStyle.Font =
                    new Font("Segoe UI", 8, FontStyle.Bold);
            }
        }

        // ════════════════════════════════════════════════════════════════
        // Filter combos
        // ════════════════════════════════════════════════════════════════
        private void PopulateFilterCombos()
        {
            if (_fullData == null) return;
            PopulateDistinct(cmbFwd, "Forwarder_name", "(All FWD)");
            PopulateDistinct(cmbContainer, "Forwarder_container", "(All Containers)");

            cmbGroupCol.Items.Clear();
            cmbGroupCol.Items.Add("(No grouping)");
            foreach (string c in new[] {
                "Forwarder_name","Forwarder_portDelivery","Forwarder_container",
                "Forwarder_term","Forwarder_carrier" })
                cmbGroupCol.Items.Add(c);
            if (cmbGroupCol.SelectedIndex < 0) cmbGroupCol.SelectedIndex = 0;

            cmbSearchCol.Items.Clear();
            cmbSearchCol.Items.Add("(All columns)");
            foreach (DataColumn c in _fullData.Columns)
                if (c.ColumnName != "Forwarder_ID")
                    cmbSearchCol.Items.Add(c.ColumnName);
            if (cmbSearchCol.SelectedIndex < 0) cmbSearchCol.SelectedIndex = 0;

            dtpFrom.Value = DateTime.Today.AddMonths(-3);
            dtpTo.Value = DateTime.Today;
        }

        private void PopulateDistinct(ComboBox cmb, string col, string allLabel)
        {
            string cur = cmb.SelectedItem?.ToString();
            cmb.Items.Clear();
            cmb.Items.Add(allLabel);
            if (_fullData != null && _fullData.Columns.Contains(col))
                foreach (var v in _fullData.AsEnumerable()
                    .Select(r => r[col]?.ToString() ?? "")
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .Distinct().OrderBy(v => v))
                    cmb.Items.Add(v);
            cmb.SelectedIndex = (cur != null && cmb.Items.Contains(cur))
                ? cmb.Items.IndexOf(cur) : 0;
        }

        private void CmbGroupCol_Changed(object sender, EventArgs e)
        {
            cmbGroupVal.Items.Clear();
            cmbGroupVal.Items.Add("(All)");
            string col = cmbGroupCol.SelectedItem?.ToString();
            if (col == null || col.StartsWith("(") || _fullData == null
                || !_fullData.Columns.Contains(col))
            { cmbGroupVal.SelectedIndex = 0; return; }
            foreach (var v in _fullData.AsEnumerable()
                .Select(r => r[col]?.ToString() ?? "")
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Distinct().OrderBy(v => v))
                cmbGroupVal.Items.Add(v);
            cmbGroupVal.SelectedIndex = 0;
        }

        // ════════════════════════════════════════════════════════════════
        // Filter logic
        // ════════════════════════════════════════════════════════════════
        private void ApplyFilters()
        {
            if (_fullData == null) return;

            string fwd = cmbFwd.SelectedItem?.ToString() ?? "";
            string cont = cmbContainer.SelectedItem?.ToString() ?? "";
            string groupCol = cmbGroupCol.SelectedItem?.ToString() ?? "";
            string groupVal = cmbGroupVal.SelectedItem?.ToString() ?? "";
            string srchCol = cmbSearchCol.SelectedItem?.ToString() ?? "";
            string keyword = txtKeyword.Text.Trim().ToUpper();
            bool useDates = chkDateFilter.Checked;
            DateTime from = dtpFrom.Value.Date;
            DateTime to = dtpTo.Value.Date.AddDays(1).AddTicks(-1);

            // Validate that groupCol is an actual column in _fullData before using it
            bool useGroupFilter = !groupCol.StartsWith("(")
                                  && !groupVal.StartsWith("(")
                                  && _fullData.Columns.Contains(groupCol);

            bool useSrchCol = !srchCol.StartsWith("(")
                              && _fullData.Columns.Contains(srchCol);

            var filtered = _fullData.Clone();
            foreach (DataRow row in _fullData.Rows)
            {
                // FWD filter
                if (!fwd.StartsWith("(") &&
                    !string.Equals(row["Forwarder_name"]?.ToString(), fwd,
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                // Container filter
                if (!cont.StartsWith("(") &&
                    !string.Equals(row["Forwarder_container"]?.ToString(), cont,
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                // Group-by value filter — only when column actually exists
                if (useGroupFilter &&
                    !string.Equals(row[groupCol]?.ToString(), groupVal,
                        StringComparison.OrdinalIgnoreCase))
                    continue;

                // Date range filter
                if (useDates)
                {
                    DateTime? cd = row["create_date"] == DBNull.Value
                        ? (DateTime?)null
                        : Convert.ToDateTime(row["create_date"]);
                    if (cd == null || cd.Value < from || cd.Value > to) continue;
                }

                // Keyword filter
                if (!string.IsNullOrEmpty(keyword))
                {
                    bool hit = false;
                    if (useSrchCol)
                    {
                        hit = row[srchCol]?.ToString().ToUpper().Contains(keyword) == true;
                    }
                    else
                    {
                        // Search all columns
                        foreach (DataColumn c in _fullData.Columns)
                            if (row[c]?.ToString().ToUpper().Contains(keyword) == true)
                            { hit = true; break; }
                    }
                    if (!hit) continue;
                }

                filtered.ImportRow(row);
            }

            if (useGroupFilter && filtered.Columns.Contains(groupCol))
                filtered.DefaultView.Sort = groupCol + " ASC";

            BindGrid(filtered);

            var parts = new List<string> { $"{filtered.Rows.Count} record(s)" };
            if (!fwd.StartsWith("(")) parts.Add($"FWD={fwd}");
            if (!cont.StartsWith("(")) parts.Add($"Container={cont}");
            if (useGroupFilter) parts.Add($"{groupCol}={groupVal}");
            if (!string.IsNullOrEmpty(keyword)) parts.Add($"Keyword='{txtKeyword.Text.Trim()}'");
            if (useDates) parts.Add($"Date {dtpFrom.Value:dd/MM/yy}–{dtpTo.Value:dd/MM/yy}");
            SetStatus(string.Join(" · ", parts), Color.SteelBlue);
        }

        private void ClearFilters()
        {
            cmbFwd.SelectedIndex = 0;
            cmbContainer.SelectedIndex = 0;
            cmbGroupCol.SelectedIndex = 0;
            cmbGroupVal.SelectedIndex = 0;
            cmbSearchCol.SelectedIndex = 0;
            txtKeyword.Text = "";
            chkDateFilter.Checked = false;
            dtpFrom.Value = DateTime.Today.AddMonths(-3);
            dtpTo.Value = DateTime.Today;
            BindGrid(_fullData);
            SetStatus($"{_fullData?.Rows.Count ?? 0} records.", Color.DarkSlateGray);
        }

        // ════════════════════════════════════════════════════════════════
        // Upload Excel — mirrors base tool logic exactly
        // ════════════════════════════════════════════════════════════════
        private void UploadExcel()
        {
            string path;
            using (var ofd = new OpenFileDialog
            {
                Filter = "Excel Files|*.xls*",
                Title = "Select Forwarder Quotation Excel"
            })
            {
                if (ofd.ShowDialog() != DialogResult.OK) return;
                path = ofd.FileName;
            }

            SetStatus("Parsing Excel…", Color.Blue);
            Application.DoEvents();

            Excel.Application xl = null;
            Excel.Workbook wb = null;
            int inserted = 0, skipped = 0;

            try
            {
                xl = new Excel.Application { Visible = false, DisplayAlerts = false };
                wb = xl.Workbooks.Open(path, ReadOnly: true);

                var fileInfo = new FileInfo(path);
                DateTime fileDate = fileInfo.CreationTime;
                string fwdName = Path.GetFileNameWithoutExtension(path);

                // Use first sheet — same as base tool fallback
                var ws = (Excel.Worksheet)wb.Sheets[1];
                var records = ParseSheet(ws, fwdName, fileDate);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);

                foreach (var m in records)
                {
                    try { _repo.Insert(m, _user.User_code); inserted++; }
                    catch { skipped++; }
                }
            }
            catch (Exception ex)
            {
                CTMessageBox.Show("Upload error:\r\n" + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (wb != null) { wb.Close(false); System.Runtime.InteropServices.Marshal.ReleaseComObject(wb); }
                if (xl != null) { xl.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(xl); }
                GC.Collect(); GC.WaitForPendingFinalizers();
            }

            LoadGrid();
            SetStatus($"Upload complete: {inserted} inserted, {skipped} skipped.", Color.DarkGreen);
        }

        // ────────────────────────────────────────────────────────────────
        // ParseSheet — mirrors base tool's LoadSingleExcelFile exactly:
        //   • Reads Value2 directly into object[,] (no COM per-cell)
        //   • FindHeaderRow: PORT/DELIVERY + CONTAINER + TOTAL/CHARGE
        //   • DetectColumnIndices: same keyword matching
        //   • GetMergedRaw: walks up the raw array for merged-cell values
        //   • Skips rows where port or container are empty
        //   • Skips rows where all charge values are 0
        // ────────────────────────────────────────────────────────────────
        private List<ForwarderQuotationModel> ParseSheet(
            Excel.Worksheet ws, string fwdName, DateTime fileDate)
        {
            var list = new List<ForwarderQuotationModel>();

            // ── Read entire sheet into array (fast, no per-cell COM) ──────
            Excel.Range used = ws.UsedRange;
            int rows = used.Rows.Count, cols = used.Columns.Count;
            object[,] raw;
            try { raw = (object[,])used.Value2; }
            catch { System.Runtime.InteropServices.Marshal.ReleaseComObject(used); return list; }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(used);
            if (raw == null) return list;

            // Helper: safe cell read
            string V(int r, int c) =>
                (r < 1 || r > rows || c < 1 || c > cols)
                    ? "" : raw[r, c]?.ToString().Trim() ?? "";

            // ── Find header row (same logic as base tool FindHeaderRow) ───
            int hdr = -1;
            for (int r = 1; r <= Math.Min(30, rows); r++)
            {
                string line = "";
                for (int c = 1; c <= Math.Min(20, cols); c++)
                    line += V(r, c).ToUpper() + " ";
                if ((line.Contains("PORT") || line.Contains("LOADING") ||
                     line.Contains("DELIVERY") || line.Contains("PLACE")) &&
                    (line.Contains("CONTAINER") || line.Contains("CONT")) &&
                    (line.Contains("TOTAL") || line.Contains("CHARGES") ||
                     line.Contains("ALL-IN")))
                { hdr = r; break; }
            }
            if (hdr < 0) return list;

            // ── Map columns (same as base tool DetectColumnIndices) ───────
            var map = new Dictionary<string, int>();
            for (int c = 1; c <= Math.Min(30, cols); c++)
            {
                string h = V(hdr, c).ToUpper().Trim();
                if (!map.ContainsKey("PORT") && (h.Contains("PORT") || h.Contains("LOADING") || h.Contains("DELIVERY") || h.Contains("PLACE"))) map["PORT"] = c;
                else if (!map.ContainsKey("TERM") && h.Contains("TERM")) map["TERM"] = c;
                else if (!map.ContainsKey("CONT") && (h.Contains("CONTAINER") || h.Contains("CONT"))) map["CONT"] = c;
                else if (!map.ContainsKey("COMMODITY") && h.Contains("COMMODITY")) map["COMMODITY"] = c;
                else if (!map.ContainsKey("HS") && h.Contains("HS")) map["HS"] = c;
                else if (!map.ContainsKey("CARRIER") && h.Contains("CARRIER")) map["CARRIER"] = c;
                else if (!map.ContainsKey("OF") && h.Contains("OF") && h.Contains("(1)")) map["OF"] = c;
                else if (!map.ContainsKey("LPOL") && h.Contains("LOCAL") && h.Contains("POL")) map["LPOL"] = c;
                else if (!map.ContainsKey("DEST") && h.Contains("DESTINATION") && h.Contains("CHARGE")) map["DEST"] = c;
                else if (!map.ContainsKey("DEL") && h.Contains("DELIVERY") && h.Contains("CHARGE")) map["DEL"] = c;
                else if (!map.ContainsKey("OTHER") && h.Contains("OTHER") && h.Contains("CHARGE")) map["OTHER"] = c;
                else if (!map.ContainsKey("TOTAL") && (h.Contains("TOTAL") || h.Contains("ALL-IN"))) map["TOTAL"] = c;
                else if (!map.ContainsKey("REMARK") && h.Contains("REMARK")) map["REMARK"] = c;
                else if (!map.ContainsKey("VOL") && h.Contains("VOL")) map["VOL"] = c;
                else if (!map.ContainsKey("VALID") && (h.Contains("VALIDITY") || h.Contains("VALID"))) map["VALID"] = c;
            }
            // Fallbacks (same as base tool)
            if (!map.ContainsKey("PORT")) map["PORT"] = 2;
            if (!map.ContainsKey("TERM")) map["TERM"] = 3;
            if (!map.ContainsKey("CONT")) map["CONT"] = 4;
            if (!map.ContainsKey("COMMODITY")) map["COMMODITY"] = 5;
            if (!map.ContainsKey("HS")) map["HS"] = 6;
            if (!map.ContainsKey("CARRIER")) map["CARRIER"] = 99;
            if (!map.ContainsKey("OF")) map["OF"] = 8;
            if (!map.ContainsKey("LPOL")) map["LPOL"] = 9;
            if (!map.ContainsKey("DEST")) map["DEST"] = 10;
            if (!map.ContainsKey("DEL")) map["DEL"] = 11;
            if (!map.ContainsKey("OTHER")) map["OTHER"] = 12;
            if (!map.ContainsKey("TOTAL")) map["TOTAL"] = 13;
            if (!map.ContainsKey("REMARK")) map["REMARK"] = 14;
            if (!map.ContainsKey("VOL")) map["VOL"] = 15;
            if (!map.ContainsKey("VALID")) map["VALID"] = 16;

            // ── Read data rows ────────────────────────────────────────────
            for (int r = hdr + 1; r <= rows; r++)
            {
                // Merged-cell aware read (mirrors GetMergedCellValue in base tool)
                string port = GetMergedRaw(raw, rows, cols, r, GetCol(map, "PORT"));
                string cont = GetMergedRaw(raw, rows, cols, r, GetCol(map, "CONT"));

                // Skip empty rows (same condition as base tool)
                if (string.IsNullOrWhiteSpace(port) || string.IsNullOrWhiteSpace(cont))
                    continue;

                decimal ParseDec(string key)
                {
                    int c = GetCol(map, key);
                    if (c < 1 || c > cols) return 0;
                    var val = raw[r, c];
                    if (val == null) return 0;
                    if (val is double d) return Math.Abs((decimal)d);
                    string s = val.ToString()
                        .Replace("USD", "").Replace("EUR", "")
                        .Replace("$", "").Replace(",", "").Trim();
                    return decimal.TryParse(s, System.Globalization.NumberStyles.Any,
                        System.Globalization.CultureInfo.InvariantCulture, out decimal res)
                        ? Math.Abs(res) : 0;
                }

                decimal of = ParseDec("OF");
                decimal lpol = ParseDec("LPOL");
                decimal dest = ParseDec("DEST");
                decimal del = ParseDec("DEL");
                decimal other = ParseDec("OTHER");
                decimal total = ParseDec("TOTAL");
                if (total <= 0) total = of + lpol + dest + del + other;

                // Skip rows with no charge data (same as base tool)
                if (total <= 0 && of == 0 && lpol == 0 && dest == 0 && del == 0 && other == 0)
                    continue;

                list.Add(new ForwarderQuotationModel
                {
                    Forwarder_ID = Guid.NewGuid().ToString(),
                    Forwarder_name = fwdName,
                    Forwarder_portDelivery = port,
                    Forwarder_term = GetMergedRaw(raw, rows, cols, r, GetCol(map, "TERM")),
                    Forwarder_container = cont,
                    Forwarder_commodity = V(r, GetCol(map, "COMMODITY")),
                    Forwarder_hsCode = V(r, GetCol(map, "HS")),
                    Forwarder_carrier = GetCol(map, "CARRIER") == 99 ? "" : V(r, GetCol(map, "CARRIER")),
                    Forwarder_of = of > 0 ? of : (decimal?)null,
                    Forwarder_localPol = lpol > 0 ? lpol : (decimal?)null,
                    Forwarder_destCharge = dest > 0 ? dest : (decimal?)null,
                    Forwarder_delivery = del > 0 ? del : (decimal?)null,
                    Forwarder_otherCharge = other > 0 ? other : (decimal?)null,
                    Forwarder_total = total > 0 ? total : (decimal?)null,
                    Forwarder_remark = V(r, GetCol(map, "REMARK")),
                    Forwarder_volumn = int.TryParse(V(r, GetCol(map, "VOL")), out int vol) ? vol : (int?)null,
                    Forwarder_validDate = V(r, GetCol(map, "VALID")),
                    create_date = fileDate,
                    create_by = _user.User_code,
                    update_date = fileDate,
                    update_by = _user.User_code
                });
            }
            return list;
        }

        // ────────────────────────────────────────────────────────────────
        // GetMergedRaw: mirrors base tool's GetMergedCellValue
        // Walks upward in the raw array to find the last non-empty value
        // (simulates merged cells without COM calls)
        // ────────────────────────────────────────────────────────────────
        private static string GetMergedRaw(object[,] raw, int rows, int cols, int r, int c)
        {
            if (c < 1 || c > cols) return "";
            for (int i = r; i >= 1; i--)
            {
                string v = raw[i, c]?.ToString().Trim() ?? "";
                if (!string.IsNullOrEmpty(v)) return v;
            }
            return "";
        }

        private static int GetCol(Dictionary<string, int> map, string key)
            => map.ContainsKey(key) ? map[key] : 0;

        // ════════════════════════════════════════════════════════════════
        // Edit
        // ════════════════════════════════════════════════════════════════
        private void OpenEditDialog()
        {
            if (!IsAdmin) return;
            if (dgv.SelectedRows.Count == 0)
            { CTMessageBox.Show("Select a row to edit.", "Info", MessageBoxButtons.OK); return; }

            var dt = (DataTable)dgv.DataSource;
            string id = dt.Rows[dgv.SelectedRows[0].Index]["Forwarder_ID"]?.ToString();
            var existing = _repo.GetById(id);

            using (var dlg = new FreightEditForm(existing, _user.User_code, isNew: false))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        _repo.Update(dlg.Result, _user.User_code);
                        LoadGrid();
                        SetStatus("Record updated.", Color.DarkGreen);
                    }
                    catch (Exception ex)
                    {
                        CTMessageBox.Show("Save error:\r\n" + ex.Message, "Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // ════════════════════════════════════════════════════════════════
        // History
        // ════════════════════════════════════════════════════════════════
        private void OpenHistory()
        {
            string id = null, displayName = null;
            if (dgv.SelectedRows.Count > 0)
            {
                var dt = (DataTable)dgv.DataSource;
                var dr = dt.Rows[dgv.SelectedRows[0].Index];
                id = dr["Forwarder_ID"]?.ToString();
                displayName = dr["Forwarder_name"]?.ToString();
            }
            using (var dlg = new FreightVersionHistoryForm(_repo, _user.User_code, id, displayName))
                if (dlg.ShowDialog(this) == DialogResult.OK) LoadGrid();
        }

        // ════════════════════════════════════════════════════════════════
        // Export
        // ════════════════════════════════════════════════════════════════
        private void ExportCSV()
        {
            if (dgv.Rows.Count == 0) { CTMessageBox.Show("No data."); return; }
            using (var sfd = new SaveFileDialog
            {
                Filter = "CSV|*.csv",
                FileName = $"Freight_{DateTime.Now:yyyyMMdd}.csv"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;
                var visCols = dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                var sb = new StringBuilder();
                sb.AppendLine(string.Join(",", visCols.Select(c => $"\"{c.HeaderText}\"")));
                foreach (DataGridViewRow row in dgv.Rows)
                    sb.AppendLine(string.Join(",",
                        visCols.Select(c => $"\"{row.Cells[c.Name].Value}\"")));
                File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                SetStatus("CSV exported.", Color.DarkGreen);
                System.Diagnostics.Process.Start(sfd.FileName);
            }
        }

        private void ExportExcel()
        {
            if (dgv.Rows.Count == 0) { CTMessageBox.Show("No data."); return; }
            using (var sfd = new SaveFileDialog
            {
                Filter = "Excel|*.xlsx",
                FileName = $"Freight_{DateTime.Now:yyyyMMdd}.xlsx"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;
                Excel.Application xl = null; Excel.Workbook wb = null;
                try
                {
                    xl = new Excel.Application { Visible = false, DisplayAlerts = false };
                    wb = xl.Workbooks.Add();
                    var ws = (Excel.Worksheet)wb.Sheets[1];
                    var visCols = dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                    for (int c = 0; c < visCols.Count; c++)
                    {
                        var cell = (Excel.Range)ws.Cells[1, c + 1];
                        cell.Value2 = visCols[c].HeaderText;
                        cell.Font.Bold = true;
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 153, 0));
                    }
                    for (int r = 0; r < dgv.Rows.Count; r++)
                        for (int c = 0; c < visCols.Count; c++)
                            ((Excel.Range)ws.Cells[r + 2, c + 1]).Value2 =
                                dgv.Rows[r].Cells[visCols[c].Name].Value?.ToString() ?? "";
                    ws.Columns.AutoFit();
                    wb.SaveAs(sfd.FileName);
                    SetStatus("Excel exported.", Color.DarkGreen);
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch (Exception ex)
                { CTMessageBox.Show("Export error:\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); }
                finally
                {
                    if (wb != null) { wb.Close(false); System.Runtime.InteropServices.Marshal.ReleaseComObject(wb); }
                    if (xl != null) { xl.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(xl); }
                    GC.Collect();
                }
            }
        }

        // ════════════════════════════════════════════════════════════════
        // UI helpers
        // ════════════════════════════════════════════════════════════════
        private void SetStatus(string msg, Color color)
        { lblStatus.Text = msg; lblStatus.ForeColor = color; }

        private static void PaintBorder(object sender, PaintEventArgs e)
        {
            var p = (Panel)sender;
            using (var pen = new System.Drawing.Pen(Color.FromArgb(220, 220, 230)))
                e.Graphics.DrawLine(pen, 0, p.Height - 1, p.Width, p.Height - 1);
        }

        private Button TB(Panel p, string text, Color bg, ref int x, int h = 28, int top = 8)
        {
            int w = Math.Max(52, text.Length * 7 + 14);
            var b = new Button
            {
                Text = text,
                Location = new Point(x, top),
                Size = new Size(w, h),
                BackColor = bg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 8f),
                Cursor = Cursors.Hand
            };
            b.FlatAppearance.BorderSize = 0;
            p.Controls.Add(b); x += w + 4;
            return b;
        }

        private ComboBox Cmb(Panel p, int w, ref int x)
        {
            var c = new ComboBox
            {
                Location = new Point(x, 11),
                Width = w,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 8f)
            };
            p.Controls.Add(c); x += w + 6;
            return c;
        }

        private DateTimePicker DTP(Panel p, ref int x)
        {
            var d = new DateTimePicker
            {
                Location = new Point(x, 10),
                Width = 100,
                Format = DateTimePickerFormat.Short,
                Font = new Font("Segoe UI", 8f)
            };
            p.Controls.Add(d); x += 106;
            return d;
        }

        private Label Lbl(string text, int x, int y) =>
            new Label
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(60, 60, 60)
            };
    }
}