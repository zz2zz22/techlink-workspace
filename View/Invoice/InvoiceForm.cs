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
using techlink_workspace.Repositories.InvoiceRepo;

namespace techlink_workspace.View.Invoice
{
    public class InvoiceForm : Form
    {
        private readonly UserModel _user;
        private readonly InvoiceRepository _repo = new InvoiceRepository();
        private DataTable _fullData;

        private int Level => _user.User_permissionLevel ?? 3;
        private bool CanImport => Level == 1;
        private bool CanDelete => Level == 1;
        private bool CanHistory => Level <= 2;
        private bool CanExport => Level <= 2;
        private bool CanSettings => Level <= 2;

        private InvoiceEditScope GetScope() =>
            Level == 1 ? InvoiceEditScope.All :
            Level == 2 ? InvoiceEditScope.ErpLocked :
                         InvoiceEditScope.LevelThree;

        // ── Toolbar ───────────────────────────────────────────────────────
        private Button btnImport, btnEdit, btnDelete,
                       btnHistory, btnExportCSV, btnExportExcel,
                       btnSettings, btnRefresh;

        // ── Filter bar ────────────────────────────────────────────────────
        private ComboBox cmbFwd, cmbContType, cmbStatus,
                               cmbGroupCol, cmbGroupVal, cmbSearchCol;
        private TextBox txtKeyword;
        private DateTimePicker dtpFrom, dtpTo;
        private CheckBox chkDate;
        private Button btnSearch, btnClear;

        private DataGridView dgv;
        private Label lblStatus;

        // ── Column headers ────────────────────────────────────────────────
        private static readonly Dictionary<string, string> HEADERS =
    new Dictionary<string, string>
{
    // ── Identity ──────────────────────────────────────────────────
    {"Invoice_Id",                   "ID"},
    {"Invoice_no",                   "ERP Invoice No"},      // row 27
    // ── ERP (rows 1-14 + brand row 6) ────────────────────────────
    {"Invoice_customerCode",         "Customer Code"},               // 1
    {"Invoice_customerName",         "Customer Name"},      // 2
    {"Invoice_customerRequestDate",  "Request Date"},     // 3
    {"Invoice_poNo",                 "PO Number"},              // 4
    {"Invoice_poDate",               "PO Date"},            // 5
    {"Invoice_brand",                "Brand"},             // 6  
    {"Invoice_saleName",             "Sale Name"},                 // 7
    {"Invoice_factoryNo",            "Factory Number"},            // 8
    {"Invoice_factoryName",          "Factory Name"},           // 9
    {"Invoice_itemCode",             "Item Code"},                // 10
    {"Invoice_itemCodeCustomers",    "Customer Item Code"},             // 11
    {"Invoice_itemName",             "Item Name"},               // 12
    {"Invoice_quantity",             "Quantity"},          // 13
    {"Invoice_unit",                 "Unit"},                  // 14
    // ── Logistics Plan (rows 15-31) ───────────────────────────────
    {"Invoice_shippingTerm",         "Shipping Term"},        // 15
    {"Invoice_paymentTerm",          "Payment Term"},         // 16
    {"Invoice_logisticPersonInCharge","NV Logistics PIC"},   // 16
    {"Invoice_logisticNote",         "NOTE"},                 // 17
    {"Invoice_factoryConfirmDate",   "Factory Confirm Date"}, // 18
    {"Invoice_shippingStatus",       "Status"},               // 19
    {"Invoice_fwdName",              "FWD Name"},             // 20
    {"Invoice_bookingNo",            "Book Number"},          // 21
    {"Invoice_contType",             "Type (Cont)"},          // 22
    {"Invoice_vgmCO",                "SI & VGM cut-off"},    // 23
    {"Invoice_cyCO",                 "CY cut-off"},          // 24
    {"Invoice_etd",                  "ETD"},                  // 25
    {"Invoice_eta",                  "ETA"},                  // 26
    {"Invoice_billType",             "Type of Bill"},         // 28
    {"Invoice_billNo",               "Bill Number"},          // 29
    {"Invoice_co",                   "CO: Yes/No"},          // 30
    {"Invoice_coNo",                 "CO Number"},            // 31
    // ── Ocean Charges (rows 32-46) ────────────────────────────────
    {"Invoice_OF",                   "OF"},                   // 32
    {"Invoice_deliveryCharges",      "Delivery Charges"},     // 33
    {"Invoice_taxes",                "Duty/Taxes"},           // 34
    {"Invoice_otherDestCharges",     "Other Dest Charges"},   // 35
    {"Invoice_thc",                  "THC"},                  // 36
    {"Invoice_blFee",                "b/l fee"},              // 37
    {"Invoice_seal",                 "Seal"},                 // 38
    {"Invoice_telexRelease",         "Telex Release"},        // 39
    {"Invoice_cfs",                  "CFS"},                  // 40
    {"Invoice_vgmFee",               "VGM"},                  // 41
    {"Invoice_ensebsams",            "ENS/EBS/AMS"},          // 42
    {"Invoice_other",                "Others"},               // 43
    {"Invoice_coFee",                "CO Fee"},               // 44
    {"Invoice_totalVND",             "Total (VND)"},          // 45
    {"Invoice_subTotalOcean",        "Sub-Total Ocean (USD)"},// 46
    // ── Invoice Status (rows 47-51) ───────────────────────────────
    {"Invoice_feeStatus",            "Status paid/Acct"},     // 47
    {"Invoice_redInvoiceNo",         "Red Invoice#"},         // 48
    {"Invoice_redInvoiceDate",       "Red Invoice Date"},     // 49
    {"Invoice_redInvoiceRecvDate",   "Red Invoice Recv"},     // 50
    {"Invoice_transferAccountantDate","Transfer to Acct"},    // 51
    // ── Customs (rows 52-64) ──────────────────────────────────────
    {"Invoice_trucking",             "Trucking"},              // 52
    {"Invoice_infrastructureFee",    "Infrastructure Fee"},   // 53
    {"Invoice_customerClearance",    "Customs Clearance"},    // 54
    {"Invoice_customFee",            "Customs Fee"},          // 55
    {"Invoice_otherCustomFee",       "Others (Custom)"},      // 56
    {"Invoice_subTotalVNDCustom",    "Sub-Total VND+Customs"},// 57
    {"Invoice_subTotalUSDCustom",    "Sub-Total USD+Customs"},// 58
    {"Invoice_grandTotalVND",        "Grand Total (VND)"},    // 59
    {"Invoice_grandTotalUSD",        "Grand Total (USD)"},    // 60
    {"Invoice_cdsNo",                "CDS No"},               // 61
    {"Invoice_cdsDate",              "CDS Date"},             // 62
    {"Invoice_cdsApproved",          "CDS Approve"},          // 63
    {"Invoice_customType",           "Mã loại hình tk"},      // 64
    {"Invoice_line",                 "Luong"},
    // ── Audit ────────────────────────────────────────────────────
    {"createdate","Created"}, {"createby","By"},
    {"updatedate","Updated"}, {"updateby","Upd By"},
};

        // ════════════════════════════════════════════════════════════════
        public InvoiceForm(UserModel user)
        {
            _user = user;
            BuildUI();
            LoadGrid();
        }

        // ════════════════════════════════════════════════════════════════
        // BUILD UI
        // ════════════════════════════════════════════════════════════════
        private void BuildUI()
        {
            Text = "Logistics Invoice";
            ClientSize = new Size(1500, 780);
            MinimumSize = new Size(1000, 500);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 8.5f);

            // ── Row 1: action buttons ─────────────────────────────────────
            var r1 = Row(44, Color.White);
            int x = 8;
            btnImport = TB(r1, "⬆ Import ERP", Color.FromArgb(0, 51, 153), ref x);
            btnEdit = TB(r1, "✎ Edit", Color.DarkOrange, ref x);
            btnDelete = TB(r1, "✕ Delete", Color.IndianRed, ref x);
            x += 8;
            btnHistory = TB(r1, "⟳ History", Color.SlateBlue, ref x);
            x += 8;
            btnExportCSV = TB(r1, "⬇ CSV", Color.DarkGreen, ref x);
            btnExportExcel = TB(r1, "⬇ Excel", Color.DarkGreen, ref x);
            x += 8;
            btnSettings = TB(r1, "⚙ Settings", Color.FromArgb(80, 80, 100), ref x);
            btnRefresh = TB(r1, "↺ Refresh", Color.FromArgb(90, 90, 90), ref x);

            // ── Row 2: filter bar ─────────────────────────────────────────
            var r2 = Row(44, Color.FromArgb(248, 248, 252));
            int x2 = 8;

            r2.Controls.Add(Lbl("FWD:", x2, 14)); x2 += 36;
            cmbFwd = Cmb(r2, 110, ref x2);

            r2.Controls.Add(Lbl("Cont:", x2, 14)); x2 += 40;
            cmbContType = Cmb(r2, 80, ref x2);

            r2.Controls.Add(Lbl("Status:", x2, 14)); x2 += 52;
            cmbStatus = Cmb(r2, 90, ref x2);

            r2.Controls.Add(Lbl("Group:", x2, 14)); x2 += 48;
            cmbGroupCol = Cmb(r2, 130, ref x2);
            cmbGroupVal = Cmb(r2, 130, ref x2);

            r2.Controls.Add(Lbl("In:", x2, 14)); x2 += 22;
            cmbSearchCol = Cmb(r2, 130, ref x2);

            txtKeyword = new TextBox
            {
                Location = new Point(x2, 11),
                Width = 140,
                Font = new Font("Segoe UI", 8.5f)
            };
            r2.Controls.Add(txtKeyword); x2 += 146;

            chkDate = new CheckBox
            {
                Text = "Date:",
                Location = new Point(x2, 13),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f)
            };
            r2.Controls.Add(chkDate); x2 += 52;

            dtpFrom = DTP(r2, ref x2);
            r2.Controls.Add(Lbl("–", x2, 14)); x2 += 14;
            dtpTo = DTP(r2, ref x2); x2 += 6;

            btnSearch = TB(r2, "🔍", Color.SteelBlue, ref x2, 26, 9, 34);
            btnClear = TB(r2, "✕ Clear", Color.DimGray, ref x2, 26, 9);

            // ── Status bar ────────────────────────────────────────────────
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

            // ── Data grid ────────────────────────────────────────────────
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
            dgv.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0) OpenEdit(); };

            Controls.Add(dgv);
            Controls.Add(lblStatus);
            Controls.Add(r2);
            Controls.Add(r1);

            // ── Permissions ───────────────────────────────────────────────
            btnImport.Enabled = CanImport;
            btnDelete.Enabled = CanDelete;
            btnHistory.Enabled = CanHistory;
            btnExportCSV.Enabled = CanExport;
            btnExportExcel.Enabled = Level == 1;
            btnSettings.Enabled = CanSettings;

            // ── Wire events ───────────────────────────────────────────────
            btnImport.Click += (s, e) => OpenImport();
            btnEdit.Click += (s, e) => OpenEdit();
            btnDelete.Click += (s, e) => DeleteSelected();
            btnHistory.Click += (s, e) => OpenHistory();
            btnExportCSV.Click += (s, e) => DoExportCSV();
            btnExportExcel.Click += (s, e) => DoExportExcel();
            btnSettings.Click += (s, e) => OpenSettings();
            btnRefresh.Click += (s, e) => LoadGrid();
            btnSearch.Click += (s, e) => ApplyFilters();
            btnClear.Click += (s, e) => ClearFilters();
            txtKeyword.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) ApplyFilters(); };
            cmbGroupCol.SelectedIndexChanged += (s, e) => RefreshGroupVal();
            cmbFwd.SelectedIndexChanged += (s, e) => ApplyFilters();
            cmbContType.SelectedIndexChanged += (s, e) => ApplyFilters();
            cmbStatus.SelectedIndexChanged += (s, e) => ApplyFilters();
            cmbGroupVal.SelectedIndexChanged += (s, e) => ApplyFilters();
        }

        // ════════════════════════════════════════════════════════════════
        // LOAD & BIND
        // ════════════════════════════════════════════════════════════════
        private void LoadGrid()
        {
            try
            {
                dgv.DataSource = null;
                _fullData = Level <= 2
                    ? _repo.GetAll()
                    : _repo.GetByPIC(_user.User_code);   // level 3 sees own invoices only
                PopulateFilterCombos();
                BindGrid(_fullData);
                SetStatus($"{_fullData.Rows.Count} record(s) loaded.", Color.DarkSlateGray);
            }
            catch (Exception ex)
            {
                SetStatus("Load error: " + ex.Message, Color.Red);
                CTMessageBox.Show("Failed to load:\r\n" + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BindGrid(DataTable dt)
        {
            dgv.DataSource = dt.Copy();

            if (dgv.Columns.Contains("Invoice_Id"))
                dgv.Columns["Invoice_Id"].Visible = false;

            foreach (var kv in HEADERS)
                if (dgv.Columns.Contains(kv.Key))
                    dgv.Columns[kv.Key].HeaderText = kv.Value;

            // Freeze Invoice No column
            if (dgv.Columns.Contains("Invoice_no"))
            {
                dgv.Columns["Invoice_no"].DisplayIndex = 0;
                dgv.Columns["Invoice_no"].Frozen = true;
                dgv.Columns["Invoice_no"].DefaultCellStyle.Font =
                    new Font("Segoe UI", 8, FontStyle.Bold);
            }

            // Highlight totals
            foreach (string col in new[]
                { "Invoice_grandTotalVND", "Invoice_grandTotalUSD", "Invoice_subTotalOcean" })
                if (dgv.Columns.Contains(col))
                {
                    dgv.Columns[col].DefaultCellStyle.BackColor = Color.LightYellow;
                    dgv.Columns[col].DefaultCellStyle.Font =
                        new Font("Segoe UI", 8, FontStyle.Bold);
                }

            // Highlight ETD-expired rows pink for level-3 awareness
            dgv.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0) return;
                var src = (DataTable)dgv.DataSource;
                if (src == null || e.RowIndex >= src.Rows.Count) return;
                var dr = src.Rows[e.RowIndex];
                if (Level == 3 && dr["Invoice_etd"] != DBNull.Value)
                {
                    if (Convert.ToDateTime(dr["Invoice_etd"]).Date <
                        DateTime.Now.AddDays(-7).Date)
                        e.CellStyle.BackColor = Color.FromArgb(255, 220, 220);
                }
            };
        }

        // ════════════════════════════════════════════════════════════════
        // FILTER COMBOS
        // ════════════════════════════════════════════════════════════════
        private void PopulateFilterCombos()
        {
            PopDist(cmbFwd, "Invoice_fwdName", "(All FWD)");
            PopDist(cmbContType, "Invoice_contType", "(All Cont)");

            cmbStatus.Items.Clear();
            cmbStatus.Items.AddRange(new object[]
                { "(All Status)", "BLANK", "BOOKED", "CANCEL", "SHIPPED", "WAITING" });
            cmbStatus.SelectedIndex = 0;

            cmbGroupCol.Items.Clear();
            cmbGroupCol.Items.Add("(No grouping)");
            foreach (string c in new[]
            {
                "Invoice_fwdName", "Invoice_contType", "Invoice_logisticPersonInCharge",
                "Invoice_billType", "Invoice_line", "Invoice_customType"
            })
                cmbGroupCol.Items.Add(c);
            if (cmbGroupCol.SelectedIndex < 0) cmbGroupCol.SelectedIndex = 0;

            cmbSearchCol.Items.Clear();
            cmbSearchCol.Items.Add("(All columns)");
            if (_fullData != null)
                foreach (DataColumn c in _fullData.Columns)
                    if (c.ColumnName != "Invoice_Id") cmbSearchCol.Items.Add(c.ColumnName);
            if (cmbSearchCol.SelectedIndex < 0) cmbSearchCol.SelectedIndex = 0;

            dtpFrom.Value = DateTime.Today.AddYears(-1);
            dtpTo.Value = DateTime.Today;
        }

        private void PopDist(ComboBox cmb, string col, string allLabel)
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

        private void RefreshGroupVal()
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
        // FILTER LOGIC
        // ════════════════════════════════════════════════════════════════
        private void ApplyFilters()
        {
            if (_fullData == null) return;

            string fwd = cmbFwd.SelectedItem?.ToString() ?? "";
            string cont = cmbContType.SelectedItem?.ToString() ?? "";
            string status = cmbStatus.SelectedItem?.ToString() ?? "";
            string groupCol = cmbGroupCol.SelectedItem?.ToString() ?? "";
            string groupVal = cmbGroupVal.SelectedItem?.ToString() ?? "";
            string srchCol = cmbSearchCol.SelectedItem?.ToString() ?? "";
            string kw = txtKeyword.Text.Trim().ToUpper();
            bool useDates = chkDate.Checked;
            var from = dtpFrom.Value.Date;
            var to = dtpTo.Value.Date.AddDays(1).AddTicks(-1);

            var statusMap = new Dictionary<string, int>
                { {"BLANK",0}, {"BOOKED",1}, {"CANCEL",2}, {"SHIPPED",3}, {"WAITING",4} };
            int? filterStatus = (!status.StartsWith("(") && statusMap.ContainsKey(status))
                ? statusMap[status] : (int?)null;

            bool useGroup = !groupCol.StartsWith("(") && !groupVal.StartsWith("(")
                            && _fullData.Columns.Contains(groupCol);
            bool useSrch = !srchCol.StartsWith("(") && _fullData.Columns.Contains(srchCol);

            var filtered = _fullData.Clone();
            foreach (DataRow row in _fullData.Rows)
            {
                if (!fwd.StartsWith("(") && !Eq(row["Invoice_fwdName"], fwd)) continue;
                if (!cont.StartsWith("(") && !Eq(row["Invoice_contType"], cont)) continue;
                if (filterStatus.HasValue &&
                    !Eq(row["Invoice_shippingStatus"], filterStatus.Value.ToString())) continue;
                if (useGroup && !Eq(row[groupCol], groupVal)) continue;
                if (useDates)
                {
                    DateTime? cd = row["createdate"] == DBNull.Value
                        ? (DateTime?)null : Convert.ToDateTime(row["createdate"]);
                    if (cd == null || cd.Value < from || cd.Value > to) continue;
                }
                if (!string.IsNullOrEmpty(kw))
                {
                    bool hit = false;
                    if (useSrch)
                        hit = row[srchCol]?.ToString().ToUpper().Contains(kw) == true;
                    else
                        foreach (DataColumn c in _fullData.Columns)
                            if (row[c]?.ToString().ToUpper().Contains(kw) == true)
                            { hit = true; break; }
                    if (!hit) continue;
                }
                filtered.ImportRow(row);
            }

            if (useGroup && filtered.Columns.Contains(groupCol))
                filtered.DefaultView.Sort = groupCol + " ASC";

            BindGrid(filtered);

            var parts = new List<string> { $"{filtered.Rows.Count} record(s)" };
            if (!fwd.StartsWith("(")) parts.Add($"FWD={fwd}");
            if (!cont.StartsWith("(")) parts.Add($"Cont={cont}");
            if (filterStatus.HasValue) parts.Add($"Status={status}");
            if (useGroup) parts.Add($"{groupCol}={groupVal}");
            if (!string.IsNullOrEmpty(kw)) parts.Add($"KW='{txtKeyword.Text.Trim()}'");
            if (useDates) parts.Add($"{dtpFrom.Value:dd/MM/yy}–{dtpTo.Value:dd/MM/yy}");
            SetStatus(string.Join(" · ", parts), Color.SteelBlue);
        }

        private void ClearFilters()
        {
            cmbFwd.SelectedIndex = cmbContType.SelectedIndex = cmbStatus.SelectedIndex =
            cmbGroupCol.SelectedIndex = cmbGroupVal.SelectedIndex =
            cmbSearchCol.SelectedIndex = 0;
            txtKeyword.Text = "";
            chkDate.Checked = false;
            BindGrid(_fullData);
            SetStatus($"{_fullData?.Rows.Count ?? 0} records.", Color.DarkSlateGray);
        }

        // ════════════════════════════════════════════════════════════════
        // CRUD ACTIONS
        // ════════════════════════════════════════════════════════════════
        private void OpenEdit()
        {
            if (dgv.SelectedRows.Count == 0)
            { CTMessageBox.Show("Select a row to edit."); return; }

            string id = ((DataTable)dgv.DataSource)
                .Rows[dgv.SelectedRows[0].Index]["Invoice_Id"]?.ToString();
            var existing = _repo.GetById(id);

            using (var dlg = new InvoiceEditForm(existing, _user.User_code, false, GetScope()))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
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

        private void DeleteSelected()
        {
            if (!CanDelete) return;
            if (dgv.SelectedRows.Count == 0)
            { CTMessageBox.Show("Select a row to delete."); return; }

            if (CTMessageBox.Show(
                    $"Delete {dgv.SelectedRows.Count} record(s)? This action is logged and reversible via History.",
                    "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
                != DialogResult.Yes) return;

            var dt = (DataTable)dgv.DataSource;
            foreach (DataGridViewRow gr in dgv.SelectedRows)
            {
                string id = dt.Rows[gr.Index]["Invoice_Id"]?.ToString();
                if (!string.IsNullOrEmpty(id))
                    _repo.Delete(id, _user.User_code);
            }
            LoadGrid();
            SetStatus("Deleted.", Color.DarkSlateGray);
        }

        private void OpenImport()
        {
            using (var dlg = new InvoiceExcelImportForm(_repo, _user.User_code))
                dlg.ShowDialog(this);
            LoadGrid();
        }

        private void OpenHistory()
        {
            string id = null, disp = null;
            if (dgv.SelectedRows.Count > 0)
            {
                var dt = (DataTable)dgv.DataSource;
                var dr = dt.Rows[dgv.SelectedRows[0].Index];
                id = dr["Invoice_Id"]?.ToString();
                disp = dr["Invoice_no"]?.ToString();
            }
            using (var dlg = new InvoiceHistoryForm(_repo, _user.User_code, id, disp))
                if (dlg.ShowDialog(this) == DialogResult.OK) LoadGrid();
        }

        private void OpenSettings()
        {
            using (var dlg = new InvoiceSettingsForm(_user.User_code))
                dlg.ShowDialog(this);
            LoadGrid();   // refresh in case PIC mapping changed visibility
        }

        // ════════════════════════════════════════════════════════════════
        // EXPORT
        // ════════════════════════════════════════════════════════════════
        private void DoExportCSV()
        {
            if (dgv.Rows.Count == 0) { CTMessageBox.Show("No data to export."); return; }

            using (var sfd = new SaveFileDialog
            { Filter = "CSV|*.csv", FileName = $"Invoice_{DateTime.Now:yyyyMMdd}.csv" })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                var vis = dgv.Columns.Cast<DataGridViewColumn>()
                             .Where(c => c.Visible).ToList();
                var sb = new StringBuilder();
                sb.AppendLine(string.Join(",", vis.Select(c => $"\"{c.HeaderText}\"")));
                foreach (DataGridViewRow row in dgv.Rows)
                    sb.AppendLine(string.Join(",",
                        vis.Select(c => $"\"{row.Cells[c.Name].Value}\"")));

                File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                SetStatus("CSV exported.", Color.DarkGreen);
                System.Diagnostics.Process.Start(sfd.FileName);
            }
        }

        private void DoExportExcel()
        {
            if (dgv.Rows.Count == 0) { CTMessageBox.Show("No data to export."); return; }

            using (var sfd = new SaveFileDialog
            { Filter = "Excel|*.xlsx", FileName = $"Invoice_{DateTime.Now:yyyyMMdd}.xlsx" })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;

                Excel.Application xl = null;
                Excel.Workbook wb = null;
                try
                {
                    xl = new Excel.Application { Visible = false, DisplayAlerts = false };
                    wb = xl.Workbooks.Add();
                    var ws = (Excel.Worksheet)wb.Sheets[1];
                    var vis = dgv.Columns.Cast<DataGridViewColumn>()
                                 .Where(c => c.Visible).ToList();

                    for (int c = 0; c < vis.Count; c++)
                    {
                        var cell = (Excel.Range)ws.Cells[1, c + 1];
                        cell.Value2 = vis[c].HeaderText;
                        cell.Font.Bold = true;
                        cell.Interior.Color =
                            System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 153, 0));
                    }

                    for (int r = 0; r < dgv.Rows.Count; r++)
                        for (int c = 0; c < vis.Count; c++)
                            ((Excel.Range)ws.Cells[r + 2, c + 1]).Value2 =
                                dgv.Rows[r].Cells[vis[c].Name].Value?.ToString() ?? "";

                    ws.Columns.AutoFit();
                    wb.SaveAs(sfd.FileName);
                    SetStatus("Excel exported.", Color.DarkGreen);
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch (Exception ex)
                {
                    CTMessageBox.Show("Export error:\r\n" + ex.Message, "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    if (wb != null)
                    {
                        wb.Close(false);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                    }
                    if (xl != null)
                    {
                        xl.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);
                    }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        // ════════════════════════════════════════════════════════════════
        // UI HELPERS
        // ════════════════════════════════════════════════════════════════
        private void SetStatus(string msg, Color color)
        { lblStatus.Text = msg; lblStatus.ForeColor = color; }

        private Panel Row(int h, Color bg)
        {
            var p = new Panel { Dock = DockStyle.Top, Height = h, BackColor = bg };
            p.Paint += (s, e) =>
            {
                using (var pen = new System.Drawing.Pen(Color.FromArgb(220, 220, 230)))
                    e.Graphics.DrawLine(pen, 0, p.Height - 1, p.Width, p.Height - 1);
            };
            return p;
        }

        private Button TB(Panel p, string text, Color bg, ref int x,
                          int h = 28, int top = 8, int fixW = 0)
        {
            int w = fixW > 0 ? fixW : Math.Max(52, text.Length * 7 + 14);
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
            p.Controls.Add(b);
            x += w + 4;
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
            p.Controls.Add(c);
            x += w + 6;
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
            p.Controls.Add(d);
            x += 106;
            return d;
        }

        private static Label Lbl(string text, int x, int y) =>
            new Label
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(60, 60, 60)
            };

        private static bool Eq(object cell, string val) =>
            string.Equals(cell?.ToString(), val, StringComparison.OrdinalIgnoreCase);
    }
}