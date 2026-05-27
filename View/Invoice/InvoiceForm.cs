using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;
using techlink_workspace.Repositories.InvoiceRepo;

namespace techlink_workspace.View.Invoice
{
    public class InvoiceForm : Form
    {
        // ── State ────────────────────────────────────────────────────────
        private readonly UserModel _user;
        private readonly InvoiceRepository _repo = new InvoiceRepository();
        private DataTable _fullData;

        // Permission: level 1=Admin, 2=Manager, 3=Normal
        private int Level => _user.User_permissionLevel ?? 3;
        private bool CanEdit => Level <= 2;
        private bool CanExport => Level <= 1;
        private bool CanDelete => Level <= 1;
        private bool CanImport => Level == 1;   // admin only

        /// <summary>
        /// Determines which sections the user can edit based on User_type name.
        /// "Logistic" → plan section only.
        /// "Custom"   → customs section only.
        /// Admin (level 1) or Manager (level 2) → all.
        /// </summary>
        private InvoiceEditScope GetEditScope()
        {
            if (Level == 1) return InvoiceEditScope.All;
            string t = (_user.User_type ?? "").ToLower();
            if (t.Contains("custom")) return InvoiceEditScope.CustomOnly;
            if (t.Contains("logistic")) return InvoiceEditScope.LogisticPlan;
            return InvoiceEditScope.All;   // manager or unrecognised → all
        }

        // ── Toolbar row 1 ────────────────────────────────────────────────
        private Button btnImportExcel, btnAdd, btnEdit, btnDelete,
                       btnHistory, btnExportCSV, btnExportExcel, btnRefresh;

        // ── Toolbar row 2 – filter bar ───────────────────────────────────
        private ComboBox cmbFwd, cmbContType, cmbEmployee, cmbFeeStatus;
        private ComboBox cmbGroupCol, cmbGroupVal;
        private ComboBox cmbSearchCol;
        private TextBox txtKeyword;
        private DateTimePicker dtpFrom, dtpTo;
        private CheckBox chkDate;
        private Button btnSearch, btnClear;

        // ── Grid + status ────────────────────────────────────────────────
        private DataGridView dgv;
        private Label lblStatus;

        // ── Friendly column headers ──────────────────────────────────────
        private static readonly Dictionary<string, string> HEADERS =
            new Dictionary<string, string>
        {
            {"Invoice_Id","ID"},{"Invoice_no","Invoice No"},
            {"Invoice_erpID","ERP ID"},{"Invoice_erpInvoiceNo","ERP Inv No"},
            {"Invoice_shippingTerm","Ship Term"},{"Invoice_paymentTerm","Pay Term"},
            {"Invoice_employee","Employee"},{"Invoice_logisticRemark","Remark"},
            {"Invoice_confirmDate","Confirm Date"},{"Invoice_fwdName","FWD"},
            {"Invoice_bookingNo","Booking No"},{"Invoice_contType","Cont Type"},
            {"Invoice_vgmCO","VGM C/O"},{"Invoice_cyCO","CY C/O"},
            {"Invoice_etd","ETD"},{"Invoice_eta","ETA"},
            {"Invoice_billType","Bill Type"},{"Invoice_billNo","Bill No"},
            {"Invoice_co","C/O"},{"Invoice_coNo","C/O No"},
            {"Invoice_OF","OF"},{"Invoice_deliveryCharges","Delivery"},
            {"Invoice_taxes","Taxes"},{"Invoice_otherDestCharges","Other Dest"},
            {"Invoice_thc","THC"},{"Invoice_blFee","B/L Fee"},
            {"Invoice_seal","Seal"},{"Invoice_telexRelease","Telex"},
            {"Invoice_cfs","CFS"},{"Invoice_vgmFee","VGM Fee"},
            {"Invoice_ensebsams","ENS/EBS/AMS"},{"Invoice_other","Other"},
            {"Invoice_totalVND","Total VND"},{"Invoice_subTotalOcean","SubTotal Ocean USD"},
            {"Invoice_coFee","C/O Fee"},{"Invoice_feeStatus","Fee Status"},
            {"Invoice_redInvoiceNo","Red Inv No"},{"Invoice_redInvoiceDate","Red Inv Date"},
            {"Invoice_redInvoiceRecvDate","Red Inv Recv"},
            {"Invoice_transferAccountDate","Transfer Acct"},
            {"Invoice_trucking","Trucking"},{"Invoice_infrastructureFee","Infra Fee"},
            {"Invoice_customerClearance","Cust Clear"},{"Invoice_customFee","Custom Fee"},
            {"Invoice_otherCustomFee","Other Custom"},
            {"Invoice_subTotalVNDCustom","SubTotal VND Custom"},
            {"Invoice_subTotalUSDCustom","SubTotal USD Custom"},
            {"Invoice_grandTotalVND","Grand Total VND"},
            {"Invoice_grandTotalUSD","Grand Total USD"},
            {"Invoice_cdsNo","CDS No"},{"Invoice_cdsDate","CDS Date"},
            {"Invoice_line","Line"},{"Invoice_customType","Custom Type"},
            {"createdate","Created"},{"createby","Created By"},
            {"updatedate","Updated"},{"updateby","Updated By"},
        };

        // ════════════════════════════════════════════════════════════════
        public InvoiceForm(UserModel user)
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
            Text = "Logistics Invoice";
            ClientSize = new Size(1500, 780);
            MinimumSize = new Size(1100, 500);
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 8.5f);

            // ── Row 1 ─────────────────────────────────────────────────────
            var r1 = Row(44, Color.White);
            int x = 8;
            btnImportExcel = TB(r1, "⬆ Import Excel", Color.FromArgb(0, 51, 153), ref x);
            btnAdd = TB(r1, "＋ Add", Color.SeaGreen, ref x);
            btnEdit = TB(r1, "✎ Edit", Color.DarkOrange, ref x);
            btnDelete = TB(r1, "✕ Delete", Color.IndianRed, ref x);
            x += 10;
            btnHistory = TB(r1, "⟳ History", Color.SlateBlue, ref x);
            x += 10;
            btnExportCSV = TB(r1, "⬇ CSV", Color.DarkGreen, ref x);
            btnExportExcel = TB(r1, "⬇ Excel", Color.DarkGreen, ref x);
            x += 10;
            btnRefresh = TB(r1, "↺ Refresh", Color.FromArgb(90, 90, 90), ref x);

            // ── Row 2 – filter bar ────────────────────────────────────────
            var r2 = Row(44, Color.FromArgb(248, 248, 252));
            int x2 = 8;

            r2.Controls.Add(Lbl("FWD:", x2, 14)); x2 += 36;
            cmbFwd = Cmb(r2, 110, ref x2);

            r2.Controls.Add(Lbl("Cont:", x2, 14)); x2 += 40;
            cmbContType = Cmb(r2, 80, ref x2);

            r2.Controls.Add(Lbl("Staff:", x2, 14)); x2 += 44;
            cmbEmployee = Cmb(r2, 100, ref x2);

            r2.Controls.Add(Lbl("Status:", x2, 14)); x2 += 52;
            cmbFeeStatus = Cmb(r2, 80, ref x2);

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

            btnSearch = TB(r2, "🔍", Color.SteelBlue, ref x2, 26, 9, 36);
            btnClear = TB(r2, "✕ Clear", Color.DimGray, ref x2, 26, 9);

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
            dgv.CellDoubleClick += (s, e) => { if (e.RowIndex >= 0 && CanEdit) OpenEdit(false); };

            Controls.Add(dgv);
            Controls.Add(lblStatus);
            Controls.Add(r2);
            Controls.Add(r1);

            // ── Permissions ──────────────────────────────────────────────
            btnImportExcel.Enabled = CanImport;   // admin only
            btnAdd.Enabled = CanEdit;
            btnEdit.Enabled = CanEdit;
            btnDelete.Enabled = CanDelete;
            btnHistory.Enabled = CanEdit;
            btnExportCSV.Enabled = CanExport;
            btnExportExcel.Enabled = CanExport;

            // ── Events ───────────────────────────────────────────────────
            btnImportExcel.Click += (s, e) => OpenImport();
            btnAdd.Click += (s, e) => OpenEdit(true);
            btnEdit.Click += (s, e) => OpenEdit(false);
            btnDelete.Click += (s, e) => DeleteSelected();
            btnHistory.Click += (s, e) => OpenHistory();
            btnExportCSV.Click += (s, e) => DoExportCSV();
            btnExportExcel.Click += (s, e) => DoExportExcel();
            btnRefresh.Click += (s, e) => LoadGrid();
            btnSearch.Click += (s, e) => ApplyFilters();
            btnClear.Click += (s, e) => ClearFilters();
            txtKeyword.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) ApplyFilters(); };
            cmbGroupCol.SelectedIndexChanged += (s, e) => RefreshGroupVal();
            cmbFwd.SelectedIndexChanged += (s, e) => ApplyFilters();
            cmbContType.SelectedIndexChanged += (s, e) => ApplyFilters();
            cmbEmployee.SelectedIndexChanged += (s, e) => ApplyFilters();
            cmbFeeStatus.SelectedIndexChanged += (s, e) => ApplyFilters();
            cmbGroupVal.SelectedIndexChanged += (s, e) => ApplyFilters();
        }

        // ════════════════════════════════════════════════════════════════
        // Load & bind
        // ════════════════════════════════════════════════════════════════
        private void LoadGrid()
        {
            try
            {
                dgv.DataSource = null;
                _fullData = _repo.GetAll();
                PopulateFilterCombos();
                BindGrid(_fullData);
                SetStatus($"{_fullData.Rows.Count} records.", Color.DarkSlateGray);
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
            var view = dt.Copy();
            dgv.DataSource = view;

            // Hide internal ID but keep accessible
            if (dgv.Columns.Contains("Invoice_Id"))
                dgv.Columns["Invoice_Id"].Visible = false;

            // Apply friendly headers
            foreach (var kv in HEADERS)
                if (dgv.Columns.Contains(kv.Key))
                    dgv.Columns[kv.Key].HeaderText = kv.Value;

            // Freeze Invoice_no (first visible)
            if (dgv.Columns.Contains("Invoice_no"))
            {
                dgv.Columns["Invoice_no"].DisplayIndex = 0;
                dgv.Columns["Invoice_no"].Frozen = true;
                dgv.Columns["Invoice_no"].DefaultCellStyle.Font =
                    new Font("Segoe UI", 8, FontStyle.Bold);
            }

            // Highlight grand total columns
            foreach (string col in new[]
                {"Invoice_grandTotalVND","Invoice_grandTotalUSD","Invoice_subTotalOcean"})
                if (dgv.Columns.Contains(col))
                {
                    dgv.Columns[col].DefaultCellStyle.BackColor = Color.LightYellow;
                    dgv.Columns[col].DefaultCellStyle.Font =
                        new Font("Segoe UI", 8, FontStyle.Bold);
                }

            // Colour fee status
            dgv.CellFormatting += (s, e) =>
            {
                if (e.RowIndex < 0 || !dgv.Columns[e.ColumnIndex].Name.Equals("Invoice_feeStatus")) return;
                int? v = e.Value as int?;
                if (v == null && e.Value != null && int.TryParse(e.Value.ToString(), out int iv)) v = iv;
                if (v == null) return;
                e.CellStyle.BackColor = v == 1 ? Color.LightGreen
                                      : v == 2 ? Color.LightYellow
                                      : Color.FromArgb(255, 200, 200);
                e.CellStyle.ForeColor = Color.Black;
            };
        }

        // ════════════════════════════════════════════════════════════════
        // Filter combos
        // ════════════════════════════════════════════════════════════════
        private void PopulateFilterCombos()
        {
            PopDist(cmbFwd, "Invoice_fwdName", "(All FWD)");
            PopDist(cmbContType, "Invoice_contType", "(All Cont)");
            PopDist(cmbEmployee, "Invoice_employee", "(All Staff)");

            cmbFeeStatus.Items.Clear();
            cmbFeeStatus.Items.AddRange(new object[]
                {"(All Status)","1 - Paid","2 - Acct","3 - Not yet"});
            cmbFeeStatus.SelectedIndex = 0;

            // Group-by columns
            cmbGroupCol.Items.Clear();
            cmbGroupCol.Items.Add("(No grouping)");
            foreach (string c in new[]
            {
                "Invoice_fwdName","Invoice_contType","Invoice_employee",
                "Invoice_shippingTerm","Invoice_billType","Invoice_line",
                "Invoice_customType","Invoice_co"
            })
                cmbGroupCol.Items.Add(c);
            if (cmbGroupCol.SelectedIndex < 0) cmbGroupCol.SelectedIndex = 0;

            // Search-in
            cmbSearchCol.Items.Clear();
            cmbSearchCol.Items.Add("(All columns)");
            foreach (DataColumn c in _fullData.Columns)
                if (c.ColumnName != "Invoice_Id")
                    cmbSearchCol.Items.Add(c.ColumnName);
            if (cmbSearchCol.SelectedIndex < 0) cmbSearchCol.SelectedIndex = 0;

            dtpFrom.Value = DateTime.Today.AddYears(-1);
            dtpTo.Value = DateTime.Today;
        }

        private void PopDist(ComboBox cmb, string col, string all)
        {
            string cur = cmb.SelectedItem?.ToString();
            cmb.Items.Clear(); cmb.Items.Add(all);
            if (_fullData != null && _fullData.Columns.Contains(col))
                foreach (var v in _fullData.AsEnumerable()
                    .Select(r => r[col]?.ToString() ?? "")
                    .Where(v => !string.IsNullOrWhiteSpace(v))
                    .Distinct().OrderBy(v => v))
                    cmb.Items.Add(v);
            cmb.SelectedIndex = cur != null && cmb.Items.Contains(cur)
                ? cmb.Items.IndexOf(cur) : 0;
        }

        private void RefreshGroupVal()
        {
            cmbGroupVal.Items.Clear(); cmbGroupVal.Items.Add("(All)");
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
            string cont = cmbContType.SelectedItem?.ToString() ?? "";
            string emp = cmbEmployee.SelectedItem?.ToString() ?? "";
            string status = cmbFeeStatus.SelectedItem?.ToString() ?? "";
            string groupCol = cmbGroupCol.SelectedItem?.ToString() ?? "";
            string groupVal = cmbGroupVal.SelectedItem?.ToString() ?? "";
            string srchCol = cmbSearchCol.SelectedItem?.ToString() ?? "";
            string kw = txtKeyword.Text.Trim().ToUpper();
            bool useDates = chkDate.Checked;
            DateTime from = dtpFrom.Value.Date;
            DateTime to = dtpTo.Value.Date.AddDays(1).AddTicks(-1);

            bool useGroup = !groupCol.StartsWith("(") && !groupVal.StartsWith("(")
                            && _fullData.Columns.Contains(groupCol);
            bool useSrch = !srchCol.StartsWith("(") && _fullData.Columns.Contains(srchCol);

            // Parse status filter
            int? filterStatus = null;
            if (!status.StartsWith("(") && status.Length > 0)
                if (int.TryParse(status.Substring(0, 1), out int fs)) filterStatus = fs;

            var filtered = _fullData.Clone();
            foreach (DataRow row in _fullData.Rows)
            {
                if (!fwd.StartsWith("(") &&
                    !Eq(row["Invoice_fwdName"], fwd)) continue;
                if (!cont.StartsWith("(") &&
                    !Eq(row["Invoice_contType"], cont)) continue;
                if (!emp.StartsWith("(") &&
                    !Eq(row["Invoice_employee"], emp)) continue;
                if (filterStatus.HasValue &&
                    !Eq(row["Invoice_feeStatus"], filterStatus.Value.ToString())) continue;
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
            if (!emp.StartsWith("(")) parts.Add($"Staff={emp}");
            if (filterStatus.HasValue) parts.Add($"Status={filterStatus}");
            if (useGroup) parts.Add($"{groupCol}={groupVal}");
            if (!string.IsNullOrEmpty(kw)) parts.Add($"Keyword='{txtKeyword.Text.Trim()}'");
            if (useDates) parts.Add($"{dtpFrom.Value:dd/MM/yy}–{dtpTo.Value:dd/MM/yy}");
            SetStatus(string.Join(" · ", parts), Color.SteelBlue);
        }

        private void ClearFilters()
        {
            cmbFwd.SelectedIndex = cmbContType.SelectedIndex =
            cmbEmployee.SelectedIndex = cmbFeeStatus.SelectedIndex =
            cmbGroupCol.SelectedIndex = cmbGroupVal.SelectedIndex =
            cmbSearchCol.SelectedIndex = 0;
            txtKeyword.Text = "";
            chkDate.Checked = false;
            BindGrid(_fullData);
            SetStatus($"{_fullData?.Rows.Count ?? 0} records.", Color.DarkSlateGray);
        }

        private static bool Eq(object cell, string val) =>
            string.Equals(cell?.ToString(), val, StringComparison.OrdinalIgnoreCase);

        // ════════════════════════════════════════════════════════════════
        // CRUD
        // ════════════════════════════════════════════════════════════════
        private void OpenEdit(bool isNew)
        {
            if (!CanEdit) return;
            InvoiceModel existing = null;
            if (!isNew)
            {
                if (dgv.SelectedRows.Count == 0)
                { CTMessageBox.Show("Select a row to edit."); return; }
                string id = ((DataTable)dgv.DataSource)
                    .Rows[dgv.SelectedRows[0].Index]["Invoice_Id"]?.ToString();
                existing = _repo.GetById(id);
            }
            using (var dlg = new InvoiceEditForm(existing, _user.User_code, isNew, GetEditScope()))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    try
                    {
                        if (isNew) _repo.Insert(dlg.Result, _user.User_code);
                        else _repo.Update(dlg.Result, _user.User_code);
                        LoadGrid();
                        SetStatus(isNew ? "Record added." : "Record updated.", Color.DarkGreen);
                    }
                    catch (Exception ex)
                    {
                        CTMessageBox.Show("Save error:\r\n" + ex.Message, "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void DeleteSelected()
        {
            if (!CanDelete) return;
            if (dgv.SelectedRows.Count == 0)
            { CTMessageBox.Show("Select a row to delete."); return; }
            if (CTMessageBox.Show($"Delete {dgv.SelectedRows.Count} record(s)?",
                "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                != DialogResult.Yes) return;

            var dt = (DataTable)dgv.DataSource;
            foreach (DataGridViewRow gr in dgv.SelectedRows)
            {
                string id = dt.Rows[gr.Index]["Invoice_Id"]?.ToString();
                if (!string.IsNullOrEmpty(id)) _repo.Delete(id, _user.User_code);
            }
            LoadGrid();
        }

        private void OpenImport()
        {
            using (var dlg = new InvoiceExcelImportForm(_repo, _user.User_code))
            {
                dlg.ShowDialog(this);
                LoadGrid();
            }
        }

        private void OpenHistory()
        {
            string id = null, dispName = null;
            if (dgv.SelectedRows.Count > 0)
            {
                var dt = (DataTable)dgv.DataSource;
                var dr = dt.Rows[dgv.SelectedRows[0].Index];
                id = dr["Invoice_Id"]?.ToString();
                dispName = dr["Invoice_no"]?.ToString();
            }
            using (var dlg = new InvoiceHistoryForm(_repo, _user.User_code, id, dispName))
                if (dlg.ShowDialog(this) == DialogResult.OK) LoadGrid();
        }

        // ════════════════════════════════════════════════════════════════
        // Export
        // ════════════════════════════════════════════════════════════════
        private void DoExportCSV()
        {
            if (dgv.Rows.Count == 0) { CTMessageBox.Show("No data."); return; }
            using (var sfd = new SaveFileDialog
            {
                Filter = "CSV|*.csv",
                FileName = $"Invoice_{DateTime.Now:yyyyMMdd}.csv"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;
                var vis = dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                var sb = new StringBuilder();
                sb.AppendLine(string.Join(",", vis.Select(c => $"\"{c.HeaderText}\"")));
                foreach (DataGridViewRow row in dgv.Rows)
                    sb.AppendLine(string.Join(",",
                        vis.Select(c => $"\"{row.Cells[c.Name].Value}\"")));
                File.WriteAllText(sfd.FileName, sb.ToString(), System.Text.Encoding.UTF8);
                SetStatus("CSV exported.", Color.DarkGreen);
                System.Diagnostics.Process.Start(sfd.FileName);
            }
        }

        private void DoExportExcel()
        {
            if (dgv.Rows.Count == 0) { CTMessageBox.Show("No data."); return; }
            using (var sfd = new SaveFileDialog
            {
                Filter = "Excel|*.xlsx",
                FileName = $"Invoice_{DateTime.Now:yyyyMMdd}.xlsx"
            })
            {
                if (sfd.ShowDialog() != DialogResult.OK) return;
                Excel.Application xl = null; Excel.Workbook wb = null;
                try
                {
                    xl = new Excel.Application { Visible = false, DisplayAlerts = false };
                    wb = xl.Workbooks.Add();
                    var ws = (Excel.Worksheet)wb.Sheets[1];
                    var vis = dgv.Columns.Cast<DataGridViewColumn>().Where(c => c.Visible).ToList();
                    for (int c = 0; c < vis.Count; c++)
                    {
                        var cell = (Excel.Range)ws.Cells[1, c + 1];
                        cell.Value2 = vis[c].HeaderText; cell.Font.Bold = true;
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.FromArgb(255, 153, 0));
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

        private static Label Lbl(string text, int x, int y) =>
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