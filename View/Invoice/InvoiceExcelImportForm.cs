using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;
using techlink_workspace.Repositories.InvoiceRepo;

namespace techlink_workspace.View.Invoice
{
    /// <summary>
    /// Excel → DB import wizard (Admin only).
    /// Step 1 — pick file / sheet.
    /// Step 2 — map Excel cols → DB cols, preview, import.
    /// On duplicate Invoice_no: UPDATE existing row instead of insert.
    /// </summary>
    public class InvoiceExcelImportForm : Form
    {
        private readonly InvoiceRepository _repo;
        private readonly string _byUser;

        // ── Friendly labels for mapping panel (DB prop → display label) ──
        // Format: "Friendly name (fieldName)"  — excludes audit + ERP cols
        private static readonly List<(string Prop, string Label)> IMPORT_FIELDS =
            new List<(string, string)>
        {
            ("Invoice_no",               "Invoice # (no)"),
            ("Invoice_shippingTerm",     "Shipping Term (shippingTerm)"),
            ("Invoice_paymentTerm",      "Payment Term (paymentTerm)"),
            ("Invoice_employee",         "NV Logistics phụ trách (employee)"),
            ("Invoice_logisticRemark",   "Lưu ý đơn hàng (logisticRemark)"),
            ("Invoice_confirmDate",      "Ngày xưởng confirm (confirmDate)"),
            ("Invoice_fwdName",          "FWD – Tên FWD (fwdName)"),
            ("Invoice_bookingNo",        "Số Booking (bookingNo)"),
            ("Invoice_contType",         "Loại cont (contType)"),
            ("Invoice_vgmCO",            "SI & VGM cut-off (vgmCO)"),
            ("Invoice_cyCO",             "CY cut-off (cyCO)"),
            ("Invoice_etd",              "ETD (etd)"),
            ("Invoice_eta",              "ETA (eta)"),
            ("Invoice_billType",         "Loại Bill (billType)"),
            ("Invoice_billNo",           "Số Bill (billNo)"),
            ("Invoice_co",               "CO (co)"),
            ("Invoice_coNo",             "Số CO (coNo)"),
            ("Invoice_OF",               "OF (OF)"),
            ("Invoice_deliveryCharges",  "Delivery charges (deliveryCharges)"),
            ("Invoice_taxes",            "Duty/Taxes (taxes)"),
            ("Invoice_otherDestCharges", "Other destination charges (otherDestCharges)"),
            ("Invoice_thc",              "THC (thc)"),
            ("Invoice_blFee",            "b/l fee (blFee)"),
            ("Invoice_seal",             "Seal (seal)"),
            ("Invoice_telexRelease",     "Telex release (telexRelease)"),
            ("Invoice_cfs",              "CFS (cfs)"),
            ("Invoice_vgmFee",           "VGM (vgmFee)"),
            ("Invoice_ensebsams",        "ENS/EBS/AMS (ensebsams)"),
            ("Invoice_other",            "OTHERS (other)"),
            ("Invoice_totalVND",         "TOTAL (VND) (totalVND)"),
            ("Invoice_subTotalOcean",    "SUB-TOTAL OCEAN (USD) (subTotalOcean)"),
            ("Invoice_coFee",            "C/O fee (coFee)"),
            ("Invoice_feeStatus",        "Status paid/Acct/not yet (feeStatus)"),
            ("Invoice_redInvoiceNo",     "RED INVOICE # (redInvoiceNo)"),
            ("Invoice_redInvoiceDate",   "RED INVOICE DATE (redInvoiceDate)"),
            ("Invoice_redInvoiceRecvDate","RED INVOICE RECEIVED DATE (redInvoiceRecvDate)"),
            ("Invoice_transferAccountDate","TRANSFER TO ACCOUNTANT (transferAccountDate)"),
            ("Invoice_trucking",         "Trucking (trucking)"),
            ("Invoice_infrastructureFee","INFRASTRUCTURE FEE (infrastructureFee)"),
            ("Invoice_customerClearance","Customs clearance (customerClearance)"),
            ("Invoice_customFee",        "Customs fee (customFee)"),
            ("Invoice_otherCustomFee",   "Others custom (otherCustomFee)"),
            ("Invoice_subTotalVNDCustom","Sub-Total Trucking+customs VND (subTotalVNDCustom)"),
            ("Invoice_subTotalUSDCustom","Sub-Total Trucking+customs USD (subTotalUSDCustom)"),
            ("Invoice_grandTotalVND",    "GRAND TOTAL VND (grandTotalVND)"),
            ("Invoice_grandTotalUSD",    "GRAND TOTAL USD (grandTotalUSD)"),
            ("Invoice_cdsNo",            "CDS NO (cdsNo)"),
            ("Invoice_cdsDate",          "CDS DATE (cdsDate)"),
            ("Invoice_line",             "luong (line)"),
            ("Invoice_customType",       "Mã loại hình tk (customType)"),
        };

        // ── Controls ─────────────────────────────────────────────────────
        private Button btnPickFile, btnLoadSheet, btnImport, btnClose;
        private Label lblFile, lblStatus;
        private ComboBox cmbSheet;
        private DataGridView dgvPreview;
        private FlowLayoutPanel flowMap;
        private DataTable _sheetData;
        private List<string> _excelCols;
        private readonly Dictionary<string, ComboBox> _mapCombos =
            new Dictionary<string, ComboBox>();

        private Excel.Application _xl;
        private Excel.Workbook _wb;

        public InvoiceExcelImportForm(InvoiceRepository repo, string byUser)
        {
            _repo = repo;
            _byUser = byUser;
            Build();
        }

        // ════════════════════════════════════════════════════════════════
        private void Build()
        {
            Text = "Import Invoice Data from Excel";
            ClientSize = new Size(1400, 780);
            MinimumSize = new Size(1100, 600);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 8.5f);
            FormClosing += (s, e) => ReleaseExcel();

            // ── Top toolbar ───────────────────────────────────────────────
            var top = new Panel { Dock = DockStyle.Top, Height = 44, BackColor = Color.White };
            top.Paint += BorderPaint;

            btnPickFile = Btn("📂 Choose File", Color.FromArgb(0, 51, 153));
            btnPickFile.Location = new Point(8, 8);
            btnPickFile.Click += BtnPickFile_Click;

            lblFile = new Label
            {
                Location = new Point(148, 14),
                AutoSize = true,
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                Text = "No file selected"
            };

            var lblSh = new Label
            {
                Text = "Sheet:",
                Location = new Point(600, 14),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
            };

            cmbSheet = new ComboBox
            {
                Location = new Point(644, 10),
                Width = 160,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 8.5f)
            };

            btnLoadSheet = Btn("Load Sheet", Color.SteelBlue);
            btnLoadSheet.Location = new Point(812, 8);
            btnLoadSheet.Enabled = false;
            btnLoadSheet.Click += BtnLoadSheet_Click;

            top.Controls.AddRange(new Control[]
                { btnPickFile, lblFile, lblSh, cmbSheet, btnLoadSheet });

            // ── Status bar ───────────────────────────────────────────────
            lblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 22,
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                BackColor = Color.LightYellow,
                ForeColor = Color.DarkSlateGray,
                TextAlign = ContentAlignment.MiddleLeft,
                Padding = new Padding(6, 0, 0, 0),
                BorderStyle = BorderStyle.FixedSingle,
                Text = "Step 1: Choose an Excel file."
            };

            // ── Bottom buttons ───────────────────────────────────────────
            var bot = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 46,
                BackColor = Color.WhiteSmoke
            };
            bot.Paint += BorderPaint;

            btnImport = Btn("⬆ Import to Database", Color.DarkGreen);
            btnImport.Location = new Point(10, 8);
            btnImport.Enabled = false;
            btnImport.Click += BtnImport_Click;

            btnClose = Btn("Close", Color.DimGray);
            btnClose.Location = new Point(200, 8);
            btnClose.Click += (s, e) => Close();
            bot.Controls.AddRange(new Control[] { btnImport, btnClose });

            // ── Split: left=mapping, right=preview ───────────────────────
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 500,
                Panel1MinSize = 400
            };

            // ── Mapping panel ─────────────────────────────────────────────
            var mappingHdr = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                Text = "  Column Mapping  (DB Field → Excel Column)",
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                BackColor = Color.FromArgb(30, 40, 60),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };

            flowMap = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(4),
                BackColor = Color.FromArgb(248, 248, 252)
            };

            var mapPanel = new Panel { Dock = DockStyle.Fill };
            mapPanel.Controls.Add(flowMap);
            mapPanel.Controls.Add(mappingHdr);
            split.Panel1.Controls.Add(mapPanel);

            // ── Preview panel (larger) ────────────────────────────────────
            var previewHdr = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                Text = "  Excel Preview (first 50 rows)",
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                BackColor = Color.FromArgb(30, 40, 60),
                ForeColor = Color.White,
                TextAlign = ContentAlignment.MiddleLeft
            };

            dgvPreview = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells,
                BackgroundColor = Color.White,
                RowHeadersVisible = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                Font = new Font("Segoe UI", 8f)
            };
            dgvPreview.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 153, 0);
            dgvPreview.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            dgvPreview.ColumnHeadersHeight = 32;
            dgvPreview.EnableHeadersVisualStyles = false;

            split.Panel2.Controls.Add(dgvPreview);
            split.Panel2.Controls.Add(previewHdr);

            Controls.Add(split);
            Controls.Add(bot);
            Controls.Add(lblStatus);
            Controls.Add(top);
        }

        // ════════════════════════════════════════════════════════════════
        // Step 1 — choose file
        // ════════════════════════════════════════════════════════════════
        private string _filePath; // store path so we can re-open for sheet read

        private void BtnPickFile_Click(object sender, EventArgs e)
        {
            using (var ofd = new OpenFileDialog
            { Filter = "Excel|*.xls*", Title = "Select Excel file to import" })
            {
                if (ofd.ShowDialog() != DialogResult.OK) return;
                _filePath = ofd.FileName;
            }

            ReleaseExcel(); // close any previously open workbook
            try
            {
                var xl = new Excel.Application { Visible = false, DisplayAlerts = false };
                var wb = xl.Workbooks.Open(_filePath, ReadOnly: true);

                cmbSheet.Items.Clear();
                foreach (Excel.Worksheet ws in wb.Sheets)
                    cmbSheet.Items.Add(ws.Name);
                if (cmbSheet.Items.Count > 0) cmbSheet.SelectedIndex = 0;

                // Release immediately — only needed sheet names
                wb.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                xl.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);
                GC.Collect();
                GC.WaitForPendingFinalizers();

                lblFile.Text = System.IO.Path.GetFileName(_filePath);
                btnLoadSheet.Enabled = true;
                SetStatus("Step 2: Choose a sheet and click Load Sheet.", Color.SteelBlue);
            }
            catch (Exception ex)
            {
                CTMessageBox.Show("Could not open file:\r\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ════════════════════════════════════════════════════════════════
        // Step 2 — load sheet
        // ════════════════════════════════════════════════════════════════
        private void BtnLoadSheet_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(_filePath) || cmbSheet.SelectedItem == null) return;
            string sheetName = cmbSheet.SelectedItem.ToString();
            try
            {
                // Re-open file just to read data, then release immediately
                Excel.Application xl = null;
                Excel.Workbook wb = null;
                try
                {
                    xl = new Excel.Application { Visible = false, DisplayAlerts = false };
                    wb = xl.Workbooks.Open(_filePath, ReadOnly: true);
                    var ws = (Excel.Worksheet)wb.Sheets[sheetName];
                    _sheetData = ReadSheet(ws, out _excelCols);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                }
                finally
                {
                    // Always release Excel before touching DB or UI
                    if (wb != null) { wb.Close(false); System.Runtime.InteropServices.Marshal.ReleaseComObject(wb); }
                    if (xl != null) { xl.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(xl); }
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

                // Preview: up to 50 rows — all COM is gone by here
                dgvPreview.DataSource = _sheetData.Rows.Count > 50
                    ? _sheetData.AsEnumerable().Take(50).CopyToDataTable()
                    : _sheetData;

                BuildMappingUI();
                btnImport.Enabled = true;
                SetStatus(
                    $"Sheet '{sheetName}' loaded: {_sheetData.Rows.Count} rows, " +
                    $"{_excelCols.Count} columns.  Step 3: Map columns then click Import.",
                    Color.DarkGreen);
            }
            catch (Exception ex)
            {
                CTMessageBox.Show("Sheet read error:\r\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ════════════════════════════════════════════════════════════════
        // Build mapping rows
        // ════════════════════════════════════════════════════════════════
        private void BuildMappingUI()
        {
            flowMap.Controls.Clear();
            _mapCombos.Clear();

            const string SKIP = "-- Skip --";
            var opts = new List<string> { SKIP };
            opts.AddRange(_excelCols);

            foreach (var (prop, label) in IMPORT_FIELDS)
            {
                var row = new Panel
                {
                    Size = new Size(490, 26),
                    Margin = new Padding(0, 1, 0, 1)
                };

                var lbl = new Label
                {
                    Text = label,
                    Location = new Point(0, 5),
                    Size = new Size(290, 18),
                    Font = new Font("Segoe UI", 8f),
                    ForeColor = Color.FromArgb(30, 30, 70)
                };

                var cmb = new ComboBox
                {
                    Location = new Point(295, 2),
                    Width = 190,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Font = new Font("Segoe UI", 8f)
                };
                cmb.Items.AddRange(opts.ToArray());

                // Always default to Skip — let user map manually to avoid confusion
                cmb.SelectedIndex = 0;

                row.Controls.AddRange(new Control[] { lbl, cmb });
                flowMap.Controls.Add(row);
                _mapCombos[prop] = cmb;
            }
        }

        // ════════════════════════════════════════════════════════════════
        // Step 3 — Import
        // ════════════════════════════════════════════════════════════════
        private void BtnImport_Click(object sender, EventArgs e)
        {
            if (_sheetData == null || _sheetData.Rows.Count == 0)
            { CTMessageBox.Show("No data to import."); return; }

            var confirm = CTMessageBox.Show(
                $"Import {_sheetData.Rows.Count} rows?\n" +
                "• New Invoice_no → INSERT\n• Existing Invoice_no → UPDATE",
                "Confirm Import", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm != DialogResult.Yes) return;

            // Must be declared here — before the loop that uses them
            var now = DateTime.Now;
            var props = typeof(InvoiceModel).GetProperties()
                            .ToDictionary(p => p.Name);

            btnImport.Enabled = false;
            SetStatus("Importing…", Color.Blue);
            Application.DoEvents();

            var errors = new System.Text.StringBuilder();
            int inserted = 0, updated = 0, failed = 0;

            foreach (DataRow row in _sheetData.Rows)
            {
                try
                {
                    var m = new InvoiceModel
                    {
                        createdate = now,
                        createby = _byUser,
                        updatedate = now,
                        updateby = _byUser
                    };

                    foreach (var (prop, _) in IMPORT_FIELDS)
                    {
                        if (!_mapCombos.TryGetValue(prop, out var cmb)) continue;
                        string excelCol = cmb.SelectedItem?.ToString();
                        if (excelCol == "-- Skip --" ||
                            !_sheetData.Columns.Contains(excelCol)) continue;
                        string raw = row[excelCol]?.ToString()?.Trim() ?? "";
                        if (string.IsNullOrEmpty(raw)) continue;
                        if (props.TryGetValue(prop, out var pi))
                            SetProperty(m, pi, raw);
                    }

                    if (string.IsNullOrWhiteSpace(m.Invoice_no))
                    { failed++; continue; }

                    var existing = _repo.GetByInvoiceNo(m.Invoice_no);
                    if (existing != null)
                    {
                        m.Invoice_Id = existing.Invoice_Id;
                        m.createdate = existing.createdate;
                        m.createby = existing.createby;
                        _repo.Update(m, _byUser);
                        updated++;
                    }
                    else
                    {
                        m.Invoice_Id = Guid.NewGuid().ToString();
                        _repo.Insert(m, _byUser);
                        inserted++;
                    }
                }
                catch (Exception ex)
                {
                    failed++;
                    if (failed <= 5)
                        errors.AppendLine(ex.Message);
                }
            }

            string msg = $"Import complete: {inserted} inserted, {updated} updated, {failed} failed.";
            if (errors.Length > 0) msg += $"\n\nFirst errors:\n{errors}";
            SetStatus(msg.Split('\n')[0], failed > 0 ? Color.OrangeRed : Color.DarkGreen);
            CTMessageBox.Show(msg, "Done", MessageBoxButtons.OK,
                failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }

        // ════════════════════════════════════════════════════════════════
        // ReadSheet — Value2 bulk read + OA date detection
        // ════════════════════════════════════════════════════════════════
        private static DataTable ReadSheet(Excel.Worksheet ws, out List<string> cols)
        {
            var range = ws.UsedRange;
            int rows = range.Rows.Count, c = range.Columns.Count;
            object[,] raw;
            try { raw = (object[,])range.Value2; }
            finally { System.Runtime.InteropServices.Marshal.ReleaseComObject(range); }

            cols = new List<string>();
            var dt = new DataTable();

            for (int j = 1; j <= c; j++)
            {
                string h = raw[1, j]?.ToString().Trim() ?? $"Col{j}";
                if (string.IsNullOrWhiteSpace(h)) h = $"Col{j}";
                string orig = h; int n = 1;
                while (dt.Columns.Contains(h)) h = $"{orig}_{n++}";
                dt.Columns.Add(h);
                cols.Add(h);
            }

            // Detect OA-date columns by sampling
            var dateCols = new HashSet<int>();
            for (int j = 1; j <= c; j++)
            {
                int hits = 0, total = 0;
                for (int r = 2; r <= Math.Min(rows, 8); r++)
                {
                    var v = raw[r, j];
                    if (v == null) continue;
                    total++;
                    if (v is double d && d > 1 && d < 2958466) hits++;
                }
                if (total > 0 && hits == total) dateCols.Add(j);
            }

            for (int r = 2; r <= rows; r++)
            {
                var dr = dt.NewRow();
                bool hasData = false;
                for (int j = 1; j <= c; j++)
                {
                    var cellVal = raw[r, j];
                    string v;
                    if (cellVal == null)
                        v = "";
                    else if (dateCols.Contains(j) && cellVal is double oaD
                             && oaD > 1 && oaD < 2958466)
                    {
                        try { v = DateTime.FromOADate(oaD).ToString("dd/MM/yyyy"); }
                        catch { v = cellVal.ToString(); }
                    }
                    else
                        v = cellVal.ToString().Trim();

                    dr[j - 1] = v;
                    if (!string.IsNullOrEmpty(v)) hasData = true;
                }
                if (hasData) dt.Rows.Add(dr);
            }
            return dt;
        }

        // ════════════════════════════════════════════════════════════════
        // SetProperty — handles OA dates, nullable int from "1.0" strings
        // ════════════════════════════════════════════════════════════════
        private static void SetProperty(InvoiceModel m, PropertyInfo pi, string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return;
            Type t = Nullable.GetUnderlyingType(pi.PropertyType) ?? pi.PropertyType;
            object val;

            if (t == typeof(string))
            {
                val = raw;
            }
            else if (t == typeof(DateTime))
            {
                if (DateTime.TryParse(raw, out var dt))
                    val = dt;
                else if (double.TryParse(raw,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture,
                    out var oa) && oa > 1 && oa < 2958466)
                    val = DateTime.FromOADate(oa);
                else return;
            }
            else if (t == typeof(double))
            {
                string clean = raw.Replace(",", "").Trim();
                if (double.TryParse(clean,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var d))
                    val = d;
                else return;
            }
            else if (t == typeof(int))
            {
                if (int.TryParse(raw, out var i))
                    val = i;
                else if (double.TryParse(raw,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var d))
                    val = (int)d;
                else return;
            }
            else return;

            pi.SetValue(m, val);
        }

        // ════════════════════════════════════════════════════════════════
        private void ReleaseExcel()
        {
            try
            {
                if (_wb != null)
                {
                    _wb.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_wb);
                    _wb = null;
                }
                if (_xl != null)
                {
                    _xl.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_xl);
                    _xl = null;
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }

        private void SetStatus(string msg, Color color)
        { lblStatus.Text = msg; lblStatus.ForeColor = color; }

        private static void BorderPaint(object sender, PaintEventArgs e)
        {
            var p = (Panel)sender;
            using (var pen = new System.Drawing.Pen(Color.FromArgb(220, 220, 230)))
                e.Graphics.DrawLine(pen, 0, p.Height - 1, p.Width, p.Height - 1);
        }

        private static Button Btn(string text, Color bg)
        {
            int w = Math.Max(80, text.Length * 7 + 16);
            var b = new Button
            {
                Text = text,
                Size = new Size(w, 28),
                BackColor = bg,
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 8f)
            };
            b.FlatAppearance.BorderSize = 0;
            return b;
        }
    }
}