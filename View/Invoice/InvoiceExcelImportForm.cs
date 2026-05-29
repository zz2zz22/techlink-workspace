using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using techlink_workspace.Controller.Logic.IDGenerate;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;
using techlink_workspace.Repositories.InvoiceRepo;
using techlink_workspace.Repositories.PICRepo;

namespace techlink_workspace.View.Invoice
{
    /// <summary>
    /// ERP Excel import – imports only the 14 ERP columns (Document 2).
    /// Each row gets a new Invoice_Id from UUIDGenerator.
    /// Invoice_logisticPersonInCharge is auto-set from PIC Management by customer code.
    /// Invoice_no defaults to Invoice_poNo and can be updated later.
    /// Duplicate detection: poNo + itemCode → UPDATE existing ERP fields.
    /// </summary>
    public class InvoiceExcelImportForm : Form
    {
        private readonly InvoiceRepository _repo;
        private readonly PICRepository _picRepo = new PICRepository();
        private readonly string _byUser;

        // ERP-only import fields (Document 2 columns)
        private static readonly List<(string Prop, string Label, string[] Keywords)> ERP_FIELDS =
            new List<(string, string, string[])>
        {
            ("Invoice_customerCode",        "Mã KH – Customer Code",
                new[]{"mã kh","customer code","ma kh"}),
            ("Invoice_customerName",        "KH Tên gọi tắt – Customer Name",
                new[]{"tên goi","customer name","ten goi"}),
            ("Invoice_customerRequestDate", "Ngày dự kiến giao hàng – Request Date",
                new[]{"ngày dự kiến","request date","du kien"}),
            ("Invoice_poNo",                "Mã ĐĐH của KH – PO Number",
                new[]{"mã đĐh","po number","po no","ddh"}),
            ("Invoice_poDate",              "Ngày ĐĐH – PO Date",
                new[]{"ngày đĐh","po date","ngay ddh"}),
            ("Invoice_brand", "Bộ Phận – Brand",
                new[]{ "bộ phận","brand","bo phan","department" }),
            ("Invoice_saleName",            "NVBH – Sale Name",
                new[]{"nvbh","sale name","salesperson","nv bh"}),
            ("Invoice_factoryNo",           "Mã Xưởng – Factory Number",
                new[]{"mã xưởng","factory no","factory number","ma xuong"}),
            ("Invoice_factoryName",         "Tên Xưởng – Factory Name",
                new[]{"tên xưởng","factory name","ten xuong"}),
            ("Invoice_itemCode",            "Mã SP – Item Code",
                new[]{"mã sp ","item code ","ma sp "}),   // trailing space avoids "mã sp kh"
            ("Invoice_itemCodeCustomers",   "Mã SP KH – Item Code Customers",
                new[]{"mã sp kh","item code customer","ma sp kh"}),
            ("Invoice_itemName",            "Tên SP – Item Name",
                new[]{"tên sp","item name","ten sp"}),
            ("Invoice_quantity",            "SL Đơn Đặt – Quantity",
                new[]{"sl đơn","quantity","sl don","sl "}),
            ("Invoice_unit",                "ĐV – Unit",
                new[]{"đv","unit","đơn vị","don vi"}),
        };

        private Button btnPickFile, btnLoadSheet, btnImport, btnClose;
        private Label lblFile, lblStatus;
        private ComboBox cmbSheet;
        private DataGridView dgvPreview;
        private FlowLayoutPanel flowMap;
        private DataTable _sheetData;
        private List<string> _excelCols;
        private readonly Dictionary<string, ComboBox> _mapCombos = new Dictionary<string, ComboBox>();
        private string _filePath;

        public InvoiceExcelImportForm(InvoiceRepository repo, string byUser)
        {
            _repo = repo; _byUser = byUser;
            Build();
        }

        private void Build()
        {
            Text = "Import ERP Invoice Data from Excel (COPR21)";
            ClientSize = new Size(1300, 700);
            MinimumSize = new Size(1000, 500);
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 8.5f);

            // ── Top toolbar ────────────────────────────────────────────────
            var top = new Panel { Dock = DockStyle.Top, Height = 44, BackColor = Color.White };
            top.Paint += BorderPaint;

            btnPickFile = Btn("📂 Choose File", Color.FromArgb(0, 51, 153));
            btnPickFile.Location = new Point(8, 8);
            btnPickFile.Click += BtnPickFile_Click;

            lblFile = new Label
            {
                Location = new Point(158, 14),
                AutoSize = true,
                ForeColor = Color.DimGray,
                Font = new Font("Segoe UI", 8f, FontStyle.Italic),
                Text = "No file selected"
            };

            top.Controls.Add(new Label
            {
                Text = "Sheet:",
                Location = new Point(560, 14),
                AutoSize = true,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
            });

            cmbSheet = new ComboBox
            {
                Location = new Point(600, 10),
                Width = 160,
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 8.5f)
            };

            btnLoadSheet = Btn("Load Sheet", Color.SteelBlue);
            btnLoadSheet.Location = new Point(770, 8);
            btnLoadSheet.Enabled = false;
            btnLoadSheet.Click += BtnLoadSheet_Click;

            top.Controls.AddRange(new Control[] { btnPickFile, lblFile, cmbSheet, btnLoadSheet });

            // ── Status bar ─────────────────────────────────────────────────
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
                Text = "Step 1: Choose the COPR21 Excel file."
            };

            // ── Bottom buttons ──────────────────────────────────────────────
            var bot = new Panel { Dock = DockStyle.Bottom, Height = 46, BackColor = Color.WhiteSmoke };
            bot.Paint += BorderPaint;

            btnImport = Btn("⬆ Import to Database", Color.DarkGreen);
            btnImport.Location = new Point(10, 8);
            btnImport.Enabled = false;
            btnImport.Click += BtnImport_Click;

            btnClose = Btn("Close", Color.DimGray);
            btnClose.Location = new Point(210, 8);
            btnClose.Click += (s, e) => Close();
            bot.Controls.AddRange(new Control[] { btnImport, btnClose });

            // ── Split: mapping | preview ────────────────────────────────────
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical,
                SplitterDistance = 480
            };

            var mapHdr = new Label
            {
                Dock = DockStyle.Top,
                Height = 28,
                Text = "  Column Mapping (ERP field → Excel column)",
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
            mapPanel.Controls.Add(mapHdr);
            split.Panel1.Controls.Add(mapPanel);

            var prevHdr = new Label
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
                Font = new Font("Segoe UI", 8f)
            };
            dgvPreview.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(255, 153, 0);
            dgvPreview.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 8, FontStyle.Bold);
            dgvPreview.ColumnHeadersHeight = 32;
            dgvPreview.EnableHeadersVisualStyles = false;
            split.Panel2.Controls.Add(dgvPreview);
            split.Panel2.Controls.Add(prevHdr);

            Controls.Add(split);
            Controls.Add(bot);
            Controls.Add(lblStatus);
            Controls.Add(top);
        }

        // ── Step 1: pick file ─────────────────────────────────────────────
        private void BtnPickFile_Click(object s, EventArgs e)
        {
            using (var ofd = new OpenFileDialog { Filter = "Excel|*.xls*", Title = "Select COPR21 Excel" })
            {
                if (ofd.ShowDialog() != DialogResult.OK) return;
                _filePath = ofd.FileName;
            }
            try
            {
                var xl = new Excel.Application { Visible = false, DisplayAlerts = false };
                var wb = xl.Workbooks.Open(_filePath, ReadOnly: true);
                cmbSheet.Items.Clear();
                foreach (Excel.Worksheet ws in wb.Sheets) cmbSheet.Items.Add(ws.Name);
                if (cmbSheet.Items.Count > 0) cmbSheet.SelectedIndex = 0;
                wb.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                xl.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xl);
                GC.Collect(); GC.WaitForPendingFinalizers();
                lblFile.Text = System.IO.Path.GetFileName(_filePath);
                btnLoadSheet.Enabled = true;
                SetStatus("Step 2: Select a sheet and click Load Sheet.", Color.SteelBlue);
            }
            catch (Exception ex)
            {
                CTMessageBox.Show("Could not open file:\r\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Step 2: load sheet ────────────────────────────────────────────
        private void BtnLoadSheet_Click(object s, EventArgs e)
        {
            if (string.IsNullOrEmpty(_filePath) || cmbSheet.SelectedItem == null) return;
            string sheet = cmbSheet.SelectedItem.ToString();
            try
            {
                Excel.Application xl = null; Excel.Workbook wb = null;
                try
                {
                    xl = new Excel.Application { Visible = false, DisplayAlerts = false };
                    wb = xl.Workbooks.Open(_filePath, ReadOnly: true);
                    var ws = (Excel.Worksheet)wb.Sheets[sheet];
                    _sheetData = ReadSheet(ws, out _excelCols);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                }
                finally
                {
                    if (wb != null) { wb.Close(false); System.Runtime.InteropServices.Marshal.ReleaseComObject(wb); }
                    if (xl != null) { xl.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(xl); }
                    GC.Collect(); GC.WaitForPendingFinalizers();
                }

                dgvPreview.DataSource = _sheetData.Rows.Count > 50
                    ? _sheetData.AsEnumerable().Take(50).CopyToDataTable()
                    : _sheetData;

                BuildMappingUI();
                btnImport.Enabled = true;
                SetStatus($"Sheet '{sheet}': {_sheetData.Rows.Count} rows. Step 3: verify mapping then Import.", Color.DarkGreen);
            }
            catch (Exception ex)
            {
                CTMessageBox.Show("Sheet error:\r\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ── Build mapping rows ────────────────────────────────────────────
        private void BuildMappingUI()
        {
            flowMap.Controls.Clear();
            _mapCombos.Clear();
            const string SKIP = "-- Skip --";
            var opts = new List<string> { SKIP };
            opts.AddRange(_excelCols);

            foreach (var (prop, label, keywords) in ERP_FIELDS)
            {
                var row = new Panel { Size = new Size(470, 28), Margin = new Padding(0, 1, 0, 1) };
                row.Controls.Add(new Label
                {
                    Text = label,
                    Location = new Point(0, 6),
                    Size = new Size(290, 18),
                    Font = new Font("Segoe UI", 8f),
                    ForeColor = Color.FromArgb(30, 30, 70)
                });
                var cmb = new ComboBox
                {
                    Location = new Point(295, 3),
                    Width = 170,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Font = new Font("Segoe UI", 8f)
                };
                cmb.Items.AddRange(opts.ToArray());

                // Auto-suggest: find first Excel column matching any keyword
                int autoIdx = 0;
                string lowerLabel = label.ToLower();
                for (int i = 1; i < opts.Count; i++)
                {
                    string colLower = opts[i].ToLower();
                    if (keywords.Any(k => colLower.Contains(k.ToLower())))
                    { autoIdx = i; break; }
                }
                cmb.SelectedIndex = autoIdx;

                row.Controls.Add(cmb);
                flowMap.Controls.Add(row);
                _mapCombos[prop] = cmb;
            }
        }

        // ── Step 3: import ────────────────────────────────────────────────
        private void BtnImport_Click(object s, EventArgs e)
        {
            if (_sheetData == null || _sheetData.Rows.Count == 0)
            { CTMessageBox.Show("No data to import."); return; }

            if (CTMessageBox.Show(
                $"Import {_sheetData.Rows.Count} rows?\n" +
                "• New poNo/itemCode → INSERT\n• Existing → UPDATE ERP fields only",
                "Confirm Import", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                != DialogResult.Yes) return;

            btnImport.Enabled = false;
            SetStatus("Importing…", Color.Blue);
            Application.DoEvents();

            var props = typeof(InvoiceModel).GetProperties().ToDictionary(p => p.Name);
            var now = DateTime.Now;
            var errors = new System.Text.StringBuilder();
            int inserted = 0, updated = 0, failed = 0;

            var records = new List<InvoiceModel>();
            foreach (DataRow row in _sheetData.Rows)
            {
                var m = new InvoiceModel { createdate = now, createby = _byUser, updatedate = now, updateby = _byUser };

                foreach (var (prop, _, _) in ERP_FIELDS)
                {
                    if (!_mapCombos.TryGetValue(prop, out var cmb)) continue;
                    string excelCol = cmb.SelectedItem?.ToString();
                    if (excelCol == "-- Skip --" || !_sheetData.Columns.Contains(excelCol)) continue;
                    string raw = row[excelCol]?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(raw)) continue;
                    if (props.TryGetValue(prop, out var pi)) SetProp(m, pi, raw);
                }

                if (string.IsNullOrWhiteSpace(m.Invoice_poNo)) { failed++; continue; }

                // Generate unique ID
                m.Invoice_Id = UUIDGenerator.getAscId();

                // Invoice_no defaults to poNo (logistics staff updates later)
                if (string.IsNullOrWhiteSpace(m.Invoice_no))
                    m.Invoice_no = m.Invoice_poNo;

                // Auto-assign PIC from customer code
                m.Invoice_logisticPersonInCharge =
                    _picRepo.FindPICByCustomerCode(m.Invoice_customerCode);

                records.Add(m);
            }

            try
            {
                var (ok, upd, fail) = _repo.BulkInsert(records, _byUser);
                inserted = ok; updated = upd; failed += fail;
            }
            catch (Exception ex)
            {
                errors.AppendLine(ex.Message);
                failed++;
            }

            string msg = $"Import complete: {inserted} inserted, {updated} updated, {failed} failed.";
            if (errors.Length > 0) msg += $"\n\nErrors:\n{errors}";
            SetStatus(msg.Split('\n')[0], failed > 0 ? Color.OrangeRed : Color.DarkGreen);
            CTMessageBox.Show(msg, "Done", MessageBoxButtons.OK,
                failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            btnImport.Enabled = true;
        }

        // ── Helpers ───────────────────────────────────────────────────────
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
                dt.Columns.Add(h); cols.Add(h);
            }

            // Detect OA-date columns
            var dateCols = new HashSet<int>();
            for (int j = 1; j <= c; j++)
            {
                int hits = 0, total = 0;
                for (int r = 2; r <= Math.Min(rows, 8); r++)
                {
                    var v = raw[r, j]; if (v == null) continue; total++;
                    if (v is double d && d > 1 && d < 2958466) hits++;
                }
                if (total > 0 && hits == total) dateCols.Add(j);
            }

            for (int r = 2; r <= rows; r++)
            {
                var dr = dt.NewRow(); bool hasData = false;
                for (int j = 1; j <= c; j++)
                {
                    var cell = raw[r, j];
                    string v;
                    if (cell == null) v = "";
                    else if (dateCols.Contains(j) && cell is double oa && oa > 1 && oa < 2958466)
                    {
                        try { v = DateTime.FromOADate(oa).ToString("dd/MM/yyyy"); }
                        catch { v = cell.ToString(); }
                    }
                    else v = cell.ToString().Trim();
                    dr[j - 1] = v;
                    if (!string.IsNullOrEmpty(v)) hasData = true;
                }
                if (hasData) dt.Rows.Add(dr);
            }
            return dt;
        }

        private static void SetProp(InvoiceModel m, PropertyInfo pi, string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return;
            var t = Nullable.GetUnderlyingType(pi.PropertyType) ?? pi.PropertyType;
            object val;
            if (t == typeof(string))
                val = raw;
            else if (t == typeof(DateTime))
            {
                if (DateTime.TryParse(raw, out var dt)) val = dt;
                else if (double.TryParse(raw, System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var oa)
                    && oa > 1 && oa < 2958466) val = DateTime.FromOADate(oa);
                else return;
            }
            else if (t == typeof(decimal))
            {
                if (decimal.TryParse(raw.Replace(",", ""),
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.InvariantCulture, out var d)) val = d;
                else return;
            }
            else return;
            pi.SetValue(m, val);
        }

        private void SetStatus(string msg, Color? color = null)
        { lblStatus.Text = msg; lblStatus.ForeColor = color ?? Color.DarkSlateGray; }

        private static void BorderPaint(object sender, System.Windows.Forms.PaintEventArgs e)
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