using System;
using System.Drawing;
using System.Windows.Forms;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;

namespace techlink_workspace.View.Freight
{
    /// <summary>Simple add/edit dialog for a single ForwarderQuotationModel row.</summary>
    public partial class FreightEditForm : Form
    {
        public ForwarderQuotationModel Result { get; private set; }

        private readonly string _employeeCode;
        private readonly bool _isNew;

        // Fields
        private TextBox txtName, txtPort, txtTerm, txtCont, txtComm, txtHs,
                        txtCarrier, txtRemark, txtValid;
        private NumericUpDown numTotal, numOf, numLpol, numDest, numDel, numOther, numVol;
        private Button btnSave, btnCancel;
        private Label lblErr;
        private string _existingId;

        public FreightEditForm(ForwarderQuotationModel existing, string employeeCode, bool isNew)
        {
            _employeeCode = employeeCode;
            _isNew = isNew;
            _existingId = existing?.Forwarder_ID;
            Build();
            if (existing != null) Populate(existing);
        }

        private void Build()
        {
            Text = _isNew ? "Add Freight Quotation" : "Edit Freight Quotation";
            ClientSize = new Size(520, 560);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9f);

            var pnl = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(16, 12, 16, 0) };

            int y = 0;
            txtName = Field(pnl, "Forwarder Name", ref y);
            txtPort = Field(pnl, "Port / Delivery", ref y);
            txtTerm = Field(pnl, "Term", ref y);
            txtCont = Field(pnl, "Container", ref y);
            txtComm = Field(pnl, "Commodity", ref y);
            txtHs = Field(pnl, "HS Code", ref y);
            txtCarrier = Field(pnl, "Carrier", ref y);
            numOf = NumField(pnl, "OF", ref y);
            numLpol = NumField(pnl, "Local POL", ref y);
            numDest = NumField(pnl, "Dest Charge", ref y);
            numDel = NumField(pnl, "Delivery", ref y);
            numOther = NumField(pnl, "Other Charge", ref y);
            numTotal = NumField(pnl, "Total (USD)", ref y);
            numVol = NumField(pnl, "Volume/month", ref y, isInt: true);
            txtValid = Field(pnl, "Valid Date", ref y);
            txtRemark = Field(pnl, "Remark", ref y, multiline: true);

            lblErr = new Label
            {
                Location = new Point(0, y + 4),
                Size = new Size(480, 20),
                ForeColor = Color.Crimson,
                Font = new Font("Segoe UI", 8.5f)
            };
            pnl.Controls.Add(lblErr);

            var pnlBtn = new Panel { Dock = DockStyle.Bottom, Height = 48, BackColor = Color.WhiteSmoke };

            btnSave = new Button
            {
                Text = "Save",
                Size = new Size(110, 32),
                Location = new Point(16, 8),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;

            btnCancel = new Button
            {
                Text = "Cancel",
                Size = new Size(90, 32),
                Location = new Point(134, 8),
                BackColor = Color.FromArgb(210, 210, 210),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            pnlBtn.Controls.AddRange(new Control[] { btnSave, btnCancel });

            Controls.Add(pnl);
            Controls.Add(pnlBtn);
            AcceptButton = btnSave;
            CancelButton = btnCancel;
        }

        private TextBox Field(Panel p, string label, ref int y,
                              bool multiline = false)
        {
            p.Controls.Add(new Label
            {
                Text = label,
                Location = new Point(0, y),
                AutoSize = true,
                ForeColor = Color.FromArgb(70, 70, 70)
            });
            y += 20;
            var tb = new TextBox
            {
                Location = new Point(0, y),
                Size = new Size(480, multiline ? 52 : 24),
                Multiline = multiline,
                MaxLength = multiline ? 2000 : 200
            };
            p.Controls.Add(tb);
            y += (multiline ? 56 : 28);
            return tb;
        }

        private NumericUpDown NumField(Panel p, string label, ref int y,
                                       bool isInt = false)
        {
            p.Controls.Add(new Label
            {
                Text = label,
                Location = new Point(0, y),
                AutoSize = true,
                ForeColor = Color.FromArgb(70, 70, 70)
            });
            y += 20;
            var n = new NumericUpDown
            {
                Location = new Point(0, y),
                Size = new Size(200, 24),
                Minimum = 0,
                Maximum = 9999999,
                DecimalPlaces = isInt ? 0 : 2
            };
            p.Controls.Add(n);
            y += 28;
            return n;
        }

        private void Populate(ForwarderQuotationModel m)
        {
            txtName.Text = m.Forwarder_name ?? "";
            txtPort.Text = m.Forwarder_portDelivery ?? "";
            txtTerm.Text = m.Forwarder_term ?? "";
            txtCont.Text = m.Forwarder_container ?? "";
            txtComm.Text = m.Forwarder_commodity ?? "";
            txtHs.Text = m.Forwarder_hsCode ?? "";
            txtCarrier.Text = m.Forwarder_carrier ?? "";
            txtRemark.Text = m.Forwarder_remark ?? "";
            txtValid.Text = m.Forwarder_validDate ?? "";
            numOf.Value = (decimal)(m.Forwarder_of ?? 0);
            numLpol.Value = (decimal)(m.Forwarder_localPol ?? 0);
            numDest.Value = (decimal)(m.Forwarder_destCharge ?? 0);
            numDel.Value = (decimal)(m.Forwarder_delivery ?? 0);
            numOther.Value = (decimal)(m.Forwarder_otherCharge ?? 0);
            numTotal.Value = (decimal)(m.Forwarder_total ?? 0);
            numVol.Value = m.Forwarder_volumn ?? 0;
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            lblErr.Text = "";
            if (string.IsNullOrWhiteSpace(txtName.Text))
            { lblErr.Text = "Forwarder name is required."; return; }

            Result = new ForwarderQuotationModel
            {
                Forwarder_ID = _isNew ? Guid.NewGuid().ToString() : _existingId,
                Forwarder_name = txtName.Text.Trim(),
                Forwarder_portDelivery = txtPort.Text.Trim(),
                Forwarder_term = txtTerm.Text.Trim(),
                Forwarder_container = txtCont.Text.Trim(),
                Forwarder_commodity = txtComm.Text.Trim(),
                Forwarder_hsCode = txtHs.Text.Trim(),
                Forwarder_carrier = txtCarrier.Text.Trim(),
                Forwarder_of = numOf.Value > 0 ? numOf.Value : (decimal?)null,
                Forwarder_localPol = numLpol.Value > 0 ? numLpol.Value : (decimal?)null,
                Forwarder_destCharge = numDest.Value > 0 ? numDest.Value : (decimal?)null,
                Forwarder_delivery = numDel.Value > 0 ? numDel.Value : (decimal?)null,
                Forwarder_otherCharge = numOther.Value > 0 ? numOther.Value : (decimal?)null,
                Forwarder_total = numTotal.Value > 0 ? numTotal.Value : (decimal?)null,
                Forwarder_remark = txtRemark.Text.Trim(),
                Forwarder_volumn = (int)numVol.Value > 0 ? (int)numVol.Value : (int?)null,
                Forwarder_validDate = txtValid.Text.Trim(),
                create_by = _employeeCode,
                create_date = DateTime.Now,
                update_by = _employeeCode,
                update_date = DateTime.Now
            };
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}