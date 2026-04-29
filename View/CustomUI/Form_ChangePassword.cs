using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;
using techlink_workspace.Controller.UI;
using techlink_workspace.Data;
using techlink_workspace.Model;

namespace techlink_workspace.View
{
    /// <summary>
    /// Modal dialog for changing the current user's password.
    /// </summary>
    public partial class Form_ChangePassword : Form
    {
        private readonly string _userId;

        private TextBox txtCurrent;
        private TextBox txtNew;
        private TextBox txtConfirm;
        private Label lblError;
        private Button btnSave;
        private Button btnCancel;

        public Form_ChangePassword(string userId)
        {
            _userId = userId;
            BuildUI();
        }

        private void BuildUI()
        {
            Text = "Change Password";
            ClientSize = new Size(360, 320);
            MinimumSize = new Size(376, 358);
            MaximumSize = new Size(376, 358);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9.5f);

            int lx = 30, fx = 30, fw = 300, y = 24;

            AddLabel("Current password", lx, y); y += 22;
            txtCurrent = AddTextBox(fx, y, fw, true); y += 50;

            AddLabel("New password", lx, y); y += 22;
            txtNew = AddTextBox(fx, y, fw, true); y += 50;

            AddLabel("Confirm new password", lx, y); y += 22;
            txtConfirm = AddTextBox(fx, y, fw, true); y += 44;

            lblError = new Label
            {
                Text = "",
                ForeColor = Color.Crimson,
                Location = new Point(lx, y),
                Size = new Size(fw, 18),
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 8.5f)
            };
            Controls.Add(lblError);
            y += 26;

            btnSave = new Button
            {
                Text = "Save",
                Location = new Point(lx, y),
                Size = new Size(140, 36),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnSave.FlatAppearance.BorderSize = 0;
            btnSave.Click += BtnSave_Click;

            btnCancel = new Button
            {
                Text = "Cancel",
                Location = new Point(190, y),
                Size = new Size(140, 36),
                BackColor = Color.FromArgb(230, 230, 230),
                ForeColor = Color.FromArgb(33, 33, 33),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10f),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.Cancel
            };
            btnCancel.FlatAppearance.BorderSize = 0;

            Controls.AddRange(new Control[] { btnSave, btnCancel });
            AcceptButton = btnSave;
            CancelButton = btnCancel;
        }

        private Label AddLabel(string text, int x, int y)
        {
            var lbl = new Label
            {
                Text = text,
                Location = new Point(x, y),
                AutoSize = true,
                ForeColor = Color.FromArgb(80, 80, 80)
            };
            Controls.Add(lbl);
            return lbl;
        }

        private TextBox AddTextBox(int x, int y, int w, bool isPassword)
        {
            var tb = new TextBox
            {
                Location = new Point(x, y),
                Size = new Size(w, 26),
                MaxLength = 50,
                UseSystemPasswordChar = isPassword
            };
            Controls.Add(tb);
            return tb;
        }

        // ── Save handler ─────────────────────────────────────
        private void BtnSave_Click(object sender, EventArgs e)
        {
            lblError.Text = "";

            string current = txtCurrent.Text;
            string newPwd = txtNew.Text;
            string confirm = txtConfirm.Text;

            if (string.IsNullOrWhiteSpace(current))
            { lblError.Text = "Please enter your current password."; return; }

            if (string.IsNullOrWhiteSpace(newPwd) || newPwd.Length < 6)
            { lblError.Text = "New password must be at least 6 characters."; return; }

            if (newPwd != confirm)
            { lblError.Text = "New passwords do not match."; return; }

            btnSave.Enabled = false;

            try
            {
                // 1. Verify current password
                if (!VerifyCurrentPassword(current))
                {
                    lblError.Text = "Current password is incorrect.";
                    return;
                }

                // 2. Update password
                UpdatePassword(newPwd);

                CTMessageBox.Show(
                    "Password changed successfully.",
                    "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            catch (Exception ex)
            {
                CTMessageBox.Show(
                    "Error updating password:\r\n" + ex.Message,
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnSave.Enabled = true;
            }
        }

        private bool VerifyCurrentPassword(string current)
        {
            string sql = @"
                SELECT COUNT(1) FROM dbo.Sys_User
                WHERE User_id = @uid AND User_password = @pwd AND User_status = 1";

            using (var conn = new SqlConnection(DatabaseUtils.GetDBConnection().ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Add("@uid", SqlDbType.NVarChar, 50).Value = _userId ?? "";
                    cmd.Parameters.Add("@pwd", SqlDbType.VarChar, 50).Value = current;
                    // NOTE: replace raw password with your hashed value when ready
                    int count = (int)cmd.ExecuteScalar();
                    return count > 0;
                }
            }
        }

        private void UpdatePassword(string newPwd)
        {
            string sql = @"
                UPDATE dbo.Sys_User
                SET User_password = @pwd
                WHERE User_id = @uid";

            using (var conn = new SqlConnection(DatabaseUtils.GetDBConnection().ConnectionString))
            {
                conn.Open();
                using (var cmd = new SqlCommand(sql, conn))
                {
                    cmd.Parameters.Add("@pwd", SqlDbType.VarChar, 50).Value = newPwd;
                    cmd.Parameters.Add("@uid", SqlDbType.NVarChar, 50).Value = _userId ?? "";
                    // NOTE: hash newPwd before storing (e.g. SHA-256) when ready
                    cmd.ExecuteNonQuery();
                }
            }
        }
    }
}