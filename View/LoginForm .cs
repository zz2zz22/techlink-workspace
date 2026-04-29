using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using techlink_workspace.Controller;
using techlink_workspace.Controller.UI;
using techlink_workspace.Model;
using techlink_workspace.View;

namespace techlink_workspace
{
    public partial class LoginForm : Form
    {
        private readonly LoginController _loginCtrl = new LoginController();


        // ── Controls ──────────────────────────────────────────
        private PictureBox picLogo;
        private Label lblTitle;
        private Label lblUser;
        private TextBox txtUser;
        private Label lblPass;
        private TextBox txtPass;
        private Button btnLogin;
        private Button btnCancel;
        private Label lblError;
        private ComboBox cboLang;

        // ── Localization strings ───────────────────────────────
        private struct Lang
        {
            public string Title, User, UserHint, Pass, PassHint,
                          Login, Cancel, ErrUser, ErrPass, ErrInvalid, Success;
        }

        private readonly Dictionary<string, Lang> _langs = new Dictionary<string, Lang>
        {
            ["EN"] = new Lang
            {
                Title = "Employee Login",
                User = "Employee Code / Username",
                UserHint = "Enter employee code or username",
                Pass = "Password",
                PassHint = "Enter password",
                Login = "Login",
                Cancel = "Cancel",
                ErrUser = "Please enter your employee code or username.",
                ErrPass = "Please enter your password.",
                ErrInvalid = "Invalid employee code or password.",
                Success = "Login successful!"
            },
            ["VI"] = new Lang
            {
                Title = "Đăng Nhập Nhân Viên",
                User = "Mã nhân viên / Tên đăng nhập",
                UserHint = "Nhập mã nhân viên hoặc tên đăng nhập",
                Pass = "Mật khẩu",
                PassHint = "Nhập mật khẩu",
                Login = "Đăng nhập",
                Cancel = "Hủy",
                ErrUser = "Vui lòng nhập mã nhân viên hoặc tên đăng nhập.",
                ErrPass = "Vui lòng nhập mật khẩu.",
                ErrInvalid = "Mã nhân viên hoặc mật khẩu không đúng.",
                Success = "Đăng nhập thành công!"
            }
        };

        private string _currentLang = "EN";

        // ── Constructor ───────────────────────────────────────

        public LoginForm()
        {
            BuildUI();
            ApplyLanguage(_currentLang);
            this.FormClosed += (s, e) => Environment.Exit(0);
        }

        // ── Build UI ──────────────────────────────────────────
        private void BuildUI()
        {
            SuspendLayout();

            Text = "Login";
            ClientSize = new Size(360, 510);
            MinimumSize = new Size(376, 548);
            MaximumSize = new Size(376, 548);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            BackColor = Color.White;
            Font = new Font("Segoe UI", 9.5f);

            // ── Language switcher (top-right) ─────────────────
            cboLang = new ComboBox
            {
                Location = new Point(258, 10),
                Size = new Size(90, 24),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 9f)
            };
            cboLang.Items.AddRange(new object[] { "🇬🇧  EN", "🇻🇳  VI" });
            cboLang.SelectedIndex = 0;
            cboLang.SelectedIndexChanged += (s, e) =>
            {
                _currentLang = cboLang.SelectedIndex == 0 ? "EN" : "VI";
                ApplyLanguage(_currentLang);
            };

            // ── Logo ─────────────────────────────────────────
            picLogo = new PictureBox
            {
                Size = new Size(209, 209),
                Location = new Point((360 - 209) / 2, 40),
                SizeMode = PictureBoxSizeMode.StretchImage,
                BorderStyle = BorderStyle.None,
                BackColor = Color.Transparent
            };
            
            picLogo.Image = Properties.Resources.CompanyLogo;

            // ── Title ─────────────────────────────────────────
            lblTitle = new Label
            {
                Font = new Font("Segoe UI", 13f, FontStyle.Bold),
                ForeColor = Color.FromArgb(33, 33, 33),
                AutoSize = false,
                Size = new Size(300, 28),
                Location = new Point(30, 262),
                TextAlign = ContentAlignment.MiddleCenter
            };

            // ── Username ──────────────────────────────────────
            lblUser = new Label
            {
                Location = new Point(30, 304),
                AutoSize = true,
                ForeColor = Color.FromArgb(80, 80, 80)
            };

            txtUser = new TextBox
            {
                Location = new Point(30, 324),
                Size = new Size(300, 26),
                MaxLength = 50
            };

            // ── Password ──────────────────────────────────────
            lblPass = new Label
            {
                Location = new Point(30, 364),
                AutoSize = true,
                ForeColor = Color.FromArgb(80, 80, 80)
            };

            txtPass = new TextBox
            {
                Location = new Point(30, 384),
                Size = new Size(300, 26),
                MaxLength = 50,
                UseSystemPasswordChar = true
            };

            // ── Error label ───────────────────────────────────
            lblError = new Label
            {
                Text = "",
                ForeColor = Color.Crimson,
                Location = new Point(30, 420),
                Size = new Size(300, 18),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 8.5f)
            };

            // ── Buttons ───────────────────────────────────────
            btnLogin = new Button
            {
                Location = new Point(30, 448),
                Size = new Size(140, 36),
                BackColor = Color.FromArgb(0, 122, 204),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10f, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnLogin.FlatAppearance.BorderSize = 0;
            btnLogin.Click += BtnLogin_Click;

            btnCancel = new Button
            {
                Location = new Point(190, 448),
                Size = new Size(140, 36),
                BackColor = Color.FromArgb(230, 230, 230),
                ForeColor = Color.FromArgb(33, 33, 33),
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10f),
                Cursor = Cursors.Hand
            };
            btnCancel.FlatAppearance.BorderSize = 0;
            btnCancel.Click += (s, e) => Application.Exit();

            AcceptButton = btnLogin;

            Controls.AddRange(new Control[]
            {
                cboLang, picLogo, lblTitle,
                lblUser, txtUser,
                lblPass, txtPass,
                lblError, btnLogin, btnCancel
            });

            ResumeLayout(false);
        }

        // ── Apply language strings ────────────────────────────
        private void ApplyLanguage(string code)
        {
            var L = _langs[code];

            lblTitle.Text = L.Title;
            lblUser.Text = L.User;
            lblPass.Text = L.Pass;
            btnLogin.Text = L.Login;
            btnCancel.Text = L.Cancel;
            lblError.Text = "";   // clear any existing error

            SetPlaceholder(txtUser, L.UserHint, false);
            SetPlaceholder(txtPass, L.PassHint, true);
        }

        // ── Placeholder helper (.NET Framework 4.7.2) ─────────
        private void SetPlaceholder(TextBox tb, string hint, bool isPassword)
        {
            // Remove old handlers by resetting text first
            tb.ForeColor = Color.Gray;
            tb.UseSystemPasswordChar = false;
            tb.Text = hint;

            // Detach previous events cleanly via Tag
            if (tb.Tag is EventHandler[] old)
            {
                tb.GotFocus -= old[0];
                tb.LostFocus -= old[1];
            }

            EventHandler onGot = (s, e) =>
            {
                if (tb.Text == hint)
                {
                    tb.Text = "";
                    tb.ForeColor = Color.Black;
                    tb.UseSystemPasswordChar = isPassword;
                }
            };

            EventHandler onLost = (s, e) =>
            {
                //if (string.IsNullOrEmpty(tb.Text))
                //{
                //    tb.UseSystemPasswordChar = false;
                //    tb.ForeColor = Color.Gray;
                //    tb.Text = hint;
                //}
            };

            tb.GotFocus += onGot;
            tb.LostFocus += onLost;
            tb.Tag = new EventHandler[] { onGot, onLost };
        }

        // ── Login logic ───────────────────────────────────────
        private void BtnLogin_Click(object sender, EventArgs e)
        {
            var L = _langs[_currentLang];
            lblError.Text = "";

            string user = txtUser.Text.Trim();
            string pass = txtPass.Text.Trim();
            var hint = _langs[_currentLang];

            // ── Input guard ──────────────────────────────────────────
            if (string.IsNullOrEmpty(user) || user == hint.UserHint)
            { lblError.Text = L.ErrUser; txtUser.Focus(); return; }

            if (string.IsNullOrEmpty(pass) || pass == hint.PassHint)
            { lblError.Text = L.ErrPass; txtPass.Focus(); return; }

            // ── Delegate to controller ───────────────────────────────
            btnLogin.Enabled = false;   // prevent double-click
            UserModel loggedInUser;
            var loginResult = _loginCtrl.Authenticate(user, pass, out loggedInUser);

            switch (loginResult)
            {
                case LoginController.LoginResult.Success:
                    AppSession.CurrentUser = loggedInUser;
                    var dashboard = new Dashboard();
                    dashboard.Show();
                    this.Hide();
                    break;

                case LoginController.LoginResult.InvalidCredentials:
                    lblError.Text = L.ErrInvalid;
                    txtPass.Clear();
                    txtPass.Focus();
                    break;

                case LoginController.LoginResult.DbError:
                    CTMessageBox.Show(
                        "Lỗi kết nối cơ sở dữ liệu!\r\nDatabase connection error!",
                        "Lỗi 弊", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
            }

            btnLogin.Enabled = true;
        }
    }
}