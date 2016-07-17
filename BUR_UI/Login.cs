using System;
using System.Windows.Forms;

namespace BUR_UI
{
    public partial class Login : Form
    {
        public string UserName { get { return txtBStaffNumber.Text; } }
        bool employeeIsValid = false;
        bool passwordIsValid = false;
        public Login()
        {
            InitializeComponent();
        }
        private void btnLogin_Click(object sender, EventArgs e)
        {
            Interface.DbLink DbLink = new Interface.DbLink();

            if (DbLink.userValidate(txtBStaffNumber.Text, txtPassword.Text))
                this.DialogResult = DialogResult.OK;
            else
            {
                MessageBox.Show(
                    "User credentials do not match. Please try again.",
                    "Access Denied",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);

                txtPassword.Clear();
            }
        }
        private void txtBStaffNumber_TextChanged(object sender, EventArgs e)
        {
            ValidateLogin();
        }
        private void ValidateLogin()
        {
            if (txtBStaffNumber.TextLength < 8 || !tryInt(txtBStaffNumber.Text))
            {
                validateMsgEmployee.Visible = true;
                employeeIsValid = false;
            }
            else
            {
                validateMsgEmployee.Visible = false;
                employeeIsValid = true;
            }
            
            doValidate();
        }
        private bool tryInt(string BStaffNumber)
        {
            try
            {
                Convert.ToInt32(BStaffNumber);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
        private void doValidate()
        {
            if (employeeIsValid && passwordIsValid)
                btnLogin.Enabled = true;
            else
                btnLogin.Enabled = false;
        }
        private void Login_Load(object sender, EventArgs e)
        {
            txtBStaffNumber.Focus();
        }
        private void txtPassword_TextChanged(object sender, EventArgs e)
        {
            validatePassword();
        }
        private void validatePassword()
        {
            if (txtPassword.TextLength < 6)
            {
                validateMsgPassword.Visible = true;
                passwordIsValid = false;
            }
            else
            {
                validateMsgPassword.Visible = false;
                passwordIsValid = true;
            }

            doValidate();
        }
    }
}
