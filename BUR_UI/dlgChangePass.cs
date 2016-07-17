using System;
using System.Windows.Forms;

namespace BUR_UI
{
    public partial class dlgChangePass : Form
    {
        private bool isMatch = false;
        private bool isOldPassValid = false;
        private bool isNewPassValid = false;
        private string StaffNumber = "";
        
        public DialogResult ShowDialog(string staffNumber)
        {
            StaffNumber = staffNumber;

            return base.ShowDialog();
        }
        public dlgChangePass()
        {
            InitializeComponent();
        }
        private void txtOldPass_TextChanged(object sender, EventArgs e)
        {
            if (txtOldPass.TextLength < 6)
            {
                isOldPassValid = false;

                lblMessage.Text = "Old password is must be 6 to 32 characters long.";
                lblMessage.Visible = true;
            }
            else
            {
                isOldPassValid = true;
            }

            checkIfValid();
        }
        private void checkIfValid()
        {
            if (isMatch && isOldPassValid && isNewPassValid)
            {
                btnOK.Enabled = true;

                lblMessage.Visible = false;
            }
            else
            {
                btnOK.Enabled = false;
            }
        }
        private void txtNewPass0_TextChanged(object sender, EventArgs e)
        {
            if (txtNewPass0.TextLength < 6)
            {
                isNewPassValid = false;

                lblMessage.Text = "New password must be 6 to 32 characters long.";
                lblMessage.Visible = true;
            }
            else
            {
                isNewPassValid = true;
            }

            checkIfValid();
        }
        private void txtNewPass1_TextChanged(object sender, EventArgs e)
        {
            if (txtNewPass0.Text != txtNewPass1.Text)
            {
                isMatch = false;

                lblMessage.Text = "New passwords do not match.";
                lblMessage.Visible = true;
            }
            else
            {
                isMatch = true;
            }

            checkIfValid();
        }
        private void btnOK_Click(object sender, EventArgs e)
        {
            Interface.DbLink dbLink = new Interface.DbLink();

            if (dbLink.userValidate(StaffNumber, txtOldPass.Text))
            {
                dbLink.ChangePassword(StaffNumber, txtNewPass0.Text);

                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("Please enter the valid old password for this account.",
                    "Access Denied", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
        }
    }
}
