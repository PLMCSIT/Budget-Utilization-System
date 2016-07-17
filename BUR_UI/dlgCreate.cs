using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BUR_UI
{
    public partial class dlgCreate : Form
    {
        string docu = "";
        public string Doc { get { return docu; } }

        public dlgCreate()
        {
            InitializeComponent();
        }

        private void btnContinue_Click(object sender, EventArgs e)
        {
            if (docu == "")
            {
                MessageBox.Show(
                    "Please select a document type.",
                    "No selected type",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);
            }
            else
                this.DialogResult = DialogResult.OK;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void lstCreate_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstCreate.SelectedItems.Count != 0)
            {
                if (lstCreate.SelectedItems[0].Text == "SAAO")
                {
                    txtDescription.Text = "Create a Statement of Appropriation,\nAllotment, and Obligation.";
                    docu = "SAAO";
                }
                else if (lstCreate.SelectedItems[0].Text == "Monthly Report")
                {
                    txtDescription.Text = "Create a Monthly Report.";
                    docu = "Monthly";
                }
            }
        }
    }
}
