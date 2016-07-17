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
    public partial class dlgSelectMonth : Form
    {
        private bool hasCmbSelected = false;
        private bool hasRdoSelected = false;
        public string rdSelected = "";

        public dlgSelectMonth()
        {
            InitializeComponent();
        }

        public int Date { get { return cmbMonth.SelectedIndex; } }
        public string Class { get { return rdSelected; } }

        private void dlgSelectMonth_Load(object sender, EventArgs e)
        {

        }

        private void cmbMonth_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbMonth.SelectedIndex > -1)
                hasCmbSelected = true;
            else
                hasCmbSelected = false;

            checkAll();
        }

        private void rdoPS_CheckedChanged(object sender, EventArgs e)
        {
            checkRdo();
        }

        private void checkRdo()
        {
            if (rdoCO.Checked)
            {
                rdSelected = "CO";
                hasRdoSelected = true;
            }
            else if (rdoFE.Checked)
            {
                rdSelected = "FE";
                hasRdoSelected = true;
            }
            else if (rdoMOOE.Checked)
            {
                rdSelected = "MOOE";
                hasRdoSelected = true;
            }
            else if (rdoPS.Checked)
            {
                rdSelected = "PS";
                hasRdoSelected = true;
            }
            else
            {
                rdSelected = "";
                hasRdoSelected = false;
            }

            checkAll();
        }

        private void checkAll()
        {
            if (hasCmbSelected && hasRdoSelected)
                btnOK.Enabled = true;
            else
                btnOK.Enabled = false;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
        }

        private void rdoMOOE_CheckedChanged(object sender, EventArgs e)
        {
            checkRdo();
        }

        private void rdoFE_CheckedChanged(object sender, EventArgs e)
        {
            checkRdo();
        }

        private void rdoCO_CheckedChanged(object sender, EventArgs e)
        {
            checkRdo();
        }
    }
}
