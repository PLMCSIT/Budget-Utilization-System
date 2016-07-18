using BUR_UI.Entities;
using BUR_UI.Interface;
using BUR_UI.Logic;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace BUR_UI
{
    public partial class Form1 : Form
    {
        DbLink link = new DbLink();
        public string docCreate = "";

        public string User_Name = "";
        public string User_Number = "";
        public string User_Pos = "";
        public bool isAdmin = false;
        public string BDHead_Number = "20030210";

        int selected = -1;

        private int token = 0;

        public Rectangle GetScreen()
        {
            return Screen.FromControl(this).Bounds;
        }
        public Form1()
        {
            InitializeComponent();
        }
        private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (MessageBox.Show("Are you sure you wan to close the program?", "Close?", MessageBoxButtons.YesNo) == DialogResult.No)
            {
                e.Cancel = true;
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Hide();
            Login loginForm = new Login();

            Typer Typer = new Typer();

            if (loginForm.ShowDialog() == DialogResult.OK)
            {
                this.Show();
                User_Name = lblUser.Text = Typer.GetSelectedStaffName(loginForm.UserName);
                User_Pos = lblPos.Text = Typer.GetPosition(loginForm.UserName);
                User_Number = loginForm.UserName;
                picPic.ImageLocation = Typer.GetUserImage(loginForm.UserName);

                isAdmin = Typer.CheckIfAdmin(loginForm.UserName);

                if (isAdmin)
                {
                    btnAdmin.Visible = true;
                    MessageBox.Show(
                        "You have ADMIN priveledges!", "Welcome, ADMIN!",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);

                    lblPos.Text += "\nSYSTEM ADMINISTRATOR";
                }
                else
                    btnAdmin.Visible = false;
            }
            else
            {
                if (MessageBox.Show(
                    "Are you sure you want to quit?",
                    "Close?",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) == DialogResult.Yes)
                    this.Close();
            }
            //formSeed();
            dataGridParticulars.DefaultCellStyle.ForeColor = Color.Black;
            List<BURModel> BURList = new List<BURModel>();
            DbLink DbLink = new DbLink();

            BURList = DbLink.FillGrid();

            FillDGrid(BURList);

            if (dataGridMain.Rows.Count > 0) token = GetLastBURNumber();
        }
        private void FillDGrid(List<BURModel> BURList)
        {
            foreach (var bur in BURList)
            {
                try
                {
                    dataGridMain.Rows.Add(
                        bur.BURNumber,
                        bur.Office,
                        bur.Payee,
                        bur.Date,
                        bur.Staff);
                } catch { }
            }
        }
        private int GetLastBURNumber()
        {
            int lastRow = dataGridMain.RowCount - 1;
            string BUR_Number = dataGridMain.Rows[lastRow].Cells[0].Value.ToString();

            BUR_Number = BUR_Number.Substring(11);

            return int.Parse(BUR_Number);
        }
        private void formSeed()
        {
            dataGridMain.Rows.Add(
                "01-2016-03-1234",
                "CET",
                "Kevin Yarnell",
                "10/03/2016 08:12:32",
                "Ho-Seong Lee");

            dataGridMain.Rows.Add(
                "01-2016-03-1235",
                "CM",
                "Dennis Johnsen",
                "10/03/2016 11:31:12",
                "Seong-Ung Bae");

            dataGridMain.Rows.Add(
                "01-2016-03-1236",
                "CS",
                "Søren Bjerg",
                "10/03/2016 13:15:01",
                "Sang-Hyeok Lee");

            dataGridMain.Rows.Add(
                "01-2016-03-1237",
                "CL",
                "Yiliang Peng",
                "10/03/2016 16:59:59",
                "Jun-Sik Bae");

            dataGridMain.Rows.Add(
                "01-2016-03-1237",
                "ICTO",
                "Bora Kim",
                "10/03/2016 16:59:59",
                "Jae-Wan Lee");
        }
        private void toolStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }
        private void btnAdd_Click(object sender, EventArgs e)
        {
                if (hasDuplicate(Convert.ToInt32(cmbCode.Text)))
                {
                    MessageBox.Show(
                        "Account Code " + cmbCode.Text + " already exists.",
                        "Warning",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                else
                {
                    dataGridParticulars.Rows.Add(
                        cmbClass.Text,
                        cmbCode.Text,
                        txtAcctName.Text,
                        numAmount.Value);

                    cmbClass.SelectedIndex = -1;
                    cmbCode.SelectedIndex = -1;
                    txtAcctName.Clear();
                    numAmount.Value = 0.00m;
                }
        }
        private bool hasDuplicate(int Code)
        {
            List<int> ExistingCodes = new List<int>();

            for (int i = 0; i < dataGridParticulars.RowCount; i++)
            {
                ExistingCodes.Add(Convert.ToInt32(dataGridParticulars.Rows[i].Cells[1].Value));
            }

            return ExistingCodes.Contains(Code);
        }
        private void btnEdit_Click(object sender, EventArgs e)
        {
            cmbClass.Text = dataGridParticulars.SelectedRows[0].Cells[0].Value.ToString();
            cmbCode.Text = dataGridParticulars.SelectedRows[0].Cells[1].Value.ToString();
            numAmount.Value = Convert.ToDecimal(dataGridParticulars.SelectedRows[0].Cells[3].Value);

            dataGridParticulars.Rows.RemoveAt(selected);
        }
        private void btnDelete_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridParticulars.Rows.RemoveAt(selected);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(),
                    "Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
        private void toolBtnCreate_Click(object sender, EventArgs e)
        {
            button4.Text = "Create";
            cmbOffice.Enabled = true;
            txtPR.Enabled = true;
            openBUR();
        }
        public void testDialog()
        {
            MessageBox.Show(
                "Dialog method has been accessed!",
                "Success",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }
        public void openBUR()
        {
            pnlMain.Visible = false;
            pnlCreate.Visible = true;
            //txtBURNumber.Text = IncrementBUR();

            List<string> Offices = link.FillOffice();
            List<string> Classes = link.FillClass();

            foreach (var office in Offices) cmbOffice.Items.Add(office);
            foreach (var cls in Classes) cmbClass.Items.Add(cls);
        }
        private void button5_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(
                "This will discard all changes made. Continue?",
                "Warning",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes)
            {
                pnlCreate.Visible = false;
                pnlMain.Visible = true;
            }

            ControlClear();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (button4.Text == "Create")
            {
                try
                {
                    if (AddBUR())
                    {
                        dataGridMain.Rows.Add(
                            txtBURNumber.Text,
                            cmbOffice.Text,
                            cmbPayee.Text,
                            DateTime.Now,
                            User_Name
                            );

                        pnlCreate.Visible = false;
                        pnlMain.Visible = true;

                        MessageBox.Show(
                            "BUR " + txtBURNumber.Text + " has been successfully created!",
                            "Success!",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);

                        ControlClear();
                    }
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                try
                {
                    EditBUR();

                    dataGridMain.SelectedRows[0].SetValues(
                        txtBURNumber.Text,
                        cmbOffice.Text,
                        cmbPayee.Text,
                        DateTime.Now,
                        User_Name);

                    pnlCreate.Visible = false;
                    pnlMain.Visible = true;

                    MessageBox.Show(
                        "BUR " + txtBURNumber.Text + " has been successfully created!",
                        "Success!",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);

                    ControlClear();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        private void EditBUR()
        {
            BURModel BUR = new BURModel();
            Typer typer = new Typer();

            BUR.BURNumber = txtBURNumber.Text;
            BUR.Office = cmbOffice.Text;
            BUR.OfficeCode = typer.GetSelectedOfficeCode(cmbOffice.Text);
            string[] Officehead = typer.GetOfficeHeadName(BUR.OfficeCode);
            BUR.OfficeheadName = Officehead[0];
            BUR.OfficeheadPos = Officehead[1];
            BUR.Payee = cmbPayee.Text;
            BUR.Payee_Number = typer.GetPayeeId(cmbPayee.Text);
            BUR.Description = txtDescription.Text;
            BUR.PRNumber = txtPR.Text;
            BUR.Staff = User_Name;
            BUR.Position = User_Pos;
            BUR.BDHead = "Lucresia C. Evangelista";
            BUR.BDHead_Pos = "Budget Officer V (Chief)";
            BUR.BStaff_Number = User_Number;
            BUR.Date = DateTime.Now.ToString();

            for (int i = 0; i < dataGridParticulars.RowCount; i++)
            {
                BUR.Particulars.Add(
                    new Items()
                    {
                        Classification = dataGridParticulars.Rows[i].Cells[0].Value.ToString(),
                        Code = dataGridParticulars.Rows[i].Cells[1].Value.ToString(),
                        Name = dataGridParticulars.Rows[i].Cells[2].Value.ToString(),
                        Amount = float.Parse(dataGridParticulars.Rows[i].Cells[3].Value.ToString()),
                        BUR_Number = txtBURNumber.Text
                    });
            }

            Context.DbInsert DbInsert = new Context.DbInsert();

            DbInsert.UpdateBUR(BUR);

            ExcelInterop Excel = new ExcelInterop();

            Excel.createBURExcel(BUR);
        }
        private void ControlClear()
        {
            cmbOffice.Items.Clear();
            cmbPayee.Items.Clear();
            txtPR.Clear();
            cmbClass.Items.Clear();
            cmbCode.Items.Clear();
            txtDescription.Clear();
            txtAcctName.Clear();
            numAmount.Value = 0.00m;
            dataGridParticulars.Rows.Clear();
        }
        private string IncrementBUR()
        {
            token++;
            int year = DateTime.Now.Year;
            int month = DateTime.Now.Month;

            string BURNumber = "01-" + year.ToString() + "-" + month.ToString("D2") + "-" + token.ToString("D4");
            return BURNumber;
        }
        private void bURToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pnlMain.Visible = false;
            pnlCreate.Visible = true;
            txtBURNumber.Text = IncrementBUR();
        }
        private void pnlCreate_Paint(object sender, PaintEventArgs e)
        {

        }
        private void cmbOffice_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbOffice.SelectedIndex == -1)
                cmbPayee.Enabled = false;
            else
                cmbPayee.Enabled = true;

            cmbPayee.Items.Clear();

            List<string> Payee = link.FillPayeeByOffice(cmbOffice.Text);

            foreach (var item in Payee) cmbPayee.Items.Add(item);
        }
        private void cmbClass_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbClass.SelectedIndex == -1)
                cmbCode.Enabled = false;
            else
                cmbCode.Enabled = true;

            cmbCode.Items.Clear();
            txtAcctName.Clear();

            List<string> Code = link.FillCodeByClass(cmbClass.Text);

            foreach (var item in Code) cmbCode.Items.Add(item);
        }
        private void cmbCode_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtAcctName.Clear();

            txtAcctName.Text = link.FillNameByCode(Convert.ToInt32(cmbCode.Text));
        }
        private void dataGridParticulars_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //selected = dataGridParticulars.CurrentCell.RowIndex;
        }
        private void dataGridParticulars_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }
        private void dataGridParticulars_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridParticulars.RowCount > 0)
                selected = dataGridParticulars.CurrentCell.RowIndex;
            else
            {
                btnDelete.Enabled = false;
                btnEdit.Enabled = false;
            }

            if (selected == -1)
            {
                btnDelete.Enabled = false;
                btnEdit.Enabled = false;
            }
            else
            {
                btnDelete.Enabled = true;
                btnEdit.Enabled = true;
            }
        }
        private void dataGridParticulars_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (dataGridParticulars.RowCount == 0)
            {
                selected = -1;
                btnEdit.Enabled = false;
                btnDelete.Enabled = false;
            }
            
        }
        public bool AddBUR()
        {
            BURModel BUR = new BURModel();
            Typer typer = new Typer();

            if (cmbOffice.SelectedIndex >= 0)
            {
                BUR.BURNumber = txtBURNumber.Text;
                BUR.Office = cmbOffice.Text;
                BUR.OfficeCode = typer.GetSelectedOfficeCode(cmbOffice.Text);
                string[] Officehead = typer.GetOfficeHeadName(BUR.OfficeCode);
                BUR.OfficeheadName = Officehead[0];
                BUR.OfficeheadPos = Officehead[1];
                BUR.Payee = cmbPayee.Text;
                BUR.Payee_Number = typer.GetPayeeId(cmbPayee.Text);
                BUR.Description = txtDescription.Text;
                BUR.PRNumber = txtPR.Text;
                BUR.Staff = User_Name;
                BUR.Position = User_Pos;
                BUR.BDHead = "Lucresia C. Evangelista";
                BUR.BDHead_Pos = "Budget Officer V (Chief)";
                BUR.BStaff_Number = User_Number;
                BUR.Date = DateTime.Now.ToString();

                for (int i = 0; i < dataGridParticulars.RowCount; i++)
                {
                    BUR.Particulars.Add(
                        new Items()
                        {
                            Classification = dataGridParticulars.Rows[i].Cells[0].Value.ToString(),
                            Code = dataGridParticulars.Rows[i].Cells[1].Value.ToString(),
                            Name = dataGridParticulars.Rows[i].Cells[2].Value.ToString(),
                            Amount = float.Parse(dataGridParticulars.Rows[i].Cells[3].Value.ToString()),
                            BUR_Number = txtBURNumber.Text
                        });
                }

                Context.DbInsert DbInsert = new Context.DbInsert();

                DbInsert.InsertBUR(BUR);

                ExcelInterop Excel = new ExcelInterop();

                Excel.createBURExcel(BUR);

                return true;
            }
            else
            {
                MessageBox.Show("Please provide for all required fields.",
                    "Review Action",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Exclamation);

                return false;
            }
        }
        private void toolBtnEdit_Click(object sender, EventArgs e)
        {
            Builder builder = new Builder();
            DbLink Link = new DbLink();
            List<AccountsModel> Acct = new List<AccountsModel>();
            List<ABModel> AB = new List<ABModel>();

            dlgCreate creator = new dlgCreate();
            if (creator.ShowDialog() == DialogResult.OK)
            {
                ExcelInterop Excel = new ExcelInterop();

                if (creator.Doc == "SAAO")
                {
                    Acct = Link.FillAccountsModel();
                    AB = Link.FillABModel();
                    List<SAAOModel> SAAO = builder.FillSAAOModel(Acct, AB);

                    SAAO = SAAO.OrderBy(a => a.Code).ToList();

                    Excel.createSAAOExcel(SAAO);
                }
                else
                {
                    dlgSelectMonth dlgMonth = new dlgSelectMonth();

                    if (dlgMonth.ShowDialog() == DialogResult.OK)
                    {
                        DateTimeFormatInfo dateTimePicker = new DateTimeFormatInfo();
                        int month = dlgMonth.Date + 1;
                        string monthName = dateTimePicker.GetMonthName(month);
                        string startDate = DateTime.Now.Year + "-" + month.ToString("D2") + "-01";
                        string endDate = DateTime.Now.Year + "-" + (month + 1).ToString("D2") + "-01";
                        Typer typer = new Typer();

                        List<SAAOModel> Monthly = new List<SAAOModel>();

                        if (dlgMonth.rdSelected == "CO")
                        {
                            Acct = Link.FillAccountsModel(typer.GetSelectedClassCode("CO"), startDate, endDate);
                            AB = Link.FillABModel(typer.GetSelectedClassCode("CO"));
                            Monthly = builder.FillMonthlyModel(Acct, AB);
                            Monthly = Monthly.OrderBy(a => a.Code).ToList();
                            Excel.createMonthlyCO(Monthly, monthName);
                        }
                        else if (dlgMonth.rdSelected == "MOOE")
                        {
                            Acct = Link.FillAccountsModel(typer.GetSelectedClassCode("MOOE"), startDate, endDate);
                            AB = Link.FillABModel(typer.GetSelectedClassCode("MOOE"));
                            Monthly = builder.FillMonthlyModel(Acct, AB);
                            Monthly = Monthly.OrderBy(a => a.Code).ToList();
                            Excel.createMonthlyMOOE(Monthly, monthName);
                        }
                        else if (dlgMonth.rdSelected == "FE")
                        {
                            Acct = Link.FillAccountsModel(typer.GetSelectedClassCode("FE"), startDate, endDate);
                            AB = Link.FillABModel(typer.GetSelectedClassCode("FE"));
                            Monthly = builder.FillMonthlyModel(Acct, AB);
                            Monthly = Monthly.OrderBy(a => a.Code).ToList();
                            Excel.createMonthlyFE(Monthly, monthName);
                        }
                        else if (dlgMonth.rdSelected == "PS")
                        {
                            Acct = Link.FillAccountsModel(typer.GetSelectedClassCode("PS"), startDate, endDate);
                            AB = Link.FillABModel(typer.GetSelectedClassCode("PS"));
                            Monthly = builder.FillMonthlyModel(Acct, AB);
                            Monthly = Monthly.OrderBy(a => a.Code).ToList();
                            Excel.createMonthlyPS(Monthly, monthName);
                        }
                    }
                }
            }
        }
        private void toolBtnPrint_Click(object sender, EventArgs e)
        {
            button4.Text = "Edit";

            DbFill DbFill = new DbFill();
            DbLink DbLink = new DbLink();
            Typer Typer = new Typer();
            BURModel BUR = new BURModel();
            Context.DbUpdate DbUpdate = new Context.DbUpdate();

            BUR = DbUpdate.FillEditor(dataGridMain.SelectedRows[0].Cells[0].Value.ToString());

            BUR.Office = Typer.GetSelectedOfficeName(BUR.OfficeCode);
            BUR.BDHead = Typer.GetSelectedBDHeadName(BUR.BDHead_Number);
            BUR.Staff = Typer.GetSelectedStaffName(BUR.Staff);
            BUR.Payee = Typer.GetSelectedPayeeName(BUR.Payee_Number);

            pnlMain.Visible = false;
            pnlCreate.Visible = true;

            txtBURNumber.Text = BUR.BURNumber;

            foreach (var classification in DbLink.FillClass())
            {
                cmbClass.Items.Add(classification);
            }

            cmbOffice.Items.Add(BUR.Office);
            cmbOffice.SelectedIndex = 0;
            cmbOffice.Enabled = false;

            cmbPayee.Items.Clear();
            cmbPayee.Items.Add(BUR.Payee);
            cmbPayee.SelectedIndex = 0;
            cmbPayee.Enabled = false;

            txtDescription.Text = BUR.Description;

            txtPR.Text = BUR.PRNumber;
            txtPR.Enabled = false;

            foreach (var item in BUR.Particulars)
            {
                dataGridParticulars.Rows.Add(
                    Typer.GetClassName(item.Code),
                    item.Code,
                    Typer.GetAcctName(item.Code),
                    item.Amount
                    );
            }
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        { }
        private bool hasRows()
        {
            if (dataGridMain.Rows.Count == 0)
            {
                MessageBox.Show(
                    "There are no rows to manipulate.",
                    "Warning",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return false;
            }

            return true;
        }
        private void btnLogOut_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(
                "Are you sure you want to log-out?",
                "Log-out?",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question) == DialogResult.Yes)
            {
                pnlAdmin.Visible = false;
                pnlMain.Visible = true;
                Form1_Load(sender, e);
            }
        }
        private void txtAcctName_TextChanged(object sender, EventArgs e)
        {
            if (txtAcctName.Text == "")
                btnAdd.Enabled = false;
            else
                btnAdd.Enabled = true;
        }
        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            dataGridMain.Rows.Clear();
            List<BURModel> BURList = new List<BURModel>();
            DbLink DbLink = new DbLink();

            BURList = DbLink.FillGrid(txtSearch.Text);

            FillDGrid(BURList);
        }
        private void dlgPrint_Load(object sender, EventArgs e)
        {

        }
        private void btnSelect_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                picUserDetail.ImageLocation = openFileDialog.FileName;
            }
        }
        private void btnAdmin_Click(object sender, EventArgs e)
        {
            DbLink dbLink = new DbLink();
            if (btnAdmin.Text == "Admin Panel")
            {
                List<UserModel> Users = new List<UserModel>();
                List<AccountGridModel> Accounts = new List<AccountGridModel>();

                btnAdmin.Text = "Main";
                pnlMain.Visible = false;
                pnlCreate.Visible = false;
                pnlAdmin.Visible = true;

                Users = dbLink.FillUserModel(Users);
                Accounts = dbLink.FillAccountGridModel(Accounts);

                FillUserGrid(Users);
                FillAccountGrid(Accounts);
            }
            else
            {
                dataGridUsers.Rows.Clear();
                dataGridAccounts.Rows.Clear();
                btnAdmin.Text = "Admin Panel";
                pnlAdmin.Visible = false;
                pnlMain.Visible = true;
            }
        }
        private void FillAccountGrid(List<AccountGridModel> accounts)
        {
            foreach (var account in accounts)
            {
                dataGridAccounts.Rows.Add(
                    account.AcctCode,
                    account.AcctName,
                    account.AcctClass
                );
            }



            //if (dataGridAccounts.RowCount >= 0)
            //    FillAccountDetails();
        }
        private void FillAccountDetails()
        {
            try
            {
                lblAcctCode.Text = dataGridAccounts.SelectedRows[0].Cells[0].Value.ToString();
                numAcctCode.Value = int.Parse(dataGridAccounts.SelectedRows[0].Cells[0].Value.ToString());
                txtEditAcctName.Text = lblAcctName.Text = dataGridAccounts.SelectedRows[0].Cells[1].Value.ToString();
                string AcctClass = lblAcctClass.Text = dataGridAccounts.SelectedRows[0].Cells[2].Value.ToString();

                switch (AcctClass)
                {
                    case "PS": cmbAcctClass.SelectedIndex = 0; break;
                    case "MOOE": cmbAcctClass.SelectedIndex = 1; break;
                    case "FE": cmbAcctClass.SelectedIndex = 2; break;
                    case "CO": cmbAcctClass.SelectedIndex = 3; break;
                }
            }
            catch
            { }
        }
        private void FillUserGrid(List<UserModel> users)
        {
            foreach (var user in users)
            {
                dataGridUsers.Rows.Add(
                    user.User_Number,
                    user.User_Name,
                    user.Discriminator,
                    user.Position,
                    user.Picture);
            }

            if (dataGridUsers.RowCount > 0)
                FillDetails();
        }
        private void dataGridUsers_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridUsers.RowCount > 0)
                FillDetails();
        }
        private void FillDetails()
        {
            try
            {
                txtStaffName.Text = dataGridUsers.SelectedRows[0].Cells[1].Value.ToString();
                txtStaffPosition.Text = dataGridUsers.SelectedRows[0].Cells[3].Value.ToString();

                if (dataGridUsers.SelectedRows[0].Cells[2].Value.ToString() == "Admin")
                    cmbType.SelectedIndex = 0;
                else
                    cmbType.SelectedIndex = 1;

                picUserDetail.ImageLocation = dataGridUsers.SelectedRows[0].Cells[4].Value.ToString();
            } catch { }
        }
        private void btnAllowEdit_Click(object sender, EventArgs e)
        {
            if (btnAllowEdit.Text == "Allow Edit")
            {
                btnAllowEdit.Text = "Save changes";
                dataGridUsers.Enabled = false;

                btnSelect.Enabled = true;
                btnChangePass.Enabled = true;
                btnDeleteUser.Enabled = true;
                txtStaffName.ReadOnly = false;
                cmbType.Enabled = true;
                txtStaffPosition.ReadOnly = false;
            }
            else
            {
                btnAllowEdit.Text = "Allow Edit";
                dataGridUsers.Enabled = true;

                btnSelect.Enabled = false;
                btnChangePass.Enabled = false;
                btnDeleteUser.Enabled = false;
                txtStaffName.ReadOnly = true;
                cmbType.Enabled = false;
                txtStaffPosition.ReadOnly = true;

                if (MessageBox.Show(
                    "Are you sure you want to save the changes to this user?",
                    "Save?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    Context.DbUpdate dbUpdate = new Context.DbUpdate();
                    DbLink dbLink = new DbLink();

                    UserModel User = new UserModel();

                    User.User_Number = dataGridUsers.SelectedRows[0].Cells[0].Value.ToString();
                    User.User_Name = txtStaffName.Text;
                    User.Discriminator = cmbType.Text;
                    User.Position = txtStaffPosition.Text;
                    User.Picture = picUserDetail.ImageLocation;

                    dbUpdate.UpdateUser(User);

                    dataGridUsers.SelectedRows[0].Cells[1].Value = User.User_Name;
                    dataGridUsers.SelectedRows[0].Cells[2].Value = User.Discriminator;
                    dataGridUsers.SelectedRows[0].Cells[3].Value = User.Position;
                    dataGridUsers.SelectedRows[0].Cells[4].Value = User.Picture;

                    MessageBox.Show(
                        "Staff " + User.User_Number + "'s details have been successfully updated!",
                        "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    FillDetails();
                }
            }
        }
        private void btnEditAccount_Click(object sender, EventArgs e)
        {
            if (btnEditAccount.Text == "Edit Account")
            {
                btnEditAccount.Text = "Save Changes";

                numAcctCode.Enabled = true;
                txtEditAcctName.ReadOnly = false;
                cmbAcctClass.Enabled = true;
            }

            else
            {
                btnEditAccount.Text = "Edit Account";

                numAcctCode.Enabled = false;
                txtEditAcctName.ReadOnly = true;
                cmbAcctClass.Enabled = false;
            }
        }
        private void dataGridAccounts_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridAccounts.RowCount > 0)
                FillAccountDetails();
        }
        private void btnChangePass_Click(object sender, EventArgs e)
        {
            dlgChangePass dlgPass = new dlgChangePass();
            string StaffNumber = dataGridUsers.SelectedRows[0].Cells[0].Value.ToString();

            if (dlgPass.ShowDialog(StaffNumber) == DialogResult.OK)
            {
                MessageBox.Show("You have successfully changed user " + txtStaffName.Text +
                    "'s password!", "Password changed",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }
    }
}
