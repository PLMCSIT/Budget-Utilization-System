namespace BUR_UI
{
    partial class dlgSelectMonth
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.cmbMonth = new System.Windows.Forms.ComboBox();
            this.grpClass = new System.Windows.Forms.GroupBox();
            this.rdoPS = new System.Windows.Forms.RadioButton();
            this.rdoMOOE = new System.Windows.Forms.RadioButton();
            this.rdoFE = new System.Windows.Forms.RadioButton();
            this.rdoCO = new System.Windows.Forms.RadioButton();
            this.grpClass.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(245, 21);
            this.label1.TabIndex = 0;
            this.label1.Text = "Generate Report for the month of:";
            // 
            // btnOK
            // 
            this.btnOK.Enabled = false;
            this.btnOK.ForeColor = System.Drawing.Color.Black;
            this.btnOK.Location = new System.Drawing.Point(176, 162);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(81, 28);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // cmbMonth
            // 
            this.cmbMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMonth.FormattingEnabled = true;
            this.cmbMonth.Items.AddRange(new object[] {
            "January",
            "February",
            "March",
            "April",
            "May",
            "June",
            "July",
            "August",
            "September",
            "October",
            "November",
            "December"});
            this.cmbMonth.Location = new System.Drawing.Point(12, 45);
            this.cmbMonth.Name = "cmbMonth";
            this.cmbMonth.Size = new System.Drawing.Size(245, 25);
            this.cmbMonth.TabIndex = 3;
            this.cmbMonth.SelectedIndexChanged += new System.EventHandler(this.cmbMonth_SelectedIndexChanged);
            // 
            // grpClass
            // 
            this.grpClass.Controls.Add(this.rdoCO);
            this.grpClass.Controls.Add(this.rdoFE);
            this.grpClass.Controls.Add(this.rdoMOOE);
            this.grpClass.Controls.Add(this.rdoPS);
            this.grpClass.Location = new System.Drawing.Point(12, 76);
            this.grpClass.Name = "grpClass";
            this.grpClass.Size = new System.Drawing.Size(245, 80);
            this.grpClass.TabIndex = 4;
            this.grpClass.TabStop = false;
            this.grpClass.Text = "Classification";
            // 
            // rdoPS
            // 
            this.rdoPS.AutoSize = true;
            this.rdoPS.Location = new System.Drawing.Point(7, 25);
            this.rdoPS.Name = "rdoPS";
            this.rdoPS.Size = new System.Drawing.Size(40, 21);
            this.rdoPS.TabIndex = 0;
            this.rdoPS.TabStop = true;
            this.rdoPS.Text = "PS";
            this.rdoPS.UseVisualStyleBackColor = true;
            this.rdoPS.CheckedChanged += new System.EventHandler(this.rdoPS_CheckedChanged);
            // 
            // rdoMOOE
            // 
            this.rdoMOOE.AutoSize = true;
            this.rdoMOOE.Location = new System.Drawing.Point(7, 52);
            this.rdoMOOE.Name = "rdoMOOE";
            this.rdoMOOE.Size = new System.Drawing.Size(65, 21);
            this.rdoMOOE.TabIndex = 1;
            this.rdoMOOE.TabStop = true;
            this.rdoMOOE.Text = "MOOE";
            this.rdoMOOE.UseVisualStyleBackColor = true;
            this.rdoMOOE.CheckedChanged += new System.EventHandler(this.rdoMOOE_CheckedChanged);
            // 
            // rdoFE
            // 
            this.rdoFE.AutoSize = true;
            this.rdoFE.Location = new System.Drawing.Point(114, 25);
            this.rdoFE.Name = "rdoFE";
            this.rdoFE.Size = new System.Drawing.Size(39, 21);
            this.rdoFE.TabIndex = 2;
            this.rdoFE.TabStop = true;
            this.rdoFE.Text = "FE";
            this.rdoFE.UseVisualStyleBackColor = true;
            this.rdoFE.CheckedChanged += new System.EventHandler(this.rdoFE_CheckedChanged);
            // 
            // rdoCO
            // 
            this.rdoCO.AutoSize = true;
            this.rdoCO.Location = new System.Drawing.Point(114, 52);
            this.rdoCO.Name = "rdoCO";
            this.rdoCO.Size = new System.Drawing.Size(44, 21);
            this.rdoCO.TabIndex = 3;
            this.rdoCO.TabStop = true;
            this.rdoCO.Text = "CO";
            this.rdoCO.UseVisualStyleBackColor = true;
            this.rdoCO.CheckedChanged += new System.EventHandler(this.rdoCO_CheckedChanged);
            // 
            // dlgSelectMonth
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.ClientSize = new System.Drawing.Size(275, 202);
            this.Controls.Add(this.grpClass);
            this.Controls.Add(this.cmbMonth);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.White;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "dlgSelectMonth";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Create Monthly Report...";
            this.Load += new System.EventHandler(this.dlgSelectMonth_Load);
            this.grpClass.ResumeLayout(false);
            this.grpClass.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.ComboBox cmbMonth;
        private System.Windows.Forms.GroupBox grpClass;
        private System.Windows.Forms.RadioButton rdoCO;
        private System.Windows.Forms.RadioButton rdoFE;
        private System.Windows.Forms.RadioButton rdoMOOE;
        private System.Windows.Forms.RadioButton rdoPS;
    }
}