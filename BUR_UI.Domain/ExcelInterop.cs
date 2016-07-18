using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Windows.Forms;

namespace BUR_UI.Logic
{
    public class ExcelInterop
    {
        private static Excel.Application ExApp = null;
        private static Excel.Workbook ExBook = null;
        private static Excel.Worksheet ExSheet = null;

        public void createSAAOExcel(List<Entities.SAAOModel> SAAO)
        {
            DateTimeFormatInfo dateFormat = new DateTimeFormatInfo();
            ExApp = new Excel.Application();
            ExApp.Visible = false;
            ExBook = ExApp.Workbooks.Open("C:\\SAAO.xlsx");
            ExSheet = (Excel.Worksheet)ExBook.Sheets[5];

            string month = dateFormat.GetMonthName(DateTime.Now.Month);

            /////////////// CAPITAL OUTLAY ///////////////

            // Month, Year
            ExSheet.Cells[7, 1] = "as of " + month + ", " + DateTime.Now.Year;

            // 202
            ExSheet.Cells[89, 6] = SAAO[0].AB;
            ExSheet.Cells[89, 9] = SAAO[0].Amount;

            // 212
            ExSheet.Cells[90, 6] = SAAO[1].AB;
            ExSheet.Cells[90, 9] = SAAO[1].Amount;

            // 221
            ExSheet.Cells[91, 6] = SAAO[2].AB;
            ExSheet.Cells[91, 9] = SAAO[2].Amount;

            // 222
            ExSheet.Cells[92, 6] = SAAO[3].AB;
            ExSheet.Cells[92, 9] = SAAO[3].Amount;

            // 223
            ExSheet.Cells[93, 6] = SAAO[4].AB;
            ExSheet.Cells[93, 9] = SAAO[4].Amount;

            // 229
            ExSheet.Cells[94, 6] = SAAO[5].AB;
            ExSheet.Cells[94, 9] = SAAO[5].Amount;

            // 231
            ExSheet.Cells[95, 6] = SAAO[6].AB;
            ExSheet.Cells[95, 9] = SAAO[6].Amount;

            // 233
            ExSheet.Cells[96, 6] = SAAO[7].AB;
            ExSheet.Cells[96, 9] = SAAO[7].Amount;

            // 235
            ExSheet.Cells[97, 6] = SAAO[8].AB;
            ExSheet.Cells[97, 9] = SAAO[8].Amount;

            // 236
            ExSheet.Cells[98, 6] = SAAO[9].AB;
            ExSheet.Cells[98, 9] = SAAO[9].Amount;

            // 240
            ExSheet.Cells[99, 6] = SAAO[10].AB;
            ExSheet.Cells[99, 9] = SAAO[10].Amount;

            /////////////// PERSONAL SERVICES ///////////////

            // 701
            ExSheet.Cells[16, 6] = SAAO[11].AB;
            ExSheet.Cells[16, 9] = SAAO[11].Amount;

            // 705
            ExSheet.Cells[18, 6] = SAAO[12].AB;
            ExSheet.Cells[18, 9] = SAAO[12].Amount;

            // 707
            ExSheet.Cells[19, 6] = SAAO[13].AB;
            ExSheet.Cells[19, 9] = SAAO[13].Amount;

            // 711
            ExSheet.Cells[20, 6] = SAAO[14].AB;
            ExSheet.Cells[20, 9] = SAAO[14].Amount;

            // 713
            ExSheet.Cells[21, 6] = SAAO[15].AB;
            ExSheet.Cells[21, 9] = SAAO[15].Amount;

            // 714
            ExSheet.Cells[22, 6] = SAAO[16].AB;
            ExSheet.Cells[22, 9] = SAAO[16].Amount;

            // 715
            ExSheet.Cells[23, 6] = SAAO[17].AB;
            ExSheet.Cells[23, 9] = SAAO[17].Amount;

            // 719
            ExSheet.Cells[24, 6] = SAAO[18].AB;
            ExSheet.Cells[24, 9] = SAAO[18].Amount;

            // 720
            ExSheet.Cells[25, 6] = SAAO[19].AB;
            ExSheet.Cells[25, 9] = SAAO[19].Amount;

            // 722
            ExSheet.Cells[26, 6] = SAAO[20].AB;
            ExSheet.Cells[26, 9] = SAAO[20].Amount;

            // 723
            ExSheet.Cells[27, 6] = SAAO[21].AB;
            ExSheet.Cells[27, 9] = SAAO[21].Amount;

            // 724
            ExSheet.Cells[28, 6] = SAAO[22].AB;
            ExSheet.Cells[28, 9] = SAAO[22].Amount;

            // 725
            ExSheet.Cells[29, 6] = SAAO[23].AB;
            ExSheet.Cells[29, 9] = SAAO[23].Amount;

            // 731
            ExSheet.Cells[30, 6] = SAAO[24].AB;
            ExSheet.Cells[30, 9] = SAAO[24].Amount;

            // 732
            ExSheet.Cells[31, 6] = SAAO[25].AB;
            ExSheet.Cells[31, 9] = SAAO[25].Amount;

            // 733
            ExSheet.Cells[32, 6] = SAAO[26].AB;
            ExSheet.Cells[32, 9] = SAAO[26].Amount;

            // 734
            ExSheet.Cells[33, 6] = SAAO[27].AB;
            ExSheet.Cells[33, 9] = SAAO[27].Amount;

            // 742
            ExSheet.Cells[34, 6] = SAAO[28].AB;
            ExSheet.Cells[34, 9] = SAAO[28].Amount;

            // 743
            ExSheet.Cells[35, 6] = SAAO[29].AB;
            ExSheet.Cells[35, 9] = SAAO[29].Amount;

            // 749
            ExSheet.Cells[36, 6] = SAAO[30].AB;
            ExSheet.Cells[36, 9] = SAAO[30].Amount;

            /////////////// MAINTENANCE & OTHER OPERATING EXPENSES ///////////////

            // 751
            ExSheet.Cells[42, 6] = SAAO[31].AB;
            ExSheet.Cells[42, 9] = SAAO[31].Amount;

            // 752
            ExSheet.Cells[43, 6] = SAAO[32].AB;
            ExSheet.Cells[43, 9] = SAAO[32].Amount;

            // 753
            ExSheet.Cells[44, 6] = SAAO[33].AB;
            ExSheet.Cells[44, 9] = SAAO[33].Amount;

            // 755
            ExSheet.Cells[45, 6] = SAAO[34].AB;
            ExSheet.Cells[45, 9] = SAAO[34].Amount;

            // 756
            ExSheet.Cells[46, 6] = SAAO[35].AB;
            ExSheet.Cells[46, 9] = SAAO[35].Amount;

            // 759
            ExSheet.Cells[47, 6] = SAAO[36].AB;
            ExSheet.Cells[47, 9] = SAAO[36].Amount;

            //760
            ExSheet.Cells[48, 6] = SAAO[37].AB;
            ExSheet.Cells[48, 9] = SAAO[37].Amount;

            //761
            ExSheet.Cells[49, 6] = SAAO[38].AB;
            ExSheet.Cells[49, 9] = SAAO[38].Amount;

            //765
            ExSheet.Cells[50, 6] = SAAO[39].AB;
            ExSheet.Cells[50, 9] = SAAO[39].Amount;

            //766
            ExSheet.Cells[51, 6] = SAAO[40].AB;
            ExSheet.Cells[51, 9] = SAAO[40].Amount;

            //767
            ExSheet.Cells[52, 6] = SAAO[41].AB;
            ExSheet.Cells[52, 9] = SAAO[41].Amount;

            //771
            ExSheet.Cells[53, 6] = SAAO[42].AB;
            ExSheet.Cells[53, 9] = SAAO[42].Amount;

            //772
            ExSheet.Cells[54, 6] = SAAO[43].AB;
            ExSheet.Cells[54, 9] = SAAO[43].Amount;

            //773
            ExSheet.Cells[55, 6] = SAAO[44].AB;
            ExSheet.Cells[55, 9] = SAAO[44].Amount;

            //774
            ExSheet.Cells[56, 6] = SAAO[45].AB;
            ExSheet.Cells[56, 9] = SAAO[45].Amount;

            //778
            ExSheet.Cells[57, 6] = SAAO[46].AB;
            ExSheet.Cells[57, 9] = SAAO[46].Amount;

            //780
            ExSheet.Cells[58, 6] = SAAO[47].AB;
            ExSheet.Cells[58, 9] = SAAO[47].Amount;

            //781
            ExSheet.Cells[59, 6] = SAAO[48].AB;
            ExSheet.Cells[59, 9] = SAAO[48].Amount;

            //782
            ExSheet.Cells[60, 6] = SAAO[49].AB;
            ExSheet.Cells[60, 9] = SAAO[49].Amount;

            //783
            ExSheet.Cells[61, 6] = SAAO[50].AB;
            ExSheet.Cells[61, 9] = SAAO[50].Amount;

            //786
            ExSheet.Cells[62, 6] = SAAO[51].AB;
            ExSheet.Cells[62, 9] = SAAO[51].Amount;

            //793
            ExSheet.Cells[63, 6] = SAAO[52].AB;
            ExSheet.Cells[63, 9] = SAAO[52].Amount;

            //796
            ExSheet.Cells[64, 6] = SAAO[53].AB;
            ExSheet.Cells[64, 9] = SAAO[53].Amount;

            //797
            ExSheet.Cells[65, 6] = SAAO[54].AB;
            ExSheet.Cells[65, 9] = SAAO[54].Amount;

            //799
            ExSheet.Cells[66, 6] = SAAO[55].AB;
            ExSheet.Cells[66, 9] = SAAO[55].Amount;

            //812
            ExSheet.Cells[67, 6] = SAAO[56].AB;
            ExSheet.Cells[67, 9] = SAAO[56].Amount;

            //821
            ExSheet.Cells[68, 6] = SAAO[57].AB;
            ExSheet.Cells[68, 9] = SAAO[57].Amount;

            //822
            ExSheet.Cells[69, 6] = SAAO[58].AB;
            ExSheet.Cells[69, 9] = SAAO[58].Amount;

            //823
            ExSheet.Cells[70, 6] = SAAO[59].AB;
            ExSheet.Cells[70, 9] = SAAO[59].Amount;

            //829
            ExSheet.Cells[71, 6] = SAAO[60].AB;
            ExSheet.Cells[71, 9] = SAAO[60].Amount;

            //833
            ExSheet.Cells[72, 6] = SAAO[61].AB;
            ExSheet.Cells[72, 9] = SAAO[61].Amount;

            //835
            ExSheet.Cells[73, 6] = SAAO[62].AB;
            ExSheet.Cells[73, 9] = SAAO[62].Amount;

            //836
            ExSheet.Cells[74, 6] = SAAO[63].AB;
            ExSheet.Cells[74, 9] = SAAO[63].Amount;

            //840
            ExSheet.Cells[75, 6] = SAAO[64].AB;
            ExSheet.Cells[75, 9] = SAAO[64].Amount;

            //841
            ExSheet.Cells[76, 6] = SAAO[65].AB;
            ExSheet.Cells[76, 9] = SAAO[65].Amount;

            //883
            ExSheet.Cells[77, 6] = SAAO[66].AB;
            ExSheet.Cells[77, 9] = SAAO[66].Amount;

            //892
            ExSheet.Cells[78, 6] = SAAO[67].AB;
            ExSheet.Cells[78, 9] = SAAO[67].Amount;

            //893
            ExSheet.Cells[79, 6] = SAAO[68].AB;
            ExSheet.Cells[79, 9] = SAAO[68].Amount;

            //969
            ExSheet.Cells[80, 6] = SAAO[69].AB;
            ExSheet.Cells[80, 9] = SAAO[69].Amount;

            ////////FINANCIAL EXPENSES

            //971
            ExSheet.Cells[86, 6] = SAAO[70].AB;
            ExSheet.Cells[86, 9] = SAAO[70].Amount;

            ExSheet.SaveAs("DBMS\\SAAO\\SAAOAsOf_" + month + ".xlsx");
            if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                ExSheet.PrintOutEx();
            ExBook.Close();
        }
        public void createBURExcel(Entities.BURModel SentBUR)
        {   
                Entities.BURModel BUR = SentBUR;

                ExApp = new Excel.Application();
                ExApp.Visible = false;
                ExBook = ExApp.Workbooks.Open("C:\\BUR.xls");
                ExSheet = (Excel.Worksheet)ExBook.Sheets[2];
                int lastRow = 20;
                float total = 0.00f;

                ExSheet.Cells[6, 7] = BUR.BURNumber;
                ExSheet.Cells[7, 2] = BUR.Payee;
                ExSheet.Cells[8, 2] = BUR.Office;
                ExSheet.Cells[12, 2] = BUR.Description + "\n" + "PR Number: " + BUR.PRNumber;

                foreach (var item in BUR.Particulars)
                {
                    ExSheet.Cells[lastRow, 2] = item.Name;
                    ExSheet.Cells[lastRow, 6] = item.Classification;
                    ExSheet.Cells[lastRow, 7] = item.Code;
                    ExSheet.Cells[lastRow, 8] = item.Amount.ToString("C2");

                    total += item.Amount;
                    lastRow++;
                }

                ExSheet.Cells[34, 8] = total;
                ExSheet.Cells[41, 2] = BUR.OfficeheadName;
                ExSheet.Cells[42, 2] = BUR.OfficeheadPos;
                ExSheet.Cells[44, 2] = DateTime.Now;

                ExSheet.Cells[41, 7] = BUR.BDHead;
                ExSheet.Cells[42, 7]    = BUR.BDHead_Pos;
                ExSheet.Cells[44, 7] = DateTime.Now;

                ExBook.SaveAs("DBMS\\BUR\\BUR_" + BUR.BURNumber + ".xls");

            if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                ExSheet.PrintOutEx();

            ExApp.Visible = true;
            ExSheet.PrintPreview();
            ExBook.Close();
        }
        public void createMonthlyCO(List<Entities.SAAOModel> Monthly, string month)
        {
            ExApp = new Excel.Application();
            ExApp.Visible = false;
            ExBook = ExApp.Workbooks.Open("C:\\SAAO.xlsx");
            ExSheet = (Excel.Worksheet)ExBook.Sheets[1];

            /////////////// CAPITAL OUTLAY ///////////////

            // Month, Year
            ExSheet.Cells[3, 1] = month + DateTime.Now.Year.ToString();

            // 202
            ExSheet.Cells[12, 6] = Monthly[0].AB;
            ExSheet.Cells[12, 9] = Monthly[0].Amount;

            // 212
            ExSheet.Cells[13, 6] = Monthly[1].AB;
            ExSheet.Cells[13, 9] = Monthly[1].Amount;

            // 221
            ExSheet.Cells[14, 6] = Monthly[2].AB;
            ExSheet.Cells[14, 9] = Monthly[2].Amount;

            // 222
            ExSheet.Cells[15, 6] = Monthly[3].AB;
            ExSheet.Cells[15, 9] = Monthly[3].Amount;

            // 223
            ExSheet.Cells[16, 6] = Monthly[4].AB;
            ExSheet.Cells[16, 9] = Monthly[4].Amount;

            // 229
            ExSheet.Cells[17, 6] = Monthly[5].AB;
            ExSheet.Cells[17, 9] = Monthly[5].Amount;

            // 231
            ExSheet.Cells[18, 6] = Monthly[6].AB;
            ExSheet.Cells[18, 9] = Monthly[6].Amount;

            // 233
            ExSheet.Cells[19, 6] = Monthly[7].AB;
            ExSheet.Cells[19, 9] = Monthly[7].Amount;

            // 235
            ExSheet.Cells[20, 6] = Monthly[8].AB;
            ExSheet.Cells[20, 9] = Monthly[8].Amount;

            // 236
            ExSheet.Cells[21, 6] = Monthly[9].AB;
            ExSheet.Cells[21, 9] = Monthly[9].Amount;

            // 240
            ExSheet.Cells[22, 6] = Monthly[10].AB;
            ExSheet.Cells[22, 9] = Monthly[10].Amount;

            // Month
            ExSheet.Cells[7, 9] = month;

            ExSheet.SaveAs("DBMS\\Monthly\\CO\\" + month + "_REPORT_CO.xlsx");

            if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                ExSheet.PrintOutEx();

            ExBook.Close();
        }
        public void createMonthlyMOOE(List<Entities.SAAOModel> Monthly, string month)
        {
            ExApp = new Excel.Application();
            ExApp.Visible = false;
            ExBook = ExApp.Workbooks.Open("C:\\SAAO.xlsx");
            ExSheet = (Excel.Worksheet)ExBook.Sheets[3];

            /////////////// MOOE ///////////////

            // Month, Year
            ExSheet.Cells[3, 1] = month + DateTime.Now.Year.ToString();

            //751
            ExSheet.Cells[12, 6] = Monthly[0].AB;
            ExSheet.Cells[12, 9] = Monthly[0].Amount;

            //752
            ExSheet.Cells[13, 6] = Monthly[1].AB;
            ExSheet.Cells[13, 9] = Monthly[1].Amount;

            //753
            ExSheet.Cells[14, 6] = Monthly[2].AB;
            ExSheet.Cells[14, 9] = Monthly[2].Amount;

            //755
            ExSheet.Cells[15, 6] = Monthly[3].AB;
            ExSheet.Cells[15, 9] = Monthly[3].Amount;

            //756
            ExSheet.Cells[16, 6] = Monthly[4].AB;
            ExSheet.Cells[16, 9] = Monthly[4].Amount;

            //759
            ExSheet.Cells[17, 6] = Monthly[5].AB;
            ExSheet.Cells[17, 9] = Monthly[5].Amount;

            //760
            ExSheet.Cells[18, 6] = Monthly[6].AB;
            ExSheet.Cells[18, 9] = Monthly[6].Amount;

            //761
            ExSheet.Cells[19, 6] = Monthly[7].AB;
            ExSheet.Cells[19, 9] = Monthly[7].Amount;

            //765
            ExSheet.Cells[20, 6] = Monthly[8].AB;
            ExSheet.Cells[20, 9] = Monthly[8].Amount;

            //766
            ExSheet.Cells[21, 6] = Monthly[9].AB;
            ExSheet.Cells[21, 9] = Monthly[9].Amount;

            //767
            ExSheet.Cells[22, 6] = Monthly[10].AB;
            ExSheet.Cells[22, 9] = Monthly[10].Amount;

            //771
            ExSheet.Cells[23, 6] = Monthly[11].AB;
            ExSheet.Cells[23, 9] = Monthly[11].Amount;

            //772
            ExSheet.Cells[24, 6] = Monthly[12].AB;
            ExSheet.Cells[24, 9] = Monthly[12].Amount;

            //773
            ExSheet.Cells[25, 6] = Monthly[13].AB;
            ExSheet.Cells[25, 9] = Monthly[13].Amount;

            //774
            ExSheet.Cells[26, 6] = Monthly[14].AB;
            ExSheet.Cells[26, 9] = Monthly[14].Amount;

            //778
            ExSheet.Cells[27, 6] = Monthly[15].AB;
            ExSheet.Cells[27, 9] = Monthly[15].Amount;

            //780
            ExSheet.Cells[28, 6] = Monthly[16].AB;
            ExSheet.Cells[28, 9] = Monthly[16].Amount;

            //781
            ExSheet.Cells[29, 6] = Monthly[17].AB;
            ExSheet.Cells[29, 9] = Monthly[17].Amount;

            //782
            ExSheet.Cells[30, 6] = Monthly[18].AB;
            ExSheet.Cells[30, 9] = Monthly[18].Amount;

            //783
            ExSheet.Cells[31, 6] = Monthly[19].AB;
            ExSheet.Cells[31, 9] = Monthly[19].Amount;

            //786
            ExSheet.Cells[32, 6] = Monthly[20].AB;
            ExSheet.Cells[32, 9] = Monthly[20].Amount;

            //793
            ExSheet.Cells[33, 6] = Monthly[21].AB;
            ExSheet.Cells[33, 9] = Monthly[21].Amount;

            //796
            ExSheet.Cells[34, 6] = Monthly[22].AB;
            ExSheet.Cells[34, 9] = Monthly[22].Amount;

            //797
            ExSheet.Cells[35, 6] = Monthly[23].AB;
            ExSheet.Cells[35, 9] = Monthly[23].Amount;

            //799
            ExSheet.Cells[36, 6] = Monthly[24].AB;
            ExSheet.Cells[36, 9] = Monthly[24].Amount;

            //812
            ExSheet.Cells[37, 6] = Monthly[25].AB;
            ExSheet.Cells[37, 9] = Monthly[25].Amount;

            //821
            ExSheet.Cells[38, 6] = Monthly[26].AB;
            ExSheet.Cells[38, 9] = Monthly[26].Amount;

            //822
            ExSheet.Cells[39, 6] = Monthly[27].AB;
            ExSheet.Cells[39, 9] = Monthly[27].Amount;

            //823
            ExSheet.Cells[40, 6] = Monthly[28].AB;
            ExSheet.Cells[40, 9] = Monthly[28].Amount;

            //829
            ExSheet.Cells[41, 6] = Monthly[29].AB;
            ExSheet.Cells[41, 9] = Monthly[29].Amount;

            //833
            ExSheet.Cells[42, 6] = Monthly[30].AB;
            ExSheet.Cells[42, 9] = Monthly[30].Amount;

            //835
            ExSheet.Cells[43, 6] = Monthly[31].AB;
            ExSheet.Cells[43, 9] = Monthly[31].Amount;

            //836
            ExSheet.Cells[44, 6] = Monthly[32].AB;
            ExSheet.Cells[44, 9] = Monthly[32].Amount;

            //840
            ExSheet.Cells[45, 6] = Monthly[33].AB;
            ExSheet.Cells[45, 9] = Monthly[33].Amount;

            //841
            ExSheet.Cells[46, 6] = Monthly[34].AB;
            ExSheet.Cells[46, 9] = Monthly[34].Amount;

            //883
            ExSheet.Cells[47, 6] = Monthly[35].AB;
            ExSheet.Cells[47, 9] = Monthly[35].Amount;

            //892
            ExSheet.Cells[48, 6] = Monthly[36].AB;
            ExSheet.Cells[48, 9] = Monthly[36].Amount;

            //893
            ExSheet.Cells[49, 6] = Monthly[37].AB;
            ExSheet.Cells[49, 9] = Monthly[37].Amount;

            //969
            ExSheet.Cells[50, 6] = Monthly[38].AB;
            ExSheet.Cells[50, 9] = Monthly[38].Amount;

            // Month
            ExSheet.Cells[7, 9] = month;

            ExSheet.SaveAs("DBMS\\Monthly\\MOOE\\" + month + "_REPORT_MOOE.xlsx");


            if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                ExSheet.PrintOutEx();

            ExBook.Close();
        }
        public void createMonthlyFE(List<Entities.SAAOModel> Monthly, string month)
        {
            ExApp = new Excel.Application();
            ExApp.Visible = false;
            ExBook = ExApp.Workbooks.Open("C:\\SAAO.xlsx");
            ExSheet = (Excel.Worksheet)ExBook.Sheets[2];

            ////////FINANCIAL EXPENSES MONTHLY

            // Month, Year
            ExSheet.Cells[3, 1] = month + DateTime.Now.Year.ToString();

            //202
            ExSheet.Cells[12, 6] = Monthly[0].AB;
            ExSheet.Cells[12, 9] = Monthly[0].Amount;

            // Month
            ExSheet.Cells[7, 9] = month;

            ExSheet.SaveAs("DBMS\\Monthly\\FE\\" + month + "_REPORT_FE.xlsx");
            if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                ExSheet.PrintOutEx();

            ExBook.Close();
        }
        public void createMonthlyPS(List<Entities.SAAOModel> Monthly, string month)
        {
            ExApp = new Excel.Application();
            ExApp.Visible = false;
            ExBook = ExApp.Workbooks.Open("C:\\SAAO.xlsx");
            ExSheet = (Excel.Worksheet)ExBook.Sheets[4];

            ////////PERSONAL SERVICES

            // Month, Year
            ExSheet.Cells[3, 1] = month + DateTime.Now.Year.ToString();

            //701
            ExSheet.Cells[12, 6] = Monthly[0].AB;
            ExSheet.Cells[12, 9] = Monthly[0].Amount;

            //705
            ExSheet.Cells[14, 6] = Monthly[1].AB;
            ExSheet.Cells[14, 9] = Monthly[1].Amount;

            //707
            ExSheet.Cells[15, 6] = Monthly[2].AB;
            ExSheet.Cells[15, 9] = Monthly[2].Amount;

            //711
            ExSheet.Cells[16, 6] = Monthly[3].AB;
            ExSheet.Cells[16, 9] = Monthly[3].Amount;

            //713
            ExSheet.Cells[17, 6] = Monthly[4].AB;
            ExSheet.Cells[17, 9] = Monthly[4].Amount;

            //714
            ExSheet.Cells[18, 6] = Monthly[5].AB;
            ExSheet.Cells[18, 9] = Monthly[5].Amount;

            //715
            ExSheet.Cells[19, 6] = Monthly[6].AB;
            ExSheet.Cells[19, 9] = Monthly[6].Amount;

            //719
            ExSheet.Cells[20, 6] = Monthly[7].AB;
            ExSheet.Cells[20, 9] = Monthly[7].Amount;

            //720
            ExSheet.Cells[21, 6] = Monthly[8].AB;
            ExSheet.Cells[21, 9] = Monthly[8].Amount;

            //722
            ExSheet.Cells[22, 6] = Monthly[9].AB;
            ExSheet.Cells[22, 9] = Monthly[9].Amount;

            //723
            ExSheet.Cells[23, 6] = Monthly[10].AB;
            ExSheet.Cells[23, 9] = Monthly[10].Amount;

            //724
            ExSheet.Cells[24, 6] = Monthly[11].AB;
            ExSheet.Cells[24, 9] = Monthly[11].Amount;

            //725
            ExSheet.Cells[25, 6] = Monthly[12].AB;
            ExSheet.Cells[25, 9] = Monthly[12].Amount;

            //731
            ExSheet.Cells[26, 6] = Monthly[13].AB;
            ExSheet.Cells[26, 9] = Monthly[13].Amount;

            //732
            ExSheet.Cells[27, 6] = Monthly[14].AB;
            ExSheet.Cells[27, 9] = Monthly[14].Amount;

            //733
            ExSheet.Cells[28, 6] = Monthly[15].AB;
            ExSheet.Cells[28, 9] = Monthly[15].Amount;

            //734
            ExSheet.Cells[29, 6] = Monthly[16].AB;
            ExSheet.Cells[29, 9] = Monthly[16].Amount;

            //742
            ExSheet.Cells[30, 6] = Monthly[17].AB;
            ExSheet.Cells[30, 9] = Monthly[17].Amount;

            //743
            ExSheet.Cells[31, 6] = Monthly[18].AB;
            ExSheet.Cells[31, 9] = Monthly[18].Amount;

            //749
            ExSheet.Cells[32, 6] = Monthly[19].AB;
            ExSheet.Cells[32, 9] = Monthly[19].Amount;


            // Month
            ExSheet.Cells[7, 9] = month;

            ExSheet.SaveAs("DBMS\\Monthly\\PS\\" + month + "_REPORT_PS.xlsx");
            if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                ExSheet.PrintOutEx();

            ExBook.Close();
        }
    }
}
