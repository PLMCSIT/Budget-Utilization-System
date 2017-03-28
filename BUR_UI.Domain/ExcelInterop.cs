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
            ExBook = ExApp.Workbooks.Open("C:\\SAAO.xls");
            ExSheet = (Excel.Worksheet)ExBook.Sheets[1];

            string month = dateFormat.GetMonthName(DateTime.Now.Month);

            /////////////// CAPITAL OUTLAY ///////////////

            // Month, Year
            ExSheet.Cells[7, 1] = "as of " + month + ", " + DateTime.Now.Year;

            //10702990
            ExSheet.Cells[95, 6] = SAAO[0].AB;
            ExSheet.Cells[95, 9] = SAAO[0].Amount;

            //10704020
            ExSheet.Cells[96, 6] = SAAO[1].AB;
            ExSheet.Cells[96, 9] = SAAO[1].Amount;

            //10705020
            ExSheet.Cells[97, 6] = SAAO[2].AB;
            ExSheet.Cells[97, 9] = SAAO[2].Amount;

            //10705030
            ExSheet.Cells[98, 6] = SAAO[3].AB;
            ExSheet.Cells[98, 9] = SAAO[3].Amount;

            //10705070
            ExSheet.Cells[99, 6] = SAAO[4].AB;
            ExSheet.Cells[99, 9] = SAAO[4].Amount;

            //10705090
            ExSheet.Cells[100, 6] = SAAO[5].AB;
            ExSheet.Cells[100, 9] = SAAO[5].Amount;

            //10705110
            ExSheet.Cells[101, 6] = SAAO[6].AB;
            ExSheet.Cells[101, 9] = SAAO[6].Amount;

            //10705130
            ExSheet.Cells[102, 6] = SAAO[7].AB;
            ExSheet.Cells[102, 9] = SAAO[7].Amount;

            //10705140
            ExSheet.Cells[103, 6] = SAAO[8].AB;
            ExSheet.Cells[103, 9] = SAAO[8].Amount;

            //10705990
            ExSheet.Cells[104, 6] = SAAO[9].AB;
            ExSheet.Cells[104, 9] = SAAO[9].Amount;

            //10707010
            ExSheet.Cells[105, 6] = SAAO[10].AB;
            ExSheet.Cells[105, 9] = SAAO[10].Amount;

            //10707020
            ExSheet.Cells[106, 6] = SAAO[11].AB;
            ExSheet.Cells[106, 9] = SAAO[11].Amount;

            //50101010
            ExSheet.Cells[16, 6] = SAAO[12].AB;
            ExSheet.Cells[16, 9] = SAAO[12].Amount;

            //50101020.1
            ExSheet.Cells[18, 6] = SAAO[13].AB;
            ExSheet.Cells[18, 9] = SAAO[13].Amount;

            //50101020.2
            ExSheet.Cells[19, 6] = SAAO[14].AB;
            ExSheet.Cells[19, 9] = SAAO[14].Amount;

            //50102010
            ExSheet.Cells[20, 6] = SAAO[15].AB;
            ExSheet.Cells[20, 9] = SAAO[15].Amount;

            //50102020
            ExSheet.Cells[21, 6] = SAAO[16].AB;
            ExSheet.Cells[21, 9] = SAAO[16].Amount;

            //50102030
            ExSheet.Cells[22, 6] = SAAO[17].AB;
            ExSheet.Cells[22, 9] = SAAO[17].Amount;

            //50102040
            ExSheet.Cells[23, 6] = SAAO[18].AB;
            ExSheet.Cells[23, 9] = SAAO[18].Amount;

            //50102050
            ExSheet.Cells[25, 6] = SAAO[19].AB;
            ExSheet.Cells[25, 9] = SAAO[19].Amount;

            //50102060
            ExSheet.Cells[26, 6] = SAAO[20].AB;
            ExSheet.Cells[26, 9] = SAAO[20].Amount;

            //50102100
            ExSheet.Cells[24, 6] = SAAO[21].AB;
            ExSheet.Cells[24, 9] = SAAO[21].Amount;

            //50102110
            ExSheet.Cells[27, 6] = SAAO[22].AB;
            ExSheet.Cells[27, 9] = SAAO[22].Amount;

            //50102120
            ExSheet.Cells[28, 6] = SAAO[23].AB;
            ExSheet.Cells[28, 9] = SAAO[23].Amount;

            //50102130
            ExSheet.Cells[29, 6] = SAAO[24].AB;
            ExSheet.Cells[29, 9] = SAAO[24].Amount;

            //50102140
            ExSheet.Cells[31, 6] = SAAO[25].AB;
            ExSheet.Cells[31, 9] = SAAO[25].Amount;

            //50102150
            ExSheet.Cells[30, 6] = SAAO[26].AB;
            ExSheet.Cells[30, 9] = SAAO[26].Amount;

            //50102990
            ExSheet.Cells[32, 6] = SAAO[27].AB;
            ExSheet.Cells[32, 9] = SAAO[27].Amount;

            //50103010
            ExSheet.Cells[33, 6] = SAAO[28].AB;
            ExSheet.Cells[33, 9] = SAAO[28].Amount;

            //50103020
            ExSheet.Cells[34, 6] = SAAO[29].AB;
            ExSheet.Cells[34, 9] = SAAO[29].Amount;

            //50103030
            ExSheet.Cells[35, 6] = SAAO[30].AB;
            ExSheet.Cells[35, 9] = SAAO[30].Amount;

            //50103040
            ExSheet.Cells[36, 6] = SAAO[31].AB;
            ExSheet.Cells[36, 9] = SAAO[31].Amount;

            //50104030
            ExSheet.Cells[37, 6] = SAAO[32].AB;
            ExSheet.Cells[37, 9] = SAAO[32].Amount;

            //50104990
            ExSheet.Cells[38, 6] = SAAO[33].AB;
            ExSheet.Cells[38, 9] = SAAO[33].Amount;

            //50201010
            ExSheet.Cells[50, 6] = SAAO[34].AB;
            ExSheet.Cells[50, 9] = SAAO[34].Amount;

            //50201020
            ExSheet.Cells[51, 6] = SAAO[35].AB;
            ExSheet.Cells[51, 9] = SAAO[35].Amount;

            //50202010
            ExSheet.Cells[52, 6] = SAAO[36].AB;
            ExSheet.Cells[52, 9] = SAAO[36].Amount;

            //50203010
            ExSheet.Cells[53, 6] = SAAO[37].AB;
            ExSheet.Cells[53, 9] = SAAO[37].Amount;

            //50203020
            ExSheet.Cells[54, 6] = SAAO[38].AB;
            ExSheet.Cells[54, 9] = SAAO[38].Amount;

            //50203070
            ExSheet.Cells[55, 6] = SAAO[39].AB;
            ExSheet.Cells[55, 9] = SAAO[39].Amount;

            //50203080
            ExSheet.Cells[56, 6] = SAAO[40].AB;
            ExSheet.Cells[56, 9] = SAAO[40].Amount;

            //50203090
            ExSheet.Cells[57, 6] = SAAO[41].AB;
            ExSheet.Cells[57, 9] = SAAO[41].Amount;

            //50203990
            ExSheet.Cells[58, 6] = SAAO[42].AB;
            ExSheet.Cells[58, 9] = SAAO[42].Amount;

            //50204010
            ExSheet.Cells[59, 6] = SAAO[43].AB;
            ExSheet.Cells[59, 9] = SAAO[43].Amount;

            //50204020
            ExSheet.Cells[60, 6] = SAAO[44].AB;
            ExSheet.Cells[60, 9] = SAAO[44].Amount;

            //50205010
            ExSheet.Cells[61, 6] = SAAO[45].AB;
            ExSheet.Cells[61, 9] = SAAO[45].Amount;

            //50205020.1
            ExSheet.Cells[62, 6] = SAAO[46].AB;
            ExSheet.Cells[62, 9] = SAAO[46].Amount;

            //50205020.2
            ExSheet.Cells[63, 6] = SAAO[47].AB;
            ExSheet.Cells[63, 9] = SAAO[47].Amount;

            //50205030
            ExSheet.Cells[64, 6] = SAAO[48].AB;
            ExSheet.Cells[64, 9] = SAAO[48].Amount;

            //50210030
            ExSheet.Cells[65, 6] = SAAO[49].AB;
            ExSheet.Cells[65, 9] = SAAO[49].Amount;

            //50211030
            ExSheet.Cells[66, 6] = SAAO[50].AB;
            ExSheet.Cells[66, 9] = SAAO[50].Amount;

            //50211990
            ExSheet.Cells[67, 6] = SAAO[51].AB;
            ExSheet.Cells[67, 9] = SAAO[51].Amount;

            //50212020
            ExSheet.Cells[68, 6] = SAAO[52].AB;
            ExSheet.Cells[68, 9] = SAAO[52].Amount;

            //50212030
            ExSheet.Cells[69, 6] = SAAO[53].AB;
            ExSheet.Cells[69, 9] = SAAO[53].Amount;

            //50213040
            ExSheet.Cells[70, 6] = SAAO[54].AB;
            ExSheet.Cells[70, 9] = SAAO[54].Amount;

            //50213050.1
            ExSheet.Cells[71, 6] = SAAO[55].AB;
            ExSheet.Cells[71, 9] = SAAO[55].Amount;

            //50213050.2
            ExSheet.Cells[72, 6] = SAAO[56].AB;
            ExSheet.Cells[72, 9] = SAAO[56].Amount;

            //50213050.3
            ExSheet.Cells[73, 6] = SAAO[57].AB;
            ExSheet.Cells[73, 9] = SAAO[57].Amount;

            //50213050.4
            ExSheet.Cells[74, 6] = SAAO[58].AB;
            ExSheet.Cells[74, 9] = SAAO[58].Amount;

            //50213050.5
            ExSheet.Cells[75, 6] = SAAO[59].AB;
            ExSheet.Cells[75, 9] = SAAO[59].Amount;

            //50213050.6
            ExSheet.Cells[76, 6] = SAAO[60].AB;
            ExSheet.Cells[76, 9] = SAAO[60].Amount;

            //50213050.7
            ExSheet.Cells[77, 6] = SAAO[61].AB;
            ExSheet.Cells[77, 9] = SAAO[61].Amount;

            //50213060
            ExSheet.Cells[78, 6] = SAAO[62].AB;
            ExSheet.Cells[78, 9] = SAAO[62].Amount;

            //50213070
            ExSheet.Cells[79, 6] = SAAO[63].AB;
            ExSheet.Cells[79, 9] = SAAO[63].Amount;

            //50216020
            ExSheet.Cells[80, 6] = SAAO[64].AB;
            ExSheet.Cells[80, 9] = SAAO[64].Amount;

            //50216030
            ExSheet.Cells[81, 6] = SAAO[65].AB;
            ExSheet.Cells[81, 9] = SAAO[65].Amount;

            //50299010
            ExSheet.Cells[82, 6] = SAAO[66].AB;
            ExSheet.Cells[82, 9] = SAAO[66].Amount;

            //50299020
            ExSheet.Cells[83, 6] = SAAO[67].AB;
            ExSheet.Cells[83, 9] = SAAO[67].Amount;

            //50299030
            ExSheet.Cells[84, 6] = SAAO[68].AB;
            ExSheet.Cells[84, 9] = SAAO[68].Amount;

            //50299050
            ExSheet.Cells[85, 6] = SAAO[69].AB;
            ExSheet.Cells[85, 9] = SAAO[69].Amount;

            //50299060
            ExSheet.Cells[86, 6] = SAAO[70].AB;
            ExSheet.Cells[86, 9] = SAAO[70].Amount;

            //50299070
            ExSheet.Cells[87, 6] = SAAO[71].AB;
            ExSheet.Cells[87, 9] = SAAO[71].Amount;

            //50299990
            ExSheet.Cells[88, 6] = SAAO[72].AB;
            ExSheet.Cells[88, 9] = SAAO[72].Amount;

            //50301040
            ExSheet.Cells[92, 6] = SAAO[73].AB;
            ExSheet.Cells[92, 9] = SAAO[73].Amount;



            try
            {
                ExSheet.SaveAs("DBMS\\SAAO\\SAAOAsOf_" + month + ".xlsx");
            }
            catch
            {

            }
            finally
            {
                if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    ExSheet.PrintOutEx();

                ExApp.Visible = true;
                ExSheet.PrintPreview();
                ExBook.Close();
            }
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
                
                if (BUR.Office != "External")
            {
                ExSheet.Cells[8, 2] = BUR.Office;
            }

                ExSheet.Cells[12, 2] = BUR.Description + "\n" + "PR Number: " + BUR.PRNumber;

                foreach (var item in BUR.Particulars)
                {
                    ExSheet.Cells[lastRow, 2] = item.Name;
                    ExSheet.Cells[lastRow, 6] = item.Classification;
                    //item.Code = item.Code.Insert(1, "-");
                    //item.Code = item.Code.Insert(4, "-");
                    //item.Code = item.Code.Insert(7, "-");
                    ExSheet.Cells[lastRow, 7] = item.Code;
                    ExSheet.Cells[lastRow, 8] = item.Amount.ToString("C2");

                    total += item.Amount;
                    lastRow++;
                }

                ExSheet.Cells[34, 8] = total;
                
                if (BUR.Office == "External")
            {
                ExSheet.Cells[41, 2] = BUR.OfficeheadName;
                ExSheet.Cells[42, 2] = BUR.OfficeheadPos;
            } else
            {
                ExSheet.Cells[41, 2] = BUR.OfficeheadName;
                ExSheet.Cells[42, 2] = BUR.OfficeheadPos;
            }

                ExSheet.Cells[44, 2] = DateTime.Now;

                ExSheet.Cells[41, 7] = BUR.BDHead;
                ExSheet.Cells[42, 7] = BUR.BDHead_Pos;
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
            ExBook = ExApp.Workbooks.Open("C:\\Monthly.xlsx");
            ExSheet = (Excel.Worksheet)ExBook.Sheets[1];

            /////////////// CAPITAL OUTLAY ///////////////

            // Month, Year
            ExSheet.Cells[3, 1] = month + DateTime.Now.Year.ToString();

            //CAPITAL OUTLAY

            //10702990
            ExSheet.Cells[12, 6] = Monthly[0].AB;
            ExSheet.Cells[12, 9] = Monthly[0].Amount;

            //10704020
            ExSheet.Cells[13, 6] = Monthly[1].AB;
            ExSheet.Cells[13, 9] = Monthly[1].Amount;

            //10705020
            ExSheet.Cells[14, 6] = Monthly[2].AB;
            ExSheet.Cells[14, 9] = Monthly[2].Amount;

            //10705030
            ExSheet.Cells[15, 6] = Monthly[3].AB;
            ExSheet.Cells[15, 9] = Monthly[3].Amount;

            //10705070
            ExSheet.Cells[16, 6] = Monthly[4].AB;
            ExSheet.Cells[16, 9] = Monthly[4].Amount;

            //10705090
            ExSheet.Cells[17, 6] = Monthly[5].AB;
            ExSheet.Cells[17, 9] = Monthly[5].Amount;

            //10705110
            ExSheet.Cells[18, 6] = Monthly[6].AB;
            ExSheet.Cells[18, 9] = Monthly[6].Amount;

            //10705130
            ExSheet.Cells[19, 6] = Monthly[7].AB;
            ExSheet.Cells[19, 9] = Monthly[7].Amount;

            //10705140
            ExSheet.Cells[20, 6] = Monthly[8].AB;
            ExSheet.Cells[20, 9] = Monthly[8].Amount;

            //10705990
            ExSheet.Cells[21, 6] = Monthly[9].AB;
            ExSheet.Cells[21, 9] = Monthly[9].Amount;

            //10707010
            ExSheet.Cells[22, 6] = Monthly[10].AB;
            ExSheet.Cells[22, 9] = Monthly[10].Amount;

            //10707020
            ExSheet.Cells[23, 6] = Monthly[11].AB;
            ExSheet.Cells[23, 9] = Monthly[11].Amount;

            try
            {
                ExSheet.SaveAs("DBMS\\Monthly\\CO\\" + month + "_REPORT_CO.xlsx");
            }
            catch
            {

            }
            finally
            {
                if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    ExSheet.PrintOutEx();

                ExApp.Visible = true;
                ExSheet.PrintPreview();
                ExBook.Close();
            }
        }
        public void createMonthlyMOOE(List<Entities.SAAOModel> Monthly, string month)
        {
            ExApp = new Excel.Application();
            ExApp.Visible = false;
            ExBook = ExApp.Workbooks.Open("C:\\Monthly.xlsx");
            ExSheet = (Excel.Worksheet)ExBook.Sheets[3];

            /////////////// MOOE ///////////////

            // Month, Year
            ExSheet.Cells[3, 1] = month + DateTime.Now.Year.ToString();

            ///////MOOE

            //50201010
            ExSheet.Cells[12, 6] = Monthly[0].AB;
            ExSheet.Cells[12, 9] = Monthly[0].Amount;

            //50201020
            ExSheet.Cells[13, 6] = Monthly[1].AB;
            ExSheet.Cells[13, 9] = Monthly[1].Amount;

            //50202010
            ExSheet.Cells[14, 6] = Monthly[2].AB;
            ExSheet.Cells[14, 9] = Monthly[2].Amount;

            //50203010
            ExSheet.Cells[15, 6] = Monthly[3].AB;
            ExSheet.Cells[15, 9] = Monthly[3].Amount;

            //50203020
            ExSheet.Cells[16, 6] = Monthly[4].AB;
            ExSheet.Cells[16, 9] = Monthly[4].Amount;

            //50203070
            ExSheet.Cells[17, 6] = Monthly[5].AB;
            ExSheet.Cells[17, 9] = Monthly[5].Amount;

            //50203080
            ExSheet.Cells[18, 6] = Monthly[6].AB;
            ExSheet.Cells[18, 9] = Monthly[6].Amount;

            //50203090
            ExSheet.Cells[19, 6] = Monthly[7].AB;
            ExSheet.Cells[19, 9] = Monthly[7].Amount;

            //50203990
            ExSheet.Cells[20, 6] = Monthly[8].AB;
            ExSheet.Cells[20, 9] = Monthly[8].Amount;

            //50204010
            ExSheet.Cells[21, 6] = Monthly[9].AB;
            ExSheet.Cells[21, 9] = Monthly[9].Amount;

            //50204020
            ExSheet.Cells[22, 6] = Monthly[10].AB;
            ExSheet.Cells[22, 9] = Monthly[10].Amount;

            //50205010
            ExSheet.Cells[23, 6] = Monthly[11].AB;
            ExSheet.Cells[23, 9] = Monthly[11].Amount;

            //50205020.1
            ExSheet.Cells[24, 6] = Monthly[12].AB;
            ExSheet.Cells[24, 9] = Monthly[12].Amount;

            //50205020.2
            ExSheet.Cells[25, 6] = Monthly[13].AB;
            ExSheet.Cells[25, 9] = Monthly[13].Amount;

            //50205030
            ExSheet.Cells[26, 6] = Monthly[14].AB;
            ExSheet.Cells[26, 9] = Monthly[14].Amount;

            //50210030
            ExSheet.Cells[27, 6] = Monthly[15].AB;
            ExSheet.Cells[27, 9] = Monthly[15].Amount;

            //50211030
            ExSheet.Cells[28, 6] = Monthly[16].AB;
            ExSheet.Cells[28, 9] = Monthly[16].Amount;

            //50211990
            ExSheet.Cells[29, 6] = Monthly[17].AB;
            ExSheet.Cells[29, 9] = Monthly[17].Amount;

            //50212020
            ExSheet.Cells[30, 6] = Monthly[18].AB;
            ExSheet.Cells[30, 9] = Monthly[18].Amount;

            //50212030
            ExSheet.Cells[31, 6] = Monthly[19].AB;
            ExSheet.Cells[31, 9] = Monthly[19].Amount;

            //50213040
            ExSheet.Cells[32, 6] = Monthly[20].AB;
            ExSheet.Cells[32, 9] = Monthly[20].Amount;

            //50213050.1
            ExSheet.Cells[33, 6] = Monthly[21].AB;
            ExSheet.Cells[33, 9] = Monthly[21].Amount;

            //50213050.2
            ExSheet.Cells[34, 6] = Monthly[22].AB;
            ExSheet.Cells[34, 9] = Monthly[22].Amount;

            //50213050.3
            ExSheet.Cells[35, 6] = Monthly[23].AB;
            ExSheet.Cells[35, 9] = Monthly[23].Amount;

            //50213050.4
            ExSheet.Cells[36, 6] = Monthly[24].AB;
            ExSheet.Cells[36, 9] = Monthly[24].Amount;

            //50213050.5
            ExSheet.Cells[37, 6] = Monthly[25].AB;
            ExSheet.Cells[37, 9] = Monthly[25].Amount;

            //50213050.6
            ExSheet.Cells[38, 6] = Monthly[26].AB;
            ExSheet.Cells[38, 9] = Monthly[26].Amount;

            //50213050.7
            ExSheet.Cells[39, 6] = Monthly[27].AB;
            ExSheet.Cells[39, 9] = Monthly[27].Amount;

            //50213060
            ExSheet.Cells[40, 6] = Monthly[28].AB;
            ExSheet.Cells[40, 9] = Monthly[28].Amount;

            //50213070
            ExSheet.Cells[41, 6] = Monthly[29].AB;
            ExSheet.Cells[41, 9] = Monthly[29].Amount;

            //50216020
            ExSheet.Cells[42, 6] = Monthly[30].AB;
            ExSheet.Cells[42, 9] = Monthly[30].Amount;

            //50216030
            ExSheet.Cells[43, 6] = Monthly[31].AB;
            ExSheet.Cells[43, 9] = Monthly[31].Amount;

            //50299010
            ExSheet.Cells[44, 6] = Monthly[32].AB;
            ExSheet.Cells[44, 9] = Monthly[32].Amount;

            //50299020
            ExSheet.Cells[45, 6] = Monthly[33].AB;
            ExSheet.Cells[45, 9] = Monthly[33].Amount;

            //50299030
            ExSheet.Cells[46, 6] = Monthly[34].AB;
            ExSheet.Cells[46, 9] = Monthly[34].Amount;

            //50299050
            ExSheet.Cells[47, 6] = Monthly[35].AB;
            ExSheet.Cells[47, 9] = Monthly[35].Amount;

            //50299060
            ExSheet.Cells[48, 6] = Monthly[36].AB;
            ExSheet.Cells[48, 9] = Monthly[36].Amount;

            //50299070
            ExSheet.Cells[49, 6] = Monthly[37].AB;
            ExSheet.Cells[49, 9] = Monthly[37].Amount;

            //50299990
            ExSheet.Cells[50, 6] = Monthly[38].AB;
            ExSheet.Cells[50, 9] = Monthly[38].Amount;

            // Month
            ExSheet.Cells[7, 9] = month;

            try
            {
                ExSheet.SaveAs("DBMS\\Monthly\\MOOE\\" + month + "_REPORT_MOOE.xlsx");
            }
            catch
            {

            }
            finally
            {
                if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    ExSheet.PrintOutEx();

                ExApp.Visible = true;
                ExSheet.PrintPreview();
                ExBook.Close();
            }
        }
        public void createMonthlyFE(List<Entities.SAAOModel> Monthly, string month)
        {
            ExApp = new Excel.Application();
            ExApp.Visible = false;
            ExBook = ExApp.Workbooks.Open("C:\\Monthly.xlsx");
            ExSheet = (Excel.Worksheet)ExBook.Sheets[2];

            ////////FINANCIAL EXPENSES MONTHLY

            // Month, Year
            ExSheet.Cells[3, 1] = month + DateTime.Now.Year.ToString();

            /////// FINANCIAL EXPENSES

            //50301040
            ExSheet.Cells[12, 6] = Monthly[0].AB;
            ExSheet.Cells[12, 9] = Monthly[0].Amount;

            // Month
            ExSheet.Cells[7, 9] = month;

            try
            {
                ExSheet.SaveAs("DBMS\\Monthly\\FE\\" + month + "_REPORT_FE.xlsx");
            }
            catch
            {

            }
            finally
            {
                if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    ExSheet.PrintOutEx();

                ExApp.Visible = true;
                ExSheet.PrintPreview();
                ExBook.Close();
            }
        }
        public void createMonthlyPS(List<Entities.SAAOModel> Monthly, string month)
        {
            ExApp = new Excel.Application();
            ExApp.Visible = false;
            ExBook = ExApp.Workbooks.Open("C:\\Monthly.xlsx");
            ExSheet = (Excel.Worksheet)ExBook.Sheets[4];

            ////////PERSONAL SERVICES

            // Month, Year
            ExSheet.Cells[3, 1] = month + DateTime.Now.Year.ToString();

            //PERSONAL SERVICES

            //50101010
            ExSheet.Cells[12, 6] = Monthly[0].AB;
            ExSheet.Cells[12, 9] = Monthly[0].Amount;

            //50101020.1
            ExSheet.Cells[14, 6] = Monthly[1].AB;
            ExSheet.Cells[14, 9] = Monthly[1].Amount;

            //50101020.2
            ExSheet.Cells[15, 6] = Monthly[2].AB;
            ExSheet.Cells[15, 9] = Monthly[2].Amount;

            //50102010
            ExSheet.Cells[16, 6] = Monthly[3].AB;
            ExSheet.Cells[16, 9] = Monthly[3].Amount;

            //50102020
            ExSheet.Cells[17, 6] = Monthly[4].AB;
            ExSheet.Cells[17, 9] = Monthly[4].Amount;

            //50102030
            ExSheet.Cells[18, 6] = Monthly[5].AB;
            ExSheet.Cells[18, 9] = Monthly[5].Amount;

            //50102040
            ExSheet.Cells[19, 6] = Monthly[6].AB;
            ExSheet.Cells[19, 9] = Monthly[6].Amount;

            //50102050
            ExSheet.Cells[21, 6] = Monthly[7].AB;
            ExSheet.Cells[21, 9] = Monthly[7].Amount;

            //50102060
            ExSheet.Cells[22, 6] = Monthly[8].AB;
            ExSheet.Cells[22, 9] = Monthly[8].Amount;

            //50102100
            ExSheet.Cells[20, 6] = Monthly[9].AB;
            ExSheet.Cells[20, 9] = Monthly[9].Amount;

            //50102110
            ExSheet.Cells[23, 6] = Monthly[10].AB;
            ExSheet.Cells[23, 9] = Monthly[10].Amount;

            //50102120
            ExSheet.Cells[24, 6] = Monthly[11].AB;
            ExSheet.Cells[24, 9] = Monthly[11].Amount;

            //50102130
            ExSheet.Cells[25, 6] = Monthly[12].AB;
            ExSheet.Cells[25, 9] = Monthly[12].Amount;

            //50102140
            ExSheet.Cells[27, 6] = Monthly[13].AB;
            ExSheet.Cells[27, 9] = Monthly[13].Amount;

            //50102150
            ExSheet.Cells[26, 6] = Monthly[14].AB;
            ExSheet.Cells[26, 9] = Monthly[14].Amount;

            //50102990
            ExSheet.Cells[28, 6] = Monthly[15].AB;
            ExSheet.Cells[28, 9] = Monthly[15].Amount;

            //50103010
            ExSheet.Cells[29, 6] = Monthly[16].AB;
            ExSheet.Cells[29, 9] = Monthly[16].Amount;

            //50103020
            ExSheet.Cells[30, 6] = Monthly[17].AB;
            ExSheet.Cells[30, 9] = Monthly[17].Amount;

            //50103030
            ExSheet.Cells[31, 6] = Monthly[18].AB;
            ExSheet.Cells[31, 9] = Monthly[18].Amount;

            //50103040
            ExSheet.Cells[32, 6] = Monthly[19].AB;
            ExSheet.Cells[32, 9] = Monthly[19].Amount;

            //50104030
            ExSheet.Cells[33, 6] = Monthly[20].AB;
            ExSheet.Cells[33, 9] = Monthly[20].Amount;

            //50104990
            ExSheet.Cells[34, 6] = Monthly[21].AB;
            ExSheet.Cells[34, 9] = Monthly[21].Amount;

            // Month
            ExSheet.Cells[7, 9] = month;

            try
            {
                ExSheet.SaveAs("DBMS\\Monthly\\PS\\" + month + "_REPORT_PS.xlsx");
            }
            catch
            {

            }
            finally
            {
                if (MessageBox.Show("Do you want to continue to printing?", "Print?",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    ExSheet.PrintOutEx();

                ExApp.Visible = true;
                ExSheet.PrintPreview();
                ExBook.Close();
            }
        }
    }
}
