using System;
using System.Data.SqlClient;
using BUR_UI.Entities;
using BUR_UI.Interface;

namespace BUR_UI.Context
{
    public class DbInsert
    {
        public void InsertBUR(BURModel BUR)
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = Properties.Resources.ConnectionString;

            using (conn)
            {
                conn.Open();

                //Insert items first.
                foreach (var item in BUR.Particulars)
                {
                    SqlCommand _comm = new SqlCommand(
                        "INSERT INTO dbo.tbl_Item (Item_Amount, Acct_Code, BUR_No) " +
                        "VALUES (" + item.Amount + ", " + item.Code + ", '" + item.BUR_Number +
                        "')",
                        conn);

                    _comm.ExecuteNonQuery();
                }

                //Insert PR Number.
                SqlCommand comm = new SqlCommand(
                    "INSERT INTO dbo.tbl_PR VALUES " +
                    "('" + BUR.PRNumber + "', '" + DateTime.Today.ToString("yyyy-MM-dd") + "')", conn);

                comm.ExecuteNonQuery();



                //Insert BUR Metadata.
                comm.CommandText =
                    "INSERT INTO dbo.tbl_BUR VALUES " +
                    "('" + BUR.BURNumber + "', '" + BUR.Date + "', '" + BUR.OfficeCode +
                    "', '20030210', '" + BUR.BStaff_Number + "', '" + BUR.Payee_Number +
                    "', '" + BUR.PRNumber + "', '" + BUR.Description + "', '" + BUR.SignatoryOfficeCode + "')";

                comm.ExecuteNonQuery();
            }
        }
        public void UpdateBUR(BURModel BUR)
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = Properties.Resources.ConnectionString;

            using (conn)
            {
                conn.Open();

                //Insert items first.
                foreach (var item in BUR.Particulars)
                {
                    SqlCommand _comm = new SqlCommand(
                        "UPDATE dbo.tbl_Item " +
                        "SET Item_Amount = " + item.Amount + ", Acct_Code = '" + item.Code + "', BUR_No = '" + item.BUR_Number +
                        "' WHERE BUR_No = '" + item.BUR_Number + "'",
                        conn);

                    _comm.ExecuteNonQuery();
                }

                //Insert BUR Metadata.
                SqlCommand comm = new SqlCommand(
                    "UPDATE dbo.tbl_BUR SET " +
                    "BUR_No = '" + BUR.BURNumber + "', BUR_FDate = '" + BUR.Date + "', Office_Code = '" + BUR.OfficeCode +
                    "', BDHead_Number = '20030210', BStaff_Number = '" + BUR.BStaff_Number + "', Employee_Number = '" + BUR.Payee_Number +
                    "', PR_Code = '" + BUR.PRNumber + "', Description = '" + BUR.Description + "'" +
                    "WHERE BUR_No = '" + BUR.BURNumber + "'", conn);

                comm.ExecuteNonQuery();
            }
        }

        public void EditAccount(AccountGridModel acct)
        {
            DbLink link = new DbLink();

            using (SqlConnection conn = link.InitSql())
            {
                // Insert to the AB table.
                SqlCommand comm = new SqlCommand("UPDATE dbo.tbl_AB SET AB_Amount = " + acct.AB + " WHERE Acct_Code = '" + acct.AcctCode + "'", conn);
                conn.Open();
                comm.ExecuteNonQuery();
            }
        }
    }
}
