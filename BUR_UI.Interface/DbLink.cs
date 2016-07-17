using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using BUR_UI.Entities;

namespace BUR_UI.Interface
{
    public class DbLink
    {
        Typer GetId = new Typer();

        public List<AccountsModel> FillAccountsModel()
        {
            List<AccountsModel> Acct = new List<AccountsModel>();
            string yearStart = DateTime.Today.Year.ToString() + "-01-01";
            string dateNow = DateTime.Now.Year.ToString() + "-" +
                DateTime.Now.Month.ToString("D2") + "-" +
                DateTime.Now.AddDays(1).Day.ToString("D2");

            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT A.Item_Amount, A.Acct_Code " +
                    "FROM dbo.tbl_Item AS A " +
                    "INNER JOIN dbo.tbl_BUR AS B " +
                    "ON A.BUR_No = B.BUR_No " +
                    "WHERE B.BUR_FDate BETWEEN '" + yearStart + "' AND '" + dateNow + "'", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Acct.Add(
                            new AccountsModel()
                            {
                                Amount = float.Parse(reader.GetDouble(0).ToString()),
                                AccountCode = reader.GetString(1)
                            });
                    }
                }
            }

            return Acct;
        }
        public List<String> FillOffice()
        {
            List<string> Offices = new List<string>();

            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Office_NameAbbr FROM dbo.tbl_A_Certified",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                        Offices.Add(reader.GetString(0));
                }

                reader.Close();
            }

            return Offices;
        }
        public void ChangePassword(string staffNumber, string newPassword)
        {
            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "UPDATE dbo.tbl_BO_Staff " +
                    "SET Password = '" + newPassword + "' " +
                    "WHERE BStaff_Number = '" + staffNumber + "'", conn);

                conn.Open();

                comm.ExecuteNonQuery();
            }
        }
        public bool userValidate(string User, string Pass)
        {
            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT * FROM dbo.tbl_BO_Staff " +
                    "WHERE BStaff_Number = '" + User + "' AND Password = '" + Pass + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    return true;
                }

                return false;
            }
        }
        public List<string> FillPR()
        {
            List<string> PRs = new List<string>();

            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT PR_Code FROM dbo.tbl_PR",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                        PRs.Add(reader.GetString(0));
                }

                reader.Close();
            }

            return PRs;
        }
        public List<string> FillPayeeByOffice(string Office_Name)
        {
            Typer Typer = new Typer();

            List<string> Payee = new List<string>();

            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Employee_Name " +
                    "FROM dbo.tbl_Payee " +
                    "WHERE Office_Code = '" + Typer.GetSelectedOfficeCode(Office_Name) + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                        Payee.Add(reader.GetString(0));
                }

                reader.Close();
            }

            return Payee;
        }     
        public List<string> FillClass()
        {
            List<string> Classes = new List<string>();

            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Acct_Class_Name FROM dbo.tbl_Classification",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                        Classes.Add(reader.GetString(0));
                }

                reader.Close();
            }

            return Classes;
        }
        public List<string> FillCodeByClass(string Class_Name)
        {
            Typer Typer = new Typer();

            List<string> Code = new List<string>();

            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Acct_Code " +
                    "FROM dbo.tbl_Particulars " +
                    "WHERE Acct_ClassId = '" + Typer.GetSelectedClassCode(Class_Name) + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Code.Add(reader.GetString(0));
                    }  
                }

                reader.Close();
            }

            return Code;
        }
        public string FillNameByCode(int Acct_Code)
        {
            string name = "";

            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Acct_Name " +
                    "FROM dbo.tbl_Particulars " +
                    "WHERE Acct_Code = '" + Acct_Code + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        name = reader.GetString(0);
                    }
                }

                reader.Close();
            }

            return name;
        }
        public List<Entities.BURModel> FillGrid()
        {
            List<Entities.BURModel> BUR = new List<Entities.BURModel>();
            Typer Typer = new Typer();

            SqlConnection conn = InitSql();

            using (conn)
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT * FROM dbo.tbl_BUR", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        BUR.Add(
                            new Entities.BURModel()
                            {
                                BURNumber = reader.GetString(1),
                                Date = reader.GetDateTime(2).ToString(),
                                Office = Typer.GetSelectedOfficeName(reader.GetString(3)),
                                Staff = Typer.GetSelectedStaffName(reader.GetString(5)),
                                Payee = Typer.GetSelectedPayeeName(reader.GetString(6))
                            });
                    }
                }
            }

            return BUR;
        }
        public SqlConnection InitSql()
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = Properties.Resources.ConnectionString;

            return conn;
        }
        public List<ABModel> FillABModel()
        {
            List<ABModel> AB = new List<ABModel>();

            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT AB_Amount, Acct_Code " +
                    "FROM dbo.tbl_AB", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        AB.Add(
                            new ABModel()
                            {
                                ApprovedBudget = float.Parse(reader.GetDouble(0).ToString()),
                                AccountCode = reader.GetString(1)
                            });
                    }
                }
            }

            return AB;
        }
        public List<AccountsModel> FillAccountsModel(int ClassId, string startDate, string endDate)
        {
            List<AccountsModel> Acct = new List<AccountsModel>();

            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT A.Item_Amount, A.Acct_Code " +
                    "FROM dbo.tbl_Item AS A " +
                    "INNER JOIN dbo.tbl_BUR AS B ON A.BUR_No = B.BUR_No " +
                    "INNER JOIN dbo.tbl_Particulars AS C ON A.Acct_Code = C.Acct_Code " +
                    "INNER JOIN dbo.tbl_Classification AS D ON D.Acct_ClassId = C.Acct_ClassId " +
                    "WHERE B.BUR_FDate BETWEEN '" + startDate + "' AND '" + endDate + "' " +
                    "AND D.Acct_ClassId = " + ClassId, conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Acct.Add(
                            new AccountsModel()
                            {
                                Amount = float.Parse(reader.GetDouble(0).ToString()),
                                AccountCode = reader.GetString(1)
                            });
                    }
                }
            }

            return Acct;
        }
        public List<ABModel> FillABModel(int ClassId)
        {
            List<ABModel> AB = new List<ABModel>();

            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT A.AB_Amount, A.Acct_Code " +
                    "FROM dbo.tbl_AB AS A " +
                    "INNER JOIN dbo.tbl_Particulars AS B ON A.Acct_Code = B.Acct_Code " +
                    "INNER JOIN dbo.tbl_Classification AS C ON B.Acct_ClassId = C.Acct_ClassId " +
                    "WHERE C.Acct_ClassId = " + ClassId, conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        AB.Add(
                            new ABModel()
                            {
                                ApprovedBudget = float.Parse(reader.GetDouble(0).ToString()),
                                AccountCode = reader.GetString(1)
                            });
                    
                    }
                }
            }

            return AB;
        }
        public List<BURModel> FillGrid(string text)
        {
            Typer typer = new Typer();
            List<BURModel> BUR = new List<BURModel>();
            Typer Typer = new Typer();

            SqlConnection conn = InitSql();

            using (conn)
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT * FROM dbo.tbl_BUR " +
                    "WHERE BUR_No LIKE '%" + text + "%'" /* "OR " +
                    "Office_Code LIKE '%" + typer.GetSelectedOfficeCode(text) + "%' OR " +
                    "Employee_Number LIKE '%" + typer.GetPayeeId(text) + "%' OR " +
                    "BStaff_Number LIKE '%" + typer.GetSelectedStaffCode(text) + "%'" */, conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        BUR.Add(
                            new BURModel()
                            {
                                BURNumber = reader.GetString(1),
                                Date = reader.GetDateTime(2).ToString(),
                                Office = Typer.GetSelectedOfficeName(reader.GetString(3)),
                                Staff = Typer.GetSelectedStaffName(reader.GetString(5)),
                                Payee = Typer.GetSelectedPayeeName(reader.GetString(6))
                            });
                    }
                }
            }

            return BUR;
        }
        public List<UserModel> FillUserModel(List<UserModel> users)
        {
            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT * FROM dbo.tbl_BO_Staff", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        users.Add(new UserModel()
                        {
                            User_Number = reader.GetString(1),
                            User_Name = reader.GetString(2),
                            Discriminator = reader.GetString(3),
                            Position = reader.GetString(5),
                            Picture = reader.GetString(6)
                        });
                    }
                }
            }

            return users;
        }
        public List<AccountGridModel> FillAccountGridModel(List<AccountGridModel> accounts)
        {
            using (SqlConnection conn = InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT A.Acct_Code, A.Acct_Name, B.Acct_Class_Name " +
                    "FROM dbo.tbl_Particulars AS A " +
                    "INNER JOIN dbo.tbl_Classification AS B " +
                    "ON A.Acct_ClassId = B.Acct_ClassId " +
                    "ORDER BY A.Acct_Code ASC", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        accounts.Add(new AccountGridModel()
                        {
                            AcctCode = reader.GetString(0),
                            AcctName = reader.GetString(1),
                            AcctClass = reader.GetString(2),
                        });
                    }
                }
            }

            return accounts;
        }
    }
}
