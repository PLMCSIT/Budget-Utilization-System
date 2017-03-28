using System.Data.SqlClient;

namespace BUR_UI.Interface
{
    public class Typer
    {
        public string GetSelectedOfficeCode(string Office_Name)
        {
            DbLink DbLink = new DbLink();
            string Office_Code = "12";

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Office_Code " +
                    "FROM dbo.tbl_A_Certified " +
                    "WHERE Office_NameAbbr = '" + Office_Name + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();


                if (reader.HasRows)
                {
                    while (reader.Read())
                        Office_Code = reader.GetString(0);
                }

                reader.Close();
            }

            return Office_Code;
        }
        public int GetSelectedClassCode(string Class_Name)
        {
            DbLink DbLink = new DbLink();
            int Class_Code = 0;

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Acct_ClassId " +
                    "FROM dbo.tbl_Classification " +
                    "WHERE Acct_Class_Name = '" + Class_Name + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();


                if (reader.HasRows)
                {
                    while (reader.Read())
                        Class_Code = reader.GetInt32(0);
                }

                reader.Close();
            }

            return Class_Code;
        }
        public string GetUserImage(string userName)
        {
            DbLink DbLink = new DbLink();
            string picURL = "";

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT  PicURL " +
                    "FROM dbo.tbl_BO_Staff " +
                    "WHERE BStaff_Number = '" + userName + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();


                if (reader.HasRows)
                {
                    while (reader.Read())
                        picURL = reader.GetString(0);
                }

                reader.Close();
            }

            return picURL;
        }
        public bool CheckIfAdmin(string userName)
        {
            DbLink dbLink = new DbLink();

            using (SqlConnection conn = dbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT * FROM dbo.tbl_BO_Staff " +
                    "WHERE BStaff_Number = '" + userName + "' AND "+
                    "Discriminator = 'Admin'", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                return reader.HasRows;
            }
        }
        public string GetPayeeId(string Payee)
        {
            DbLink DbLink = new DbLink();
            string Employee_Number = "";

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Employee_Number " +
                    "FROM dbo.tbl_Payee " +
                    "WHERE Employee_Name = '" + Payee + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();


                if (reader.HasRows)
                {
                    while (reader.Read())
                        Employee_Number = reader.GetString(0);
                } else
                {
                    reader.Close();
                    comm = new SqlCommand(
                    "SELECT Employee_Number " +
                    "FROM dbo.tbl_Ext_Payee " +
                    "WHERE Employee_Name = '" + Payee + "'",
                    conn);

                    reader = comm.ExecuteReader();


                    if (reader.HasRows)
                    {
                        while (reader.Read())
                            Employee_Number = reader.GetString(0);
                    }
                }

                reader.Close();
            }

            return Employee_Number;
        }
        public string GetSelectedOfficeName(string Office_Code)
        {
            DbLink DbLink = new DbLink();
            string Office_Name = "";

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Office_NameAbbr " +
                    "FROM dbo.tbl_A_Certified " +
                    "WHERE Office_Code = '" + Office_Code + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();


                if (reader.HasRows)
                {
                    while (reader.Read())
                        Office_Name = reader.GetString(0);
                }

                reader.Close();
            }

            return Office_Name;
        }
        public string GetPosition(string userName)
        {
            DbLink DbLink = new DbLink();
            string Password = "";

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Position FROM dbo.tbl_BO_Staff WHERE BStaff_Number = '" + userName + "'", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Password = reader.GetString(0);
                    }
                }
            }

            return Password;
        }
        public string GetSelectedStaffName(string BStaff_Number)
        {
            DbLink DbLink = new DbLink();
            string BStaff_Name = "";

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT BStaff_Name " +
                    "FROM dbo.tbl_BO_Staff " +
                    "WHERE BStaff_Number = '" + BStaff_Number + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                        BStaff_Name = reader.GetString(0);
                }

                reader.Close();
            }

            return BStaff_Name;
        }
        public string GetSelectedPayeeName(string Payee_Number)
        {
            DbLink DbLink = new DbLink();
            string Payee = "";

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Employee_Name " +
                    "FROM dbo.tbl_Payee " +
                    "WHERE Employee_Number = '" + Payee_Number + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                        Payee = reader.GetString(0);
                }

                reader.Close();
            }

            if (Payee == "") Payee = Payee_Number;

            return Payee;
        }
        public string GetSelectedBDHeadName(string BDHead_Number)
        {
            DbLink DbLink = new DbLink();
            string BDHead_Name = "";

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT BDHead_Name " +
                    "FROM dbo.tbl_B_Certified " +
                    "WHERE BDHead_Number = '" + BDHead_Number + "'",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                        BDHead_Name = reader.GetString(0);
                }

                reader.Close();
            }

            return BDHead_Name;
        }
        public string GetClassName(string code)
        {
            DbLink DbLink = new DbLink();

            string Acct_Class_Name = "";
            
            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT A.Acct_Class_Name " + 
                    "FROM dbo.tbl_Classification AS A " +
                    "INNER JOIN dbo.tbl_Particulars AS B " +
                    "ON A.Acct_ClassId = B.Acct_ClassId " +
                    "WHERE B.Acct_Code = '" + code + "'", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Acct_Class_Name = reader.GetString(0);
                    }
                }
            }

            return Acct_Class_Name;
        }
        public string GetAcctName(string code)
        {
            DbLink DbLink = new DbLink();
            string Acct_Name = "";

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Acct_Name " +
                    "FROM dbo.tbl_Particulars " +
                    "WHERE Acct_Code = '" + code + "'", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Acct_Name = reader.GetString(0);
                    }
                }
            }

            return Acct_Name;
        }
        public string[] GetOfficeHeadName(string officeCode)
        {
            DbLink DbLink = new DbLink();
            string[] Officehead = new string[2];

            using (SqlConnection conn = DbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT A.Officehead_Name, A.Officehead_Pos " +
                    "FROM dbo.tbl_Officehead AS A " +
                    "INNER JOIN dbo.tbl_A_Certified AS B " +
                    "ON A.OfficeheadId = B.OfficeheadId " +
                    "WHERE B.Office_Code = '" + officeCode + "'", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Officehead[0] = reader.GetString(0);
                        Officehead[1] = reader.GetString(1);
                    }
                }
            }

            return Officehead;
        }
        public string GetSelectedStaffCode(string Staff_Name)
        {
            DbLink dbLink = new DbLink();

            string Staff_Number = "";

            using (SqlConnection conn = dbLink.InitSql())
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT BStaff_Number FROM dbo.tbl_BO_Staff WHERE " +
                    "BStaff_Name LIKE '%" + Staff_Name + "%'", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        Staff_Number = reader.GetString(0);
                    }
                }
            }

            return Staff_Number;
        }
    }
}
