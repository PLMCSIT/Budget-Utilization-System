using System.Data.SqlClient;
using BUR_UI.Entities;

namespace BUR_UI.Context
{
    public class DbUpdate
    {
        public BURModel FillEditor(string BUR_Number)
        {
            BURModel BUR = new Entities.BURModel();

            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = Properties.Resources.ConnectionString;

            using (conn)
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT * FROM dbo.tbl_BUR " +
                    "WHERE BUR_No = '" + BUR_Number + "'", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        BUR.BURNumber = reader.GetString(1);
                        BUR.Date = reader.GetDateTime(2).ToString();
                        BUR.OfficeCode = reader.GetString(3);
                        BUR.BDHead_Number = reader.GetString(4);
                        BUR.BStaff_Number = reader.GetString(5);
                        BUR.Payee_Number = reader.GetString(6);
                        BUR.PRNumber = reader.GetString(7);
                        BUR.Description = reader.GetString(8);
                    }
                }
            }

            BUR = FillParticulars(BUR);

            return BUR;
        }
        public BURModel FillParticulars(BURModel BUR)
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = Properties.Resources.ConnectionString;

            using (conn)
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT * FROM dbo.tbl_Item " +
                    "WHERE BUR_No = '" + BUR.BURNumber + "'", conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        BUR.Particulars.Add(
                            new Entities.Items()
                            {
                                Amount = float.Parse(reader.GetDouble(1).ToString()),
                                Code = reader.GetString(2),
                                BUR_Number = reader.GetString(3)
                            });
                    }
                }
            }

            return BUR;
        }
        public void UpdateUser(UserModel user)
        {

            using (SqlConnection conn = new SqlConnection())
            {
                conn.ConnectionString = Properties.Resources.ConnectionString;

                SqlCommand comm = new SqlCommand(
                    "UPDATE dbo.tbl_BO_Staff " +
                    "SET BStaff_Name = '" + user.User_Name + "', " +
                    "Discriminator = '" + user.Discriminator + "', " +
                    "Position = '" + user.Position + "', " +
                    "PicURL = '" + user.Picture + "'" +
                    "WHERE BStaff_Number = '" + user.User_Number + "'", conn);

                conn.Open();

                comm.ExecuteNonQuery();
            }
        }
    }
}
