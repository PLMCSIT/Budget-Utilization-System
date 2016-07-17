using System.Data.SqlClient;

namespace BUR_UI.Context
{
    public class DbDelete
    {
        public void DeleteBUR(string BUR_Number)
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = Properties.Resources.ConnectionString;

            using (conn)
            {
                // Delete actual BUR.
                SqlCommand comm = new SqlCommand(
                    "DELETE FROM dbo.tbl_BUR WHERE BUR_No = '" + BUR_Number + "'", conn);

                conn.Open();

                comm.ExecuteNonQuery();

                // Delete BUR Items.
                comm.CommandText = "DELETE FROM dbo.tbl_Item WHERE BUR_No = '" + BUR_Number + "'";

                comm.ExecuteNonQuery();
            }
        }
    }
}
