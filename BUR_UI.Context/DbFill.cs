using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BUR_UI.Interface
{
    public class DbFill
    {
        public List<string> FillOffice()
        {
            SqlConnection conn = new SqlConnection();
            conn.ConnectionString = Context.Properties.Resources.ConnectionString;

            using (conn)
            {
                SqlCommand comm = new SqlCommand(
                    "SELECT Office_Name FROM dbo.tbl_A_Certified",
                    conn);

                conn.Open();

                SqlDataReader reader = comm.ExecuteReader();

                List<string> Offices = new List<string>();

                if (reader.HasRows)
                {
                    while (reader.Read())
                        Offices.Add(reader.GetString(0));
                }

                return Offices;
            }
        }


        


    }
}
