using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;

namespace efind
{
    public class _SQLConnection
    {
        //string constr = ConfigurationManager.ConnectionStrings["constr"].ConnectionString;

        //public static string constr ="";
        public static SqlConnection conn;

        public _SQLConnection()
        {
            conn = new SqlConnection();
        }

        public string BuildSQLConnectionString(string datasource, string database, string user, string password)
        {
            System.Data.SqlClient.SqlConnectionStringBuilder builder =
            new System.Data.SqlClient.SqlConnectionStringBuilder();
            
            builder.DataSource  = datasource ;
            builder.IntegratedSecurity  = true;
            //builder.InitialCatalog  = "AdventureWorks;NewValue=Bad";
            builder.InitialCatalog = database;
            builder.UserID = user;
            builder.Password = password;
            builder["Trusted_Connection"] = true;
            //Console.WriteLine(builder.ConnectionString);
            return builder.ConnectionString; 
        }

        public void ConnectToSqL(string datasource, string database, string user, string password)
        {
            try
            {

                conn.ConnectionString = BuildSQLConnectionString(datasource, database, user, password);
            }

            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void SQLDisconnect()
        {
            try
            {
                conn.Close();
            }
            catch (Exception ex)
            {

            }
        }
        public void GetData()
        {
            SqlCommand command1 = new SqlCommand();
            SqlDataReader reader = null;
            
            try
            {
                //SQL Command : select * from obj as a left join field as b on a.obj_id = b.obj_id and a.obj_type ='D' order by a.obj_id;

                // 1 obj ID may map to many obj row -> with different indexing, 
                command1.CommandText = "Select EmployeeID, Title, BirthDate From HumanResources.Employee";
                command1.CommandType = CommandType.Text;
                command1.Connection = conn;
                //reader = command1.ExecuteReader();

                using (SqlDataReader sdr = command1.ExecuteReader())
                {
                    //Create a new DataSet.
                    DataSet ds = new DataSet();
                    ds.Tables.Add("Customers");

                    //Load DataReader into the DataTable.
                    ds.Tables[0].Load(sdr);
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
