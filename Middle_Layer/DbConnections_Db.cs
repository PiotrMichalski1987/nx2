using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace WrkWebApp.Middle_Layer
{
    public class DbConnections_Db
    {

        private static string dbName = "prg_data.mdf";
        private static string dbPath = System.AppDomain.CurrentDomain.BaseDirectory + "App_Data\\" + dbName;

        private static string connString = "Data Source=(LocalDB)\\MSSQLLocalDB;AttachDbFilename=" + dbPath + ";Integrated Security=True";


        private DbConnection conn;
        private DbCommand cmd;
        private List<DbParameter> param;


        internal void addParameter(string name, object value)
        {
            param.Add(new SqlParameter(name, value));
        }


        public DbConnections_Db()
        {
            conn = new SqlConnection(connString);
            cmd = new SqlCommand();
            cmd.Connection = conn;
            param = new List<DbParameter>();
        }




        public int executeNonQuery(string command)
        {
            setupSqlCommand(command);

            try
            {
                conn.Open();
                return cmd.ExecuteNonQuery();
            }
            catch
            {
                return -1; //failure
            }
            finally
            {
                conn.Close();
            }
        }

        public int executeScalar(string command)
        {
            setupSqlCommand(command);

            try
            {
                conn.Open();

                return (int)cmd.ExecuteScalar();
            }
            catch
            {
                return -1;

            }
            finally
            {
                conn.Close();
            }
        }

        public DataTable executeReader(string command)
        {
            setupSqlCommand(command);

            DataTable table = new DataTable();

            try
            {
                conn.Open();
                table.Load(cmd.ExecuteReader()); //automatically closes the data reader
                return table;
            }
            catch
            {


                return null; //failure
            }
            finally
            {
                conn.Close();
            }
        }

        public DataTable fillDataTable(string command)
        {
            setupSqlCommand(command);

            DataTable table = new DataTable();

            DbDataAdapter adapter = new SqlDataAdapter();
            adapter.SelectCommand = cmd;

            try
            {
                adapter.Fill(table); // Fill will open and close connection (so long as it hasn't been opened before calling it)

                return table;
            }
            catch
            {
                return null; //failure
            }
        }

        private void setupSqlCommand(string command)
        {
            cmd.CommandText = command;
            //comm.CommandType = CommandType.StoredProcedure; //UNCOMMENT IF USING STORED PROCEDURES

            cmd.Parameters.Clear(); //clear params from any previously executed command

            foreach (DbParameter dbP in param)
            {
                cmd.Parameters.Add(dbP);
            }

            param.Clear(); //clear list of params ready for next command
        }
    }
}