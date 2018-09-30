using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace WrkWebApp.Middle_Layer
{
    public class ClientDb
    {
        public int Id { set; get; }
        public String Name {set;get;}

        private DbConnections_Db conn;


        public ClientDb()
        {
            conn = new DbConnections_Db();
        }

        public DataTable SelectAllClients()
        {
            string command = "SELECT * FROM Client";

            DataTable tb = conn.executeReader(command);

            return tb;

        }


        public bool AddName()
        {
            conn.addParameter("@Name", Name);
            

            string command = "INSERT INTO Client " +
               "(Name) " +
           "VALUES " +
           "(@Name)";

            return conn.executeNonQuery(command) > 0;

        }

    }
}