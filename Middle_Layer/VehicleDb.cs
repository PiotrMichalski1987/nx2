using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace WrkWebApp.Middle_Layer
{
    public class VehicleDb
    {

        private DbConnections_Db conn;
        public string Code { set; get; }
        public int Id { set; get; }


        public VehicleDb()
        {
            conn = new DbConnections_Db();
        }


        public DataTable SelectVehicleById()
        {
            conn.addParameter("@Id", Id);
            
     

            string comm = "SELECT * FROM Vehicle WHERE Id=@Id";

            DataTable dt = conn.executeReader(comm);

            return dt;


        }

        public DataTable SelectAllVehicles()
        {
            string command = "SELECT * FROM Vehicle";

            DataTable tb = conn.executeReader(command);

            return tb;

        }

    }
}