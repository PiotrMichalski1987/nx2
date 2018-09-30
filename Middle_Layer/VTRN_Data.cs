using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace WrkWebApp.Middle_Layer
{
    public class VTRN_Data
    {

        private DbConnections_Db conn;
        public int Id { set; get; }
        public int Veh { set; get; }
        public string Vtrn_Veh_Code { set; get; }

        public float Vtrn_Monies { set; get; }

        public DateTime Vtrn_Date_Driver { set; get; }


        public VTRN_Data()
        {
            conn = new DbConnections_Db();
        }


        public DataTable EmptyDataTable()
        {
            Id = -1;
            conn.addParameter("@Id", Id);

            String comm = "SELECT * FROM VTRN_Data WHERE Id=@Id";
            DataTable dt = new DataTable();

            dt = conn.executeReader(comm);

            return dt;
        }


        public bool AddRow()
        {
            conn.addParameter("@Veh", Veh);
            conn.addParameter("@Vtrn_Veh_Code", Vtrn_Veh_Code);
            conn.addParameter("@Vtrn_Monies", Vtrn_Monies);
            conn.addParameter("@Vtrn_Date_Driver", Vtrn_Date_Driver);

            string command = "INSERT INTO VTRN_Data " +
               "(Veh, Vtrn_Veh_Code, Vtrn_Monies, Vtrn_Date_Driver  ) " +
           "VALUES " +
           "(@Veh, @Vtrn_Veh_Code, @Vtrn_Monies, @Vtrn_Date_Driver  )";

            return conn.executeNonQuery(command) > 0;

        }

        public DataTable SelectUsingVehAndDate()
        {
            conn.addParameter("@Veh", Veh);
            //conn.addParameter("@Vtrn_Date_Driver", Vtrn_Date_Driver.ToString("dd/MM/yyyy"));

            ///string comm = "SELECT * FROM VTRN_Data WHERE Veh=@Veh AND FORMAT(Vtrn_Date_Driver , 'dd/MM/yyyy')=@Vtrn_Date_Driver";
            string comm = "SELECT * FROM VTRN_Data WHERE Veh=@Veh ";
            DataTable dt = conn.executeReader(comm);

            return dt;


        }

        



    }
}