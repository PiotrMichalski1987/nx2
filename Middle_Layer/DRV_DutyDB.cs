using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace WrkWebApp.Middle_Layer
{
    public class DRV_DutyDB
    {
        private DbConnections_Db conn;

        public int Id { set; get; }

        public int Veh { set; get; }
        public string VehCode { set; get; }

        public int Drv { set; get; }

        public DateTime Date { set; get; }
        public TimeSpan Duty_Start { set; get; }
        public TimeSpan Duty_End { set; get; }
        public TimeSpan Duty_Time { set; get; }

        public float Total_Km { set; get; }

        public float MPG { set; get; }

        public float Co2_Kg { set; get; }

        public string DrvName { set; get; }

        public string _drivers_Employment_Type { set; get; }

        public float _drivers_Overtime_Rate { set; get; }

        public float _drivers_Standard_Rate { set; get; }


        public DRV_DutyDB()
        {
            conn = new DbConnections_Db();
           
        }

        public DataTable SelectRowsUsingDateAndVeh()
        {
            conn.addParameter("@Date", Date.ToString("dd/MM/yyyy"));
            conn.addParameter("@Veh", Veh);

            string cmd = "SELECT * FROM DRV_Duty WHERE Veh=@Veh AND FORMAT(Date , 'dd/MM/yyyy')=@Date";
            // string command = "SELECT * FROM Costing WHERE FORMAT(Date , 'dd/MM/yyyy') BETWEEN @DateFrom AND @DateTo";
            DataTable dt = conn.executeReader(cmd);


            return dt;

        }


        public DataTable SelectRowsUsingDateVehAndDrv()
        {
            conn.addParameter("@Date", Date.ToString("dd/MM/yyyy"));
            conn.addParameter("@Veh", Veh);
            conn.addParameter("@Drv", Drv);

            string cmd = "SELECT * FROM DRV_Duty WHERE Veh=@Veh AND FORMAT(Date , 'dd/MM/yyyy')=@Date AND Drv=@Drv";
            DataTable dt = conn.executeReader(cmd);


            return dt;

        }


        public DataTable EmptyDataTable()
        {
            Id = -1;
            conn.addParameter("@Id", Id);

            String comm = "SELECT * FROM DRV_Duty WHERE Id=@Id";
            DataTable dt = new DataTable();

            dt = conn.executeReader(comm);

            return dt;
        }

        public DataTable SelectByDrvDateVeh()
        {
            conn.addParameter("@Date", Date.ToString("dd/MM/yyyy"));
            conn.addParameter("@Veh", Veh);
            conn.addParameter("@Drv", Drv);

            string comm = "SELECT * FROM DRV_Duty WHERE Drv=@Drv AND Veh=@Veh AND FORMAT(Date , 'dd/MM/yyyy')=@Date";

            DataTable dt = conn.executeReader(comm);

            return dt;


        }



        public bool AddRow()
        {
            conn.addParameter("@Date", Date);
            conn.addParameter("@Veh", Veh);
            conn.addParameter("@VehCode", VehCode);
            conn.addParameter("@Drv", Drv);
            conn.addParameter("@Duty_Start", Duty_Start);
            conn.addParameter("@Duty_End", Duty_End);
            conn.addParameter("@Duty_Time", Duty_Time);
            conn.addParameter("@Total_Km", Total_Km);
            conn.addParameter("@MPG", MPG);
            conn.addParameter("@Co2_Kg", Co2_Kg);
            conn.addParameter("@DrvName", DrvName);


            conn.addParameter("@_drivers_Overtime_Rate", _drivers_Overtime_Rate);
            conn.addParameter("@_drivers_Standard_Rate", _drivers_Standard_Rate);
            conn.addParameter("@_drivers_Employment_Type", _drivers_Employment_Type);




            /*
            string command = "INSERT INTO DRV_Duty " +
                "(Drv, Date, Veh, VehCode, Duty_Start, Duty_End, Duty_Time, Total_Km, MPG, Co2_Kg, DrvName) " +
            "VALUES " +
            "(@Drv, @Date, @Veh, @VehCode, @Duty_Start, @Duty_End, @Duty_Time, @Total_Km, @MPG, @Co2_Kg, @DrvName )"; */


            string command = "INSERT INTO DRV_Duty " +
                "(Drv, Date, Veh, VehCode, Duty_Start, Duty_End, Duty_Time, Total_Km, MPG, Co2_Kg, DrvName, _drivers_Overtime_Rate, _drivers_Standard_Rate, _drivers_Employment_Type) " +
            "VALUES " +
            "(@Drv, @Date, @Veh, @VehCode, @Duty_Start, @Duty_End, @Duty_Time, @Total_Km, @MPG, @Co2_Kg, @DrvName, @_drivers_Overtime_Rate, @_drivers_Standard_Rate, @_drivers_Employment_Type )";

            return conn.executeNonQuery(command) > 0;
        }









        



    }
}