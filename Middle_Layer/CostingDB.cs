using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace WrkWebApp.Middle_Layer
{
    public class CostingDB
    {

        
        private DbConnections_Db conn;
        public int Id { set; get; }
        public int Veh { set; get; }

        public float Average_Consumption_Lper100Km { set; get; }
        public float Average_Consumption_MPG { set; get; }
        public float Total_Distance { set; get; }
        public float Total_Fuel_Used { get; set; }
        public float Total_Cost_of_running { get; set; }

        public float Diesel_Cost_per_l { get; set; }

        public float Target_Consumption { set; get; }
        public float AddBlue_Percentage_Per_L { get; set; }

        public float AddBlue_Cost_Per_L { set; get; }

        public float Approximate_Adblue_L { set; get; }

        public float Approximate_Addblue_Cost { set; get; }

        public string VehCode { set; get; }

        public DateTime Date { set; get; }

        public DateTime DateFrom { set; get; }

        public DateTime DateTo { set; get; }

        public float Current_Consumption { set; get; }

        public CostingDB()
        {
             conn = new DbConnections_Db();
        }

        public DataTable EmptyDataTable()
        {
            Id = -1;
            conn.addParameter("@Id", Id);

            String comm = "SELECT * FROM Costing WHERE Id=@Id";
            DataTable dt = new DataTable();

            dt = conn.executeReader(comm);

            return dt;
        }


 

        public bool AddRowTest()
        {
            conn.addParameter("@Date", Date);
            conn.addParameter("@Veh", Veh);
            conn.addParameter("@VehCode", VehCode);
            conn.addParameter("@Total_Distance", Total_Distance);

            conn.addParameter("@Total_Fuel_Used" , Total_Fuel_Used);
            conn.addParameter("@Diesel_Cost_per_l", Diesel_Cost_per_l);
            conn.addParameter("@Target_Consumption", Target_Consumption);
            conn.addParameter("@AddBlue_Percentage_Per_L", AddBlue_Percentage_Per_L);
            conn.addParameter("@Approximate_Addblue_Cost", Approximate_Addblue_Cost);
            conn.addParameter("@Approximate_Adblue_L", Approximate_Adblue_L);
            conn.addParameter("@Average_Consumption_Lper100Km", Average_Consumption_Lper100Km);
            conn.addParameter("@Average_Consumption_MPG", Average_Consumption_MPG);
            


            

            string command = "INSERT INTO Costing " +
                "(Date,Veh, VehCode, Total_Distance, Total_Fuel_Used, Diesel_Cost_per_l, Target_Consumption, AddBlue_Percentage_Per_L, Approximate_Addblue_Cost, Approximate_Adblue_L, Average_Consumption_Lper100Km, Average_Consumption_MPG  ) " +
            "VALUES " +
            "(@Date, @Veh, @VehCode, @Total_Distance, @Total_Fuel_Used, @Diesel_Cost_per_l, @Target_Consumption, @AddBlue_Percentage_Per_L, @Approximate_Addblue_Cost, @Approximate_Adblue_L, @Average_Consumption_Lper100Km, @Average_Consumption_MPG  )";


            return conn.executeNonQuery(command) > 0;

        }

        public DataTable SelectAllRowsByGivenDate()
        {

            conn.addParameter("@DateFrom",  DateFrom.ToString("dd/MM/yyyy"));
            conn.addParameter("@DateTo", DateTo.ToString("dd/MM/yyyy"));


            DataTable dt = new DataTable();

            string command = "SELECT * FROM Costing WHERE FORMAT(Date , 'dd/MM/yyyy') BETWEEN @DateFrom AND @DateTo";
            
            dt = conn.executeReader(command);

            return dt;


        }


        public DataTable SelectRowByVehAndDate()
        {
            conn.addParameter("@Date", Date.ToString("dd/MM/yyyy"));
            conn.addParameter("@Veh", Veh);



            DataTable dt = new DataTable();

            string command = "SELECT * FROM Costing WHERE FORMAT(Date , 'dd/MM/yyyy') = @Date AND Veh=@Veh";

            dt = conn.executeReader(command);

            return dt;




        }














    }
}