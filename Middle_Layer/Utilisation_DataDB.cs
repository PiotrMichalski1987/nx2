using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace WrkWebApp.Middle_Layer
{
    public class Utilisation_DataDB
    {

        private DbConnections_Db conn;


        public int Id { set; get; }

        public int Veh { set; get; }

        public int Client { set; get; }


        public DateTime Man_Date_Drv { set; get; }

        public int Man_Number { set; get; }

        public int Man_Total_Packs { set; get; }

        public int Man_Total_Jobs { set; get; }

  

        public int Bkg_Number { set; get; }

        public string Bkg_Customer_Code { set; get; }

        public int Bkg_Cons_Packs { set; get; }

        public string Cons_Delivery_Postcode { set; get; }

        public int Bkg_Cons_Weight{set;get;}

        public float Bkg_Cons_Price { set; get; }

        public float Man_Total_Revenue { set; get; }

        public int Bkg_Satus { set; get; }


        public string Man_Veh_Code { set; get; }

        public DateTime DateFrom { set; get; }

        public DateTime DateTo { set; get; }


        public Utilisation_DataDB()
        {
            conn = new DbConnections_Db();
        }

        public bool AddRow()
        {
            //conn.addParameter("@Id", Id);
            conn.addParameter("@Man_Date_Drv", Man_Date_Drv);
            conn.addParameter("@Veh", Veh);
            conn.addParameter("@Man_Number", Man_Number);
            conn.addParameter("@Man_Total_Packs", Man_Total_Packs);
            conn.addParameter("@Man_Total_Jobs", Man_Total_Jobs);
            conn.addParameter("@Man_Veh_Code", Man_Veh_Code);
            conn.addParameter("@Bkg_Number", Bkg_Number);
            conn.addParameter("@Bkg_Customer_Code", Bkg_Customer_Code);
            conn.addParameter("@Bkg_Cons_Packs", Bkg_Cons_Packs);
            conn.addParameter("@Bkg_Cons_Weight", Bkg_Cons_Weight);
            conn.addParameter("@Bkg_Cons_Price", Bkg_Cons_Price);
            conn.addParameter("@Man_Total_Revenue", Man_Total_Revenue);
            conn.addParameter("@Client", Client);
            conn.addParameter("@Cons_Delivery_Postcode", Cons_Delivery_Postcode);
            conn.addParameter("@Bkg_Status", Bkg_Satus);

            /*
            string command = "INSERT INTO Utilisation_Data " +
               "(Man_Date_Drv, Veh, Man_Number, Man_Total_Packs, Man_Total_Jobs, Man_Veh_Code, Bkg_Number, Bkg_Customer_Code, Bkg_Cons_Packs, Bkg_Cons_Weight, Bkg_Cons_Price, Man_Total_Revenue  ) " +
           "VALUES " +
           "(@Man_Date_Drv, @Veh, @Man_Number, @Man_Total_Packs, @Man_Total_Jobs, @Man_Veh_Code, @Bkg_Number, @Bkg_Customer_Code, @Bkg_Cons_Packs, @Bkg_Cons_Weight, @Bkg_Cons_Price, @Man_Total_Revenue   )"; */


            string command = "INSERT INTO Utilisation_Data " +
               "(Man_Date_Drv, Veh, Man_Number, Man_Total_Packs, Man_Total_Jobs, Man_Veh_Code, Bkg_Number, Bkg_Customer_Code, Bkg_Cons_Packs, Bkg_Cons_Weight, Bkg_Cons_Price, Man_Total_Revenue, Cons_Delivery_Postcode, Bkg_Status  ) " +
           "VALUES " +
           "(@Man_Date_Drv, @Veh, @Man_Number, @Man_Total_Packs, @Man_Total_Jobs, @Man_Veh_Code, @Bkg_Number, @Bkg_Customer_Code, @Bkg_Cons_Packs, @Bkg_Cons_Weight, @Bkg_Cons_Price, @Man_Total_Revenue, @Cons_Delivery_Postcode, @Bkg_Status  )";

            return conn.executeNonQuery(command) > 0;
        }

        public DataTable SelectUsingVehAndDate()
        {
            conn.addParameter("@Veh", Veh);
            conn.addParameter("@Man_Date_Drv", Man_Date_Drv.ToString("dd/MM/yyyy"));

            string comm = "SELECT * FROM Utilisation_Data WHERE Veh=@Veh AND FORMAT(Man_Date_Drv , 'dd/MM/yyyy')=@Man_Date_Drv";
            DataTable dt = conn.executeReader(comm);

            return dt;


        }


        public DataTable SelectDateRange()
        {
            conn.addParameter("@DateFrom", DateFrom.ToString("dd/MM/yyyy"));
            conn.addParameter("@DateTo", DateTo.ToString("dd/MM/yyyy"));

            string comm = "SELECT DISTINCT Man_Date_Drv FROM Utilisation_Data WHERE FORMAT(Man_Date_Drv, 'dd/MM/yyyy') BETWEEN @DateFrom AND @DateTo ";
            //string comm = "SELECT DISTINCT CONVERT(VARCHAR, Man_Date_Drv, 103) as Man_Date_Drv FROM Utilisation_Data ";
            DataTable dt = conn.executeReader(comm);

            return dt;

        }

        public DataTable SelectAllRowsBetweenDates()
        {

            conn.addParameter("@DateFrom", DateFrom.ToString("dd/MM/yyyy"));
            conn.addParameter("@DateTo", DateTo.ToString("dd/MM/yyyy"));


            

            string command = "SELECT * FROM Utilisation_Data WHERE FORMAT(Man_Date_Drv , 'dd/MM/yyyy') BETWEEN @DateFrom AND @DateTo";

            DataTable dt = conn.executeReader(command);

            return dt;


        }

        public DataTable SelectAllWithManifestNumber()
        {
            conn.addParameter("@Man_Number", Man_Number);
            //conn.addParameter("@Bkg_Number", Bkg_Number);


            string cmd = "SELECT * FROM Utilisation_Data WHERE Man_Number=@Man_Number ";


            DataTable dt = conn.executeReader(cmd);

            return dt;




        }

        



        public DataTable SelectUniqueManifestsBetweenDates2()
        {
            conn.addParameter("@DateFrom", DateFrom.ToString("dd/MM/yyyy"));
            conn.addParameter("@DateTo", DateTo.ToString("dd/MM/yyyy"));

            DataTable dt = new DataTable();

            string command = "SELECT DISTINCT * FROM Utilisation_Data WHERE FORMAT(Man_Date_Drv , 'dd/MM/yyyy') BETWEEN @DateFrom AND @DateTo";

            dt = conn.executeReader(command);

            return dt;


        }

        public DataTable SelectUniqueManifestsBetweenDates3()
        {
            conn.addParameter("@DateFrom", DateFrom.ToString("dd/MM/yyyy"));
            conn.addParameter("@DateTo", DateTo.ToString("dd/MM/yyyy"));

            DataTable dt = new DataTable();

            string command = "SELECT DISTINCT Man_Number, Man_Veh_Code FROM Utilisation_Data WHERE FORMAT(Man_Date_Drv , 'dd/MM/yyyy') BETWEEN @DateFrom AND @DateTo";

            dt = conn.executeReader(command);

            return dt;


        }

        public DataTable SelectMan_Total_Pack()
        {
            
            conn.addParameter("@Man_Number", Man_Number);

            DataTable dt = new DataTable();

            string command = "SELECT Man_Total_Packs FROM Utilisation_Data WHERE Man_Number=@Man_Number";
            dt = conn.executeReader(command);

            return dt;

        }

        public DataTable SelectUniqueJobsBetweenDatesAndByManifest()
        {
            conn.addParameter("@DateFrom", DateFrom.ToString("dd/MM/yyyy"));
            conn.addParameter("@DateTo", DateTo.ToString("dd/MM/yyyy"));
            conn.addParameter("@Man_Number", Man_Number);

            DataTable dt = new DataTable();

            string command = "SELECT DISTINCT Bkg_Number FROM Utilisation_Data WHERE FORMAT(Man_Date_Drv , 'dd/MM/yyyy') BETWEEN @DateFrom AND @DateTo AND Man_Number=@Man_Number ";

            dt = conn.executeReader(command);

            return dt;


        }

        public DataTable SelectUniqueRowsBetweenDatesAndByJob()
        {
            conn.addParameter("@DateFrom", DateFrom.ToString("dd/MM/yyyy"));
            conn.addParameter("@DateTo", DateTo.ToString("dd/MM/yyyy"));
            conn.addParameter("@Bkg_Number", Bkg_Number);

            DataTable dt = new DataTable();

            string command = "SELECT * FROM Utilisation_Data WHERE FORMAT(Man_Date_Drv , 'dd/MM/yyyy') BETWEEN @DateFrom AND @DateTo AND Bkg_Number=@Bkg_Number ";

            dt = conn.executeReader(command);

            return dt;


        }

        public DataTable SelectUniqueRowBetweenDatesByManifest()
        {
            conn.addParameter("@DateFrom", DateFrom.ToString("dd/MM/yyyy"));
            conn.addParameter("@DateTo", DateTo.ToString("dd/MM/yyyy"));
            conn.addParameter("@Man_Number", Man_Number);

            DataTable dt = new DataTable();

            string command = "SELECT * FROM Utilisation_Data WHERE FORMAT(Man_Date_Drv , 'dd/MM/yyyy') BETWEEN @DateFrom AND @DateTo AND Man_Number=@Man_Number ";

            dt = conn.executeReader(command);

            return dt;


        }


        public DataTable EmptyDataTable()
        {
            Id = -1;
            conn.addParameter("@Id", Id);

            String comm = "SELECT * FROM Utilisation_Data WHERE Id=@Id";
            DataTable dt = new DataTable();

            dt = conn.executeReader(comm);

            return dt;
        }

























    }
}