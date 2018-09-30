using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;

namespace WrkWebApp.Middle_Layer
{
    public class DriversDb
    {
        private DbConnections_Db conn;

        public int Id { set; get; }

        public string First_Name { set; get; }

        public string Second_Name { set; get; }

        public string Type_Of_Employment { set; get; }

        public float Overtime_Rate { set; get; }

        public float Standard_Rate { set; get; }

        public string drv_card { set; get; }


        public DriversDb()
        {
            conn = new DbConnections_Db();
        }

        public bool AddRow()
        {
            conn.addParameter("@drv_card", drv_card);
            conn.addParameter("@First_Name", First_Name);
            conn.addParameter("@Second_Name", Second_Name);
            conn.addParameter("@Type_Of_Employment", Type_Of_Employment);
            conn.addParameter("@Overtime_Rate", Overtime_Rate);
            conn.addParameter("@Standard_Rate", Standard_Rate);

            string command = "INSERT INTO Drivers " +
               "(drv_card, First_Name, Second_Name, Type_Of_Employment, Overtime_Rate, Standard_Rate  ) " +
           "VALUES " +
           "(@drv_card, @First_Name, @Second_Name, @Type_Of_Employment, @Overtime_Rate, @Standard_Rate  )";

            return conn.executeNonQuery(command)>0;

        }

        public DataTable SelectAllDrivers()
        {
            string command = "SELECT * FROM Drivers";

            DataTable tb = conn.executeReader(command);

            return tb;

        }


        public DataTable SelectDistinctTypes()
        {
            string command = "SELECT DISTINCT Type_Of_Employment FROM Drivers";

            DataTable tb = conn.executeReader(command);

            return tb;

        }

        public bool UpdateOverTimeRate()
        {
            conn.addParameter("Id", Id);
            conn.addParameter("@Overtime_Rate", Overtime_Rate);

            string command = "UPDATE Drivers SET Overtime_Rate=@Overtime_Rate  WHERE Id=@Id";
            return conn.executeNonQuery(command) > 0;
        }

        public bool UpdateOverTimeRateBasedOnType()
        {
    
            conn.addParameter("Type_Of_Employment", Type_Of_Employment);
            conn.addParameter("@Overtime_Rate", Overtime_Rate);

            string command = "UPDATE Drivers SET Overtime_Rate=@Overtime_Rate  WHERE Type_Of_Employment=@Type_Of_Employment";
            return conn.executeNonQuery(command) > 0;
        }

        public bool UpdateStandardRateBasedOnType()
        {
           
            conn.addParameter("Type_Of_Employment", Type_Of_Employment);
            conn.addParameter("@Standard_Rate", Standard_Rate);

            string command = "UPDATE Drivers SET Standard_Rate=@Standard_Rate  WHERE Type_Of_Employment=@Type_Of_Employment";
            return conn.executeNonQuery(command) > 0;
        }

        public DataTable SelectDriverById()
        {
            conn.addParameter("Id", Id);

            string command = "SELECT * FROM Drivers WHERE Id=@Id";
            DataTable dt = conn.executeReader(command);
            return dt;


        }


        public bool UpdateStandardRate()
        {
            conn.addParameter("Id", Id);

            conn.addParameter("@Standard_Rate", Standard_Rate);
            string command = "UPDATE Drivers SET Standard_Rate=@Standard_Rate  WHERE Id=@Id";
            return conn.executeNonQuery(command) > 0;
        }




        public DataTable SelectIdByFirstAndSecondName()
        {
            conn.addParameter("@First_Name", First_Name);
            conn.addParameter("@Second_Name", Second_Name);


            string command = "SELECT Id From Drivers WHERE Second_Name=@Second_Name AND First_Name=@First_Name";

            DataTable tb = conn.executeReader(command);

            return tb;

        }

        public void SplitNamesFromReport(string nameFromReport)
        {
            

            char[] deli = new char[] { ',' };
            string[] names = nameFromReport.Split(deli, StringSplitOptions.RemoveEmptyEntries);


            bool done = false;
            string nm = "";
            foreach (char elem in names[1])
            {
                if (!done)
                {
                    if (char.IsLetter(elem))
                    {
                        nm += elem;
                        done = true;
                    }
                }
                else
                {
                    nm += elem;
                }
            }

            First_Name = nm;
            Second_Name = names[0];
        }



        public bool driverExists()
        {



            conn.addParameter("@First_Name", First_Name);
            conn.addParameter("@Second_Name", Second_Name);

            String command = "SELECT COUNT(Second_Name)FROM Drivers WHERE First_Name=@First_Name AND Second_Name=@Second_Name";

            int result = conn.executeScalar(command);

            return result > 0 ;
        }



    }
}