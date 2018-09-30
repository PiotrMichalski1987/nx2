using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WrkWebApp.Middle_Layer
{
    public class ClientDateDb
    {


        public int Id { set; get; }
        public int Client { set; get; }

        public float ProfitLoss { set; get; }

        public float TotalCost { set; get; }

        public DateTime Date { set; get; }


       

        private DbConnections_Db conn;


        public bool DateExists()
        {
            conn.addParameter("@Date", Date);
            String command = "SELECT COUNT(Date)FROM ClientData WHERE Date=@Date";

            int result = conn.executeScalar(command);

            return result > 0;
        }



    }
}