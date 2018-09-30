using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using WrkWebApp.Middle_Layer;
using WrkWebApp.Utility;

namespace WrkWebApp.Presentation_Layer
{
    public partial class UploadReportsScreen : System.Web.UI.Page
    {
        private static DataTable emptyFixedCost;
        private static DataTable emptyCosting;
        private static DataTable emptyDrv_Duty;
        private DataTable emptyUtilisation_Data;
        private OleDbConnection myCon;
        private OleDbCommand cmd;
        private DataTable vehiclesTb;

        private VTRN_Data vtrn_data = new VTRN_Data();
        private CostingDB cst = new CostingDB();
        private DRV_DutyDB drv = new DRV_DutyDB();
        private Utilisation_DataDB utl = new Utilisation_DataDB();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                Session["uploadCount"] = 1;
            }

            if (Page.IsPostBack && Session.Count != 0) 
            {

            }
            if(Page.IsPostBack)
            {
                lblInfo.Text = "";
            }
            
        }


        protected void btnCnfr_Click(object sender, EventArgs e)
        {
            DateTime selectedDate = cldSelectDate.SelectedDate;
           


            
            
            


           if (FlUplUpload.HasFile)
           {
                string sr = FlUplUpload.PostedFile.ContentType;
                try
               {
                    VehicleDb vh = new VehicleDb();
                     vehiclesTb = vh.SelectAllVehicles();
                    

                    if (sr == "application/vnd.ms-excel" || sr == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    {
                        string flName = Session["uploadCount"].ToString();

                        FlUplUpload.SaveAs(Server.MapPath("~/App_Data/") + flName + Path.GetExtension(FlUplUpload.FileName));
                        lblInfo.Text = "Upload successful.";



                        string path = Server.MapPath("~/App_Data/" + flName);
                         myCon = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;MAXSCANROWS=0;HDR=NO;IMEX=1\" ");
                        try
                        {

                            
                            if (Int32.Parse(Session["uploadCount"].ToString()) == 1)
                            {
                                
                                Processing_UploadCostingReport();                              
                           
                            }
                            else if (Int32.Parse(Session["uploadCount"].ToString()) == 2)
                            {
                                Processing_UploatDriversDutyReport();
                                                             
                            }
                            else if (Int32.Parse(Session["uploadCount"].ToString()) == 3)
                            {

                                Processing_UploadUtilisationReport();                              
   
                            }
                            else if (Int32.Parse(Session["uploadCount"].ToString()) == 4)
                            {

                                Processing_UploadFixedCostReport();
                                
                            }
                        }
                        catch
                        {
                            lblInfo.Text = "Error:";
                        }
                        finally
                        {
                            myCon.Close();
                        }

                        
                        /*
                        myCon.Open();
                        if ((myCon.State & ConnectionState.Open) > 0)
                        {
                            Response.Write("Connection OK!");
                            myCon.Close();
                        }
                        else
                        {
                            Response.Write("Connection no good!");
                        }   */                 

                    }
                    else
                    {
                        
                        lblInfo.Text = " Wrong extension";
                        lblInfo.Text = sr;
                    }
                }
               catch (Exception ex)
               {
                    
                    lblInfo.Text = "Error occured while attempting to upload file: " + ex.Message;
               }

                
           }
        }

        private void Processing_UploadCostingReport()
        {

            emptyCosting = cst.EmptyDataTable();
            


            try
            { 
            

               
                myCon.Open();
                DataTable dtSchema = myCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                String sheet1 = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                myCon.Close();





                //OleDbCommand cmd;
                //cmd = new OleDbCommand("SELECT * FROM " + "[Costs$]", myCon);
                cmd = new OleDbCommand("SELECT * FROM " + "["+sheet1+"]", myCon);
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                oleda.SelectCommand = cmd;
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                oleda.Fill(dt);

                GridView1.DataSource = dt;
                GridView1.DataBind();

                string dateRangeString = dt.Rows[0]["F7"].ToString();

                string date = "";

                //Obtaining date from date range cell.
                for (int z = 12; z <= 21; z++)
                {
                    if (dateRangeString[z] == '-')
                    {
                        date += '/';
                    }
                    else
                    {
                        date += dateRangeString[z];
                    }

                }


                DateTime newDate = DateTime.ParseExact(date, "yyyy/MM/dd", CultureInfo.InvariantCulture);




                bool flag = true;
                int i = 7;
                while (flag)
                {

                    if (i > dt.Rows.Count - 1 || dt.Rows[i]["F2"].ToString() == "" || dt.Rows[i]["F2"].ToString() == null)
                    {
                        Response.Write("Breaking while: null or empty string!!");
                        break;
                    }
                    else
                    {

                        string vehCode = dt.Rows[i]["F1"].ToString();

                        string shortCode = "";
                        shortCode += vehCode[4];
                        shortCode += vehCode[5];
                        shortCode += vehCode[6];

                        for (int a = 0; a <= vehiclesTb.Rows.Count - 1; a++)
                        {
                            string codeInTb = vehiclesTb.Rows[a]["Code"].ToString();
                            string shortCodeInTb = "";

                            shortCodeInTb += codeInTb[2];
                            shortCodeInTb += codeInTb[3];
                            shortCodeInTb += codeInTb[4];

                            if (shortCode.Equals(shortCodeInTb))
                            {
                                /*
                                cst.Date = newDate;
                                cst.Veh = Int32.Parse(vehiclesTb.Rows[a]["Id"].ToString());
                                cst.Total_Distance = float.Parse(dt.Rows[i]["F2"].ToString());
                                cst.Total_Fuel_Used = float.Parse(dt.Rows[i]["F3"].ToString());

                                cst.Diesel_Cost_per_l = float.Parse(dt.Rows[0]["F2"].ToString());
                                cst.Target_Consumption = float.Parse(dt.Rows[1]["F2"].ToString());
                                cst.AddBlue_Percentage_Per_L = float.Parse(dt.Rows[2]["F2"].ToString());
                                cst.Approximate_Addblue_Cost = float.Parse(dt.Rows[3]["F2"].ToString());

                                cst.Average_Consumption_Lper100Km = 0;// cst.Total_Distance * cst.Total_Fuel_Used / 100;

                                cst.Average_Consumption_MPG = 0;// (cst.Total_Distance * (float)0.621371192) / (cst.Total_Fuel_Used * (float)0.219969157);  //test and possibly change to double // there is an issue keep an eye on it
                                cst.Total_Cost_of_running = cst.Total_Fuel_Used * cst.Diesel_Cost_per_l;

                                cst.Approximate_Adblue_L = cst.Total_Fuel_Used * cst.AddBlue_Percentage_Per_L / 100; */


                                DataRow item = emptyCosting.NewRow();

                                item["Date"] = newDate;
                                item["VehCode"] = vehCode;
                                item["Veh"] = Int32.Parse(vehiclesTb.Rows[a]["Id"].ToString());
                                item["Total_Distance"] = float.Parse(dt.Rows[i]["F2"].ToString());
                                item["Total_Fuel_Used"] = float.Parse(dt.Rows[i]["F3"].ToString());
                                item["Diesel_Cost_per_l"] = float.Parse(dt.Rows[0]["F2"].ToString());
                                item["Target_Consumption"] = float.Parse(dt.Rows[1]["F2"].ToString());
                                item["AddBlue_Percentage_Per_L"] = float.Parse(dt.Rows[2]["F2"].ToString());
                                item["Approximate_Addblue_Cost"] = float.Parse(dt.Rows[3]["F2"].ToString());
                                item["Total_Cost_of_running"] = float.Parse(dt.Rows[i]["F3"].ToString()) * float.Parse(dt.Rows[0]["F2"].ToString()); //cst.Total_Cost_of_running = cst.Total_Fuel_Used * cst.Diesel_Cost_per_l;
                                item["Approximate_Adblue_L"] = float.Parse(dt.Rows[i]["F3"].ToString()) * float.Parse(dt.Rows[2]["F2"].ToString()) / 100;

                                emptyCosting.Rows.Add(item);



                                //cst.AddRowTes t();

                            }

                        }
                        i++;
                    }

                }
                System.Diagnostics.Debug.WriteLine("empty costing: " + emptyCosting.Rows.Count);
                
                Session["uploadCount"] = 2;
            }
            catch (Exception exx)
            {

            }



        }

        private void Processing_UploatDriversDutyReport()
        {

            emptyDrv_Duty = drv.EmptyDataTable();
            

            try
            {

                DriversDb driv = new DriversDb();
                DataTable drivers = driv.SelectAllDrivers();


                myCon.Open();
                DataTable dtSchema = myCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                String sheet1 = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                myCon.Close();

                //OleDbCommand cmd;
                // cmd = new OleDbCommand("SELECT * FROM " + "[1$]", myCon); // temporary
                 cmd = new OleDbCommand("SELECT * FROM " + "[" + sheet1 + "]", myCon); // temporary

                OleDbDataAdapter oleda = new OleDbDataAdapter();
                oleda.SelectCommand = cmd;
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                oleda.Fill(dt);

                GridView1.DataSource = dt;
                GridView1.DataBind();

                bool test = true;
                int p = 1;
                bool drvExist = true;
                while (test && drvExist)
                {

                    if (p > dt.Rows.Count - 1)
                    {
                        Response.Write("Breaking while: null or empty string!!");
                        break;
                    }
                    else
                    {
                        string nameFromRport = dt.Rows[p]["F1"].ToString();

                        char[] deli = new char[] { ',' };
                        string[] names = nameFromRport.Split(deli, StringSplitOptions.RemoveEmptyEntries);


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



                        names[0].TrimEnd();
                        names[0].TrimStart();



                        driv.First_Name = nm;
                        driv.Second_Name = names[0];

                        drvExist = driv.driverExists();
                        Response.Write("drvExist:    " + drvExist);

                        if (!drvExist)
                        {
                            Response.Write("[Second_NAme]:" + names[0] + "[First_Name]:" + nm);
                            break;
                        }

                    }
                    p++;
                }
                if (!drvExist)
                {
                    lblInfo.Text = "Number of drivers is inconsistent, please upload new Driver Report";
                }
                else
                {
                    int drvId = -1;
                    int vehId = -1;
                    string vehCode = "";

                    bool flag = true;
                    int i = 1;
                    while (flag)
                    {
                        ;
                        if (i > dt.Rows.Count - 1)
                        {
                            Response.Write("Breaking while: null or empty string!!");
                            break;
                        }
                        else
                        {
                            DRV_DutyDB drvDuty = new DRV_DutyDB();


                            //drvDuty.VehCode = dt.Rows[i]["F5"].ToString();
                            vehCode = dt.Rows[i]["F5"].ToString();

                            string shortCode = "";
                            shortCode += vehCode[4];
                            shortCode += vehCode[5];
                            shortCode += vehCode[6];

                            for (int a = 0; a <= vehiclesTb.Rows.Count - 1; a++)
                            {
                                string codeInTb = vehiclesTb.Rows[a]["Code"].ToString();
                                string shortCodeInTb = "";

                                shortCodeInTb += codeInTb[2];
                                shortCodeInTb += codeInTb[3];
                                shortCodeInTb += codeInTb[4];

                                if (shortCode.Equals(shortCodeInTb))
                                {
                                    drvDuty.Veh = Int32.Parse(vehiclesTb.Rows[a]["Id"].ToString());
                                    vehId = Int32.Parse(vehiclesTb.Rows[a]["Id"].ToString());

                                    drvDuty.DrvName = dt.Rows[i]["F1"].ToString();

                                    driv.SplitNamesFromReport(drvDuty.DrvName);
                                    DataTable driverId = driv.SelectIdByFirstAndSecondName();
                                    if (driverId != null)
                                    {

                                        drvId = Int32.Parse(driverId.Rows[0]["Id"].ToString());
                                        drvDuty.Drv = drvId;
                                    }

                                    driv.Id = drvDuty.Drv;
                                    DataTable driverData = driv.SelectDriverById();


                                    Response.Write("ID:   " + drvDuty.Drv);



                                    float _drivers_standard_Rate = 0;
                                    float _drivers_overtime_Rate = 0;
                                    string _drives_Employment_Type = "No type";
                                    if (driverData != null)
                                    {
                                        
                                         _drivers_standard_Rate = float.Parse(driverData.Rows[0]["Standard_Rate"].ToString());
                                         _drivers_overtime_Rate = float.Parse(driverData.Rows[0]["Overtime_Rate"].ToString());
                                         _drives_Employment_Type = driverData.Rows[0]["Type_Of_Employment"].ToString();

                                        drvDuty._drivers_Employment_Type = _drives_Employment_Type;
                                        drvDuty._drivers_Overtime_Rate = _drivers_standard_Rate;
                                        drvDuty._drivers_Standard_Rate = _drivers_overtime_Rate;

                                    }


                                    string date = dt.Rows[i]["F6"].ToString();
                                    drvDuty.Date = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture);


                                    string time = dt.Rows[i]["F7"].ToString();
                                    drvDuty.Duty_Start = (TimeSpan) DateTime.ParseExact(time, "HH:mm", CultureInfo.InvariantCulture).TimeOfDay;

                                    time = dt.Rows[i]["F8"].ToString();
                                    drvDuty.Duty_End = (TimeSpan)DateTime.ParseExact(time, "HH:mm", CultureInfo.InvariantCulture).TimeOfDay;

                                    time = dt.Rows[i]["F9"].ToString();
                                    drvDuty.Duty_Time = (TimeSpan) DateTime.ParseExact(time, "HH:mm", CultureInfo.InvariantCulture).TimeOfDay;

                                    drvDuty.Total_Km = float.Parse(dt.Rows[i]["F14"].ToString());

                                    drvDuty.MPG = 0;//float.Parse(dt.Rows[i]["F15"].ToString());

                                    drvDuty.Co2_Kg = 0;// float.Parse(dt.Rows[i]["F16"].ToString());


                                    DataRow item = emptyDrv_Duty.NewRow();
                                    item["_drivers_Employment_Type"] = _drives_Employment_Type;
                                    item["_drivers_overtime_Rate"] = _drivers_overtime_Rate;
                                    item["_drivers_standard_Rate"] = _drivers_standard_Rate;

                                    item["DrvName"] = drvDuty.DrvName;
                                    item["Id"] = -1;
                                    item["Veh"] = vehId;
                                    item["Drv"] = drvId;
                                    item["Date"] = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture); ;
                                    item["VehCode"] = vehCode;
                                    time = dt.Rows[i]["F7"].ToString();
                                    string dat = DateTime.ParseExact(time, "HH:mm", CultureInfo.InvariantCulture).ToString();
                                    item["Duty_Start"] = DateTime.ParseExact(time, "HH:mm", CultureInfo.InvariantCulture).TimeOfDay; ;

                                    time = dt.Rows[i]["F8"].ToString();
                                    item["Duty_End"] = DateTime.ParseExact(time, "HH:mm", CultureInfo.InvariantCulture).TimeOfDay; ;

                                    time = dt.Rows[i]["F9"].ToString();
                                    item["Duty_Time"] = DateTime.ParseExact(time, "HH:mm", CultureInfo.InvariantCulture).TimeOfDay; ;

                                    item["Total_Km"] = float.Parse(dt.Rows[i]["F14"].ToString()); ;
                                    item["MPG"] = 0;
                                    item["Co2_Kg"] = 0;

                                    emptyDrv_Duty.Rows.Add(item);



                                    //drvDuty.AddRow();
                                }
                            }

                            i++;
                        }

                    }

                    
                    System.Diagnostics.Debug.WriteLine("empty drv: " + emptyDrv_Duty.Rows.Count);
                    Session["uploadCount"] = 3;
                }
            }
            catch (Exception exx)
            {

            }

        }

        private void Processing_UploadUtilisationReport()
        {

            emptyUtilisation_Data = utl.EmptyDataTable();
            

            try
            {
                myCon.Open();
                DataTable dtSchema = myCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                String sheet1 = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                myCon.Close();



                //OleDbCommand cmd;
                //cmd = new OleDbCommand("SELECT * FROM " + "[1$]", myCon); // temporary
                cmd = new OleDbCommand("SELECT * FROM " + "[" + sheet1 + "]", myCon); // temporary
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                oleda.SelectCommand = cmd;
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                oleda.Fill(dt);



                GridView1.DataSource = dt;
                GridView1.DataBind();


                bool flag = true;
                int i = 1;
                while (flag)
                {

                    if (i > dt.Rows.Count - 1)
                    {
                        Response.Write("Breaking while: null or empty string!!");
                        break;
                    }
                    else
                    {


                        Utilisation_DataDB util = new Utilisation_DataDB();


                        util.Man_Veh_Code = dt.Rows[i]["F5"].ToString();

                        if (util.Man_Veh_Code.Length < 3)
                        {
                            util.Man_Veh_Code = "XXXXXXX";
                        }


                        string shortCode = "";


                        shortCode += util.Man_Veh_Code[2];
                        shortCode += util.Man_Veh_Code[3];
                        shortCode += util.Man_Veh_Code[4];

                        for (int a = 0; a <= vehiclesTb.Rows.Count - 1; a++)
                        {
                            string codeInTb = vehiclesTb.Rows[a]["Code"].ToString();
                            string shortCodeInTb = "";

                            shortCodeInTb += codeInTb[2];
                            shortCodeInTb += codeInTb[3];
                            shortCodeInTb += codeInTb[4];

                            if (shortCode.Equals(shortCodeInTb))
                            {
                                util.Veh = Int32.Parse(vehiclesTb.Rows[a]["Id"].ToString());
                                string date = dt.Rows[i]["F1"].ToString();
                                util.Man_Date_Drv = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                util.Man_Date_Drv = util.Man_Date_Drv.AddDays(2);
                                DateTime dte = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                dte = dte.AddDays(2);

                                util.Man_Number = Int32.Parse(dt.Rows[i]["F2"].ToString());

                                util.Man_Total_Packs = Int32.Parse(dt.Rows[i]["F3"].ToString());


                                util.Man_Total_Jobs = Int32.Parse(dt.Rows[i]["F4"].ToString());

                                util.Bkg_Number = Int32.Parse(dt.Rows[i]["F6"].ToString());

                                util.Bkg_Customer_Code = dt.Rows[i]["F9"].ToString();

                                util.Bkg_Cons_Packs = Int32.Parse(dt.Rows[i]["F10"].ToString());

                                util.Bkg_Cons_Weight = Int32.Parse(dt.Rows[i]["F11"].ToString());

                                util.Bkg_Cons_Price = float.Parse(dt.Rows[i]["F12"].ToString());

                                util.Man_Total_Revenue = float.Parse(dt.Rows[i]["F14"].ToString());

                                util.Bkg_Satus = Int32.Parse(dt.Rows[i]["F19"].ToString());


                                util.Cons_Delivery_Postcode = dt.Rows[i]["F17"].ToString();

                                DataRow item = emptyUtilisation_Data.NewRow();
                                item["Veh"] = Int32.Parse(vehiclesTb.Rows[a]["Id"].ToString()); ;
                                item["Man_Date_Drv"] = dte;
                                item["Man_Veh_Code"] = util.Man_Veh_Code;
                                item["Man_Total_Packs"] = Int32.Parse(dt.Rows[i]["F3"].ToString());
                                item["Man_Total_Jobs"] = Int32.Parse(dt.Rows[i]["F4"].ToString());
                                item["Man_Number"] = Int32.Parse(dt.Rows[i]["F2"].ToString());
                                item["Bkg_Number"] = Int32.Parse(dt.Rows[i]["F6"].ToString());
                                item["Bkg_Customer_Code"] = dt.Rows[i]["F9"].ToString(); ;
                                item["Bkg_Cons_Packs"] = Int32.Parse(dt.Rows[i]["F10"].ToString()); ;
                                item["Bkg_Cons_Weight"] = Int32.Parse(dt.Rows[i]["F11"].ToString()); ;
                                item["Bkg_Cons_Price"] = float.Parse(dt.Rows[i]["F12"].ToString()); ;
                                item["Man_Total_Revenue"] = float.Parse(dt.Rows[i]["F14"].ToString()); ;
                                item["Cons_Delivery_Postcode"] = dt.Rows[i]["F17"].ToString(); ;
                                item["Bkg_Status"] = Int32.Parse(dt.Rows[i]["F19"].ToString());
      


                                emptyUtilisation_Data.Rows.Add(item);

                                //util.AddRow();
                            }
                        }

                        i++;
                        Session["uploadCount"] = 4;

                    }

                }

                AddDataToDB();
                System.Diagnostics.Debug.WriteLine("empty utl: " + emptyUtilisation_Data.Rows.Count);
                Session["uploadCount"] = 4;
            }
            catch (Exception exx)
            {

            }

        }

        private void Processing_UploadFixedCostReport()
        {
            emptyFixedCost = vtrn_data.EmptyDataTable();

            try
            {

                //OleDbCommand cmd;
                cmd = new OleDbCommand("SELECT * FROM " + "[1$]", myCon); // temporary
                OleDbDataAdapter oleda = new OleDbDataAdapter();
                oleda.SelectCommand = cmd;
                DataSet ds = new DataSet();
                DataTable dt = new DataTable();
                oleda.Fill(dt);

                GridView1.DataSource = dt;
                GridView1.DataBind();


                bool flag = true;
                int i = 1;
                while (flag)
                {

                    if (i > dt.Rows.Count)
                    {
                        Response.Write("Breaking while: null or empty string!!");
                        break;
                    }
                    else
                    {


                        VTRN_Data vtrndt = new VTRN_Data();

                        vtrndt.Vtrn_Veh_Code = dt.Rows[i]["F2"].ToString();

                        string shortCode = "";
                        shortCode += vtrndt.Vtrn_Veh_Code[2];
                        shortCode += vtrndt.Vtrn_Veh_Code[3];
                        shortCode += vtrndt.Vtrn_Veh_Code[4];

                        for (int a = 0; a <= vehiclesTb.Rows.Count - 1; a++)
                        {
                            string codeInTb = vehiclesTb.Rows[a]["Code"].ToString();
                            string shortCodeInTb = "";

                            shortCodeInTb += codeInTb[2];
                            shortCodeInTb += codeInTb[3];
                            shortCodeInTb += codeInTb[4];

                            if (shortCode.Equals(shortCodeInTb))
                            {
                                vtrndt.Veh = Int32.Parse(vehiclesTb.Rows[a]["Id"].ToString());
                                vtrndt.Vtrn_Monies = float.Parse(dt.Rows[i]["F1"].ToString());
                                string date = dt.Rows[i]["F3"].ToString();
                                vtrndt.Vtrn_Date_Driver = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                vtrndt.Vtrn_Date_Driver = vtrndt.Vtrn_Date_Driver.AddDays(2);

                                DataRow item = emptyFixedCost.NewRow();
                                item["Veh"] = Int32.Parse(vehiclesTb.Rows[a]["Id"].ToString());
                                item["Vtrn_Monies"] = float.Parse(dt.Rows[i]["F1"].ToString());

                                string d = dt.Rows[i]["F3"].ToString();
                                DateTime dtme = DateTime.ParseExact(date, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                                dtme.AddDays(2);
                                item["Vtrn_Date_Driver"] = dtme;

                                //vtrndt.AddRow();
                            }

                        }

                        i++;
                    }

                }

                Session["uploadCount"] = 0;
            }
            catch (Exception exx)
            {

            }



        }

 

        private void AddDataToDB()
        {
            

            foreach(DataRow elem in emptyCosting.Rows)
            {
                

                CostingDB cst = new CostingDB();
                string date = elem["Date"].ToString();
                cst.VehCode = elem["VehCode"].ToString();
                // cst.Date = DateTime.ParseExact(date , "yyyy/MM/dd", CultureInfo.InvariantCulture);
                cst.Date = (DateTime) elem["Date"];
                cst.Veh = Int32.Parse(elem["Veh"].ToString());
                cst.Total_Distance = float.Parse(elem["Total_Distance"].ToString());
                cst.Total_Fuel_Used = float.Parse(elem["Total_Fuel_Used"].ToString());
                cst.Diesel_Cost_per_l = float.Parse(elem["Diesel_Cost_per_l"].ToString());
                cst.Target_Consumption = float.Parse(elem["Target_Consumption"].ToString());
                cst.AddBlue_Percentage_Per_L = float.Parse(elem["AddBlue_Percentage_Per_L"].ToString());
                cst.Approximate_Addblue_Cost = float.Parse(elem["Approximate_Addblue_Cost"].ToString());
                cst.Average_Consumption_Lper100Km = 0;// cst.Total_Distance * cst.Total_Fuel_Used / 100;
                cst.Average_Consumption_MPG = 0;// (cst.Total_Distance * (float)0.621371192) / (cst.Total_Fuel_Used * (float)0.219969157);  //test and possibly change to double // there is an issue keep an eye on it
                cst.Total_Cost_of_running = cst.Total_Fuel_Used * cst.Diesel_Cost_per_l;
                cst.Approximate_Adblue_L = cst.Total_Fuel_Used * cst.AddBlue_Percentage_Per_L / 100;
                cst.AddRowTest();

            }

            
            foreach(DataRow elem in emptyDrv_Duty.Rows)
            {
                DRV_DutyDB drvDuty = new DRV_DutyDB();
                drvDuty._drivers_Employment_Type = elem["_drivers_Employment_Type"].ToString();
                drvDuty._drivers_Overtime_Rate = float.Parse(elem["_drivers_Overtime_Rate"].ToString());
                drvDuty._drivers_Standard_Rate = float.Parse(elem["_drivers_Standard_Rate"].ToString());
                drvDuty.DrvName = elem["DrvName"].ToString(); 
                drvDuty.VehCode = elem["VehCode"].ToString();
                drvDuty.Veh = Int32.Parse(elem["Veh"].ToString());
                drvDuty.Drv = Int32.Parse(elem["Drv"].ToString());
                drvDuty.Date = (DateTime)elem["Date"];
                drvDuty.Duty_Start = (TimeSpan)elem["Duty_Start"];          
                drvDuty.Duty_End = (TimeSpan)elem["Duty_End"];
                drvDuty.Duty_Time = (TimeSpan)elem["Duty_Time"];
                drvDuty.Total_Km = float.Parse(elem["Total_Km"].ToString());
                drvDuty.MPG =  float.Parse(elem["MPG"].ToString());
                drvDuty.Co2_Kg = float.Parse(elem["Co2_Kg"].ToString());
                drvDuty.AddRow();

            } 


            foreach(DataRow elem in emptyUtilisation_Data.Rows)
            {
                Utilisation_DataDB util = new Utilisation_DataDB();
                util.Man_Veh_Code = elem["Man_Veh_Code"].ToString();
                util.Veh = Int32.Parse(elem["Veh"].ToString());            
                util.Man_Date_Drv = (DateTime)elem["Man_Date_Drv"];
                util.Man_Number = Int32.Parse(elem["Man_Number"].ToString());
                util.Man_Total_Packs = Int32.Parse(elem["Man_Total_Packs"].ToString());
                util.Man_Total_Jobs = Int32.Parse(elem["Man_Total_Jobs"].ToString());
                util.Bkg_Number = Int32.Parse(elem["Bkg_Number"].ToString());
                util.Bkg_Customer_Code = elem["Bkg_Customer_Code"].ToString();
                util.Bkg_Cons_Packs = Int32.Parse(elem["Bkg_Cons_Packs"].ToString());
                util.Bkg_Cons_Weight = Int32.Parse(elem["Bkg_Cons_Weight"].ToString());
                util.Bkg_Cons_Price = float.Parse(elem["Bkg_Cons_Price"].ToString());
                util.Man_Total_Revenue = float.Parse(elem["Man_Total_Revenue"].ToString());
                util.Cons_Delivery_Postcode = elem["Cons_Delivery_Postcode"].ToString();
                util.Bkg_Satus = Int32.Parse(elem["Bkg_Status"].ToString());
                util.AddRow();


            }

            /*
            foreach( DataRow elem in emptyFixedCost.Rows )
            {

                VTRN_Data vtrndt = new VTRN_Data();


                vtrndt.Veh = Int32.Parse(elem["Veh"].ToString());
                vtrndt.Vtrn_Monies = float.Parse(elem["Vtrn_Monies"].ToString());
                vtrndt.Vtrn_Date_Driver = (DateTime)elem["Vtrn_Date_Time"];

                vtrndt.AddRow();

            }*/







        }

   

        protected void GridView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        
    }
}