using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WrkWebApp.Diagram_DB;
using WrkWebApp.Middle_Layer;
using System.IO.Compression;

namespace WrkWebApp.Presentation_Layer
{
    public partial class ProduceReportsScreen : System.Web.UI.Page
    {
        private OleDbConnection myCon;
        private OleDbConnection myCon2;
        private OleDbCommand cmd;
        private static string reportOutput;
        private static string path;
        private static string from_to;
        private static string name;

        public object Log { get; private set; }

        //private static FileInfo fl;

        public void SetUpExcelConnection()
        {
             reportOutput = "reportOutput.xls";
             path = Server.MapPath("~/App_Data/");

            

            if (File.Exists(path + reportOutput))
            {
                File.Delete(path + reportOutput);
            }

            
            File.Copy(path + "report.xls", path + reportOutput);
            //fl= fl = new FileInfo(path + reportOutput);
            
           

            this.myCon = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + reportOutput + ";Extended Properties=\"Excel 8.0;MAXSCANROWS=0;HDR=NO\" ");
            this.myCon2 = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + reportOutput + ";Extended Properties=\"Excel 8.0;MAXSCANROWS=0;HDR=YES\" ");
        }



        protected void Page_Load(object sender, EventArgs e)
        {
            if (!Page.IsPostBack)
            {
                ddlSelectReport.Items.Add(ListItem.FromString("Profit and loss per vehicle type"));//0
                ddlSelectReport.Items.Add(ListItem.FromString("Profit and loss per vehicle code"));//1
                ddlSelectReport.Items.Add(ListItem.FromString("Customer profit and loss by manifest"));//2
                ddlSelectReport.Items.Add(ListItem.FromString("Profit and loss by Client"));//3
                ddlSelectReport.Items.Add(ListItem.FromString("Profit and loss by postcode")); //4
                ddlSelectReport.Items.Add(ListItem.FromString("Drivers Report")); //5
                ddlSelectReport.Items.Add(ListItem.FromString("All Reports")); //6
            }

        }

        protected void btnConfrm_Click(object sender, EventArgs e)
        {
            



            string zipName = "";
           // FileInfo fl = new FileInfo(path + reportOutput);
          //  FileInfo fl2 = new FileInfo()

            if (ddlSelectReport.SelectedIndex == -1)
            {
                lblInfo.Text = "Please select report";
            }
            else if (cldFrom.SelectedDate.Date == DateTime.MinValue || cldTo.SelectedDate.Date == DateTime.MinValue)
            {
                lblInfo.Text = "Please select date range";
            }
            else
            {
                lblInfo.Text = "";
                string from = cldFrom.SelectedDate.ToString("dd/MM/yyyy");
                string to = cldTo.SelectedDate.ToString("dd/MM/yyyy");


                this.SetUpExcelConnection();
                //string name = "";
                //string from_to = "";
                

                from_to = "From_"  + from + "_to_"  + to   ;
                from_to = from_to.Replace('/', '-');
                //from_to = "";

                if (ddlSelectReport.SelectedIndex == 1)
                {
                    //this.Processing_ProfitAndLossByVehicleCode(from, to);
                    //this.Processing_BasedOnManReport(from, to, 1);
                    Processing_New(from, to, 1);
                    name = "Code_";
                    name += from_to;
                    //ZipFile();

                    //CreateReportFile(path, name);

                    

                }
                else if (ddlSelectReport.SelectedIndex == 0)
                {
                    //this.Processing_ProfitAndLossByVehicleType(from, to);
                    Processing_New(from, to, 0);
                    name = "Type_";
                    name += from_to ;
                    zipName = name + ".zip";

                    if (!File.Exists(path + name + ".xls") && !File.Exists(path + zipName))
                    {
                        File.Copy(path + "reportOutput.xls", path + name + ".xls");
                        File.Copy(path + "Reports_NxGroup.zip", path + zipName);
                    }

                    





                    

                    


                    //string zipName = from_to + ".zip";
                    //File.Copy(path + "Reports_NxGroup.zip", path + zipName);
                    //fl = fl = new FileInfo(path + zipName);
                    //fl = CreateReportFile(path, name);


                    // File.Copy(path + "reportOutput.xls", path + name + ".xls");
                    // zipName = name + ".zip";                 
                    // File.Copy(path + "Reports_NxGroup.zip", path + zipName);


                    // PrepareZipFile(path, from_to);
                    //ZipFile(path, zipName, name);

                    // fl = Download_Dwn(path, name);

                }
                else if (ddlSelectReport.SelectedIndex == 2)
                {
                    //this.Processing_CustomerProfitAndLossByManifest_VER2(from,to);
                    //this.Processing_BasedOnManReport(from, to, 2);
                    Processing_New(from, to, 2);
                    name = "Man_";
                    name += from_to + ".xls";
                    CreateReportFile(path, name);
                    Download_Dwn(path, name);
                }
                else if(ddlSelectReport.SelectedIndex == 3)
                {
                    //this.Processing_ProfitLossByClient(from, to);
                    Processing_New(from, to, 3);
                    //this.Processing_BasedOnManReport(from, to, 3);
                    name = "Client_";
                    name += from_to + ".xls";
                    CreateReportFile(path, name);
                    Download_Dwn(path, name);

                }
                else if (ddlSelectReport.SelectedIndex == 4)
                {
                   

                    Processing_New(from, to, 4);
                    //this.Processing_BasedOnManReport(from, to, 4);
                    // this.Processing_RunsToPostcodes(from, to);
                    name = "Postcode_";
                    name += from_to + ".xls";
                    CreateReportFile(path, name);
                    Download_Dwn(path, name);

                }
                else if(ddlSelectReport.SelectedIndex == 6)
                {
                    //this is here because I do not know why some process holds the file

                    FileInfo fileName = new FileInfo(path + "reportOutput.xls");

                    zipName = "NxTransportReports_" + from_to + ".zip";
                    File.Copy(path + "Reports_NxGroup.zip", path + zipName);
                    



                    /*
                    Processing_New(from, to, 0);                  
                    name = "Type_";
                    name += from_to;     
                    


                    File.Copy(path + "reportOutput.xls", path + name + ".xls");
                    this.ZipFile(path, zipName, name + ".xls");
                    File.Delete(path + name + ".xls");
                    name = "";

                    if(FileIsLocked(path + "reportOutput.xls"))
                    {

                        System.Diagnostics.Debug.WriteLine("LOCKED");
                      


                    }
                   


                    Processing_New(from, to, 1);
                    name = "Code_" + from_to;
                    File.Copy(path + "reportOutput.xls", path + name + ".xls");
                    this.ZipFile(path, zipName, name + ".xls");
                    File.Delete(path + name + ".xls");
                    name = "";
                    if (FileIsLocked(path + "reportOutput.xls"))
                    {

                        System.Diagnostics.Debug.WriteLine("LOCKED");



                    }

                    */

                    Processing_New(from, to, 2);
                    name = "Man_" + from_to ;
                    File.Copy(path + "reportOutput.xls", path + name + ".xls");
                    this.ZipFile(path, zipName, name + ".xls");
                    File.Delete(path + name + ".xls");
                    name = "";
                    if (FileIsLocked(path + "reportOutput.xls"))
                    {

                        System.Diagnostics.Debug.WriteLine("LOCKED");



                    }


                    /*
                    Processing_New(from, to, 3);
                    name = "Client_" + from_to ;
                    File.Copy(path + "reportOutput.xls", path + name + ".xls");
                    this.ZipFile(path, zipName, name + ".xls");
                    File.Delete(path + name + ".xls");
                    name = "";
                    if (FileIsLocked(path + "reportOutput.xls"))
                    {

                        System.Diagnostics.Debug.WriteLine("LOCKED");



                    }


                    Processing_New(from, to, 4);
                    name = "Postcode_" + from_to;
                    File.Copy(path + "reportOutput.xls", path + name + ".xls");
                    this.ZipFile(path, zipName, name + ".xls");
                    File.Delete(path + name + ".xls");
                    name = "";
                    if (FileIsLocked(path + "reportOutput.xls"))
                    {

                        System.Diagnostics.Debug.WriteLine("LOCKED");



                    }


                    */

                }



                FileInfo fl = new FileInfo(path + name + ".xls");
                FileInfo fl2 = new FileInfo(path + zipName);

                Response.AddHeader("Content-Disposition", "attachment; filename=" + fl2.Name);
                Response.TransmitFile(fl2.FullName);
                //Response.Close();

                Response.Flush();

                if(File.Exists(path+zipName))
                {

                    File.Delete(path + zipName);


                }



                if(File.Exists(path + name + ".xls"))
                {
                    File.Delete(path + name + ".xls");

                }

                if(File.Exists(path + zipName))
                {
                    File.Delete(path + zipName);
                }

              


                
                


                //string uri = path;
                //string stringWebResorce = uri + reportOutput;


                //WebClient webClient = new WebClient();

                //webClient.DownloadFile(fl.FullName);

                //Response.Clear();
                myCon.Close();
                myCon2.Close();
                myCon.Dispose();
                myCon2.Dispose();
                

                //fl.Refresh();


                //this  is wrong, tmp fix
              

            }
        }

        public void ZipFile(string relativePath, string zipName, string reportName)
        {
            

            //string zipPath = path + "/Reports_NxGroup.zip";
            

            using (FileStream zipFileToOpen = new FileStream(relativePath + zipName, FileMode.Open))           
            using (ZipArchive archive = new ZipArchive(zipFileToOpen, ZipArchiveMode.Update))
            {
                archive.CreateEntryFromFile(relativePath + reportName, reportName);           
            }


        }

        bool WaitForFile(string fullname)
        {
            int numTries = 0;

            while (true)
            {
                ++numTries;
                try
                {
                    using (FileStream fs = new FileStream(fullname, FileMode.Open, FileAccess.ReadWrite, FileShare.None, 100))
                    {

                        fs.ReadByte();
                        break;

                    }
                }
                catch (Exception ex)
                {

                    System.Diagnostics.Debug.WriteLine("Cannot access the file");


                    if (numTries > 100)
                    {
                        System.Diagnostics.Debug.WriteLine(" Waiting for file, giving up after x tries");

                        return false;


                    }

                    System.Threading.Thread.Sleep(500);
                }
            }
            return true;
        }



        public bool FileIsLocked(string filepath)
        {
            bool isLocked = false;
            System.IO.FileStream fs;

            try
            {
                fs = System.IO.File.Open(filepath, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.Read, System.IO.FileShare.None);
                fs.Close();
            }
            catch (System.IO.IOException ex)
            {
                isLocked = true;
            }
            return isLocked;

        }

        public FileInfo CreateReportFile(string path , string name)
        {

            File.Copy(path + "reportOutput.xls", path + name);
            FileInfo fl = fl = new FileInfo(path + name);
            return fl;

        }

        public void PrepareZipFile(string path, string name)
        {


            string zipName = name + ".zip";
            File.Copy(path + "Reports_NxGroup.zip", path + zipName);
        }

        public FileInfo Download_Dwn(string path, string name)
        {

            //fl = fl = new FileInfo(path + name);

            //this  is wrong, tmp fix
              FileInfo fl = fl = new FileInfo(path + name);
            return fl;

            

        }

        

        public void Processing_ProfitAndLossByVehicleType(string from, string to)
        {

            cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + from.ToString() + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close();

            cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close();


            cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9 ) VALUES (" + "'" + "VehicleCode" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "FixedCost" + "'" + "," + "'" + "AdblueCost" + "'" + "," + "'" + "TotalDistance" + "'" + "," + "'" + "FuelUsed" + "'" + "," + "'" + "FuelCost" + "'" + "," + "'" + "DriveCosts" + "'" + "," + "'" + "ProfitLoss" + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close();

            float approximateAdblue = 0;
            float profit_loss = 0;
            float driverCost = 0;
            float addblueTotal = 0;
            float totalCostOfRunning = 0;
            float revenue = 0;
            float fixedCost = 0;
            float totalDistance = 0;
            float totalFuelUsed = 0;

            string vehType = "";

            float totalDistance18T = 0;
            float totalDistance7T = 0;
            float totalDistanceUnitT = 0;

            float totalFuelUsed18T = 0;
            float totalFuelUsed7T = 0;
            float totalFuelUsedUnit = 0;

            float profitLoss18T = 0;
            float profitLoss7T = 0;
            float profitLossUnitT = 0;

            float driverCost18T = 0;
            float driverCost7T = 0;
            float driverCostUnit = 0;

            float addblueTotal18T = 0;
            float addblueTotal7T = 0;
            float addblueTotalUnit = 0;

            float totalCostOfRunning18T = 0;
            float totalCostOfRunning7T = 0;
            float totalCostOfRunningUnit = 0;

            float revenue18T = 0;
            float revenue7T = 0;
            float revenueUnit = 0;

            float fixedCost18T = 0;
            float fixedCost7T = 0;
            float fixedCostUnit = 0;


            CostingDB cst = new CostingDB();
            cst.DateFrom = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            cst.DateTo = DateTime.ParseExact(to, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            VehicleDb vehDB = new VehicleDb();
            DataTable vehiclesDT = vehDB.SelectAllVehicles();

            if (vehiclesDT.Rows.Count > 0 && vehiclesDT != null)
            {
                revenue = 0;
                for (int u = 0; u <= vehiclesDT.Rows.Count - 1; u++)
                {

                    string vehCode = vehiclesDT.Rows[u]["Code"].ToString();

                    float dieselCostPerL = 0;
                    float addBluePercentage = 0;
                    float targetConsumption = 0;
                    float approximateAddblueL = 0;
                    float addBlueCost = 0;
                    int veh = -1;

                    DateTime _date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    int id = Int32.Parse(vehiclesDT.Rows[u]["Id"].ToString());

                    CostingDB costingRow = new CostingDB();
                    veh = id;
                    costingRow.Veh = id;
                    costingRow.Date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                    DataTable costingRowDt = costingRow.SelectRowByVehAndDate();


                    Middle_Layer.VTRN_Data vtrnData = new Middle_Layer.VTRN_Data();
                    vtrnData.Veh = veh;
                    vtrnData.Vtrn_Date_Driver = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);


                    Utilisation_DataDB util = new Utilisation_DataDB();
                    util.Man_Date_Drv = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    util.Veh = veh;


                    if (costingRowDt != null && costingRowDt.Rows.Count > 0)
                    {

                        totalDistance = float.Parse(costingRowDt.Rows[0]["Total_Distance"].ToString());
                        totalFuelUsed = float.Parse(costingRowDt.Rows[0]["Total_Fuel_Used"].ToString());
                        dieselCostPerL = float.Parse(costingRowDt.Rows[0]["diesel_Cost_per_l"].ToString());
                        addBluePercentage = float.Parse(costingRowDt.Rows[0]["AddBlue_Percentage_Per_L"].ToString());
                        addBlueCost = float.Parse(costingRowDt.Rows[0]["Approximate_Addblue_Cost"].ToString());
                        targetConsumption = float.Parse(costingRowDt.Rows[0]["Target_Consumption"].ToString());
                        approximateAddblueL = float.Parse(costingRowDt.Rows[0]["Approximate_Adblue_L"].ToString());
                        veh = int.Parse(costingRowDt.Rows[0]["Veh"].ToString());


                        DRV_DutyDB drvDt = new DRV_DutyDB();
                        drvDt.Date = _date;
                        drvDt.VehCode = vehCode;
                        drvDt.Veh = veh;

                        DataTable drivDuty = drvDt.SelectRowsUsingDateAndVeh();


                        if (drivDuty != null && drivDuty.Rows.Count > 0)
                        {
                            int min = 0;
                            int hr = 0;
                            int minPlusHr = 0;
                            float totalHours = 0;
                            float overtimeHours = 0;
                            float standardHours = 0;
                            float overtimeEarnings = 0;
                            float standardEarnings = 0;

                            for (int a = 0; a <= drivDuty.Rows.Count - 1; a++)
                            {
                                DateTime time = DateTime.ParseExact(drivDuty.Rows[a]["Duty_Time"].ToString(), "HH:mm:ss", CultureInfo.InvariantCulture);
                                min += time.Minute;
                                hr += time.Hour;

                                minPlusHr = hr * 60 + min;
                                totalHours = (float)minPlusHr / 60;


                                if (overtimeHours > 0)
                                {
                                    overtimeEarnings = overtimeHours * float.Parse(drivDuty.Rows[a]["_drivers_Overtime_Rate"].ToString());
                                    standardHours = totalHours - overtimeHours;
                                    standardEarnings = standardHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings + overtimeEarnings;
                                }
                                else
                                {
                                    standardEarnings = totalHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings;
                                }
                            }
                        }
                    }

                    DataTable vtrn = vtrnData.SelectUsingVehAndDate();
                    if (vtrn != null && vtrn.Rows.Count > 0)
                    {
                        fixedCost = float.Parse(vtrn.Rows[0]["Vtrn_Monies"].ToString());
                    }
                    else
                    {
                        fixedCost = 0;
                    }

                    DataTable ut = util.SelectUsingVehAndDate();

                    if (ut != null && ut.Rows.Count > 0)
                    {
                        revenue = float.Parse(ut.Rows[0]["Man_Total_Revenue"].ToString());

                    }
                    else
                    {
                        revenue = 0;
                    }


                    addblueTotal = approximateAddblueL * addBlueCost;
                    totalCostOfRunning = totalFuelUsed * dieselCostPerL;
                    profit_loss = revenue - (fixedCost + addblueTotal + totalFuelUsed + totalCostOfRunning + driverCost);

                    VehicleDb vehDb = new VehicleDb();
                    vehDb.Id = veh;

                    DataTable vehDT = vehDb.SelectVehicleById();
                    if (vehDT != null)
                    {
                        vehType = vehDT.Rows[0]["Type"].ToString();

                    }
                    else
                    {
                        System.Diagnostics.Debug.Write("In here");
                    }

                    if (vehType.Contains("18T"))
                    {
                        revenue18T += revenue;
                        fixedCost18T += fixedCost;
                        addblueTotal18T += addblueTotal;
                        totalDistance18T += totalDistance;
                        totalFuelUsed18T += totalFuelUsed;
                        totalCostOfRunning18T += totalCostOfRunning;
                        driverCost18T += driverCost;
                        profitLoss18T += profit_loss;


                    }
                    else if (vehType.Contains("7.5T"))
                    {
                        revenue7T += revenue;
                        fixedCost7T += fixedCost;
                        addblueTotal7T += addblueTotal;
                        totalDistance7T += totalDistance;
                        totalFuelUsed7T += totalFuelUsed;
                        totalCostOfRunning7T += totalCostOfRunning;
                        driverCost7T += driverCost;
                        profitLoss7T += profit_loss;
                    }
                    else if (vehType.Contains("Unit"))
                    {
                        revenueUnit += revenue;
                        fixedCostUnit += fixedCost;
                        addblueTotalUnit += addblueTotal;
                        totalDistanceUnitT += totalDistance;
                        totalFuelUsedUnit += totalFuelUsed;
                        totalCostOfRunningUnit += totalCostOfRunning;
                        driverCostUnit += driverCost;
                        profitLossUnitT += profit_loss;
                    }




                    driverCost = 0;
                    profit_loss = 0;
                    driverCost = 0;
                    addblueTotal = 0;
                    approximateAdblue = 0;
                    totalCostOfRunning = 0;
                    revenue = 0;
                    fixedCost = 0;
                    totalDistance = 0;
                    totalFuelUsed = 0;


                }
            }
            cmd = new OleDbCommand("INSERT INTO [1$] ( F1, F2, F3, F4, F5, F6, F7, F8, F9 ) VALUES ( " + "'" + "18T" + "'" + "," + revenue18T + "," + fixedCost18T + "," + addblueTotal18T + "," + totalDistance18T + "," + totalFuelUsed18T + "," + totalCostOfRunning18T + "," + driverCost18T + "," + profitLoss18T + ")");

            myCon.Open();
            cmd.Connection = myCon;
            cmd.ExecuteNonQuery();
            myCon.Close();



            cmd = new OleDbCommand("INSERT INTO [1$] ( F1, F2, F3, F4, F5, F6, F7, F8, F9 ) VALUES ( " + "'" + "7T" + "'" + "," + revenue7T + "," + fixedCost7T + "," + addblueTotal7T + "," + totalDistance7T + "," + totalFuelUsed7T + "," + totalCostOfRunning7T + "," + driverCost7T + "," + profitLoss7T + ")");
            myCon.Open();
            cmd.Connection = myCon;
            cmd.ExecuteNonQuery();
            myCon.Close();


            cmd = new OleDbCommand("INSERT INTO [1$] ( F1, F2, F3, F4, F5, F6, F7, F8, F9 ) VALUES ( " + "'" + "Unit" + "'" + "," + revenueUnit + "," + fixedCostUnit + "," + addblueTotalUnit + "," + totalDistanceUnitT + "," + totalFuelUsedUnit + "," + totalCostOfRunningUnit + "," + driverCostUnit + "," + profitLossUnitT + ")");
            myCon.Open();
            cmd.Connection = myCon;
            cmd.ExecuteNonQuery();
            myCon.Close();

        }



        public void Processing_ProfitAndLossByVehicleCode(string from, string to)
        {

            cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + from.ToString() + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close();

            cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close();

            cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9 ) VALUES (" + "'" + "VehicleCode" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "FixedCost" + "'" + "," + "'" + "AdblueCost" + "'" + "," + "'" + "TotalDistance" + "'" + "," + "'" + "FuelUsed" + "'" + "," + "'" + "FuelCost" + "'" + "," + "'" + "DriveCosts" + "'" + "," + "'" + "ProfitLoss" + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close();
            
            float _profit_loss = 0;
            float _driverCost = 0;
            float _addblueTotal = 0;
            float _totalCostOfRunning = 0;
            float _revenue = 0;
            float _fixedCost = 0;

            CostingDB cst = new CostingDB();
            cst.DateFrom = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            cst.DateTo = DateTime.ParseExact(to, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            VehicleDb vehDB = new VehicleDb();
            DataTable vehiclesDT = vehDB.SelectAllVehicles();

            if (vehiclesDT.Rows.Count > 0 && vehiclesDT != null)
            {
                _revenue = 0;
                for (int u = 0; u <= vehiclesDT.Rows.Count - 1; u++)
                {

                    string _vehCode = vehiclesDT.Rows[u]["Code"].ToString();
                    float _totalDistance = 0;
                    float _totalFuelUsed = 0;
                    float _dieselCostPerL = 0;
                    float _addBluePercentage = 0;
                    float _targetConsumption = 0;
                    float _approximateAddblueL = 0;
                    float _addBlueCost = 0;
                    int _veh = -1;

                    DateTime _date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    int _id = Int32.Parse(vehiclesDT.Rows[u]["Id"].ToString());

                    CostingDB costingRow = new CostingDB();
                    _veh = _id;
                    costingRow.Veh = _id;
                    costingRow.Date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                    DataTable costingRowDt = costingRow.SelectRowByVehAndDate();


                    Middle_Layer.VTRN_Data vtrnData = new Middle_Layer.VTRN_Data();
                    vtrnData.Veh = _veh;
                    vtrnData.Vtrn_Date_Driver = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);


                    Utilisation_DataDB util = new Utilisation_DataDB();
                    util.Man_Date_Drv = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                    util.Veh = _veh;


                    if (costingRowDt != null && costingRowDt.Rows.Count > 0)
                    {

                        _totalDistance = float.Parse(costingRowDt.Rows[0]["Total_Distance"].ToString());
                        _totalFuelUsed = float.Parse(costingRowDt.Rows[0]["Total_Fuel_Used"].ToString());
                        _dieselCostPerL = float.Parse(costingRowDt.Rows[0]["diesel_Cost_per_l"].ToString());
                        _addBluePercentage = float.Parse(costingRowDt.Rows[0]["AddBlue_Percentage_Per_L"].ToString());
                        _addBlueCost = float.Parse(costingRowDt.Rows[0]["Approximate_Addblue_Cost"].ToString());
                        _targetConsumption = float.Parse(costingRowDt.Rows[0]["Target_Consumption"].ToString());
                        _approximateAddblueL = float.Parse(costingRowDt.Rows[0]["Approximate_Adblue_L"].ToString());
                        _veh = int.Parse(costingRowDt.Rows[0]["Veh"].ToString());


                        DRV_DutyDB drvDt = new DRV_DutyDB();
                        drvDt.Date = _date;
                        drvDt.VehCode = _vehCode;
                        drvDt.Veh = _veh;

                        DataTable drivDuty = drvDt.SelectRowsUsingDateAndVeh();

                        System.Diagnostics.Debug.Write("driv duty count before = " + drivDuty.Rows.Count);
                        if (drivDuty != null && drivDuty.Rows.Count > 0)
                        {
                            int min = 0;
                            int hr = 0;
                            int minPlusHr = 0;
                            float totalHours = 0;
                            float overtimeHours = 0;
                            float standardHours = 0;
                            float overtimeEarnings = 0;
                            float standardEarnings = 0;

                            for (int a = 0; a <= drivDuty.Rows.Count - 1; a++)
                            {

                                DateTime time = DateTime.ParseExact(drivDuty.Rows[a]["Duty_Time"].ToString(), "HH:mm:ss", CultureInfo.InvariantCulture);
                                min += time.Minute;
                                hr += time.Hour;

                                minPlusHr = hr * 60 + min;
                                totalHours = (float)minPlusHr / 60;

                                if (overtimeHours > 0)
                                {
                                    overtimeEarnings = overtimeHours * float.Parse(drivDuty.Rows[a]["_drivers_Overtime_Rate"].ToString());
                                    standardHours = totalHours - overtimeHours;
                                    standardEarnings = standardHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    _driverCost = standardEarnings + overtimeEarnings;
                                }
                                else
                                {
                                    standardEarnings = totalHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    _driverCost = standardEarnings;
                                }
                            }
                        }
                    }


                    DataTable vtrn = vtrnData.SelectUsingVehAndDate();

                    if (vtrn != null && vtrn.Rows.Count > 0)
                    {
                        _fixedCost = float.Parse(vtrn.Rows[0]["Vtrn_Monies"].ToString());
                    }
                    else
                    {
                        _fixedCost = 0;
                    }

                    DataTable ut = util.SelectUsingVehAndDate();

                    if (ut != null && ut.Rows.Count > 0)
                    {
                        _revenue = float.Parse(ut.Rows[0]["Man_Total_Revenue"].ToString());

                    }
                    else
                    {
                        _revenue = 0;
                    }


                    _addblueTotal = _approximateAddblueL * _addBlueCost;
                    _totalCostOfRunning = _totalFuelUsed * _dieselCostPerL;
                    _profit_loss = _revenue - (_fixedCost + _addblueTotal + _totalFuelUsed + _totalCostOfRunning + _driverCost);

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1, F2, F3, F4, F5, F6, F7, F8, F9 ) VALUES ( " + "'" + _vehCode + "'" + "," + _revenue + "," + _fixedCost + "," + _addblueTotal + "," + _totalDistance + "," + _totalFuelUsed + "," + _totalCostOfRunning + "," + _driverCost + "," + _profit_loss + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    _driverCost = 0;
                }
            }

        }



        public void Processing_CustomerProfitAndLossByManifest(string from, string to)
        {
            //description
            ClientDateDb cldt = new ClientDateDb();
            cldt.Date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);


            Utilisation_DataDB utilDataDB = new Utilisation_DataDB();
            utilDataDB.DateFrom = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            utilDataDB.DateTo = DateTime.ParseExact(to, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            DataTable utilData = utilDataDB.SelectAllRowsBetweenDates();


            DataTable uniqueMan = utilDataDB.SelectUniqueManifestsBetweenDates3();


            //tmp, it is a mess
            DataView dv = uniqueMan.DefaultView ;
            dv.Sort = "Man_Veh_Code ASC";
            DataTable dt = dv.ToTable();
            uniqueMan = dt;

            Response.Write("uniqueMan = " + uniqueMan.Rows.Count);
            Response.Write("utilData = " + utilData.Rows.Count);

            cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + from.ToString() + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close();

            if (uniqueMan != null && uniqueMan.Rows.Count > 0)
            {
                //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12 ) VALUES (" + "'" + "Manifest" + "'" + "," + "'" + "Veh Reg" + "'" + "," + "'" + "Customer" + "'" + "," + "'" + "Postcode" + "'" + "," + "'" + "Pallets" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Fuel" + "'" + "," + "'" + "Adblue Cost" + "'" + "," + "'" + "Drivers Costs" + "'" + "," + "'" + "Fixed Costs" + "'" + "," + "'" + "Profit" + "'" + "," + "'" + "Utilisation" + "'" + ")");
                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12, F13 ) VALUES (" + "'" + "Manifest" + "'" + "," + "'" + "Veh Reg" + "'" + "," + "'" + "Customer" + "'" + "," + "'" + "Postcode" + "'" + "," + "'" + "Pallets" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Fuel" + "'" + "," + "'" + "Adblue Cost" + "'" + "," + "'" + "Drivers Costs" + "'" + "," + "'" + "Fixed Costs" + "'" + ","+ "'" + "Total Cost" + "'" + "," + "'" + "Profit" + "'" + "," + "'" + "Utilisation" + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close();

                for (int i = 0; i <= uniqueMan.Rows.Count - 1; i++)
                {
                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    float tot_utilisation = 0;
                    float prop_profit = 0;
                    float tot_profit = 0;
                    float tot_adblue = 0;
                    float fixedCost = 0;
                    float driverCost = 0;
                    float prop_cost = 0;
                    float tot_cost = 0;


                    

                    utilDataDB.Man_Number = Int32.Parse(uniqueMan.Rows[i]["Man_Number"].ToString());

                    DataTable manJobs = utilDataDB.SelectUniqueRowBetweenDatesByManifest();
                    if (manJobs != null && manJobs.Rows.Count > 0)
                    {

                        int man = Int32.Parse(manJobs.Rows[0]["Man_Number"].ToString());
                        string vehicle = manJobs.Rows[0]["Man_Veh_Code"].ToString();
                        float tot_revenue = float.Parse(manJobs.Rows[0]["Man_Total_Revenue"].ToString());
                        int tot_packs = Int32.Parse(manJobs.Rows[0]["Man_Total_Packs"].ToString());
                        int job_nbr = Int32.Parse(manJobs.Rows[0]["Bkg_Number"].ToString());
                        int veho = Int32.Parse(manJobs.Rows[0]["Veh"].ToString());


                        string vehType = "Error";
                        VehicleDb vehDb = new VehicleDb();
                        vehDb.Id = veho;
                        DataTable vehDT = vehDb.SelectVehicleById();

                        string d = manJobs.Rows[0]["Man_Date_Drv"].ToString();
                        DateTime date = DateTime.ParseExact(d, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);



                        Middle_Layer.VTRN_Data fixedDb = new Middle_Layer.VTRN_Data();
                        fixedDb.Veh = veho;
                        fixedDb.Vtrn_Date_Driver = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                        DataTable fixedDT = fixedDb.SelectUsingVehAndDate();

                        if (fixedDT != null && fixedDT.Rows.Count > 0)
                        {
                            fixedCost = float.Parse(fixedDT.Rows[0]["Vtrn_Monies"].ToString());
                        }

                        if (vehDT != null)
                        {
                            vehType = vehDT.Rows[0]["Type"].ToString();
                            if (vehType.Contains("Unit"))
                            {
                                tot_utilisation = ((tot_packs / (float)26) * (float)100);
                            }
                            else if (vehType.Contains("18T"))
                            {
                                tot_utilisation = ((tot_packs / (float)14) * (float)100);
                            }
                            else if (vehType.Contains("7.5T"))
                            {
                                tot_utilisation = ((tot_packs / (float)8) * (float)100);
                            }

                            tot_utilisation = float.Parse(tot_utilisation.ToString("0.00"));
                        }


                        DRV_DutyDB drvDt = new DRV_DutyDB();
                        drvDt.Date = date;
                        drvDt.VehCode = vehicle;
                        drvDt.Veh = veho;

                        DataTable drivDuty = drvDt.SelectRowsUsingDateAndVeh();


                        if (drivDuty != null && drivDuty.Rows.Count > 0)
                        {

                            int min = 0;
                            int hr = 0;
                            int minPlusHr = 0;
                            float totalHours = 0;
                            float overtimeHours = 0;
                            float standardHours = 0;
                            float overtimeEarnings = 0;
                            float standardEarnings = 0;

                            for (int a = 0; a <= drivDuty.Rows.Count - 1; a++)
                            {
                                DateTime time = DateTime.ParseExact(drivDuty.Rows[a]["Duty_Time"].ToString(), "HH:mm:ss", CultureInfo.InvariantCulture);
                                min += time.Minute;
                                hr += time.Hour;

                                minPlusHr = hr * 60 + min;
                                totalHours = (float)minPlusHr / 60;


                                if (overtimeHours > 0)
                                {
                                    overtimeEarnings = overtimeHours * float.Parse(drivDuty.Rows[a]["_drivers_Overtime_Rate"].ToString());
                                    standardHours = totalHours - overtimeHours;
                                    standardEarnings = standardHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings + overtimeEarnings;
                                }
                                else
                                {
                                    standardEarnings = totalHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings;
                                }

                                driverCost = float.Parse(driverCost.ToString("0.00"));
                            }
                        }

                        CostingDB cst = new CostingDB();
                        cst.Veh = veho;
                        cst.Date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);


                        DataTable cstRowDT = cst.SelectRowByVehAndDate();

                        float tot_fuel = 0;
                        float tot_fuel_used = 0;
                        float diesel_cost = 0;
                        if (cstRowDT != null && cstRowDT.Rows.Count > 0)
                        {
                            tot_fuel_used = float.Parse(cstRowDT.Rows[0]["Total_Fuel_Used"].ToString());
                            diesel_cost = float.Parse(cstRowDT.Rows[0]["Diesel_Cost_per_l"].ToString());

                            tot_fuel_used = float.Parse(tot_fuel_used.ToString("0.00"));
                            diesel_cost = float.Parse(diesel_cost.ToString("0.00"));


                            tot_fuel = diesel_cost * tot_fuel_used;
                            tot_fuel = float.Parse(tot_fuel.ToString("0.00"));

                            float adblue_L = 0;
                            float adblue_cost = 0;

                            adblue_L = float.Parse(cstRowDT.Rows[0]["Approximate_Adblue_L"].ToString());
                            adblue_cost = float.Parse(cstRowDT.Rows[0]["Approximate_Addblue_Cost"].ToString());

                            adblue_L = float.Parse(adblue_L.ToString("0.00"));
                            adblue_cost = float.Parse(adblue_cost.ToString("0.00"));

                            tot_adblue = adblue_L * adblue_cost;
                        }

                        tot_profit = tot_revenue - (tot_fuel + tot_adblue + driverCost + fixedCost);
                        tot_cost = tot_fuel + tot_adblue + driverCost + fixedCost;

                        tot_cost = float.Parse(tot_cost.ToString("0.0"));
                        tot_profit = float.Parse(tot_profit.ToString("0.00"));
                        tot_adblue = float.Parse(tot_adblue.ToString("0.00"));


                        //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12 ) VALUES (" + "'" + man + "'" + "," + "'" + vehicle + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + tot_packs + "'" + "," + "'" + tot_revenue + "'" + "," + "'" + tot_fuel + "'" + "," + "'" + tot_adblue + "'" + "," + "'" + driverCost + "'" + "," + "'" + fixedCost + "'" + "," + "'" + tot_profit + "'" + "," + "'" + tot_utilisation + "'" + ")");
                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + man + "'" + "," + "'" + vehicle + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + tot_packs + "'" + "," + "'" + tot_revenue + "'" + "," + "'" + tot_fuel + "'" + "," + "'" + tot_adblue + "'" + "," + "'" + driverCost + "'" + "," + "'" + fixedCost + "'" + "," + "'"+  tot_cost + "'" + "," + "'" + tot_profit + "'" + "," + "'" + tot_utilisation + "'" + ")");
                        myCon.Open();
                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();
                        myCon.Close();

                        DataTable jobs = utilDataDB.SelectUniqueJobsBetweenDatesAndByManifest();
                        if (jobs != null && jobs.Rows.Count > 0)
                        {

                            for (int y = 0; y <= jobs.Rows.Count - 1; y++)
                            {
                                utilDataDB.Bkg_Number = Int32.Parse(jobs.Rows[y]["Bkg_Number"].ToString());
                                DataTable finalrw = utilDataDB.SelectUniqueRowsBetweenDatesAndByJob();


                                string postcode = finalrw.Rows[0]["Cons_Delivery_Postcode"].ToString();

                                
                                int job = Int32.Parse(finalrw.Rows[0]["Bkg_Number"].ToString());
                                string customer = finalrw.Rows[0]["Bkg_Customer_Code"].ToString();
                                int pallets = Int32.Parse(finalrw.Rows[0]["Bkg_Cons_Packs"].ToString());
                                float revenue = float.Parse(finalrw.Rows[0]["Bkg_Cons_Price"].ToString());

                                revenue = float.Parse(revenue.ToString("0.00"));


                                float prop_utilisation = 0;
                                if (finalrw != null && finalrw.Rows.Count > 0)
                                {

                                    if (vehType.Contains("Unit"))
                                    {
                                        prop_utilisation = ((pallets / (float)26) * (float)100);
                                    }
                                    else if (vehType.Contains("18T"))
                                    {
                                        prop_utilisation = ((pallets / (float)14) * (float)100);
                                    }
                                    else if (vehType.Contains("7.5T"))
                                    {
                                        prop_utilisation = ((pallets / (float)8) * (float)100);
                                    }
                                    prop_utilisation = float.Parse(prop_utilisation.ToString("0.00"));

                                    float prop_fixedCost = (prop_utilisation / (float)tot_utilisation) * fixedCost;
                                    float prop_driverCost = (prop_utilisation / (float)tot_utilisation) * driverCost;
                                    float prop_fuelCost = (prop_utilisation / (float)tot_utilisation) * tot_fuel;
                                    float prop_adblue = (prop_utilisation / (float)tot_utilisation) * tot_adblue;

                                    prop_profit = revenue - (prop_fuelCost + prop_adblue + prop_driverCost + prop_fixedCost);
                                    prop_cost = prop_fuelCost + prop_adblue + prop_driverCost + prop_fixedCost;

                                    revenue = float.Parse(revenue.ToString("0.00"));
                                    prop_profit = float.Parse(prop_profit.ToString("0.00"));
                                    prop_fixedCost = float.Parse(prop_fixedCost.ToString("0.00"));
                                    prop_driverCost = float.Parse(prop_driverCost.ToString("0.00"));
                                    prop_fuelCost = float.Parse(prop_fuelCost.ToString("0.00"));
                                    prop_adblue = float.Parse(prop_adblue.ToString("0.00"));
                                    prop_cost = float.Parse(prop_cost.ToString("0.00"));


                                    //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12 ) VALUES (" + "'" + "  " + "'" + "," + "'" + "  " + "'" + "," + "'" + customer + "'" + "," + "'" + " " + "'" + "," + "'" + pallets + "'" + "," + "'" + revenue + "'" + "," + "'" + prop_fuelCost + "'" + "," + "'" + prop_adblue + "'" + "," + "'" + prop_driverCost + "'" + "," + "'" + prop_fixedCost + "'" + "," + "'" + prop_profit + "'" + "," + "'" + prop_utilisation + "'" + ")");
                                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + "  " + "'" + "," + "'" + "  " + "'" + "," + "'" + customer + "'" + "," + "'" + postcode + "'" + "," + "'" + pallets + "'" + "," + "'" + revenue + "'" + "," + "'" + prop_fuelCost + "'" + "," + "'" + prop_adblue + "'" + "," + "'" + prop_driverCost + "'" + "," + "'" + prop_fixedCost + "'" + ","+ "'" + prop_cost + "'" + "," + "'" + prop_profit + "'" + "," + "'" + prop_utilisation + "'" + ")");

                                    myCon.Open();
                                    cmd.Connection = myCon;

                                    cmd.ExecuteNonQuery();
                                    myCon.Close();

                                }
                            }
                        }
                    }
                }
            }


        }


        public void Processing_CustomerProfitAndLossByManifest_VER3(string from, string to)
        {






        }


        public void Processing_CustomerProfitAndLossByManifest_VER2(string from, string to)
        {
            //description
            


            List<UtilisationReportHelperClass> listUtilData = new List<UtilisationReportHelperClass>();


            Utilisation_DataDB utilDataDB = new Utilisation_DataDB();
            utilDataDB.DateFrom = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            utilDataDB.DateTo = DateTime.ParseExact(to, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            DataTable utilData = utilDataDB.SelectAllRowsBetweenDates();


            DataTable uniqueMan = utilDataDB.SelectUniqueManifestsBetweenDates3();


            //tmp, it is a mess
            DataView dv = uniqueMan.DefaultView;
            dv.Sort = "Man_Veh_Code ASC";
            DataTable dt = dv.ToTable();
            uniqueMan = dt;

  

           /* cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + from.ToString() + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close(); */

            if (uniqueMan != null && uniqueMan.Rows.Count > 0)
            {
                //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12 ) VALUES (" + "'" + "Manifest" + "'" + "," + "'" + "Veh Reg" + "'" + "," + "'" + "Customer" + "'" + "," + "'" + "Postcode" + "'" + "," + "'" + "Pallets" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Fuel" + "'" + "," + "'" + "Adblue Cost" + "'" + "," + "'" + "Drivers Costs" + "'" + "," + "'" + "Fixed Costs" + "'" + "," + "'" + "Profit" + "'" + "," + "'" + "Utilisation" + "'" + ")");
                /*cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12, F13 ) VALUES (" + "'" + "Manifest" + "'" + "," + "'" + "Veh Reg" + "'" + "," + "'" + "Customer" + "'" + "," + "'" + "Postcode" + "'" + "," + "'" + "Pallets" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Fuel" + "'" + "," + "'" + "Adblue Cost" + "'" + "," + "'" + "Drivers Costs" + "'" + "," + "'" + "Fixed Costs" + "'" + "," + "'" + "Total Cost" + "'" + "," + "'" + "Profit" + "'" + "," + "'" + "Utilisation" + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close(); */
                string currentVehicle = ""; 
                

                for (int i = 0; i <= uniqueMan.Rows.Count - 1; i++)
                {
                    UtilisationReportHelperClass utl = new UtilisationReportHelperClass();



                    /*
                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close(); */

                    float tot_utilisation = 0;
                    float prop_profit = 0;
                    float tot_profit = 0;
                    float tot_adblue = 0;
                    float fixedCost = 0;
                    float driverCost = 0;
                    float prop_cost = 0;
                    float tot_cost = 0;

                    


                    utilDataDB.Man_Number = Int32.Parse(uniqueMan.Rows[i]["Man_Number"].ToString());

                    DataTable manJobs = utilDataDB.SelectUniqueRowBetweenDatesByManifest();
                    if (manJobs != null && manJobs.Rows.Count > 0)
                    {

                        int man = Int32.Parse(manJobs.Rows[0]["Man_Number"].ToString());
                        string vehicle = manJobs.Rows[0]["Man_Veh_Code"].ToString();
                        float tot_revenue = float.Parse(manJobs.Rows[0]["Man_Total_Revenue"].ToString());
                        int tot_packs = Int32.Parse(manJobs.Rows[0]["Man_Total_Packs"].ToString());
                        int job_nbr = Int32.Parse(manJobs.Rows[0]["Bkg_Number"].ToString());
                        int veho = Int32.Parse(manJobs.Rows[0]["Veh"].ToString());

                       


                        string vehType = "Error";
                        VehicleDb vehDb = new VehicleDb();
                        vehDb.Id = veho;
                        DataTable vehDT = vehDb.SelectVehicleById();

                        string d = manJobs.Rows[0]["Man_Date_Drv"].ToString();
                        DateTime date = DateTime.ParseExact(d, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);



                        Middle_Layer.VTRN_Data fixedDb = new Middle_Layer.VTRN_Data();
                        fixedDb.Veh = veho;
                        fixedDb.Vtrn_Date_Driver = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                        DataTable fixedDT = fixedDb.SelectUsingVehAndDate();

                        if (fixedDT != null && fixedDT.Rows.Count > 0)
                        {
                            fixedCost = float.Parse(fixedDT.Rows[0]["Vtrn_Monies"].ToString());
                        }

                        if (vehDT != null)
                        {
                            vehType = vehDT.Rows[0]["Type"].ToString();
                            if (vehType.Contains("Unit"))
                            {
                                utl.VehicleCapacity = 26;
                                tot_utilisation = ((tot_packs / (float)26) * (float)100);

                                if(tot_utilisation>100)
                                {
                                    tot_utilisation = 100;

                                }

                            }
                            else if (vehType.Contains("18T"))
                            {
                                utl.VehicleCapacity = 14;
                                tot_utilisation = ((tot_packs / (float)14) * (float)100);

                                if (tot_utilisation > 100)
                                {
                                    tot_utilisation = 100;

                                }
                            }
                            else if (vehType.Contains("7.5T"))
                            {
                                utl.VehicleCapacity = 8;
                                tot_utilisation = ((tot_packs / (float)8) * (float)100);

                                if (tot_utilisation > 100)
                                {
                                    tot_utilisation = 100;

                                }
                            }

                            tot_utilisation = float.Parse(tot_utilisation.ToString("0.00"));
                        }


                        DRV_DutyDB drvDt = new DRV_DutyDB();
                        drvDt.Date = date;
                        drvDt.VehCode = vehicle;
                        drvDt.Veh = veho;

                        DataTable drivDuty = drvDt.SelectRowsUsingDateAndVeh();


                        if (drivDuty != null && drivDuty.Rows.Count > 0)
                        {

                            int min = 0;
                            int hr = 0;
                            int minPlusHr = 0;
                            float totalHours = 0;
                            float overtimeHours = 0;
                            float standardHours = 0;
                            float overtimeEarnings = 0;
                            float standardEarnings = 0;

                            for (int a = 0; a <= drivDuty.Rows.Count - 1; a++)
                            {
                                DateTime time = DateTime.ParseExact(drivDuty.Rows[a]["Duty_Time"].ToString(), "HH:mm:ss", CultureInfo.InvariantCulture);
                                min += time.Minute;
                                hr += time.Hour;

                                minPlusHr = hr * 60 + min;
                                totalHours = (float)minPlusHr / 60;


                                if (overtimeHours > 0)
                                {
                                    overtimeEarnings = overtimeHours * float.Parse(drivDuty.Rows[a]["_drivers_Overtime_Rate"].ToString());
                                    standardHours = totalHours - overtimeHours;
                                    standardEarnings = standardHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings + overtimeEarnings;
                                }
                                else
                                {
                                    standardEarnings = totalHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings;
                                }

                                driverCost = float.Parse(driverCost.ToString("0.00"));
                            }
                        }

                        CostingDB cst = new CostingDB();
                        cst.Veh = veho;
                        cst.Date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);


                        DataTable cstRowDT = cst.SelectRowByVehAndDate();

                        float tot_fuel = 0;
                        float tot_fuel_used = 0;
                        float diesel_cost = 0;
                        if (cstRowDT != null && cstRowDT.Rows.Count > 0)
                        {
                            tot_fuel_used = float.Parse(cstRowDT.Rows[0]["Total_Fuel_Used"].ToString());
                            diesel_cost = float.Parse(cstRowDT.Rows[0]["Diesel_Cost_per_l"].ToString());

                            tot_fuel_used = float.Parse(tot_fuel_used.ToString("0.00"));
                            diesel_cost = float.Parse(diesel_cost.ToString("0.00"));


                            tot_fuel = diesel_cost * tot_fuel_used;
                            tot_fuel = float.Parse(tot_fuel.ToString("0.00"));

                            float adblue_L = 0;
                            float adblue_cost = 0;

                            adblue_L = float.Parse(cstRowDT.Rows[0]["Approximate_Adblue_L"].ToString());
                            adblue_cost = float.Parse(cstRowDT.Rows[0]["Approximate_Addblue_Cost"].ToString());

                            adblue_L = float.Parse(adblue_L.ToString("0.00"));
                            adblue_cost = float.Parse(adblue_cost.ToString("0.00"));

                            tot_adblue = adblue_L * adblue_cost;
                        }

                        tot_profit = tot_revenue - (tot_fuel + tot_adblue + driverCost + fixedCost);
                        tot_cost = tot_fuel + tot_adblue + driverCost + fixedCost;

                        tot_cost = float.Parse(tot_cost.ToString("0.0"));
                        tot_profit = float.Parse(tot_profit.ToString("0.00"));
                        tot_adblue = float.Parse(tot_adblue.ToString("0.00"));

                        /*
                        if (currentVehicle.Equals(vehicle))
                        {
                            vehicle = "";
                            fixedCost = 0;
                        }
                        else
                        {
                            currentVehicle = vehicle;
                            
                        } */


                        utl.ManifestNumber = man;
                        utl.VehCode = vehicle;
                        utl.PalletCount = tot_packs;
                        utl.Revenue = tot_revenue;
                        utl.Fuel = tot_fuel;
                        utl.AddBlueCost = tot_adblue;
                        utl.DriverCost = driverCost;
                        utl.FixedCost = fixedCost;
                        utl.TotalCost = tot_cost;
                        utl.Profit = tot_profit;
                        utl.Utilisation = tot_utilisation;

                        listUtilData.Add(utl);



                        //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12 ) VALUES (" + "'" + man + "'" + "," + "'" + vehicle + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + tot_packs + "'" + "," + "'" + tot_revenue + "'" + "," + "'" + tot_fuel + "'" + "," + "'" + tot_adblue + "'" + "," + "'" + driverCost + "'" + "," + "'" + fixedCost + "'" + "," + "'" + tot_profit + "'" + "," + "'" + tot_utilisation + "'" + ")");
                        /*cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + man + "'" + "," + "'" + vehicle + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + tot_packs + "'" + "," + "'" + tot_revenue + "'" + "," + "'" + tot_fuel + "'" + "," + "'" + tot_adblue + "'" + "," + "'" + driverCost + "'" + "," + "'" + fixedCost + "'" + "," + "'" + tot_cost + "'" + "," + "'" + tot_profit + "'" + "," + "'" + tot_utilisation + "'" + ")");
                        myCon.Open();
                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();
                        myCon.Close();*/

                        

                        DataTable jobs = utilDataDB.SelectUniqueJobsBetweenDatesAndByManifest();
                        if (jobs != null && jobs.Rows.Count > 0)
                        {

                            for (int y = 0; y <= jobs.Rows.Count - 1; y++)
                            {
                                UtilisationReportHelperClass utilElem = new UtilisationReportHelperClass();

                                utilDataDB.Bkg_Number = Int32.Parse(jobs.Rows[y]["Bkg_Number"].ToString());
                                DataTable finalrw = utilDataDB.SelectUniqueRowsBetweenDatesAndByJob();


                                string postcode = finalrw.Rows[0]["Cons_Delivery_Postcode"].ToString();
                                int status = Int32.Parse(finalrw.Rows[0]["Bkg_Status"].ToString());

                                int job = Int32.Parse(finalrw.Rows[0]["Bkg_Number"].ToString());
                                string customer = finalrw.Rows[0]["Bkg_Customer_Code"].ToString();
                                int pallets = Int32.Parse(finalrw.Rows[0]["Bkg_Cons_Packs"].ToString());
                                float revenue = float.Parse(finalrw.Rows[0]["Bkg_Cons_Price"].ToString());

                                revenue = float.Parse(revenue.ToString("0.00"));


                                float prop_utilisation = 0;
                                if (finalrw != null && finalrw.Rows.Count > 0)
                                {

                                    if (vehType.Contains("Unit"))
                                    {
                                        if(tot_utilisation==100)
                                        {
                                            prop_utilisation = ((pallets / (float)tot_packs) * (float)100);
                                        }
                                        else
                                        {
                                            prop_utilisation = ((pallets / (float)26) * (float)100);
                                        }

                                        
                                        
                                    }
                                    else if (vehType.Contains("18T"))
                                    {

                                        if (tot_utilisation == 100)
                                        {
                                            prop_utilisation = ((pallets / (float)tot_packs) * (float)100);
                                        }
                                        else
                                        {
                                            prop_utilisation = ((pallets / (float)14) * (float)100);
                                        }
                                        
                                        

                                    }
                                    else if (vehType.Contains("7.5T"))
                                    {
                                        if (tot_utilisation == 100)
                                        {
                                            prop_utilisation = ((pallets / (float)tot_packs) * (float)100);
                                        }
                                        else
                                        {
                                            prop_utilisation = ((pallets / (float)8) * (float)100);
                                        }

                                        
                                        
                                    }
                                    prop_utilisation = float.Parse(prop_utilisation.ToString("0.00"));

                                    float prop_fixedCost = (prop_utilisation / (float)tot_utilisation) * fixedCost;
                                    float prop_driverCost = (prop_utilisation / (float)tot_utilisation) * driverCost;
                                    float prop_fuelCost = (prop_utilisation / (float)tot_utilisation) * tot_fuel;
                                    float prop_adblue = (prop_utilisation / (float)tot_utilisation) * tot_adblue;

                                    prop_profit = revenue - (prop_fuelCost + prop_adblue + prop_driverCost + prop_fixedCost);
                                    prop_cost = prop_fuelCost + prop_adblue + prop_driverCost + prop_fixedCost;

                                    revenue = float.Parse(revenue.ToString("0.00"));
                                    prop_profit = float.Parse(prop_profit.ToString("0.00"));
                                    prop_fixedCost = float.Parse(prop_fixedCost.ToString("0.00"));
                                    prop_driverCost = float.Parse(prop_driverCost.ToString("0.00"));
                                    prop_fuelCost = float.Parse(prop_fuelCost.ToString("0.00"));
                                    prop_adblue = float.Parse(prop_adblue.ToString("0.00"));
                                    prop_cost = float.Parse(prop_cost.ToString("0.00"));


                                    utilElem.Customer = customer;
                                    utilElem.PostCode = postcode;
                                    utilElem.PalletCount = pallets;
                                    utilElem.Revenue = revenue;
                                    utilElem.Fuel = prop_fuelCost;
                                    utilElem.AddBlueCost = prop_adblue;
                                    utilElem.DriverCost = prop_driverCost;
                                    utilElem.FixedCost = prop_fixedCost;
                                    utilElem.TotalCost = prop_cost;
                                    utilElem.Profit = prop_profit;
                                    utilElem.Utilisation = prop_utilisation;
                                    utilElem.Status = status;

                                    /*
                                    if (utilElem.Customer == "ITLUGGAG")
                                    {
                                        utilElem.PalletCount = utilElem.PalletCount / 2;
                                    }*/

                                    utl.list.Add(utilElem);
                                    

                                    //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12 ) VALUES (" + "'" + "  " + "'" + "," + "'" + "  " + "'" + "," + "'" + customer + "'" + "," + "'" + " " + "'" + "," + "'" + pallets + "'" + "," + "'" + revenue + "'" + "," + "'" + prop_fuelCost + "'" + "," + "'" + prop_adblue + "'" + "," + "'" + prop_driverCost + "'" + "," + "'" + prop_fixedCost + "'" + "," + "'" + prop_profit + "'" + "," + "'" + prop_utilisation + "'" + ")");
                                   /* cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + "  " + "'" + "," + "'" + "  " + "'" + "," + "'" + customer + "'" + "," + "'" + postcode + "'" + "," + "'" + pallets + "'" + "," + "'" + revenue + "'" + "," + "'" + prop_fuelCost + "'" + "," + "'" + prop_adblue + "'" + "," + "'" + prop_driverCost + "'" + "," + "'" + prop_fixedCost + "'" + "," + "'" + prop_cost + "'" + "," + "'" + prop_profit + "'" + "," + "'" + prop_utilisation + "'" + ")");

                                    myCon.Open();
                                    cmd.Connection = myCon;

                                    cmd.ExecuteNonQuery();
                                    myCon.Close(); */

                                }
                            }
                        }
                    }
                }
            }


            VehicleDb vh = new VehicleDb();
            DataTable allVehivles = vh.SelectAllVehicles();

           
            foreach (DataRow elem in allVehivles.Rows )
            {
                string code = elem["Code"].ToString();
                int palletCount = 0;
                float revenue = 0;
                float fuel = 0;
                float adblue = 0;
                float driverCost = 0;
                float totalCost = 0;
                float profit = 0;
                float fixedCost = 0;




                foreach (UtilisationReportHelperClass e in listUtilData)
                {
                    if(e.VehCode.Equals(code))
                    {
                        driverCost += e.DriverCost;
                        profit += e.Profit;
                        fixedCost += e.FixedCost;
                        revenue += e.Revenue;
                        adblue += e.AddBlueCost;
                        totalCost += e.TotalCost;
                        fuel += e.Fuel;
                        palletCount += e.PalletCount;
                    }   
                    
                }
            

                foreach (UtilisationReportHelperClass e in listUtilData)
                {
                    if (e.VehCode.Equals(code))
                    {
                        e.SummedPalletCount = palletCount;
                        e.SummedTotalCost = totalCost;
                        e.SummedAdblue = adblue;
                        e.SummedDriverCost = driverCost;
                        e.SummedProfit = profit;
                        e.SummedFixed = fixedCost;
                        e.SummedFuel = fuel;
                        e.SummedRevenue = revenue;
                        
                    }

                }

                foreach (UtilisationReportHelperClass e in listUtilData)
                {
                    foreach (UtilisationReportHelperClass o in e.list)
                    {


                        o.FixedCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.FixedCost;
                        o.FixedCost = float.Parse(o.FixedCost.ToString("0.00"));

                        o.Fuel = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.Fuel;
                        o.Fuel = float.Parse(o.Fuel.ToString("0.00"));

                        o.AddBlueCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.AddBlueCost;
                        o.AddBlueCost = float.Parse(o.AddBlueCost.ToString("0.00"));

                        o.DriverCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.DriverCost;
                        o.DriverCost = float.Parse(o.DriverCost.ToString("0.00"));

                        o.TotalCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.TotalCost;
                        o.TotalCost = float.Parse(o.TotalCost.ToString("0.00"));

                        o.Profit = o.Revenue - o.TotalCost;

                    }

                }






            }

            


            cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + from.ToString() + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close();


            cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12, F13 ) VALUES (" + "'" + "Manifest" + "'" + "," + "'" + "Veh Reg" + "'" + "," + "'" + "Customer" + "'" + "," + "'" + "Postcode" + "'" + "," + "'" + "Pallets" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Fuel" + "'" + "," + "'" + "Adblue Cost" + "'" + "," + "'" + "Drivers Costs" + "'" + "," + "'" + "Fixed Costs" + "'" + "," + "'" + "Total Cost" + "'" + "," + "'" + "Profit" + "'" + "," + "'" + "Utilisation" + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close();


            /*
                      if (currentVehicle.Equals(vehicle))
                      {
                          vehicle = "";
                          fixedCost = 0;
                      }
                      else
                      {
                          currentVehicle = vehicle;

                      } */

            string currentveh = "";
            bool flag = false;
            foreach (UtilisationReportHelperClass u in listUtilData)
            {

                if (currentveh.Equals(u.VehCode))
                {
                    u.VehCode = "";
                    //fixedCost = 0;
                    flag = true;
                }
                else
                {
                    flag = false;
                    currentveh = u.VehCode;

                }








                myCon.Open();
                if (!flag)
                {


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");

                    
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    //myCon.Close();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + " " + "'" + "," + "'" + u.VehCode + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + u.SummedPalletCount + "'" + "," + "'" + u.SummedRevenue + "'" + "," + "'" + u.Fuel  + "'" + "," + "'" + u.AddBlueCost + "'" + "," + "'" + u.DriverCost + "'" + "," + "'" + u.FixedCost + "'" + "," + "'" + u.TotalCost + "'" + "," + "'" + (u.SummedRevenue-u.TotalCost) + "'" + "," + "'" + " " + "'" + ")");
                    //myCon.Open();
                    cmd.Connection = myCon;
                    cmd.ExecuteNonQuery();
                    //myCon.Close();




                }
                else
                {

                    

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "                                                                                           " + "'" + ")");
                   // myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    //myCon.Close();


                    


                   
                }

                cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + u.ManifestNumber + "'" + ")");
               // myCon.Open();

                cmd.Connection = myCon;
                cmd.ExecuteNonQuery();
               // myCon.Close();


                foreach (UtilisationReportHelperClass y in u.list)
                {
                    System.Diagnostics.Debug.WriteLine("pallet count y = " + y.PalletCount);
                    //prop_utilisation = ((pallets / (float)26) * (float)100);

                    
                    /*
                    float newFixedCost = ((float)y.PalletCount / (float)u.SummedPalletCount) * u.FixedCost;
                    newFixedCost = float.Parse(newFixedCost.ToString("0.00"));

                    float newFuelCost = ((float)y.PalletCount / (float)u.SummedPalletCount) * u.Fuel;
                    newFuelCost = float.Parse(newFuelCost.ToString("0.00"));

                    float newAddblueCost = ((float)y.PalletCount / (float)u.SummedPalletCount) * u.AddBlueCost;
                    newAddblueCost = float.Parse(newAddblueCost.ToString("0.00"));

                    float newDriverCost = ((float)y.PalletCount / (float)u.SummedPalletCount) * u.DriverCost;
                    newDriverCost = float.Parse(newDriverCost.ToString("0.00"));

                    float newTotalCost = ((float)y.PalletCount / (float)u.SummedPalletCount) * u.TotalCost;
                    newTotalCost = float.Parse(newTotalCost.ToString("0.00")); */


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + "  " + "'" + "," + "'" + "  " + "'" + "," + "'" + y.Customer + "'" + "," + "'" + y.PostCode + "'" + "," + "'" + y.PalletCount + "'" + "," + "'" + y.Revenue + "'" + "," + "'" + y.Fuel + "'" + "," + "'" + y.AddBlueCost + "'" + "," + "'" + y.DriverCost+ "'" + "," + "'" + y.FixedCost + "'" + "," + "'" + y.TotalCost + "'" + "," + "'" + y.Profit + "'" + "," + "'" + y.Utilisation + "'" + ")");

                   //myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                   // myCon.Close();


                }

                myCon.Close();
            }


        }

        public void Processing_ProfitLossByClient_VER2(string from, string to)
        {
            //description



            List<UtilisationReportHelperClass> listUtilData = new List<UtilisationReportHelperClass>();


            Utilisation_DataDB utilDataDB = new Utilisation_DataDB();
            utilDataDB.DateFrom = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            utilDataDB.DateTo = DateTime.ParseExact(to, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            DataTable utilData = utilDataDB.SelectAllRowsBetweenDates();


            DataTable uniqueMan = utilDataDB.SelectUniqueManifestsBetweenDates3();


            //tmp, it is a mess
            DataView dv = uniqueMan.DefaultView;
            dv.Sort = "Man_Veh_Code ASC";
            DataTable dt = dv.ToTable();
            uniqueMan = dt;



            /* cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + from.ToString() + "'" + ")");

             myCon.Open();
             cmd.Connection = myCon;

             cmd.ExecuteNonQuery();
             myCon.Close(); */

            if (uniqueMan != null && uniqueMan.Rows.Count > 0)
            {
                //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12 ) VALUES (" + "'" + "Manifest" + "'" + "," + "'" + "Veh Reg" + "'" + "," + "'" + "Customer" + "'" + "," + "'" + "Postcode" + "'" + "," + "'" + "Pallets" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Fuel" + "'" + "," + "'" + "Adblue Cost" + "'" + "," + "'" + "Drivers Costs" + "'" + "," + "'" + "Fixed Costs" + "'" + "," + "'" + "Profit" + "'" + "," + "'" + "Utilisation" + "'" + ")");
                /*cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12, F13 ) VALUES (" + "'" + "Manifest" + "'" + "," + "'" + "Veh Reg" + "'" + "," + "'" + "Customer" + "'" + "," + "'" + "Postcode" + "'" + "," + "'" + "Pallets" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Fuel" + "'" + "," + "'" + "Adblue Cost" + "'" + "," + "'" + "Drivers Costs" + "'" + "," + "'" + "Fixed Costs" + "'" + "," + "'" + "Total Cost" + "'" + "," + "'" + "Profit" + "'" + "," + "'" + "Utilisation" + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close(); */
                string currentVehicle = "";


                for (int i = 0; i <= uniqueMan.Rows.Count - 1; i++)
                {
                    UtilisationReportHelperClass utl = new UtilisationReportHelperClass();



                    /*
                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close(); */

                    float tot_utilisation = 0;
                    float prop_profit = 0;
                    float tot_profit = 0;
                    float tot_adblue = 0;
                    float fixedCost = 0;
                    float driverCost = 0;
                    float prop_cost = 0;
                    float tot_cost = 0;




                    utilDataDB.Man_Number = Int32.Parse(uniqueMan.Rows[i]["Man_Number"].ToString());

                    DataTable manJobs = utilDataDB.SelectUniqueRowBetweenDatesByManifest();
                    if (manJobs != null && manJobs.Rows.Count > 0)
                    {

                        int man = Int32.Parse(manJobs.Rows[0]["Man_Number"].ToString());
                        string vehicle = manJobs.Rows[0]["Man_Veh_Code"].ToString();
                        float tot_revenue = float.Parse(manJobs.Rows[0]["Man_Total_Revenue"].ToString());
                        int tot_packs = Int32.Parse(manJobs.Rows[0]["Man_Total_Packs"].ToString());
                        int job_nbr = Int32.Parse(manJobs.Rows[0]["Bkg_Number"].ToString());
                        int veho = Int32.Parse(manJobs.Rows[0]["Veh"].ToString());




                        string vehType = "Error";
                        VehicleDb vehDb = new VehicleDb();
                        vehDb.Id = veho;
                        DataTable vehDT = vehDb.SelectVehicleById();

                        string d = manJobs.Rows[0]["Man_Date_Drv"].ToString();
                        DateTime date = DateTime.ParseExact(d, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);



                        Middle_Layer.VTRN_Data fixedDb = new Middle_Layer.VTRN_Data();
                        fixedDb.Veh = veho;
                        fixedDb.Vtrn_Date_Driver = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                        DataTable fixedDT = fixedDb.SelectUsingVehAndDate();

                        if (fixedDT != null && fixedDT.Rows.Count > 0)
                        {
                            fixedCost = float.Parse(fixedDT.Rows[0]["Vtrn_Monies"].ToString());
                        }

                        if (vehDT != null)
                        {
                            vehType = vehDT.Rows[0]["Type"].ToString();
                            if (vehType.Contains("Unit"))
                            {
                                utl.VehicleCapacity = 26;
                                tot_utilisation = ((tot_packs / (float)26) * (float)100);

                                if (tot_utilisation > 100)
                                {
                                    tot_utilisation = 100;

                                }

                            }
                            else if (vehType.Contains("18T"))
                            {
                                utl.VehicleCapacity = 14;
                                tot_utilisation = ((tot_packs / (float)14) * (float)100);

                                if (tot_utilisation > 100)
                                {
                                    tot_utilisation = 100;

                                }
                            }
                            else if (vehType.Contains("7.5T"))
                            {
                                utl.VehicleCapacity = 8;
                                tot_utilisation = ((tot_packs / (float)8) * (float)100);

                                if (tot_utilisation > 100)
                                {
                                    tot_utilisation = 100;

                                }
                            }

                            tot_utilisation = float.Parse(tot_utilisation.ToString("0.00"));
                        }


                        DRV_DutyDB drvDt = new DRV_DutyDB();
                        drvDt.Date = date;
                        drvDt.VehCode = vehicle;
                        drvDt.Veh = veho;

                        DataTable drivDuty = drvDt.SelectRowsUsingDateAndVeh();


                        if (drivDuty != null && drivDuty.Rows.Count > 0)
                        {

                            int min = 0;
                            int hr = 0;
                            int minPlusHr = 0;
                            float totalHours = 0;
                            float overtimeHours = 0;
                            float standardHours = 0;
                            float overtimeEarnings = 0;
                            float standardEarnings = 0;

                            for (int a = 0; a <= drivDuty.Rows.Count - 1; a++)
                            {
                                DateTime time = DateTime.ParseExact(drivDuty.Rows[a]["Duty_Time"].ToString(), "HH:mm:ss", CultureInfo.InvariantCulture);
                                min += time.Minute;
                                hr += time.Hour;

                                minPlusHr = hr * 60 + min;
                                totalHours = (float)minPlusHr / 60;


                                if (overtimeHours > 0)
                                {
                                    overtimeEarnings = overtimeHours * float.Parse(drivDuty.Rows[a]["_drivers_Overtime_Rate"].ToString());
                                    standardHours = totalHours - overtimeHours;
                                    standardEarnings = standardHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings + overtimeEarnings;
                                }
                                else
                                {
                                    standardEarnings = totalHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings;
                                }

                                driverCost = float.Parse(driverCost.ToString("0.00"));
                            }
                        }

                        CostingDB cst = new CostingDB();
                        cst.Veh = veho;
                        cst.Date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);


                        DataTable cstRowDT = cst.SelectRowByVehAndDate();

                        float tot_fuel = 0;
                        float tot_fuel_used = 0;
                        float diesel_cost = 0;
                        if (cstRowDT != null && cstRowDT.Rows.Count > 0)
                        {
                            tot_fuel_used = float.Parse(cstRowDT.Rows[0]["Total_Fuel_Used"].ToString());
                            diesel_cost = float.Parse(cstRowDT.Rows[0]["Diesel_Cost_per_l"].ToString());

                            tot_fuel_used = float.Parse(tot_fuel_used.ToString("0.00"));
                            diesel_cost = float.Parse(diesel_cost.ToString("0.00"));


                            tot_fuel = diesel_cost * tot_fuel_used;
                            tot_fuel = float.Parse(tot_fuel.ToString("0.00"));

                            float adblue_L = 0;
                            float adblue_cost = 0;

                            adblue_L = float.Parse(cstRowDT.Rows[0]["Approximate_Adblue_L"].ToString());
                            adblue_cost = float.Parse(cstRowDT.Rows[0]["Approximate_Addblue_Cost"].ToString());

                            adblue_L = float.Parse(adblue_L.ToString("0.00"));
                            adblue_cost = float.Parse(adblue_cost.ToString("0.00"));

                            tot_adblue = adblue_L * adblue_cost;
                        }

                        tot_profit = tot_revenue - (tot_fuel + tot_adblue + driverCost + fixedCost);
                        tot_cost = tot_fuel + tot_adblue + driverCost + fixedCost;

                        tot_cost = float.Parse(tot_cost.ToString("0.0"));
                        tot_profit = float.Parse(tot_profit.ToString("0.00"));
                        tot_adblue = float.Parse(tot_adblue.ToString("0.00"));

                        /*
                        if (currentVehicle.Equals(vehicle))
                        {
                            vehicle = "";
                            fixedCost = 0;
                        }
                        else
                        {
                            currentVehicle = vehicle;
                            
                        } */


                        utl.ManifestNumber = man;
                        utl.VehCode = vehicle;
                        utl.PalletCount = tot_packs;
                        utl.Revenue = tot_revenue;
                        utl.Fuel = tot_fuel;
                        utl.AddBlueCost = tot_adblue;
                        utl.DriverCost = driverCost;
                        utl.FixedCost = fixedCost;
                        utl.TotalCost = tot_cost;
                        utl.Profit = tot_profit;
                        utl.Utilisation = tot_utilisation;
                     

                        listUtilData.Add(utl);



                        //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12 ) VALUES (" + "'" + man + "'" + "," + "'" + vehicle + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + tot_packs + "'" + "," + "'" + tot_revenue + "'" + "," + "'" + tot_fuel + "'" + "," + "'" + tot_adblue + "'" + "," + "'" + driverCost + "'" + "," + "'" + fixedCost + "'" + "," + "'" + tot_profit + "'" + "," + "'" + tot_utilisation + "'" + ")");
                        /*cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + man + "'" + "," + "'" + vehicle + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + tot_packs + "'" + "," + "'" + tot_revenue + "'" + "," + "'" + tot_fuel + "'" + "," + "'" + tot_adblue + "'" + "," + "'" + driverCost + "'" + "," + "'" + fixedCost + "'" + "," + "'" + tot_cost + "'" + "," + "'" + tot_profit + "'" + "," + "'" + tot_utilisation + "'" + ")");
                        myCon.Open();
                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();
                        myCon.Close();*/



                        DataTable jobs = utilDataDB.SelectUniqueJobsBetweenDatesAndByManifest();
                        if (jobs != null && jobs.Rows.Count > 0)
                        {

                            for (int y = 0; y <= jobs.Rows.Count - 1; y++)
                            {
                                UtilisationReportHelperClass utilElem = new UtilisationReportHelperClass();

                                utilDataDB.Bkg_Number = Int32.Parse(jobs.Rows[y]["Bkg_Number"].ToString());
                                DataTable finalrw = utilDataDB.SelectUniqueRowsBetweenDatesAndByJob();


                                string postcode = finalrw.Rows[0]["Cons_Delivery_Postcode"].ToString();


                                int job = Int32.Parse(finalrw.Rows[0]["Bkg_Number"].ToString());
                                string customer = finalrw.Rows[0]["Bkg_Customer_Code"].ToString();
                                int pallets = Int32.Parse(finalrw.Rows[0]["Bkg_Cons_Packs"].ToString());
                                float revenue = float.Parse(finalrw.Rows[0]["Bkg_Cons_Price"].ToString());
                                int status = Int32.Parse(finalrw.Rows[0]["Bkg_Status"].ToString());

                                revenue = float.Parse(revenue.ToString("0.00"));


                                float prop_utilisation = 0;
                                if (finalrw != null && finalrw.Rows.Count > 0)
                                {

                                    if (vehType.Contains("Unit"))
                                    {
                                        if (tot_utilisation == 100)
                                        {
                                            prop_utilisation = ((pallets / (float)tot_packs) * (float)100);
                                        }
                                        else
                                        {
                                            prop_utilisation = ((pallets / (float)26) * (float)100);
                                        }



                                    }
                                    else if (vehType.Contains("18T"))
                                    {

                                        if (tot_utilisation == 100)
                                        {
                                            prop_utilisation = ((pallets / (float)tot_packs) * (float)100);
                                        }
                                        else
                                        {
                                            prop_utilisation = ((pallets / (float)14) * (float)100);
                                        }



                                    }
                                    else if (vehType.Contains("7.5T"))
                                    {
                                        if (tot_utilisation == 100)
                                        {
                                            prop_utilisation = ((pallets / (float)tot_packs) * (float)100);
                                        }
                                        else
                                        {
                                            prop_utilisation = ((pallets / (float)8) * (float)100);
                                        }



                                    }
                                    prop_utilisation = float.Parse(prop_utilisation.ToString("0.00"));

                                    float prop_fixedCost = (prop_utilisation / (float)tot_utilisation) * fixedCost;
                                    float prop_driverCost = (prop_utilisation / (float)tot_utilisation) * driverCost;
                                    float prop_fuelCost = (prop_utilisation / (float)tot_utilisation) * tot_fuel;
                                    float prop_adblue = (prop_utilisation / (float)tot_utilisation) * tot_adblue;

                                    prop_profit = revenue - (prop_fuelCost + prop_adblue + prop_driverCost + prop_fixedCost);
                                    prop_cost = prop_fuelCost + prop_adblue + prop_driverCost + prop_fixedCost;

                                    revenue = float.Parse(revenue.ToString("0.00"));
                                    prop_profit = float.Parse(prop_profit.ToString("0.00"));
                                    prop_fixedCost = float.Parse(prop_fixedCost.ToString("0.00"));
                                    prop_driverCost = float.Parse(prop_driverCost.ToString("0.00"));
                                    prop_fuelCost = float.Parse(prop_fuelCost.ToString("0.00"));
                                    prop_adblue = float.Parse(prop_adblue.ToString("0.00"));
                                    prop_cost = float.Parse(prop_cost.ToString("0.00"));


                                    utilElem.Customer = customer;
                                    utilElem.PostCode = postcode;
                                    utilElem.PalletCount = pallets;
                                    utilElem.Revenue = revenue;
                                    utilElem.Fuel = prop_fuelCost;
                                    utilElem.AddBlueCost = prop_adblue;
                                    utilElem.DriverCost = prop_driverCost;
                                    utilElem.FixedCost = prop_fixedCost;
                                    utilElem.TotalCost = prop_cost;
                                    utilElem.Profit = prop_profit;
                                    utilElem.Utilisation = prop_utilisation;
                                    utilElem.Status = status;

                                    /*
                                    if (utilElem.Customer == "ITLUGGAG")
                                    {
                                        utilElem.PalletCount = utilElem.PalletCount / 2;
                                    }*/

                                    utl.list.Add(utilElem);


                                    //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12 ) VALUES (" + "'" + "  " + "'" + "," + "'" + "  " + "'" + "," + "'" + customer + "'" + "," + "'" + " " + "'" + "," + "'" + pallets + "'" + "," + "'" + revenue + "'" + "," + "'" + prop_fuelCost + "'" + "," + "'" + prop_adblue + "'" + "," + "'" + prop_driverCost + "'" + "," + "'" + prop_fixedCost + "'" + "," + "'" + prop_profit + "'" + "," + "'" + prop_utilisation + "'" + ")");
                                    /* cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + "  " + "'" + "," + "'" + "  " + "'" + "," + "'" + customer + "'" + "," + "'" + postcode + "'" + "," + "'" + pallets + "'" + "," + "'" + revenue + "'" + "," + "'" + prop_fuelCost + "'" + "," + "'" + prop_adblue + "'" + "," + "'" + prop_driverCost + "'" + "," + "'" + prop_fixedCost + "'" + "," + "'" + prop_cost + "'" + "," + "'" + prop_profit + "'" + "," + "'" + prop_utilisation + "'" + ")");

                                     myCon.Open();
                                     cmd.Connection = myCon;

                                     cmd.ExecuteNonQuery();
                                     myCon.Close(); */

                                }
                            }
                        }
                    }
                }
            }


            VehicleDb vh = new VehicleDb();
            DataTable allVehivles = vh.SelectAllVehicles();


            foreach (DataRow elem in allVehivles.Rows)
            {
                string code = elem["Code"].ToString();
                int palletCount = 0;
                float revenue = 0;
                float fuel = 0;
                float adblue = 0;
                float driverCost = 0;
                float totalCost = 0;
                float profit = 0;
                float fixedCost = 0;




                foreach (UtilisationReportHelperClass e in listUtilData)
                {
                    if (e.VehCode.Equals(code))
                    {
                        driverCost += e.DriverCost;
                        profit += e.Profit;
                        fixedCost += e.FixedCost;
                        revenue += e.Revenue;
                        adblue += e.AddBlueCost;
                        totalCost += e.TotalCost;
                        fuel += e.Fuel;
                        palletCount += e.PalletCount;
                    }

                }


                foreach (UtilisationReportHelperClass e in listUtilData)
                {
                    if (e.VehCode.Equals(code))
                    {
                        e.SummedPalletCount = palletCount;
                        e.SummedTotalCost = totalCost;
                        e.SummedAdblue = adblue;
                        e.SummedDriverCost = driverCost;
                        e.SummedProfit = profit;
                        e.SummedFixed = fixedCost;
                        e.SummedFuel = fuel;
                        e.SummedRevenue = revenue;

                    }

                }

                foreach (UtilisationReportHelperClass e in listUtilData)
                {
                    foreach (UtilisationReportHelperClass o in e.list)
                    {


                        o.FixedCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.FixedCost;
                        o.FixedCost = float.Parse(o.FixedCost.ToString("0.00"));

                        o.Fuel = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.Fuel;
                        o.Fuel = float.Parse(o.Fuel.ToString("0.00"));

                        o.AddBlueCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.AddBlueCost;
                        o.AddBlueCost = float.Parse(o.AddBlueCost.ToString("0.00"));

                        o.DriverCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.DriverCost;
                        o.DriverCost = float.Parse(o.DriverCost.ToString("0.00"));

                        o.TotalCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.TotalCost;
                        o.TotalCost = float.Parse(o.TotalCost.ToString("0.00"));

                        o.Profit = o.Revenue - o.TotalCost;

                    }

                }






            }


            List<ClientNames> ls = new List<ClientNames>();

            foreach(UtilisationReportHelperClass elem in listUtilData)
            {
                foreach (UtilisationReportHelperClass elem2 in elem.list)
                {

                    if(ls.Count == 0)
                    {
                        ClientNames c = new ClientNames();
                        c.Client = elem2.Customer;
                        c.Postcode = elem2.PostCode;
                        c.ProfitLoss = elem2.Profit;
                        c.TotalCost = elem.TotalCost;

                        if (elem2.Status == 1)
                        {
                            c.CountOther++;
                        }
                        else if (elem2.Status == 2)
                        {
                            c.CountRefused++;


                        }
                        else if (elem2.Status == 3)
                        {
                            c.CountSuccess++;
                        }
                        else if (elem2.Status == 7)
                        {
                            c.CountCancelled++;
                        }
                        else
                        {
                            c.CountOther++;
                        }
                    }
                    else
                    {
                        int p = 0;
                        foreach (ClientNames ff in ls)
                        {
                            p++;
                            if (elem2.Customer.Equals(ff.Client))
                            {
                                ClientNames c = new ClientNames();
                                c.Client = elem2.Customer;
                                c.Postcode = elem2.PostCode;
                                c.ProfitLoss = elem2.Profit;
                                c.TotalCost = elem.TotalCost;

                                if (elem2.Status == 1)
                                {
                                    c.CountOther++;
                                }
                                else if (elem2.Status == 2)
                                {
                                    c.CountRefused++;


                                }
                                else if (elem2.Status == 3)
                                {
                                    c.CountSuccess++;
                                }
                                else if (elem2.Status == 7)
                                {
                                    c.CountCancelled++;
                                }
                                else
                                {
                                    c.CountOther++;
                                }
                                ls.Add(c);
                                break;


                            }
                            else if (ls.Count == p)
                            {

                                ClientNames c = new ClientNames();
                                c.Client = elem2.Customer;
                                c.Postcode = elem2.PostCode;
                                c.ProfitLoss = elem2.Profit;
                                c.TotalCost = elem.TotalCost;

                                if (elem2.Status == 1)
                                {
                                    c.CountOther++;
                                }
                                else if (elem2.Status == 2)
                                {
                                    c.CountRefused++;


                                }
                                else if (elem2.Status == 3)
                                {
                                    c.CountSuccess++;
                                }
                                else if (elem2.Status == 7)
                                {
                                    c.CountCancelled++;
                                }
                                else
                                {
                                    c.CountOther++;
                                }
                                ls.Add(c);
                                break;
                            }


                        }


                    

                            


                    }

                }


            }



            cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + from.ToString() + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;

            cmd.ExecuteNonQuery();
            myCon.Close();


            cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12, F13 ) VALUES (" + "'" + "Customer" + "'" + ","  + "'" + "Profit" + "'" + "'" + "Customer" + "'" + "'" + "Customer" + "'" + "'" + "Customer" + "'" + ")");

            myCon.Open();
            cmd.Connection = myCon;
             
            cmd.ExecuteNonQuery();
            myCon.Close();


            /*
                      if (currentVehicle.Equals(vehicle))
                      {
                          vehicle = "";
                          fixedCost = 0;
                      }
                      else
                      {
                          currentVehicle = vehicle;

                      } */

            string currentveh = "";
            bool flag = false;
            foreach (UtilisationReportHelperClass u in listUtilData)
            {

                if (currentveh.Equals(u.VehCode))
                {
                    u.VehCode = "";
                    //fixedCost = 0;
                    flag = true;
                }
                else
                {
                    flag = false;
                    currentveh = u.VehCode;

                }








                myCon.Open();
                if (!flag)
                {


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");


                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    //myCon.Close();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + " " + "'" + "," + "'" + u.VehCode + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + u.SummedPalletCount + "'" + "," + "'" + u.SummedRevenue + "'" + "," + "'" + u.Fuel + "'" + "," + "'" + u.AddBlueCost + "'" + "," + "'" + u.DriverCost + "'" + "," + "'" + u.FixedCost + "'" + "," + "'" + u.TotalCost + "'" + "," + "'" + (u.SummedRevenue - u.TotalCost) + "'" + "," + "'" + " " + "'" + ")");
                    //myCon.Open();
                    cmd.Connection = myCon;
                    cmd.ExecuteNonQuery();
                    //myCon.Close();




                }
                else
                {



                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "                                                                                           " + "'" + ")");
                    // myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    //myCon.Close();






                }

                cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + u.ManifestNumber + "'" + ")");
                // myCon.Open();

                cmd.Connection = myCon;
                cmd.ExecuteNonQuery();
                // myCon.Close();


                foreach (UtilisationReportHelperClass y in u.list)
                {
                    System.Diagnostics.Debug.WriteLine("pallet count y = " + y.PalletCount);
                    //prop_utilisation = ((pallets / (float)26) * (float)100);


                    /*
                    float newFixedCost = ((float)y.PalletCount / (float)u.SummedPalletCount) * u.FixedCost;
                    newFixedCost = float.Parse(newFixedCost.ToString("0.00"));

                    float newFuelCost = ((float)y.PalletCount / (float)u.SummedPalletCount) * u.Fuel;
                    newFuelCost = float.Parse(newFuelCost.ToString("0.00"));

                    float newAddblueCost = ((float)y.PalletCount / (float)u.SummedPalletCount) * u.AddBlueCost;
                    newAddblueCost = float.Parse(newAddblueCost.ToString("0.00"));

                    float newDriverCost = ((float)y.PalletCount / (float)u.SummedPalletCount) * u.DriverCost;
                    newDriverCost = float.Parse(newDriverCost.ToString("0.00"));

                    float newTotalCost = ((float)y.PalletCount / (float)u.SummedPalletCount) * u.TotalCost;
                    newTotalCost = float.Parse(newTotalCost.ToString("0.00")); */


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + "  " + "'" + "," + "'" + "  " + "'" + "," + "'" + y.Customer + "'" + "," + "'" + y.PostCode + "'" + "," + "'" + y.PalletCount + "'" + "," + "'" + y.Revenue + "'" + "," + "'" + y.Fuel + "'" + "," + "'" + y.AddBlueCost + "'" + "," + "'" + y.DriverCost + "'" + "," + "'" + y.FixedCost + "'" + "," + "'" + y.TotalCost + "'" + "," + "'" + y.Profit + "'" + "," + "'" + y.Utilisation + "'" + ")");

                    //myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    // myCon.Close();


                }

                myCon.Close();
            }


        }



        public void Processing_ProfitLossByClient(string from, string to)
        {

            //description
            List<UtilisationReportHelperClass> listUtilData = new List<UtilisationReportHelperClass>();
            List<ClientNames> clientNames = new List<ClientNames>();

            Utilisation_DataDB utilDataDB = new Utilisation_DataDB();
            utilDataDB.DateFrom = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            utilDataDB.DateTo = DateTime.ParseExact(to, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            DataTable utilData = utilDataDB.SelectAllRowsBetweenDates();
            DataTable uniqueMan = utilDataDB.SelectUniqueManifestsBetweenDates3();

            if (uniqueMan != null && uniqueMan.Rows.Count > 0)
            {

                for (int i = 0; i <= uniqueMan.Rows.Count - 1; i++)
                {                  

                    float tot_utilisation = 0;
                    float prop_profit = 0;
                    float tot_profit = 0;
                    float tot_adblue = 0;
                    float fixedCost = 0;
                    float driverCost = 0;

                    utilDataDB.Man_Number = Int32.Parse(uniqueMan.Rows[i]["Man_Number"].ToString());

                    DataTable manJobs = utilDataDB.SelectUniqueRowBetweenDatesByManifest();
                    if (manJobs != null && manJobs.Rows.Count > 0)
                    {

                        int man = Int32.Parse(manJobs.Rows[0]["Man_Number"].ToString());
                        string vehicle = manJobs.Rows[0]["Man_Veh_Code"].ToString();
                        float tot_revenue = float.Parse(manJobs.Rows[0]["Man_Total_Revenue"].ToString());
                        int tot_packs = Int32.Parse(manJobs.Rows[0]["Man_Total_Packs"].ToString());
                        int job_nbr = Int32.Parse(manJobs.Rows[0]["Bkg_Number"].ToString());
                        int veho = Int32.Parse(manJobs.Rows[0]["Veh"].ToString());

                        tot_revenue = float.Parse(tot_revenue.ToString("0.00"));


                        string vehType = "Error";
                        VehicleDb vehDb = new VehicleDb();
                        vehDb.Id = veho;
                        DataTable vehDT = vehDb.SelectVehicleById();

                        string d = manJobs.Rows[0]["Man_Date_Drv"].ToString();
                        DateTime date = DateTime.ParseExact(d, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);



                        Middle_Layer.VTRN_Data fixedDb = new Middle_Layer.VTRN_Data();
                        fixedDb.Veh = veho;
                        fixedDb.Vtrn_Date_Driver = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                        DataTable fixedDT = fixedDb.SelectUsingVehAndDate();

                        if (fixedDT != null && fixedDT.Rows.Count > 0)
                        {
                            fixedCost = float.Parse(fixedDT.Rows[0]["Vtrn_Monies"].ToString());
                            fixedCost = float.Parse(fixedCost.ToString("0.00"));
                        }

                        if (vehDT != null)
                        {
                            vehType = vehDT.Rows[0]["Type"].ToString();
                            if (vehType.Contains("Unit"))
                            {
                                tot_utilisation = ((tot_packs / (float)26) * (float)100);
                            }
                            else if (vehType.Contains("18T"))
                            {
                                tot_utilisation = ((tot_packs / (float)14) * (float)100);
                            }
                            else if (vehType.Contains("7.5T"))
                            {
                                tot_utilisation = ((tot_packs / (float)8) * (float)100);
                            }
                        }
                        tot_utilisation = float.Parse(tot_utilisation.ToString("0.00"));


                        DRV_DutyDB drvDt = new DRV_DutyDB();
                        drvDt.Date = date;
                        drvDt.VehCode = vehicle;
                        drvDt.Veh = veho;

                        DataTable drivDuty = drvDt.SelectRowsUsingDateAndVeh();


                        if (drivDuty != null && drivDuty.Rows.Count > 0)
                        {
                            int min = 0;
                            int hr = 0;
                            int minPlusHr = 0;
                            float totalHours = 0;
                            float overtimeHours = 0;
                            float standardHours = 0;
                            float overtimeEarnings = 0;
                            float standardEarnings = 0;

                            for (int a = 0; a <= drivDuty.Rows.Count - 1; a++)
                            {
                                DateTime time = DateTime.ParseExact(drivDuty.Rows[a]["Duty_Time"].ToString(), "HH:mm:ss", CultureInfo.InvariantCulture);
                                min += time.Minute;
                                hr += time.Hour;

                                minPlusHr = hr * 60 + min;
                                totalHours = (float)minPlusHr / 60;


                                if (overtimeHours > 0)
                                {
                                    overtimeEarnings = overtimeHours * float.Parse(drivDuty.Rows[a]["_drivers_Overtime_Rate"].ToString());
                                    standardHours = totalHours - overtimeHours;
                                    standardEarnings = standardHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings + overtimeEarnings;
                                }
                                else
                                {
                                    standardEarnings = totalHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings;
                                }

                                driverCost = float.Parse(driverCost.ToString("0.00"));

                            }
                        }

                        CostingDB cst = new CostingDB();
                        cst.Veh = veho;
                        //--------------------------------------cst.Date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        cst.Date = DateTime.ParseExact(d, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                        DataTable cstRowDT = cst.SelectRowByVehAndDate();

                        float tot_fuel = 0;
                        float tot_fuel_used = 0;
                        float diesel_cost = 0;
                        if (cstRowDT != null && cstRowDT.Rows.Count > 0)
                        {
                            tot_fuel_used = float.Parse(cstRowDT.Rows[0]["Total_Fuel_Used"].ToString());
                            diesel_cost = float.Parse(cstRowDT.Rows[0]["Diesel_Cost_per_l"].ToString());

                            tot_fuel_used = float.Parse(tot_fuel_used.ToString("0.00"));
                            diesel_cost = float.Parse(diesel_cost.ToString("0.00"));


                            tot_fuel = diesel_cost * tot_fuel_used;

                            float adblue_L = 0;
                            float adblue_cost = 0;

                            adblue_L = float.Parse(cstRowDT.Rows[0]["Approximate_Adblue_L"].ToString());
                            adblue_cost = float.Parse(cstRowDT.Rows[0]["Approximate_Addblue_Cost"].ToString());

                            adblue_L = float.Parse(adblue_L.ToString("0.00"));
                            adblue_cost = float.Parse(adblue_cost.ToString("0.00"));

                            tot_adblue = adblue_L * adblue_cost;
                        }

                        tot_profit = tot_revenue - (tot_fuel + tot_adblue + driverCost + fixedCost);

                        tot_profit = float.Parse(tot_profit.ToString("0.00"));




                        DataTable jobs = utilDataDB.SelectUniqueJobsBetweenDatesAndByManifest();
                        if (jobs != null && jobs.Rows.Count > 0)
                        {

                            for (int y = 0; y <= jobs.Rows.Count - 1; y++)
                            {
                                utilDataDB.Bkg_Number = Int32.Parse(jobs.Rows[y]["Bkg_Number"].ToString());
                                DataTable finalrw = utilDataDB.SelectUniqueRowsBetweenDatesAndByJob();

                                int job = Int32.Parse(finalrw.Rows[0]["Bkg_Number"].ToString());
                                string customer = finalrw.Rows[0]["Bkg_Customer_Code"].ToString();
                                int pallets = Int32.Parse(finalrw.Rows[0]["Bkg_Cons_Packs"].ToString());
                                float revenue = float.Parse(finalrw.Rows[0]["Bkg_Cons_Price"].ToString());

                                revenue = float.Parse(revenue.ToString("0.00"));

                                float prop_utilisation = 0;
                                if (finalrw != null && finalrw.Rows.Count > 0)
                                {

                                    if (vehType.Contains("Unit"))
                                    {
                                        prop_utilisation = ((pallets / (float)26) * (float)100);
                                    }
                                    else if (vehType.Contains("18T"))
                                    {
                                        prop_utilisation = ((pallets / (float)14) * (float)100);
                                    }
                                    else if (vehType.Contains("7.5T"))
                                    {
                                        prop_utilisation = ((pallets / (float)8) * (float)100);
                                    }
                                    prop_utilisation = float.Parse(prop_utilisation.ToString("0.00"));

                                    float prop_fixedCost = (prop_utilisation / (float)tot_utilisation) * fixedCost;
                                    float prop_driverCost = (prop_utilisation / (float)tot_utilisation) * driverCost;
                                    float prop_fuelCost = (prop_utilisation / (float)tot_utilisation) * tot_fuel;
                                    float prop_adblue = (prop_utilisation / (float)tot_utilisation) * tot_adblue;

                                    prop_profit = revenue - (prop_fuelCost + prop_adblue + prop_driverCost + prop_fixedCost);


                                    revenue = float.Parse(revenue.ToString("0.00"));
                                    prop_profit = float.Parse(prop_profit.ToString("0.00"));
                                    prop_fixedCost = float.Parse(prop_fixedCost.ToString("0.00"));
                                    prop_driverCost = float.Parse(prop_driverCost.ToString("0.00"));
                                    prop_fuelCost = float.Parse(prop_fuelCost.ToString("0.00"));
                                    prop_adblue = float.Parse(prop_adblue.ToString("0.00"));

                                    int c = -1;
                                    bool flag = false;

                                    if (clientNames.Count == 0)
                                    {
                                        ClientNames cl = new ClientNames();
                                        cl.Client = customer;
                                        cl.ProfitLoss = prop_profit;
                                        cl.ProfitLoss = float.Parse( cl.ProfitLoss.ToString("0.00") );
                                        clientNames.Add(cl);

                                    }
                                    else if (clientNames.Count>0)
                                    {
                                        foreach (ClientNames elem in clientNames)
                                        {
                                            System.Diagnostics.Debug.WriteLine("In here");
                                            c++;
                                            if (elem.Client.Equals(customer))
                                            {
                                                System.Diagnostics.Debug.WriteLine("In here 2");
                                                elem.ProfitLoss += float.Parse(prop_profit.ToString("0.00"));
                                                elem.ProfitLoss = float.Parse(elem.ProfitLoss.ToString("0.00"));
                                                flag = false;
                                                break;
                                            }
                                            else if (c >= clientNames.Count - 1)
                                            {
                                                System.Diagnostics.Debug.WriteLine("In here 3");
                                                flag = true;
                                            }

                                        }

                                        if (flag)
                                        {
                                            System.Diagnostics.Debug.WriteLine("In here 4");
                                            ClientNames cl = new ClientNames();
                                            cl.Client = customer;
                                            cl.ProfitLoss =+float.Parse(prop_profit.ToString("0.00"));
                                            cl.ProfitLoss = float.Parse(cl.ProfitLoss.ToString("0.00"));
                                            clientNames.Add(cl);
                                        }




                                    }


                                    


                                }
                            }
                        }
                    }
                }
            }


            cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2 ) VALUES (" + "'" + "Client Name" + "'" + "," + "'" + "ProfitLoss" + "'" +  ")");
            myCon.Open();
            cmd.Connection = myCon;
            cmd.ExecuteNonQuery();
            myCon.Close();



            
            foreach (ClientNames elem in clientNames)
            {
                myCon.Open();
                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2 ) VALUES (" + "'" + elem.Client.ToString() + "'" + "," + "'" + elem.ProfitLoss.ToString() + "'" + ")");
                cmd.Connection = myCon;
                cmd.ExecuteNonQuery();
                myCon.Close();
            }
            
        }


        public void  Processing_DetailedDriversReport(string from, string to)
        {
            DRV_DutyDB drv_duty = new DRV_DutyDB();
            drv_duty.Date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);



            DataTable dt = drv_duty.SelectRowsUsingDateVehAndDrv();


        






        }


        public void Processing_New(string from, string to, int report)
        {





            if (WaitForFile(path + reportOutput) == true)
            {

                if (File.Exists(path + reportOutput))
                {

                    File.Delete(path + reportOutput);


                }





                SetUpExcelConnection();

                /*
                string calculationSheet2 = "INSERT INTO [Calculations] (Client) VALUES (@Client) ";
               // myCon.Open();
                cmd = new OleDbCommand(calculationSheet2);
                cmd.Parameters.AddWithValue("@Client", 2);
                cmd.Connection = myCon2;
                cmd.ExecuteNonQuery(); */

                /*
                string calculationSheet3 = "CREATE TABLE [Calculations] (yy int) ";
                //myCon2.Open();
                cmd = new OleDbCommand(calculationSheet3);
                cmd.Connection = myCon2;
                cmd.ExecuteNonQuery();
                //myCon.Close();

                string calculationSheet4 = "INSERT INTO [Calculations] (yy) VALUES (@id) ";
                // myCon.Open();
                cmd = new OleDbCommand(calculationSheet4);
                cmd.Parameters.AddWithValue("@id", 2);
                cmd.Connection = myCon2;
                cmd.ExecuteNonQuery();

        */
                #region New_Processing Declarations

                Utilisation_DataDB utilDataDB = new Utilisation_DataDB();
                utilDataDB.DateFrom = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                utilDataDB.DateTo = DateTime.ParseExact(to, "dd/MM/yyyy", CultureInfo.InvariantCulture);


                DataTable dates = utilDataDB.SelectDateRange();
                VehicleDb vehDb = new VehicleDb();
                DataTable vehicles = vehDb.SelectAllVehicles();
                Utilisation_DataDB util = new Utilisation_DataDB();
                DataTable uniqueMan;

                List<HelperClass_Level1> list_lvl1 = new List<HelperClass_Level1>();






                DataView dv0 = dates.DefaultView;
                //dv0.Sort = "Man_Date_Drv";
                DataTable d = dv0.ToTable();
                dates = d;

                DateTime dateTime;


                #endregion



                foreach (DataRow date in dates.Rows)
                {


                    dateTime = (DateTime)date["Man_Date_Drv"];
                    System.Diagnostics.Debug.WriteLine(dateTime.ToString());


                    foreach (DataRow veh in vehicles.Rows)
                    {
                        System.Diagnostics.Debug.WriteLine("veh list = " + vehicles.Rows.Count);

                        util.DateFrom = (DateTime)date["Man_Date_Drv"];
                        util.DateTo = (DateTime)date["Man_Date_Drv"];
                        util.Man_Date_Drv = (DateTime)date["Man_Date_Drv"];
                        util.Veh = Int32.Parse(veh["Id"].ToString());
                        uniqueMan = util.SelectUsingVehAndDate();

                        //tmp, it is a mess
                        DataView dv = uniqueMan.DefaultView;


                        dv.Sort = "Man_Veh_Code ASC";
                        DataTable dt = dv.ToTable(true, "Man_Number", "Man_Total_Revenue", "Man_Total_Packs");
                        //DataTable dt2 = dv.ToTable(true, "Man_Number", "Bkg_Cons_Packs");
                        uniqueMan = dt;

                        System.Diagnostics.Debug.WriteLine("man list = " + uniqueMan.Rows.Count);

                        HelperClass_Level1 lvl1 = new HelperClass_Level1();
                        lvl1.RevenueTotal = (float)uniqueMan.AsEnumerable().Sum(r => r.Field<double>("Man_Total_Revenue"));
                        lvl1.PalletTotal = uniqueMan.AsEnumerable().Sum(r => r.Field<int>("Man_Total_Packs"));
                        // lvl1.Alternative_PalletTotal = dt2.AsEnumerable().Sum(r => r.Field<int>("Bkg_Cons_Packs"));
                        //lvl1.Alternative_PalletTotal = 0;
                        lvl1.VehicleCode = veh["Code"].ToString();

                        #region New_Processing fixedCostCalculations
                        ////fixed cost
                        Middle_Layer.VTRN_Data fixedDb = new Middle_Layer.VTRN_Data();
                        fixedDb.Vtrn_Date_Driver = (DateTime)date["Man_Date_Drv"];
                        fixedDb.Veh = Int32.Parse(veh["Id"].ToString());
                        DataTable fixedDT = fixedDb.SelectUsingVehAndDate();

                        if (fixedDT != null && fixedDT.Rows.Count > 0)
                        {
                            lvl1.FixedCostTotal = float.Parse(fixedDT.Rows[0]["Vtrn_Monies"].ToString());
                        }

                        #endregion

                        #region New_Processing DriverCost 
                        /////driver costs
                        DRV_DutyDB drvDt = new DRV_DutyDB();
                        drvDt.Date = (DateTime)date["Man_Date_Drv"];
                        drvDt.VehCode = veh["Code"].ToString();
                        drvDt.Veh = Int32.Parse(veh["Id"].ToString());

                        DataTable drivDuty = drvDt.SelectRowsUsingDateAndVeh();

                        lvl1.Date = drvDt.Date;

                        float driverCost = 0;
                        if (drivDuty != null && drivDuty.Rows.Count > 0)
                        {

                            int min = 0;
                            int hr = 0;
                            int minPlusHr = 0;
                            float totalHours = 0;
                            float overtimeHours = 0;
                            float standardHours = 0;
                            float overtimeEarnings = 0;
                            float standardEarnings = 0;

                            for (int a = 0; a <= drivDuty.Rows.Count - 1; a++)
                            {
                                DateTime time = DateTime.ParseExact(drivDuty.Rows[a]["Duty_Time"].ToString(), "HH:mm:ss", CultureInfo.InvariantCulture);
                                min += time.Minute;
                                hr += time.Hour;

                                minPlusHr = hr * 60 + min;
                                totalHours = (float)minPlusHr / 60;
                                lvl1.DrvCost_TotalHr = totalHours;
                                lvl1.DrvCost_OvertimeRate = float.Parse(drivDuty.Rows[a]["_drivers_Overtime_Rate"].ToString());
                                lvl1.DrvCost_StandardRate = float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());

                                if (totalHours >= 8)
                                {
                                    overtimeHours = totalHours - 8;
                                    overtimeEarnings = overtimeHours * float.Parse(drivDuty.Rows[a]["_drivers_Overtime_Rate"].ToString());
                                    standardHours = 8;
                                    standardEarnings = standardHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings + overtimeEarnings;
                                }
                                else
                                {
                                    standardHours = totalHours;
                                    standardEarnings = totalHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings;
                                }

                                driverCost = float.Parse(driverCost.ToString("0.00"));
                                lvl1.DriversCostTotal = driverCost;
                                lvl1.DrvCost_StandardHours = standardHours;
                                lvl1.DrvCost_OvertimeHours = overtimeHours;
                                lvl1.DrvCost_OvertimeEarnings = overtimeEarnings;
                                lvl1.DrvCost_StandardEarnings = standardEarnings;
                            }


                        }

                        #endregion


                        #region New_Processing Costing
                        ///////////////Costs
                        CostingDB cst = new CostingDB();
                        cst.Veh = Int32.Parse(veh["Id"].ToString());

                        cst.Date = (DateTime)date["Man_Date_Drv"];

                        DataTable cstRowDT = cst.SelectRowByVehAndDate();

                        float tot_fuel = 0;
                        float tot_fuel_used = 0;
                        float diesel_cost = 0;
                        if (cstRowDT != null && cstRowDT.Rows.Count > 0)
                        {
                            tot_fuel_used = float.Parse(cstRowDT.Rows[0]["Total_Fuel_Used"].ToString());
                            diesel_cost = float.Parse(cstRowDT.Rows[0]["Diesel_Cost_per_l"].ToString());

                            tot_fuel_used = float.Parse(tot_fuel_used.ToString("0.00"));
                            diesel_cost = float.Parse(diesel_cost.ToString("0.00"));


                            lvl1.FuelTotal = diesel_cost * tot_fuel_used;
                            lvl1.FuelTotal = float.Parse(lvl1.FuelTotal.ToString("0.00"));

                            float adblue_L = 0;
                            float adblue_cost = 0;

                            adblue_L = float.Parse(cstRowDT.Rows[0]["Approximate_Adblue_L"].ToString());
                            adblue_cost = float.Parse(cstRowDT.Rows[0]["Approximate_Addblue_Cost"].ToString());

                            adblue_L = float.Parse(adblue_L.ToString("0.00"));
                            adblue_cost = float.Parse(adblue_cost.ToString("0.00"));
                            lvl1.AdblueTotal = adblue_L * adblue_cost;

                        }

                        #endregion

                        lvl1.ProfitTotal = lvl1.RevenueTotal - (lvl1.FuelTotal + lvl1.AdblueTotal + lvl1.DriversCostTotal + lvl1.FixedCostTotal);
                        lvl1.TotalCost = lvl1.FuelTotal + lvl1.AdblueTotal + lvl1.DriversCostTotal + lvl1.FixedCostTotal;

                        lvl1.TotalCost = float.Parse(lvl1.TotalCost.ToString("0.0"));
                        lvl1.ProfitTotal = float.Parse(lvl1.ProfitTotal.ToString("0.00"));
                        lvl1.AdblueTotal = float.Parse(lvl1.AdblueTotal.ToString("0.00"));


                        string type = veh["Type"].ToString();
                        lvl1.VehType = type;


                        //in case there is bug while taking data from db (total pallet count is sometimes 0, causing havoc)


                        list_lvl1.Add(lvl1);

                        foreach (DataRow man in uniqueMan.Rows)
                        {






                            HelperClass_Level2 lvl2 = new HelperClass_Level2();
                            lvl2.ManifestNumber = Int32.Parse(man["Man_Number"].ToString());
                            lvl2.RevenueManTotal = float.Parse(man["Man_Total_Revenue"].ToString());
                            //lvl2.RevenueManTotal = float.Parse(man["Man_Total_Revenue"].ToString());
                            util.Man_Number = Int32.Parse(man["Man_Number"].ToString());

                            DataTable t = util.SelectMan_Total_Pack();



                            DataTable jobs = util.SelectAllWithManifestNumber();

                            int tot_packs = Int32.Parse(t.Rows[0]["Man_Total_Packs"].ToString());


                            //alternative pallet count
                            int alt_pallet_total = 0;
                            foreach (DataRow job in jobs.Rows)
                            {
                                alt_pallet_total += Int32.Parse(job["Bkg_Cons_Packs"].ToString());
                            }

                            if (tot_packs <= 0 && alt_pallet_total > 0)
                            {
                                tot_packs = alt_pallet_total;
                                lvl1.PalletTotal = alt_pallet_total;

                            }



                            //lvl1.RevenueTotal = (float)jobs.AsEnumerable().Sum(r => r.Field<double>("Man_Total_Revenue"));


                            //utilisation
                            string vehType = veh["Type"].ToString();
                            lvl1.VehType = vehType;
                            if (vehType.Contains("Unit"))
                            {
                                // lvl1.VehType = "Unit";
                                lvl2.VehicleCapacity = 26;
                                lvl2.UtilisationTotal = ((tot_packs / (float)26) * (float)100);
                                if (lvl2.UtilisationTotal > 100)
                                {
                                    lvl2.UtilisationTotal = 100;
                                }
                            }
                            else if (vehType.Contains("18T"))
                            {
                                //lvl1.VehType = "18T";
                                lvl2.VehicleCapacity = 14;
                                lvl2.UtilisationTotal = ((tot_packs / (float)14) * (float)100);

                                if (lvl2.UtilisationTotal > 100)
                                {
                                    lvl2.UtilisationTotal = 100;
                                }
                            }
                            else if (vehType.Contains("7.5T"))
                            {
                                //lvl1.VehType="7.5T";
                                lvl2.VehicleCapacity = 8;
                                lvl2.UtilisationTotal = ((tot_packs / (float)8) * (float)100);
                                if (lvl2.UtilisationTotal > 100)
                                {
                                    lvl2.UtilisationTotal = 100;
                                }
                            }

                            lvl2.UtilisationTotal = float.Parse(lvl2.UtilisationTotal.ToString("0.00"));
                            lvl1.list.Add(lvl2);
                            foreach (DataRow job in jobs.Rows)
                            {

                                HelperClass_Level3 lvl3 = new HelperClass_Level3();
                                lvl3.BookingNumber = Int32.Parse(job["Bkg_Number"].ToString());
                                lvl3.Postcode = job["Cons_Delivery_Postcode"].ToString();
                                lvl3.PalletCount = Int32.Parse(job["Bkg_Cons_Packs"].ToString());
                                lvl3.Customer = job["Bkg_Customer_Code"].ToString();
                                lvl3.BookingStatus = Int32.Parse(job["Bkg_Status"].ToString());

                                lvl3.Utilisation = 0;


                                if (vehType.Contains("Unit"))
                                {
                                    if (lvl2.UtilisationTotal == 100)
                                    {
                                        lvl3.Utilisation = ((lvl3.PalletCount / (float)tot_packs) * (float)100);
                                        lvl3.UtilisationToShow = ((lvl3.PalletCount / (float)tot_packs) * (float)100);
                                    }
                                    else
                                    {
                                        lvl3.Utilisation = ((lvl3.PalletCount / (float)26) * (float)100);
                                        lvl3.UtilisationToShow = ((lvl3.PalletCount / (float)26) * (float)100);
                                    }
                                }
                                else if (vehType.Contains("18T"))
                                {

                                    if (lvl2.UtilisationTotal == 100)
                                    {
                                        lvl3.Utilisation = ((lvl3.PalletCount / (float)tot_packs) * (float)100);
                                        lvl3.UtilisationToShow = ((lvl3.PalletCount / (float)tot_packs) * (float)100);

                                    }
                                    else
                                    {
                                        lvl3.Utilisation = ((lvl3.PalletCount / (float)14) * (float)100);
                                        lvl3.UtilisationToShow = ((lvl3.PalletCount / (float)14) * (float)100);
                                    }

                                }
                                else if (vehType.Contains("7.5T"))
                                {
                                    if (lvl2.UtilisationTotal == 100)
                                    {
                                        lvl3.Utilisation = ((lvl3.PalletCount / (float)tot_packs) * (float)100);
                                        lvl3.UtilisationToShow = ((lvl3.PalletCount / (float)tot_packs) * (float)100);
                                    }
                                    else
                                    {
                                        lvl3.Utilisation = ((lvl3.PalletCount / (float)8) * (float)100);
                                        lvl3.UtilisationToShow = ((lvl3.PalletCount / (float)8) * (float)100);
                                    }



                                }
                                lvl3.Utilisation = float.Parse(lvl3.Utilisation.ToString("0.00"));
                                lvl3.UtilisationToShow = float.Parse(lvl3.UtilisationToShow.ToString("0.00"));

                                lvl3.FixedCost = (lvl3.PalletCount / (float)lvl1.PalletTotal) * lvl1.FixedCostTotal;
                                lvl3.DriverCost = (lvl3.PalletCount / (float)lvl1.PalletTotal) * lvl1.DriversCostTotal;
                                lvl3.Fuel = (lvl3.PalletCount / (float)lvl1.PalletTotal) * lvl1.FuelTotal;
                                lvl3.Adblue = (lvl3.PalletCount / (float)lvl1.PalletTotal) * lvl1.AdblueTotal;
                                lvl3.Revenue = (lvl3.Utilisation / (float)lvl2.UtilisationTotal) * lvl2.RevenueManTotal;

                                lvl3.Profit = lvl3.Revenue - (lvl3.Fuel + lvl3.Adblue + lvl3.DriverCost + lvl3.FixedCost);
                                lvl3.TotalCost = lvl3.Fuel + lvl3.Adblue + lvl3.DriverCost + lvl3.FixedCost;

                                lvl3.Revenue = float.Parse(lvl3.Revenue.ToString("0.00"));
                                lvl3.Profit = float.Parse(lvl3.Profit.ToString("0.00"));
                                lvl3.FixedCost = float.Parse(lvl3.FixedCost.ToString("0.00"));
                                lvl3.DriverCost = float.Parse(lvl3.DriverCost.ToString("0.00"));
                                lvl3.Fuel = float.Parse(lvl3.Fuel.ToString("0.00"));
                                lvl3.Adblue = float.Parse(lvl3.Adblue.ToString("0.00"));
                                lvl3.TotalCost = float.Parse(lvl3.TotalCost.ToString("0.00"));
                                lvl3.Date = dateTime;
                                lvl3.VehCode = lvl1.VehicleCode;
                                lvl3.ManNumber = lvl2.ManifestNumber;
                                lvl3.VehType = lvl1.VehType;

                                lvl2.list.Add(lvl3);

                            }
                        }



                    }




                }


                //to get lvl2 and lvl3 in cases where there is only lvl1
                foreach (HelperClass_Level1 yy in list_lvl1)
                {
                    if (yy.list.Count == 0)
                    {
                        HelperClass_Level2 hl2 = new HelperClass_Level2();
                        yy.list.Add(hl2);

                        HelperClass_Level3 hl3 = new HelperClass_Level3();
                        hl3.DriverCost = yy.DriversCostTotal;
                        hl3.FixedCost = yy.FixedCostTotal;
                        hl3.Fuel = yy.FuelTotal;
                        hl3.Adblue = yy.AdblueTotal;
                        hl3.Revenue = yy.RevenueTotal;
                        hl3.Profit = yy.ProfitTotal;
                        hl3.TotalCost = yy.TotalCost;
                        hl3.Postcode = "       ";
                        hl3.Customer = "       ";
                        hl3.VehCode = "        ";
                        hl3.VehType = "        ";
                        hl2.list.Add(hl3);





                    }
                }







                if (report == 2)
                {
                    string calculationSheet = "CREATE TABLE [ProportionCalculations] (Veh string, Dt string, PalletCount string, PalletTotal string, FixedTotal string, DriverCostTotal string, FuelCostTotal string, AdblueTotal string, RevenueTotal string, ProfitTotal string, TotalCost string, Prop_FixedCost string, Prop_DriverCost string, Prop_FuelCost string, Prop_AdblueCost string, Prop_Profit string, Prop_TotalCost string, Manifest string, Job string) ";
                    myCon2.Open();
                    cmd = new OleDbCommand(calculationSheet);
                    cmd.Connection = myCon2;
                    cmd.ExecuteNonQuery();
                    //myCon.Close();
                    myCon2.Close();

                    string calculationSheet2 = "CREATE TABLE [DriverCostData] (Veh string, Dt string, OvertimeRate string, StandardRate string, TotalHr string, StandardHours string, OvertimeHours string, StandardEarnings string, OvertimeEarnings string, DriverTotalCost string) ";
                    myCon2.Open();
                    cmd = new OleDbCommand(calculationSheet2);
                    cmd.Connection = myCon2;
                    cmd.ExecuteNonQuery();
                    //myCon.Close();
                    myCon2.Close();



                    System.Diagnostics.Debug.WriteLine("List size =" + list_lvl1.Count);


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2, F3, F4 ) VALUES (" + "'" + "From" + "'" + "," + "'" + from.ToString() + "'" + "," + "'" + "to" + "'" + "," + "'" + to.ToString() + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12, F13 ) VALUES (" + "'" + "Manifest" + "'" + "," + "'" + "Veh Reg" + "'" + "," + "'" + "Customer" + "'" + "," + "'" + "Postcode" + "'" + "," + "'" + "Pallets" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Fuel" + "'" + "," + "'" + "Adblue Cost" + "'" + "," + "'" + "Drivers Costs" + "'" + "," + "'" + "Fixed Costs" + "'" + "," + "'" + "Total Cost" + "'" + "," + "'" + "Profit" + "'" + "," + "'" + "Utilisation" + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    myCon.Open();
                    myCon2.Open();
                    foreach (HelperClass_Level1 elem in list_lvl1)
                    {
                        /*
                        if(elem.list.Count == 0)
                        {
                            HelperClass_Level2 hl2 = new HelperClass_Level2();
                            elem.list.Add(hl2);

                            HelperClass_Level3 hl3 =new HelperClass_Level3();
                            hl3.DriverCost = elem.DriversCostTotal;
                            hl3.FixedCost = elem.FixedCostTotal;
                            hl3.Fuel = elem.FuelTotal;
                            hl3.Adblue = elem.AdblueTotal;
                            hl3.Revenue = elem.RevenueTotal;
                            hl3.Profit = elem.ProfitTotal;
                            hl3.TotalCost = elem.TotalCost;
                            hl2.list.Add(hl3);



                        }*/

                        elem.DrvCost_OvertimeRate = float.Parse(elem.DrvCost_OvertimeRate.ToString("0.00"));
                        elem.DrvCost_StandardRate = float.Parse(elem.DrvCost_StandardRate.ToString("0.00"));
                        elem.DrvCost_TotalHr = float.Parse(elem.DrvCost_TotalHr.ToString("0.00"));
                        elem.DrvCost_StandardHours = float.Parse(elem.DrvCost_StandardHours.ToString("0.00"));
                        elem.DrvCost_OvertimeHours = float.Parse(elem.DrvCost_OvertimeHours.ToString("0.00"));
                        elem.DrvCost_StandardEarnings = float.Parse(elem.DrvCost_StandardEarnings.ToString("0.00"));
                        elem.DrvCost_OvertimeEarnings = float.Parse(elem.DrvCost_OvertimeEarnings.ToString("0.00"));
                        elem.DriversCostTotal = float.Parse(elem.DriversCostTotal.ToString("0.00"));


                        string command = "INSERT INTO [DriverCostData] (Veh, Dt, OvertimeRate, StandardRate, TotalHr, StandardHours, OvertimeHours, StandardEarnings, OvertimeEarnings,  DriverTotalCost) " +
                               "VALUES (@Veh, @Dt, @OvertimeRate, @StandardRate, @TotalHr, @StandardHours, @OvertimeHours, @StandardEarnings, @OvertimeEarnings, @DriverTotalCost) ";
                        // myCon.Open();
                        cmd = new OleDbCommand(command);
                        cmd.Parameters.AddWithValue("@Veh", elem.VehicleCode.ToString());
                        cmd.Parameters.AddWithValue("@Dt", elem.Date.ToString("dd/MM/yyyy"));
                        cmd.Parameters.AddWithValue("@OvertimeRate", elem.DrvCost_OvertimeRate.ToString());
                        cmd.Parameters.AddWithValue("@StandardRate", elem.DrvCost_StandardRate.ToString());
                        cmd.Parameters.AddWithValue("@TotalHr", elem.DrvCost_TotalHr.ToString());
                        cmd.Parameters.AddWithValue("@StandardHours", elem.DrvCost_StandardHours.ToString());
                        cmd.Parameters.AddWithValue("@OvertimeHours", elem.DrvCost_OvertimeHours.ToString());
                        cmd.Parameters.AddWithValue("@StandardEarnings", elem.DrvCost_StandardEarnings.ToString());
                        cmd.Parameters.AddWithValue("@OvertimeEarnings", elem.DrvCost_OvertimeEarnings.ToString());
                        cmd.Parameters.AddWithValue("@DriverTotalCost", elem.DriversCostTotal.ToString());
                        cmd.Connection = myCon2;
                        cmd.ExecuteNonQuery();








                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");


                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();
                        //myCon.Close();


                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + " " + "'" + "," + "'" + elem.VehicleCode + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + elem.PalletTotal + "'" + "," + "'" + elem.RevenueTotal + "'" + "," + "'" + elem.FuelTotal + "'" + "," + "'" + elem.AdblueTotal + "'" + "," + "'" + elem.DriversCostTotal + "'" + "," + "'" + elem.FixedCostTotal + "'" + "," + "'" + elem.TotalCost + "'" + "," + "'" + elem.ProfitTotal + "'" + "," + "'" + " " + "'" + ")");
                        //myCon.Open();
                        cmd.Connection = myCon;
                        cmd.ExecuteNonQuery();
                        //myCon.Close();


                        foreach (HelperClass_Level2 elem2 in elem.list)
                        {

                            cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + elem2.ManifestNumber + "'" + ")");
                            // myCon.Open();

                            cmd.Connection = myCon;
                            cmd.ExecuteNonQuery();
                            // myCon.Close();

                            foreach (HelperClass_Level3 elem3 in elem2.list)
                            {
                                elem.FixedCostTotal = float.Parse(elem.FixedCostTotal.ToString("0.00"));
                                elem.DriversCostTotal = float.Parse(elem.DriversCostTotal.ToString("0.00"));
                                elem.FuelTotal = float.Parse(elem.FuelTotal.ToString("0.00"));
                                elem.AdblueTotal = float.Parse(elem.AdblueTotal.ToString("0.00"));
                                elem.RevenueTotal = float.Parse(elem.RevenueTotal.ToString("0.00"));
                                elem.ProfitTotal = float.Parse(elem.ProfitTotal.ToString("0.00"));
                                elem.TotalCost = float.Parse(elem.TotalCost.ToString("0.00"));
                                elem3.FixedCost = float.Parse(elem3.FixedCost.ToString("0.00"));
                                elem3.DriverCost = float.Parse(elem3.DriverCost.ToString("0.00"));
                                elem3.Fuel = float.Parse(elem3.Fuel.ToString("0.00"));
                                elem3.Adblue = float.Parse(elem3.Adblue.ToString("0.00"));
                                elem3.Profit = float.Parse(elem3.Profit.ToString("0.00"));
                                elem3.TotalCost = float.Parse(elem3.TotalCost.ToString("0.00"));



                                command = "INSERT INTO [ProportionCalculations] (Veh, Dt, PalletCount, PalletTotal, FixedTotal,  DriverCostTotal, FuelCostTotal, AdblueTotal, RevenueTotal, ProfitTotal, TotalCost, Prop_FixedCost, Prop_DriverCost, Prop_FuelCost, Prop_AdblueCost, Prop_Profit, Prop_TotalCost, Manifest, Job) " +
                               "VALUES (@Veh, @Dt, @PalletCount, @PalletTotal, @FixedTotal, @DriverCostTotal, @FuelCostTotal, @AdblueTotal, @RevenueTotal, @ProfitTotal, @TotalCost, @Prop_FixedCost, @Prop_DriverCost, @Prop_FuelCost, @Prop_AdblueCost, @Prop_Profit, @Prop_TotalCost, @Manifest, @Job) ";
                                // myCon.Open();
                                cmd = new OleDbCommand(command);
                                cmd.Parameters.AddWithValue("@Veh", elem.VehicleCode.ToString());
                                cmd.Parameters.AddWithValue("@Dt", elem.Date.ToString("dd/MM/yyyy"));
                                cmd.Parameters.AddWithValue("@PalletCount", elem3.PalletCount.ToString());
                                cmd.Parameters.AddWithValue("@PalletTotal", elem.PalletTotal.ToString());

                                cmd.Parameters.AddWithValue("@FixedTotal", elem.FixedCostTotal.ToString());
                                cmd.Parameters.AddWithValue("@DriverCostTotal", elem.DriversCostTotal.ToString());
                                cmd.Parameters.AddWithValue("@FuelCostTotal", elem.FuelTotal.ToString());
                                cmd.Parameters.AddWithValue("@AdblueTotal", elem.AdblueTotal.ToString());
                                cmd.Parameters.AddWithValue("@RevenueTotal", elem.RevenueTotal.ToString());
                                cmd.Parameters.AddWithValue("@ProfitTotal", elem.ProfitTotal.ToString());
                                cmd.Parameters.AddWithValue("@TotalCost", elem.TotalCost.ToString());
                                cmd.Parameters.AddWithValue("@Prop_Fixed", elem3.FixedCost.ToString());
                                cmd.Parameters.AddWithValue("@Prop_DriverCost", elem3.DriverCost.ToString());
                                cmd.Parameters.AddWithValue("@Prop_FuelCost", elem3.Fuel.ToString());
                                cmd.Parameters.AddWithValue("@Prop_AdblueCost", elem3.Adblue.ToString());
                                //cmd.Parameters.AddWithValue("@Prop_RevenueCost", elem3.Revenue);
                                cmd.Parameters.AddWithValue("@Prop_Profit", elem3.Profit.ToString());
                                cmd.Parameters.AddWithValue("@Prop_TotalCost", elem3.TotalCost.ToString());
                                cmd.Parameters.AddWithValue("@Manifest", elem3.ManNumber.ToString());
                                cmd.Parameters.AddWithValue("@Job", elem3.BookingNumber.ToString());




                                cmd.Connection = myCon2;
                                cmd.ExecuteNonQuery();


                                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + "  " + "'" + "," + "'" + "  " + "'" + "," + "'" + elem3.Customer + "'" + "," + "'" + elem3.Postcode + "'" + "," + "'" + elem3.PalletCount + "'" + "," + "'" + elem3.Revenue + "'" + "," + "'" + elem3.Fuel + "'" + "," + "'" + elem3.Adblue + "'" + "," + "'" + elem3.DriverCost + "'" + "," + "'" + elem3.FixedCost + "'" + "," + "'" + elem3.TotalCost + "'" + "," + "'" + elem3.Profit + "'" + "," + "'" + elem3.UtilisationToShow + "'" + ")");

                                //myCon.Open();
                                cmd.Connection = myCon;

                                cmd.ExecuteNonQuery();
                                // myCon.Close();


                            }

                        }



                    }
                    myCon2.Close();
                    myCon.Close();

                }
                else if (report == 4)
                {

                    List<HelperClass_Level3> newList = new List<HelperClass_Level3>();


                    foreach (HelperClass_Level1 hlp1 in list_lvl1)
                    {
                        /*
                        if(hlp1.list.Count == 0)
                        {
                             HelperClass_Level3 toAdd = new HelperClass_Level3();


                                    toAdd.Revenue = hlp1.RevenueTotal;
                                    toAdd.Postcode = " ";
                                    toAdd.Customer = " ";
                                    toAdd.Count++;
                                    //toAdd.PltCnt =+hlp3.PalletCount;
                                    toAdd.Profit = hlp1.ProfitTotal;
                                    toAdd.TotalCost = hlp1.TotalCost;

                                    toAdd.PalletCount = hlp1.PalletTotal;
                                    toAdd.BookingStatus = -1;
                                    toAdd.CountSuccess = 0;
                                    toAdd.CountCancelled = 0;
                                    toAdd.CountOther = 0;
                                    toAdd.CountRefused = 0;

                            newList.Add(toAdd);
                            toAdd.ForCalculations.Add(toAdd);


                        }
                        else*/
                        {


                            foreach (HelperClass_Level2 hlp2 in hlp1.list)
                            {
                                List<HelperClass_Level3> tmp_ToAdd = new List<HelperClass_Level3>();

                                foreach (HelperClass_Level3 hlp3 in hlp2.list)
                                {
                                    string psotcode = "";

                                    if (hlp3.Postcode.Length <= 0)
                                    {
                                        psotcode = "UNKNOWN";
                                    }
                                    else
                                    {
                                        psotcode += hlp3.Postcode[0];
                                        psotcode += hlp3.Postcode[1];
                                        psotcode += hlp3.Postcode[2];
                                        psotcode += hlp3.Postcode[3];
                                    }



                                    if (newList.Count == 0)
                                    {

                                        HelperClass_Level3 toAdd = new HelperClass_Level3();


                                        toAdd.Revenue = hlp3.Revenue;
                                        toAdd.Postcode = psotcode;
                                        toAdd.Count++;
                                        //toAdd.PltCnt =+hlp3.PalletCount;
                                        toAdd.Profit = hlp3.Profit;
                                        toAdd.TotalCost = hlp3.TotalCost;

                                        toAdd.PalletCount = hlp3.PalletCount;
                                        toAdd.BookingStatus = hlp3.BookingStatus;
                                        toAdd.CountSuccess = hlp3.CountSuccess;
                                        toAdd.CountCancelled = hlp3.CountCancelled;
                                        toAdd.CountOther = hlp3.CountOther;
                                        toAdd.CountRefused = hlp3.CountRefused;

                                        if (hlp3.BookingStatus == 1)
                                        {
                                            toAdd.CountOther++;
                                        }
                                        else if (hlp3.BookingStatus == 2)
                                        {
                                            toAdd.CountRefused++;
                                        }
                                        else if (hlp3.BookingStatus == 3)
                                        {
                                            toAdd.CountSuccess++;
                                        }
                                        else if (hlp3.BookingStatus == 7)
                                        {
                                            toAdd.CountCancelled++;
                                        }
                                        else
                                        {
                                            toAdd.CountOther++;
                                        }

                                        newList.Add(toAdd);
                                        toAdd.ForCalculations.Add(hlp3);
                                        //break;


                                    }
                                    else
                                    {



                                        int c = 0;
                                        foreach (HelperClass_Level3 elem in newList)
                                        {


                                            if (c == newList.Count - 1 && !elem.Postcode.Equals(psotcode))
                                            {

                                                HelperClass_Level3 toAdd = new HelperClass_Level3();

                                                //string psotcode = "";
                                                //elem.PltCnt = +hlp3.PalletCount;
                                                toAdd.Revenue = hlp3.Revenue;
                                                toAdd.Postcode = psotcode;
                                                toAdd.Count++;
                                                toAdd.Profit = hlp3.Profit;
                                                toAdd.TotalCost = hlp3.TotalCost;
                                                toAdd.PalletCount = hlp3.PalletCount;

                                                if (hlp3.BookingStatus == 1)
                                                {
                                                    toAdd.CountOther++;
                                                }
                                                else if (hlp3.BookingStatus == 2)
                                                {
                                                    toAdd.CountRefused++;
                                                }
                                                else if (hlp3.BookingStatus == 3)
                                                {
                                                    toAdd.CountSuccess++;
                                                }
                                                else if (hlp3.BookingStatus == 7)
                                                {
                                                    toAdd.CountCancelled++;
                                                }
                                                else
                                                {
                                                    toAdd.CountOther++;
                                                }
                                                newList.Add(toAdd);
                                                toAdd.ForCalculations.Add(hlp3);
                                                break;



                                            }

                                            else if (elem.Postcode.Equals(psotcode))
                                            {
                                                //elem.PalletCount += hlp3.PalletCount;
                                                elem.Revenue += hlp3.Revenue;
                                                elem.PalletCount += hlp3.PalletCount;
                                                elem.TotalCost += hlp3.TotalCost;
                                                elem.Profit += hlp3.Profit;
                                                //elem.PltCnt = +hlp3.PalletCount;
                                                elem.Count++;
                                                elem.BookingStatus += hlp3.BookingStatus;
                                                elem.CountSuccess += hlp3.CountSuccess;
                                                elem.CountCancelled += hlp3.CountCancelled;
                                                elem.CountOther += hlp3.CountOther;
                                                elem.CountRefused += hlp3.CountRefused;


                                                if (hlp3.BookingStatus == 1)
                                                {
                                                    elem.CountOther++;
                                                }
                                                else if (hlp3.BookingStatus == 2)
                                                {
                                                    elem.CountRefused++;
                                                }
                                                else if (hlp3.BookingStatus == 3)
                                                {
                                                    elem.CountSuccess++;
                                                }
                                                else if (hlp3.BookingStatus == 7)
                                                {
                                                    elem.CountCancelled++;
                                                }
                                                else
                                                {
                                                    elem.CountOther++;
                                                }

                                                elem.ForCalculations.Add(hlp3);
                                                break;


                                            }
                                            c++;





                                        }

                                    }



                                }






                            }




                        }









                    }

                    newList = newList.OrderByDescending(x => x.Profit).ToList();





                    string calculationSheet = "CREATE TABLE [Calculations] (Postcode string,Client string, Profit string, Revenue string, TotalCost string, DriverCost string, FixedCost string, Fuel string, BookingNumber string, Manifest string, BookingStatus string, PalletCount string, Dt string) ";
                    myCon2.Open();
                    cmd = new OleDbCommand(calculationSheet);
                    cmd.Connection = myCon2;
                    cmd.ExecuteNonQuery();
                    //myCon.Close();
                    myCon2.Close();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2, F3, F4 ) VALUES (" + "'" + "From" + "'" + "," + "'" + from.ToString() + "'" + "," + "'" + "to" + "'" + "," + "'" + to.ToString() + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10 ) VALUES (" + "'" + "Postcode" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Total Costs" + "'" + "," + "'" + "Profit " + "'" + "," + "'" + "No of pallets" + "'" + "," + "'" + "no of jobs for customer" + "'" + "," + "'" + "% Success" + "'" + "," + "'" + "% Refused" + "'" + "," + "'" + "% Cancelled" + "'" + "," + "'" + "% Other reasons" + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    float t_profit = 0;
                    float t_Revenue = 0;
                    float t_TotalCost = 0;
                    float t_DriverCost = 0;
                    float t_FixedCost = 0;
                    float t_Fuel = 0;


                    myCon.Open();
                    foreach (HelperClass_Level3 hh in newList)
                    {
                        t_profit += hh.Profit;
                        t_Revenue += hh.Revenue;
                        t_TotalCost += hh.TotalCost;
                        t_DriverCost += hh.DriverCost;
                        t_FixedCost += hh.FixedCost;
                        t_Fuel += hh.Fuel;

                        float prc_succ = ((float)hh.CountSuccess / (float)hh.Count) * 100;
                        float prc_ref = ((float)hh.CountRefused / (float)hh.Count) * 100;
                        float prc_cancel = (float)(hh.CountCancelled / (float)hh.Count) * 100;
                        float prc_other = ((float)hh.CountOther / (float)hh.Count) * 100;

                        prc_succ = float.Parse(prc_succ.ToString("0.00"));
                        prc_ref = float.Parse(prc_ref.ToString("0.00"));
                        prc_cancel = float.Parse(prc_cancel.ToString("0.00"));
                        prc_other = float.Parse(prc_other.ToString("0.00"));

                        //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + hh.Customer + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Counter + "'" + "," + "'" + ((hh.CountSuccess/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountRefused/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountCancelled/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountOther/hh.Counter)*100)  + "'" + ")");
                        //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + hh.Customer + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Counter + "'" + "," + "'" + hh.CountSuccess  + "'" + "," + "'" + hh.CountRefused   + "'" + "," + "'" + hh.CountCancelled  + "'" + "," + "'" + hh.CountOther  + "'" + ")");
                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10) VALUES (" + "'" + hh.Postcode + "'" + "," + "'" + hh.Revenue + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.PalletCount + "'" + "," + "'" + hh.Count + "'" + "," + "'" + prc_succ + "'" + "," + "'" + prc_ref + "'" + "," + "'" + prc_cancel + "'" + "," + "'" + prc_other + "'" + ")");

                        // myCon.Open();
                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();
                        //myCon.Close();

                    }

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10) VALUES (" + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + ")");

                    // myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10) VALUES (" + "'" + "Totals" + "'" + "," + "'" + t_Revenue + "'" + "," + "'" + t_TotalCost + "'" + "," + "'" + t_profit + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + ")");

                    // myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();




                    myCon.Close();

                    myCon2.Open();

                    foreach (HelperClass_Level3 elem in newList)
                    {



                        foreach (HelperClass_Level3 hlp in elem.ForCalculations)
                        {

                            string command = "INSERT INTO [Calculations] (Postcode, Client, Profit, Revenue, TotalCost, DriverCost, FixedCost, Fuel, BookingNumber, Manifest, BookingStatus, PalletCount, Dt) " +
                                "VALUES (@Postcode, @Client, @Profit, @Revenue, @TotalCost, @DriverCost, @FixedCost, @Fuel, @BookingNumber, @Manifest, @BookingStatus, @PalletCount, @Dt) ";
                            // myCon.Open();
                            cmd = new OleDbCommand(command);
                            cmd.Parameters.AddWithValue("@Postcode", hlp.Postcode.ToString());
                            cmd.Parameters.AddWithValue("@Client", hlp.Customer.ToString());
                            cmd.Parameters.AddWithValue("@Profit", hlp.Profit.ToString());
                            cmd.Parameters.AddWithValue("@Rvenue", hlp.Revenue.ToString());
                            cmd.Parameters.AddWithValue("@TotalCost", hlp.TotalCost.ToString());
                            cmd.Parameters.AddWithValue("@DriverCost", hlp.DriverCost.ToString());
                            cmd.Parameters.AddWithValue("@FixedCost", hlp.FixedCost.ToString());
                            cmd.Parameters.AddWithValue("@Fuel", hlp.Fuel.ToString());
                            cmd.Parameters.AddWithValue("@Manifest", hlp.ManNumber.ToString());
                            cmd.Parameters.AddWithValue("@BookingNumber", hlp.BookingNumber.ToString());
                            cmd.Parameters.AddWithValue("@BookingStatus", hlp.BookingStatus.ToString());
                            cmd.Parameters.AddWithValue("@PalletCount", hlp.PalletCount.ToString());
                            cmd.Parameters.AddWithValue("@Date", hlp.Date.ToString("dd/MM/yyyy"));




                            cmd.Connection = myCon2;
                            cmd.ExecuteNonQuery();

                        }



                    }



                    myCon2.Close();

                }
                else if (report == 3)
                {

                    List<HelperClass_Level3> newList = new List<HelperClass_Level3>();


                    foreach (HelperClass_Level1 hlp1 in list_lvl1)
                    {
                        /*
                        if (hlp1.list.Count == 0)
                        {
                            HelperClass_Level3 toAdd = new HelperClass_Level3();


                            toAdd.Revenue = hlp1.RevenueTotal;
                            toAdd.Customer = " ";
                            toAdd.Postcode = " ";
                            toAdd.Count++;
                            //toAdd.PltCnt =+hlp3.PalletCount;
                            toAdd.Profit = hlp1.ProfitTotal;
                            toAdd.TotalCost = hlp1.TotalCost;

                            toAdd.PalletCount = hlp1.PalletTotal;
                            toAdd.BookingStatus = -1;
                            toAdd.CountSuccess = 0;
                            toAdd.CountCancelled = 0;
                            toAdd.CountOther = 0;
                            toAdd.CountRefused = 0;

                            newList.Add(toAdd);
                            toAdd.ForCalculations.Add(toAdd);


                        }
                        else*/
                        {

                            foreach (HelperClass_Level2 hlp2 in hlp1.list)
                            {


                                foreach (HelperClass_Level3 hlp3 in hlp2.list)
                                {


                                    if (newList.Count == 0)
                                    {

                                        HelperClass_Level3 toAdd = new HelperClass_Level3();


                                        toAdd.Revenue = hlp3.Revenue;
                                        toAdd.Customer = hlp3.Customer;
                                        toAdd.Count++;
                                        //toAdd.PltCnt =+hlp3.PalletCount;
                                        toAdd.Profit = hlp3.Profit;
                                        toAdd.TotalCost = hlp3.TotalCost;

                                        toAdd.PalletCount = hlp3.PalletCount;
                                        toAdd.BookingStatus = hlp3.BookingStatus;
                                        toAdd.CountSuccess = hlp3.CountSuccess;
                                        toAdd.CountCancelled = hlp3.CountCancelled;
                                        toAdd.CountOther = hlp3.CountOther;
                                        toAdd.CountRefused = hlp3.CountRefused;

                                        if (hlp3.BookingStatus == 1)
                                        {
                                            toAdd.CountOther++;
                                        }
                                        else if (hlp3.BookingStatus == 2)
                                        {
                                            toAdd.CountRefused++;
                                        }
                                        else if (hlp3.BookingStatus == 3)
                                        {
                                            toAdd.CountSuccess++;
                                        }
                                        else if (hlp3.BookingStatus == 7)
                                        {
                                            toAdd.CountCancelled++;
                                        }
                                        else
                                        {
                                            toAdd.CountOther++;
                                        }

                                        newList.Add(toAdd);
                                        toAdd.ForCalculations.Add(hlp3);
                                        //break;


                                    }
                                    else
                                    {



                                        int c = 0;
                                        foreach (HelperClass_Level3 elem in newList)
                                        {


                                            if (c == newList.Count - 1 && !elem.Customer.Equals(hlp3.Customer))
                                            {

                                                HelperClass_Level3 toAdd = new HelperClass_Level3();

                                                //string psotcode = "";
                                                //elem.PltCnt = +hlp3.PalletCount;
                                                toAdd.Revenue = hlp3.Revenue;
                                                toAdd.Customer = hlp3.Customer;
                                                toAdd.Count++;
                                                toAdd.Profit = hlp3.Profit;
                                                toAdd.TotalCost = hlp3.TotalCost;
                                                toAdd.PalletCount = hlp3.PalletCount;

                                                if (hlp3.BookingStatus == 1)
                                                {
                                                    toAdd.CountOther++;
                                                }
                                                else if (hlp3.BookingStatus == 2)
                                                {
                                                    toAdd.CountRefused++;
                                                }
                                                else if (hlp3.BookingStatus == 3)
                                                {
                                                    toAdd.CountSuccess++;
                                                }
                                                else if (hlp3.BookingStatus == 7)
                                                {
                                                    toAdd.CountCancelled++;
                                                }
                                                else
                                                {
                                                    toAdd.CountOther++;
                                                }
                                                newList.Add(toAdd);
                                                toAdd.ForCalculations.Add(hlp3);
                                                break;



                                            }

                                            else if (elem.Customer.Equals(hlp3.Customer))
                                            {
                                                //elem.PalletCount += hlp3.PalletCount;
                                                elem.Revenue += hlp3.Revenue;
                                                elem.PalletCount += hlp3.PalletCount;
                                                elem.TotalCost += hlp3.TotalCost;
                                                elem.Profit += hlp3.Profit;
                                                //elem.PltCnt = +hlp3.PalletCount;
                                                elem.Count++;
                                                elem.BookingStatus += hlp3.BookingStatus;
                                                elem.CountSuccess += hlp3.CountSuccess;
                                                elem.CountCancelled += hlp3.CountCancelled;
                                                elem.CountOther += hlp3.CountOther;
                                                elem.CountRefused += hlp3.CountRefused;


                                                if (hlp3.BookingStatus == 1)
                                                {
                                                    elem.CountOther++;
                                                }
                                                else if (hlp3.BookingStatus == 2)
                                                {
                                                    elem.CountRefused++;
                                                }
                                                else if (hlp3.BookingStatus == 3)
                                                {
                                                    elem.CountSuccess++;
                                                }
                                                else if (hlp3.BookingStatus == 7)
                                                {
                                                    elem.CountCancelled++;
                                                }
                                                else
                                                {
                                                    elem.CountOther++;
                                                }

                                                elem.ForCalculations.Add(hlp3);
                                                break;


                                            }
                                            c++;





                                        }

                                    }



                                }


                            }


                        }








                    }

                    string calculationSheet = "CREATE TABLE [Calculations] (Postcode string,Client string, Profit string, Revenue string, TotalCost string, DriverCost string, FixedCost string, Fuel string, BookingNumber string, Manifest string, BookingStatus string, PalletCount string, Dt string) ";
                    myCon2.Open();
                    cmd = new OleDbCommand(calculationSheet);
                    cmd.Connection = myCon2;
                    cmd.ExecuteNonQuery();
                    //myCon.Close();
                    myCon2.Close();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2, F3, F4 ) VALUES (" + "'" + "From" + "'" + "," + "'" + from.ToString() + "'" + "," + "'" + "to" + "'" + "," + "'" + to.ToString() + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10 ) VALUES (" + "'" + "Customer" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Total Costs" + "'" + "," + "'" + "Profit " + "'" + "," + "'" + "No of pallets" + "'" + "," + "'" + "no of jobs for customer" + "'" + "," + "'" + "% Success" + "'" + "," + "'" + "% Refused" + "'" + "," + "'" + "% Cancelled" + "'" + "," + "'" + "% Other reasons" + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    newList = newList.OrderByDescending(x => x.Profit).ToList();

                    float t_profit = 0;
                    float t_Revenue = 0;
                    float t_TotalCost = 0;
                    float t_DriverCost = 0;
                    float t_FixedCost = 0;
                    float t_Fuel = 0;

                    myCon.Open();
                    foreach (HelperClass_Level3 hh in newList)
                    {

                        t_profit += hh.Profit;
                        t_Revenue += hh.Revenue;
                        t_TotalCost += hh.TotalCost;
                        t_DriverCost += hh.DriverCost;
                        t_FixedCost += hh.FixedCost;
                        t_Fuel += hh.Fuel;

                        float prc_succ = ((float)hh.CountSuccess / (float)hh.Count) * 100;
                        float prc_ref = ((float)hh.CountRefused / (float)hh.Count) * 100;
                        float prc_cancel = (float)(hh.CountCancelled / (float)hh.Count) * 100;
                        float prc_other = ((float)hh.CountOther / (float)hh.Count) * 100;

                        prc_succ = float.Parse(prc_succ.ToString("0.00"));
                        prc_ref = float.Parse(prc_ref.ToString("0.00"));
                        prc_cancel = float.Parse(prc_cancel.ToString("0.00"));
                        prc_other = float.Parse(prc_other.ToString("0.00"));

                        //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + hh.Customer + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Counter + "'" + "," + "'" + ((hh.CountSuccess/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountRefused/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountCancelled/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountOther/hh.Counter)*100)  + "'" + ")");
                        //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + hh.Customer + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Counter + "'" + "," + "'" + hh.CountSuccess  + "'" + "," + "'" + hh.CountRefused   + "'" + "," + "'" + hh.CountCancelled  + "'" + "," + "'" + hh.CountOther  + "'" + ")");
                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10) VALUES (" + "'" + hh.Customer + "'" + "," + "'" + hh.Revenue + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.PalletCount + "'" + "," + "'" + hh.Count + "'" + "," + "'" + prc_succ + "'" + "," + "'" + prc_ref + "'" + "," + "'" + prc_cancel + "'" + "," + "'" + prc_other + "'" + ")");

                        // myCon.Open();
                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();
                        //myCon.Close();








                    }
                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10) VALUES (" + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + ")");

                    // myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10) VALUES (" + "'" + "Totals" + "'" + "," + "'" + t_Revenue + "'" + "," + "'" + t_TotalCost + "'" + "," + "'" + t_profit + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + ")");

                    // myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();

                    myCon.Close();

                    myCon2.Open();

                    foreach (HelperClass_Level3 elem in newList)
                    {



                        foreach (HelperClass_Level3 hlp in elem.ForCalculations)
                        {

                            string command = "INSERT INTO [Calculations] (Postcode, Client, Profit, Revenue, TotalCost, DriverCost, FixedCost, Fuel, BookingNumber, Manifest, BookingStatus, PalletCount, Dt) " +
                                "VALUES (@Postcode, @Client, @Profit, @Revenue, @TotalCost, @DriverCost, @FixedCost, @Fuel, @BookingNumber, @Manifest, @BookingStatus, @PalletCount, @Dt) ";
                            // myCon.Open();
                            cmd = new OleDbCommand(command);
                            cmd.Parameters.AddWithValue("@Postcode", hlp.Postcode.ToString());
                            cmd.Parameters.AddWithValue("@Client", hlp.Customer.ToString());
                            cmd.Parameters.AddWithValue("@Profit", hlp.Profit.ToString());
                            cmd.Parameters.AddWithValue("@Rvenue", hlp.Revenue.ToString());
                            cmd.Parameters.AddWithValue("@TotalCost", hlp.TotalCost.ToString());
                            cmd.Parameters.AddWithValue("@DriverCost", hlp.DriverCost.ToString());
                            cmd.Parameters.AddWithValue("@FixedCost", hlp.FixedCost.ToString());
                            cmd.Parameters.AddWithValue("@Fuel", hlp.Fuel.ToString());
                            cmd.Parameters.AddWithValue("@Manifest", hlp.ManNumber.ToString());
                            cmd.Parameters.AddWithValue("@BookingNumber", hlp.BookingNumber.ToString());
                            cmd.Parameters.AddWithValue("@BookingStatus", hlp.BookingStatus.ToString());
                            cmd.Parameters.AddWithValue("@PalletCount", hlp.PalletCount.ToString());
                            cmd.Parameters.AddWithValue("@Date", hlp.Date.ToString("dd/MM/yyyy"));




                            cmd.Connection = myCon2;
                            cmd.ExecuteNonQuery();

                        }



                    }
                    myCon2.Close();

                }
                else if (report == 1)
                {
                    List<HelperClass_Level1> newList = new List<HelperClass_Level1>();
                    List<HelperClass_Level3> newList3 = new List<HelperClass_Level3>();

                    string calculationSheet = "CREATE TABLE [Calculations] (VehCode string, FixedCost string, Adblue string, DriverCost string, Revenue string, TotalCost string, Profit string, Manifest string, Job string, Dt string) ";
                    myCon2.Open();
                    cmd = new OleDbCommand(calculationSheet);
                    cmd.Connection = myCon2;
                    cmd.ExecuteNonQuery();
                    //myCon.Close();
                    myCon2.Close();


                    foreach (HelperClass_Level1 elem in list_lvl1)
                    {
                        /*if (elem.list.Count==0)
                        {
                             HelperClass_Level3 toAdd = new HelperClass_Level3();
                                    toAdd.VehCode= elem.VehicleCode;
                                    toAdd.VehType = elem.VehType;
                                    toAdd.Profit = elem.ProfitTotal;
                                    toAdd.TotalCost = elem.TotalCost;
                                    // toAdd.TotalCost = elem.TotalCost;
                                    toAdd.Adblue = elem.AdblueTotal;
                                    toAdd.DriverCost = elem.DriversCostTotal;
                                    toAdd.FixedCost = elem.FixedCostTotal;
                                    toAdd.Fuel = elem.FuelTotal;
                                    toAdd.Revenue = elem.RevenueTotal;

                            toAdd.ForCalculations.Add(toAdd);
                                    newList3.Add(toAdd);



                        }
                        else*/
                        {

                            foreach (HelperClass_Level2 elem2 in elem.list)
                            {
                                foreach (HelperClass_Level3 elem3 in elem2.list)
                                {

                                    if (newList3.Count == 0)
                                    {
                                        HelperClass_Level3 toAdd = new HelperClass_Level3();
                                        toAdd.VehCode = elem.VehicleCode;
                                        toAdd.VehType = elem.VehType;
                                        toAdd.Profit = elem3.Profit;
                                        toAdd.TotalCost = elem3.TotalCost;
                                        // toAdd.TotalCost = elem.TotalCost;
                                        toAdd.Adblue = elem3.Adblue;
                                        toAdd.DriverCost = elem3.DriverCost;
                                        toAdd.FixedCost = elem3.FixedCost;
                                        toAdd.Fuel = elem3.Fuel;
                                        toAdd.Revenue = elem3.Revenue;

                                        toAdd.ForCalculations.Add(elem3);
                                        newList3.Add(toAdd);

                                    }
                                    else
                                    {

                                        int c = 0;
                                        foreach (HelperClass_Level3 newListElem3 in newList3)
                                        {
                                            if (elem3.VehCode.Equals(newListElem3.VehCode))
                                            {
                                                newListElem3.Profit += elem3.Profit;
                                                newListElem3.TotalCost += elem3.TotalCost;
                                                newListElem3.Adblue += elem3.Adblue;
                                                newListElem3.DriverCost += elem3.DriverCost;
                                                newListElem3.FixedCost += elem3.FixedCost;
                                                newListElem3.Fuel += elem3.Fuel;
                                                newListElem3.Revenue += elem3.Revenue;
                                                newListElem3.ForCalculations.Add(elem3);
                                                break;

                                            }
                                            else if (!elem3.VehCode.Equals(newListElem3.VehCode) && c == newList3.Count - 1)
                                            {

                                                HelperClass_Level3 toAdd = new HelperClass_Level3();
                                                toAdd.VehCode = elem.VehicleCode;
                                                toAdd.Profit = elem3.Profit;
                                                toAdd.TotalCost = elem3.TotalCost;

                                                toAdd.Adblue = elem3.Adblue;
                                                toAdd.DriverCost = elem3.DriverCost;
                                                toAdd.FixedCost = elem3.FixedCost;
                                                toAdd.Fuel = elem3.Fuel;
                                                toAdd.Revenue = elem3.Revenue;

                                                toAdd.ForCalculations.Add(elem3);
                                                newList3.Add(toAdd);
                                                break;

                                            }

                                            c++;



                                        }
                                    }







                                }




                            }



                            if (newList.Count == 0)
                            {
                                HelperClass_Level1 toAdd = new HelperClass_Level1();
                                toAdd.VehicleCode = elem.VehicleCode;
                                toAdd.VehType = elem.VehType;
                                toAdd.ProfitTotal = elem.ProfitTotal;
                                toAdd.TotalCost = elem.TotalCost;
                                // toAdd.TotalCost = elem.TotalCost;
                                toAdd.AdblueTotal = elem.AdblueTotal;
                                toAdd.DriversCostTotal = elem.DriversCostTotal;
                                toAdd.FixedCostTotal = elem.FixedCostTotal;
                                toAdd.FuelTotal = elem.FuelTotal;
                                toAdd.RevenueTotal = elem.RevenueTotal;

                                newList.Add(toAdd);

                            }
                            else
                            {

                                int c = 0;
                                foreach (HelperClass_Level1 newListElem in newList)
                                {
                                    if (elem.VehicleCode.Equals(newListElem.VehicleCode))
                                    {
                                        newListElem.ProfitTotal += elem.ProfitTotal;
                                        newListElem.TotalCost += elem.TotalCost;
                                        newListElem.AdblueTotal += elem.AdblueTotal;
                                        newListElem.DriversCostTotal += elem.DriversCostTotal;
                                        newListElem.FixedCostTotal += elem.FixedCostTotal;
                                        newListElem.FuelTotal += elem.FuelTotal;
                                        newListElem.RevenueTotal += elem.RevenueTotal;
                                        break;

                                    }
                                    else if (!elem.VehicleCode.Equals(newListElem.VehicleCode) && c == newList.Count - 1)
                                    {

                                        HelperClass_Level1 toAdd = new HelperClass_Level1();
                                        toAdd.VehType = elem.VehType;
                                        toAdd.ProfitTotal = elem.ProfitTotal;
                                        toAdd.TotalCost = elem.TotalCost;
                                        toAdd.VehicleCode = elem.VehicleCode;
                                        toAdd.AdblueTotal = elem.AdblueTotal;
                                        toAdd.DriversCostTotal = elem.DriversCostTotal;
                                        toAdd.FixedCostTotal = elem.FixedCostTotal;
                                        toAdd.FuelTotal = elem.FuelTotal;
                                        toAdd.RevenueTotal = elem.RevenueTotal;


                                        newList.Add(toAdd);
                                        break;

                                    }
                                    c++;



                                }
                            }



                        }



                    }

                    float t_profit = 0;
                    float t_Revenue = 0;
                    float t_TotalCost = 0;
                    float t_DriverCost = 0;
                    float t_FixedCost = 0;
                    float t_Fuel = 0;
                    float t_adblue = 0;

                    myCon2.Open();
                    foreach (HelperClass_Level3 hlpelem in newList3)
                    {


                        foreach (HelperClass_Level3 calculationElem in hlpelem.ForCalculations)
                        {


                            string command = "INSERT INTO [Calculations] (VehCode, FixedCost, Adblue, DriverCost, Revenue, TotalCost, Profit, Manifest, Job, Dt) " +
                                "VALUES (@VehCode, @FixedCost, @Adblue, @DriverCost, @Revenue, @TotalCost, @Profit, @Manifest, @Job, @Dt) ";
                            // myCon.Open();
                            cmd = new OleDbCommand(command);
                            cmd.Parameters.AddWithValue("@VehCode", hlpelem.VehCode.ToString());
                            cmd.Parameters.AddWithValue("@FixedCost", calculationElem.FixedCost.ToString());
                            cmd.Parameters.AddWithValue("@Adblue", calculationElem.Adblue.ToString());
                            cmd.Parameters.AddWithValue("@DriverCost", calculationElem.DriverCost.ToString());
                            cmd.Parameters.AddWithValue("@Revenue", calculationElem.Revenue.ToString());
                            cmd.Parameters.AddWithValue("@TotalCost", calculationElem.TotalCost.ToString());
                            cmd.Parameters.AddWithValue("@Profit", calculationElem.Profit.ToString());
                            cmd.Parameters.AddWithValue("@Manifest", calculationElem.ManNumber.ToString());
                            cmd.Parameters.AddWithValue("@Job", calculationElem.BookingNumber.ToString());
                            cmd.Parameters.AddWithValue("@Dt", calculationElem.Date.ToString("dd/MM/yyyy"));
                            //cmd.Parameters.AddWithValue("@Dt", calculationElem.Date.ToString("dd/MM/yyyy"));
                            cmd.Connection = myCon2;
                            cmd.ExecuteNonQuery();



                        }

                    }
                    myCon2.Close();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2, F3, F4 ) VALUES (" + "'" + "From" + "'" + "," + "'" + from.ToString() + "'" + "," + "'" + "to" + "'" + "," + "'" + to.ToString() + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + "VehicleCode" + "'" + "," + "'" + "Fixed Cost" + "'" + "," + "'" + "Adblue cost" + "'" + "," + "'" + "Fuel Cost" + "'" + "," + "'" + "Driver Cost" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Total Cost" + "'" + "," + "'" + "Profit" + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    myCon.Open();
                    foreach (HelperClass_Level3 i in newList3)
                    {
                        t_profit += i.Profit;
                        t_Revenue += i.Revenue;
                        t_TotalCost += i.TotalCost;
                        t_DriverCost += i.DriverCost;
                        t_FixedCost += i.FixedCost;
                        t_Fuel += i.Fuel;
                        t_adblue += i.Adblue;


                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1, F2, F3, F4, F5, F6, F7, F8 ) VALUES ( " + "'" + i.VehCode + "'" + "," + i.FixedCost + "," + i.Adblue + "," + i.Fuel + "," + i.DriverCost + "," + i.Revenue + "," + i.TotalCost + "," + i.Profit + ")");


                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();



                    }
                    myCon.Close();

                    myCon.Open();

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES ( " + "'" + "  " + "'" + ")");


                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1, F2, F3, F4, F5, F6, F7, F8 ) VALUES ( " + "'" + "Totals" + "'" + "," + t_FixedCost + "," + t_adblue + "," + t_Fuel + "," + t_DriverCost + "," + t_Revenue + "," + t_TotalCost + "," + t_profit + ")");


                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();

                    myCon.Close();





                }
                else if (report == 0)
                {
                    List<HelperClass_Level1> newList = new List<HelperClass_Level1>();
                    List<HelperClass_Level3> newList3 = new List<HelperClass_Level3>();

                    string calculationSheet = "CREATE TABLE [Calculations] (VehType string , VehCode string, FixedCost string, Adblue string, DriverCost string, Revenue string, TotalCost string, Profit string, Manifest string, Job string, Dt string) ";
                    myCon2.Open();
                    cmd = new OleDbCommand(calculationSheet);
                    cmd.Connection = myCon2;
                    cmd.ExecuteNonQuery();
                    //myCon.Close();
                    myCon2.Close();


                    foreach (HelperClass_Level1 elem in list_lvl1)
                    {

                        if (elem.list.Count == 0)
                        {
                            HelperClass_Level3 toAdd = new HelperClass_Level3();
                            toAdd.VehCode = elem.VehicleCode;
                            toAdd.VehType = elem.VehType;
                            toAdd.Profit = elem.ProfitTotal;
                            toAdd.TotalCost = elem.TotalCost;
                            // toAdd.TotalCost = elem.TotalCost;
                            toAdd.Adblue = elem.AdblueTotal;
                            toAdd.DriverCost = elem.DriversCostTotal;
                            toAdd.FixedCost = elem.FixedCostTotal;
                            toAdd.Fuel = elem.FuelTotal;
                            toAdd.Revenue = elem.RevenueTotal;
                            if (newList3.Count == 0)
                            {


                                toAdd.ForCalculations.Add(toAdd);
                                newList3.Add(toAdd);

                            }
                            else
                            {
                                foreach (HelperClass_Level3 el in newList3)
                                {

                                    if (el.VehType.Equals(elem.VehType))
                                    {
                                        el.Profit += toAdd.Profit;
                                        el.TotalCost += toAdd.TotalCost;
                                        el.Adblue += toAdd.Adblue;
                                        el.DriverCost += toAdd.DriverCost;
                                        el.FixedCost += toAdd.FixedCost;
                                        el.Fuel += toAdd.Fuel;
                                        el.Revenue += toAdd.Revenue;
                                        el.ForCalculations.Add(toAdd);

                                    }


                                }
                            }




                        }
                        else
                        {

                            foreach (HelperClass_Level2 elem2 in elem.list)
                            {
                                foreach (HelperClass_Level3 elem3 in elem2.list)
                                {

                                    if (newList3.Count == 0)
                                    {
                                        HelperClass_Level3 toAdd = new HelperClass_Level3();
                                        toAdd.VehCode = elem.VehicleCode;
                                        toAdd.VehType = elem.VehType;
                                        toAdd.Profit = elem3.Profit;
                                        toAdd.TotalCost = elem3.TotalCost;
                                        // toAdd.TotalCost = elem.TotalCost;
                                        toAdd.Adblue = elem3.Adblue;
                                        toAdd.DriverCost = elem3.DriverCost;
                                        toAdd.FixedCost = elem3.FixedCost;
                                        toAdd.Fuel = elem3.Fuel;
                                        toAdd.Revenue = elem3.Revenue;

                                        toAdd.ForCalculations.Add(elem3);
                                        newList3.Add(toAdd);

                                    }
                                    else
                                    {

                                        int c = 0;
                                        foreach (HelperClass_Level3 newListElem3 in newList3)
                                        {
                                            if (elem3.VehType.Equals(newListElem3.VehType))
                                            {
                                                newListElem3.Profit += elem3.Profit;
                                                newListElem3.TotalCost += elem3.TotalCost;
                                                newListElem3.Adblue += elem3.Adblue;
                                                newListElem3.DriverCost += elem3.DriverCost;
                                                newListElem3.FixedCost += elem3.FixedCost;
                                                newListElem3.Fuel += elem3.Fuel;
                                                newListElem3.Revenue += elem3.Revenue;
                                                newListElem3.ForCalculations.Add(elem3);
                                                break;

                                            }
                                            else if (!elem3.VehType.Equals(newListElem3.VehType) && c == newList3.Count - 1)
                                            {

                                                HelperClass_Level3 toAdd = new HelperClass_Level3();
                                                toAdd.VehCode = elem.VehicleCode;
                                                toAdd.VehType = elem.VehType;
                                                toAdd.Profit = elem3.Profit;
                                                toAdd.TotalCost = elem3.TotalCost;

                                                toAdd.Adblue = elem3.Adblue;
                                                toAdd.DriverCost = elem3.DriverCost;
                                                toAdd.FixedCost = elem3.FixedCost;
                                                toAdd.Fuel = elem3.Fuel;
                                                toAdd.Revenue = elem3.Revenue;

                                                toAdd.ForCalculations.Add(elem3);
                                                newList3.Add(toAdd);
                                                break;

                                            }

                                            c++;



                                        }
                                    }







                                }




                            }



                            if (newList.Count == 0)
                            {
                                HelperClass_Level1 toAdd = new HelperClass_Level1();
                                toAdd.VehicleCode = elem.VehicleCode;
                                toAdd.VehType = elem.VehType;
                                toAdd.ProfitTotal = elem.ProfitTotal;
                                toAdd.TotalCost = elem.TotalCost;
                                // toAdd.TotalCost = elem.TotalCost;
                                toAdd.AdblueTotal = elem.AdblueTotal;
                                toAdd.DriversCostTotal = elem.DriversCostTotal;
                                toAdd.FixedCostTotal = elem.FixedCostTotal;
                                toAdd.FuelTotal = elem.FuelTotal;
                                toAdd.RevenueTotal = elem.RevenueTotal;

                                newList.Add(toAdd);

                            }
                            else
                            {

                                int c = 0;
                                foreach (HelperClass_Level1 newListElem in newList)
                                {
                                    if (elem.VehicleCode.Equals(newListElem.VehicleCode))
                                    {
                                        newListElem.ProfitTotal += elem.ProfitTotal;
                                        newListElem.TotalCost += elem.TotalCost;
                                        newListElem.AdblueTotal += elem.AdblueTotal;
                                        newListElem.DriversCostTotal += elem.DriversCostTotal;
                                        newListElem.FixedCostTotal += elem.FixedCostTotal;
                                        newListElem.FuelTotal += elem.FuelTotal;
                                        newListElem.RevenueTotal += elem.RevenueTotal;
                                        break;

                                    }
                                    else if (!elem.VehicleCode.Equals(newListElem.VehicleCode) && c == newList.Count - 1)
                                    {

                                        HelperClass_Level1 toAdd = new HelperClass_Level1();
                                        toAdd.VehType = elem.VehType;
                                        toAdd.ProfitTotal = elem.ProfitTotal;
                                        toAdd.TotalCost = elem.TotalCost;
                                        toAdd.VehicleCode = elem.VehicleCode;
                                        toAdd.AdblueTotal = elem.AdblueTotal;
                                        toAdd.DriversCostTotal = elem.DriversCostTotal;
                                        toAdd.FixedCostTotal = elem.FixedCostTotal;
                                        toAdd.FuelTotal = elem.FuelTotal;
                                        toAdd.RevenueTotal = elem.RevenueTotal;


                                        newList.Add(toAdd);
                                        break;

                                    }
                                    c++;



                                }
                            }

                        }




                    }







                    myCon2.Open();
                    foreach (HelperClass_Level3 hlpelem in newList3)
                    {
                        foreach (HelperClass_Level3 calculationElem in hlpelem.ForCalculations)
                        {




                            string command = "INSERT INTO [Calculations] (VehType, VehCode, FixedCost, Adblue, DriverCost, Revenue, TotalCost, Profit, Manifest, Job, Dt) " +
                                "VALUES (@VehType, @VehCode, @FixedCost, @Adblue, @DriverCost, @Revenue, @TotalCost, @Profit, @Manifest, @Job, @Dt) ";
                            // myCon.Open();
                            cmd = new OleDbCommand(command);
                            cmd.Parameters.AddWithValue("@VehType", hlpelem.VehType.ToString());
                            cmd.Parameters.AddWithValue("@VehCode", calculationElem.VehCode.ToString());
                            cmd.Parameters.AddWithValue("@FixedCost", calculationElem.FixedCost.ToString());
                            cmd.Parameters.AddWithValue("@Adblue", calculationElem.Adblue.ToString());
                            cmd.Parameters.AddWithValue("@DriverCost", calculationElem.DriverCost.ToString());
                            cmd.Parameters.AddWithValue("@Revenue", calculationElem.Revenue.ToString());
                            cmd.Parameters.AddWithValue("@TotalCost", calculationElem.TotalCost.ToString());
                            cmd.Parameters.AddWithValue("@Profit", calculationElem.Profit.ToString());
                            cmd.Parameters.AddWithValue("@Manifest", calculationElem.ManNumber.ToString());
                            cmd.Parameters.AddWithValue("@Job", calculationElem.BookingNumber.ToString());
                            cmd.Parameters.AddWithValue("@Dt", calculationElem.Date.ToString("dd/MM/yyyy"));
                            //cmd.Parameters.AddWithValue("@Dt", calculationElem.Date.ToString("dd/MM/yyyy"));
                            cmd.Connection = myCon2;
                            cmd.ExecuteNonQuery();



                        }

                    }
                    myCon2.Close();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2, F3, F4 ) VALUES (" + "'" + "From" + "'" + "," + "'" + from.ToString() + "'" + "," + "'" + "to" + "'" + "," + "'" + to.ToString() + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + "VehicleCode" + "'" + "," + "'" + "Fixed Cost" + "'" + "," + "'" + "Adblue cost" + "'" + "," + "'" + "Fuel Cost" + "'" + "," + "'" + "Driver Cost" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Total Cost" + "'" + "," + "'" + "Profit" + "'" + ")");

                    myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    myCon.Close();

                    float t_profit = 0;
                    float t_Revenue = 0;
                    float t_TotalCost = 0;
                    float t_DriverCost = 0;
                    float t_FixedCost = 0;
                    float t_Fuel = 0;
                    float t_adblue = 0;

                    myCon.Open();

                    foreach (HelperClass_Level3 i in newList3)
                    {

                        t_profit += i.Profit;
                        t_Revenue += i.Revenue;
                        t_TotalCost += i.TotalCost;
                        t_DriverCost += i.DriverCost;
                        t_FixedCost += i.FixedCost;
                        t_Fuel += i.Fuel;
                        t_adblue += i.Adblue;

                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1, F2, F3, F4, F5, F6, F7, F8 ) VALUES ( " + "'" + i.VehType + "'" + "," + i.FixedCost + "," + i.Adblue + "," + i.Fuel + "," + i.DriverCost + "," + i.Revenue + "," + i.TotalCost + "," + i.Profit + ")");


                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();



                    }
                    myCon.Close();

                    myCon.Open();

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES ( " + "'" + "  " + "'" + ")");


                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();


                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1, F2, F3, F4, F5, F6, F7, F8 ) VALUES ( " + "'" + "Totals" + "'" + "," + t_FixedCost + "," + t_adblue + "," + t_Fuel + "," + t_DriverCost + "," + t_Revenue + "," + t_TotalCost + "," + t_profit + ")");


                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();


                    myCon.Close();






                }
                /////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////////////////////////////////////////////////
                /////////////////////////////////////////////////////////////////////////////////////////




                myCon.Close();
                myCon2.Close();
                myCon.Dispose();
                myCon2.Dispose();

            }
        }
        public class HelperClass_Level1
        {
            
            public float DrvCost_OvertimeEarnings { set; get; }

            public float DrvCost_StandardEarnings { set; get; }

            public float DrvCost_TotalHr { set; get; }
            public float DrvCost_StandardRate { set; get; }

            public float DrvCost_OvertimeHours { set; get; }

            public float DrvCost_StandardHours { set; get; }

            public float DrvCost_OvertimeRate { set; get; }

            public int Alternative_PalletTotal { set; get; }

            public string VehType  {set;get;}
            public DateTime Date { set; get; }
            public String VehicleCode { set; get; }
            public int PalletTotal { set; get; }

            public float RevenueTotal { set; get; }

            public float ProfitTotal { set; get; }
            public float FuelTotal { set; get; }

            public float DriversCostTotal { set; get; }

            public float FixedCostTotal { set; get; }

            public float AdblueTotal { set; get; }
            public float TotalCost { set; get; }           

            public List<HelperClass_Level2> list = new List<HelperClass_Level2>();
        }

        public class HelperClass_Level2
        {
            public int ManifestNumber { set; get; }

            public float RevenueManTotal { set; get; }
            public float UtilisationTotal { set; get; }

            public int VehicleCapacity { set; get; }
        
            public List<HelperClass_Level3> list = new List<HelperClass_Level3>();
        }

        public class HelperClass_Level3
        {
            public HelperClass_Level3()
            {
                PltCnt = 0;
                Count = 0;
                CountSuccess = 0;
                CountCancelled = 0;
                CountRefused = 0;
                CountOther = 0;
                ForCalculations = new List<HelperClass_Level3>();



            }
            public int PltCnt {set;get;}

            public string VehType { set; get; }

            public List<HelperClass_Level3> ForCalculations;

            public int ManNumber { set; get; }

            public string VehCode { set; get; }
            public int BookingNumber { set; get; }
            public int PalletCount { set; get; }

            public string Postcode { set; get; }

            public string Customer { set; get; }

            public int BookingStatus { set; get; }

            public float Revenue { set; get; }

            public float Adblue { set; get; }

            public float Fuel { set; get; }

            public float DriverCost { set; get; }

            public float Profit { set; get; }

            public float TotalCost { set; get; }

            public float FixedCost { set; get; }

            public int Count { set; get; }

            public float Utilisation { set; get; }
            public float UtilisationToShow { set; get; }

            public int CountSuccess { set; get; }
            public int CountCancelled { set; get; }

            public int CountRefused { set; get; }

            public int CountOther { set; get; }

            public DateTime Date { set; get; }




        }


        public void Processing_BasedOnManReport(string from, string to, int report)
        {
            //description

            List<UtilisationReportHelperClass> listUtilData = new List<UtilisationReportHelperClass>();
      
            Utilisation_DataDB utilDataDB = new Utilisation_DataDB();
            utilDataDB.DateFrom = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            utilDataDB.DateTo = DateTime.ParseExact(to, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            DataTable utilData = utilDataDB.SelectAllRowsBetweenDates();


            DataTable uniqueMan = utilDataDB.SelectUniqueManifestsBetweenDates3();


            //tmp, it is a mess
            DataView dv = uniqueMan.DefaultView;
            dv.Sort = "Man_Veh_Code ASC";
            DataTable dt = dv.ToTable();
            uniqueMan = dt;


            if (uniqueMan != null && uniqueMan.Rows.Count > 0)
            {
                string currentVehicle = "";


                for (int i = 0; i <= uniqueMan.Rows.Count - 1; i++)
                {
                    UtilisationReportHelperClass utl = new UtilisationReportHelperClass();

                    float tot_utilisation = 0;
                    float prop_profit = 0;
                    float tot_profit = 0;
                    float tot_adblue = 0;
                    float fixedCost = 0;
                    float driverCost = 0;
                    float prop_cost = 0;
                    float tot_cost = 0;




                    utilDataDB.Man_Number = Int32.Parse(uniqueMan.Rows[i]["Man_Number"].ToString());

                    DataTable manJobs = utilDataDB.SelectUniqueRowBetweenDatesByManifest();
                    if (manJobs != null && manJobs.Rows.Count > 0)
                    {

                        int man = Int32.Parse(manJobs.Rows[0]["Man_Number"].ToString());
                        string vehicle = manJobs.Rows[0]["Man_Veh_Code"].ToString();
                        float tot_revenue = float.Parse(manJobs.Rows[0]["Man_Total_Revenue"].ToString());
                        int tot_packs = Int32.Parse(manJobs.Rows[0]["Man_Total_Packs"].ToString());
                        int job_nbr = Int32.Parse(manJobs.Rows[0]["Bkg_Number"].ToString());
                        int veho = Int32.Parse(manJobs.Rows[0]["Veh"].ToString());

                        string vehType = "Error";
                        VehicleDb vehDb = new VehicleDb();
                        vehDb.Id = veho;
                        DataTable vehDT = vehDb.SelectVehicleById();

                        string d = manJobs.Rows[0]["Man_Date_Drv"].ToString();
                        DateTime date = DateTime.ParseExact(d, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                        Middle_Layer.VTRN_Data fixedDb = new Middle_Layer.VTRN_Data();
                        fixedDb.Veh = veho;
                        fixedDb.Vtrn_Date_Driver = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                        DataTable fixedDT = fixedDb.SelectUsingVehAndDate();

                        if (fixedDT != null && fixedDT.Rows.Count > 0)
                        {
                            fixedCost = float.Parse(fixedDT.Rows[0]["Vtrn_Monies"].ToString());
                        }

                        if (vehDT != null)
                        {
                            vehType = vehDT.Rows[0]["Type"].ToString();
                            if (vehType.Contains("Unit"))
                            {
                                utl.VehicleCapacity = 26;
                                tot_utilisation = ((tot_packs / (float)26) * (float)100);

                                if (tot_utilisation > 100)
                                {
                                    tot_utilisation = 100;

                                }

                            }
                            else if (vehType.Contains("18T"))
                            {
                                utl.VehicleCapacity = 14;
                                tot_utilisation = ((tot_packs / (float)14) * (float)100);

                                if (tot_utilisation > 100)
                                {
                                    tot_utilisation = 100;

                                }
                            }
                            else if (vehType.Contains("7.5T"))
                            {
                                utl.VehicleCapacity = 8;
                                tot_utilisation = ((tot_packs / (float)8) * (float)100);

                                if (tot_utilisation > 100)
                                {
                                    tot_utilisation = 100;

                                }
                            }

                            tot_utilisation = float.Parse(tot_utilisation.ToString("0.00"));
                        }


                        DRV_DutyDB drvDt = new DRV_DutyDB();
                        drvDt.Date = date;
                        drvDt.VehCode = vehicle;
                        drvDt.Veh = veho;

                        DataTable drivDuty = drvDt.SelectRowsUsingDateAndVeh();

                        if (drivDuty != null && drivDuty.Rows.Count > 0)
                        {

                            int min = 0;
                            int hr = 0;
                            int minPlusHr = 0;
                            float totalHours = 0;
                            float overtimeHours = 0;
                            float standardHours = 0;
                            float overtimeEarnings = 0;
                            float standardEarnings = 0;

                            for (int a = 0; a <= drivDuty.Rows.Count - 1; a++)
                            {
                                DateTime time = DateTime.ParseExact(drivDuty.Rows[a]["Duty_Time"].ToString(), "HH:mm:ss", CultureInfo.InvariantCulture);
                                min += time.Minute;
                                hr += time.Hour;

                                minPlusHr = hr * 60 + min;
                                totalHours = (float)minPlusHr / 60;


                                if (overtimeHours > 0)
                                {
                                    overtimeEarnings = overtimeHours * float.Parse(drivDuty.Rows[a]["_drivers_Overtime_Rate"].ToString());
                                    standardHours = totalHours - overtimeHours;
                                    standardEarnings = standardHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings + overtimeEarnings;
                                }
                                else
                                {
                                    standardEarnings = totalHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings;
                                }

                                driverCost = float.Parse(driverCost.ToString("0.00"));
                            }
                        }

                        CostingDB cst = new CostingDB();
                        cst.Veh = veho;
                        cst.Date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);


                        DataTable cstRowDT = cst.SelectRowByVehAndDate();

                        float tot_fuel = 0;
                        float tot_fuel_used = 0;
                        float diesel_cost = 0;
                        if (cstRowDT != null && cstRowDT.Rows.Count > 0)
                        {
                            tot_fuel_used = float.Parse(cstRowDT.Rows[0]["Total_Fuel_Used"].ToString());
                            diesel_cost = float.Parse(cstRowDT.Rows[0]["Diesel_Cost_per_l"].ToString());

                            tot_fuel_used = float.Parse(tot_fuel_used.ToString("0.00"));
                            diesel_cost = float.Parse(diesel_cost.ToString("0.00"));


                            tot_fuel = diesel_cost * tot_fuel_used;
                            tot_fuel = float.Parse(tot_fuel.ToString("0.00"));

                            float adblue_L = 0;
                            float adblue_cost = 0;

                            adblue_L = float.Parse(cstRowDT.Rows[0]["Approximate_Adblue_L"].ToString());
                            adblue_cost = float.Parse(cstRowDT.Rows[0]["Approximate_Addblue_Cost"].ToString());

                            adblue_L = float.Parse(adblue_L.ToString("0.00"));
                            adblue_cost = float.Parse(adblue_cost.ToString("0.00"));

                            tot_adblue = adblue_L * adblue_cost;
                        }

                        tot_profit = tot_revenue - (tot_fuel + tot_adblue + driverCost + fixedCost);
                        tot_cost = tot_fuel + tot_adblue + driverCost + fixedCost;

                        tot_cost = float.Parse(tot_cost.ToString("0.0"));
                        tot_profit = float.Parse(tot_profit.ToString("0.00"));
                        tot_adblue = float.Parse(tot_adblue.ToString("0.00"));

                        utl.ManifestNumber = man;
                        utl.VehCode = vehicle;
                        utl.PalletCount = tot_packs;
                        utl.Revenue = tot_revenue;
                        utl.Fuel = tot_fuel;
                        utl.AddBlueCost = tot_adblue;
                        utl.DriverCost = driverCost;
                        utl.FixedCost = fixedCost;
                        utl.TotalCost = tot_cost;
                        utl.Profit = tot_profit;
                        utl.Utilisation = tot_utilisation;

                        listUtilData.Add(utl);

                        DataTable jobs = utilDataDB.SelectUniqueJobsBetweenDatesAndByManifest();
                        if (jobs != null && jobs.Rows.Count > 0)
                        {

                            for (int y = 0; y <= jobs.Rows.Count - 1; y++)
                            {
                                UtilisationReportHelperClass utilElem = new UtilisationReportHelperClass();

                                utilDataDB.Bkg_Number = Int32.Parse(jobs.Rows[y]["Bkg_Number"].ToString());
                                DataTable finalrw = utilDataDB.SelectUniqueRowsBetweenDatesAndByJob();


                                string postcode = finalrw.Rows[0]["Cons_Delivery_Postcode"].ToString();
                                int status = Int32.Parse(finalrw.Rows[0]["Bkg_Status"].ToString());

                                int job = Int32.Parse(finalrw.Rows[0]["Bkg_Number"].ToString());
                                string customer = finalrw.Rows[0]["Bkg_Customer_Code"].ToString();
                                int pallets = Int32.Parse(finalrw.Rows[0]["Bkg_Cons_Packs"].ToString());
                                float revenue = float.Parse(finalrw.Rows[0]["Bkg_Cons_Price"].ToString());

                                revenue = float.Parse(revenue.ToString("0.00"));


                                float prop_utilisation = 0;
                                if (finalrw != null && finalrw.Rows.Count > 0)
                                {

                                    if (vehType.Contains("Unit"))
                                    {
                                        if (tot_utilisation == 100)
                                        {
                                            prop_utilisation = ((pallets / (float)tot_packs) * (float)100);
                                        }
                                        else
                                        {
                                            prop_utilisation = ((pallets / (float)26) * (float)100);
                                        }
                                    }
                                    else if (vehType.Contains("18T"))
                                    {

                                        if (tot_utilisation == 100)
                                        {
                                            prop_utilisation = ((pallets / (float)tot_packs) * (float)100);
                                        }
                                        else
                                        {
                                            prop_utilisation = ((pallets / (float)14) * (float)100);
                                        }

                                    }
                                    else if (vehType.Contains("7.5T"))
                                    {
                                        if (tot_utilisation == 100)
                                        {
                                            prop_utilisation = ((pallets / (float)tot_packs) * (float)100);
                                        }
                                        else
                                        {
                                            prop_utilisation = ((pallets / (float)8) * (float)100);
                                        }



                                    }
                                    prop_utilisation = float.Parse(prop_utilisation.ToString("0.00"));

                                    float prop_fixedCost = (prop_utilisation / (float)tot_utilisation) * fixedCost;
                                    float prop_driverCost = (prop_utilisation / (float)tot_utilisation) * driverCost;
                                    float prop_fuelCost = (prop_utilisation / (float)tot_utilisation) * tot_fuel;
                                    float prop_adblue = (prop_utilisation / (float)tot_utilisation) * tot_adblue;

                                    prop_profit = revenue - (prop_fuelCost + prop_adblue + prop_driverCost + prop_fixedCost);
                                    prop_cost = prop_fuelCost + prop_adblue + prop_driverCost + prop_fixedCost;

                                    revenue = float.Parse(revenue.ToString("0.00"));
                                    prop_profit = float.Parse(prop_profit.ToString("0.00"));
                                    prop_fixedCost = float.Parse(prop_fixedCost.ToString("0.00"));
                                    prop_driverCost = float.Parse(prop_driverCost.ToString("0.00"));
                                    prop_fuelCost = float.Parse(prop_fuelCost.ToString("0.00"));
                                    prop_adblue = float.Parse(prop_adblue.ToString("0.00"));
                                    prop_cost = float.Parse(prop_cost.ToString("0.00"));


                                    utilElem.Customer = customer;
                                    utilElem.PostCode = postcode;
                                    utilElem.PalletCount = pallets;
                                    utilElem.Revenue = revenue;
                                    utilElem.Fuel = prop_fuelCost;
                                    utilElem.AddBlueCost = prop_adblue;
                                    utilElem.DriverCost = prop_driverCost;
                                    utilElem.FixedCost = prop_fixedCost;
                                    utilElem.TotalCost = prop_cost;
                                    utilElem.Profit = prop_profit;
                                    utilElem.Utilisation = prop_utilisation;
                                    utilElem.Status = status;

                                    utl.list.Add(utilElem);

                                }
                            }
                        }
                    }
                }
            }


            VehicleDb vh = new VehicleDb();
            DataTable allVehivles = vh.SelectAllVehicles();


            foreach (DataRow elem in allVehivles.Rows)
            {
                string code = elem["Code"].ToString();
                int palletCount = 0;
                float revenue = 0;
                float fuel = 0;
                float adblue = 0;
                float driverCost = 0;
                float totalCost = 0;
                float profit = 0;
                float fixedCost = 0;

                foreach (UtilisationReportHelperClass e in listUtilData)
                {
                    if (e.VehCode.Equals(code))
                    {
                        driverCost += e.DriverCost;
                        profit += e.Profit;
                        fixedCost += e.FixedCost;
                        revenue += e.Revenue;
                        adblue += e.AddBlueCost;
                        totalCost += e.TotalCost;
                        fuel += e.Fuel;
                        palletCount += e.PalletCount;
                    }

                }


                foreach (UtilisationReportHelperClass e in listUtilData)
                {
                    if (e.VehCode.Equals(code))
                    {
                        e.SummedPalletCount = palletCount;
                        e.SummedTotalCost = totalCost;
                        e.SummedAdblue = adblue;
                        e.SummedDriverCost = driverCost;
                        e.SummedProfit = profit;
                        e.SummedFixed = fixedCost;
                        e.SummedFuel = fuel;
                        e.SummedRevenue = revenue;

                    }

                }

                foreach (UtilisationReportHelperClass e in listUtilData)
                {
                    foreach (UtilisationReportHelperClass o in e.list)
                    {


                        o.FixedCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.FixedCost;
                        o.FixedCost = float.Parse(o.FixedCost.ToString("0.00"));

                        o.Fuel = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.Fuel;
                        o.Fuel = float.Parse(o.Fuel.ToString("0.00"));

                        o.AddBlueCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.AddBlueCost;
                        o.AddBlueCost = float.Parse(o.AddBlueCost.ToString("0.00"));

                        o.DriverCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.DriverCost;
                        o.DriverCost = float.Parse(o.DriverCost.ToString("0.00"));

                        o.TotalCost = ((float)o.PalletCount / (float)e.SummedPalletCount) * e.TotalCost;
                        o.TotalCost = float.Parse(o.TotalCost.ToString("0.00"));

                        o.Profit = o.Revenue - o.TotalCost;

                    }

                }






            }


            if(report == 2)
            {
                
                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2, F3, F4 ) VALUES (" + "'" + "From" + "'" + "," + "'" + from.ToString() + "'" + "," + "'" + "to" + "'" + "," + "'" + to.ToString() + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close();


                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12, F13 ) VALUES (" + "'" + "Manifest" + "'" + "," + "'" + "Veh Reg" + "'" + "," + "'" + "Customer" + "'" + "," + "'" + "Postcode" + "'" + "," + "'" + "Pallets" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Fuel" + "'" + "," + "'" + "Adblue Cost" + "'" + "," + "'" + "Drivers Costs" + "'" + "," + "'" + "Fixed Costs" + "'" + "," + "'" + "Total Cost" + "'" + "," + "'" + "Profit" + "'" + "," + "'" + "Utilisation" + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close();

                string currentveh = "";
                bool flag = false;
                myCon.Open();
                foreach (UtilisationReportHelperClass u in listUtilData)
                {

                    if (currentveh.Equals(u.VehCode))
                    {
                        u.VehCode = "";
                        //fixedCost = 0;
                        flag = true;
                    }
                    else
                    {
                        flag = false;
                        currentveh = u.VehCode;

                    }

                    
                    if (!flag)
                    {


                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "________________________________________________________________________________________________________________________________________________" + "'" + ")");


                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();
                        //myCon.Close();


                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + " " + "'" + "," + "'" + u.VehCode + "'" + "," + "'" + " " + "'" + "," + "'" + " " + "'" + "," + "'" + u.SummedPalletCount + "'" + "," + "'" + u.SummedRevenue + "'" + "," + "'" + u.Fuel + "'" + "," + "'" + u.AddBlueCost + "'" + "," + "'" + u.DriverCost + "'" + "," + "'" + u.FixedCost + "'" + "," + "'" + u.TotalCost + "'" + "," + "'" + (u.SummedRevenue - u.TotalCost) + "'" + "," + "'" + " " + "'" + ")");
                        //myCon.Open();
                        cmd.Connection = myCon;
                        cmd.ExecuteNonQuery();
                        //myCon.Close();

                    }
                    else
                    {

                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + "                                                                                           " + "'" + ")");
                        // myCon.Open();
                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();
                        //myCon.Close();

                    }

                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1 ) VALUES (" + "'" + u.ManifestNumber + "'" + ")");
                    // myCon.Open();

                    cmd.Connection = myCon;
                    cmd.ExecuteNonQuery();
                    // myCon.Close();


                    foreach (UtilisationReportHelperClass y in u.list)
                    {

                        cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8,F9,F10,F11,F12,F13 ) VALUES (" + "'" + "  " + "'" + "," + "'" + "  " + "'" + "," + "'" + y.Customer + "'" + "," + "'" + y.PostCode + "'" + "," + "'" + y.PalletCount + "'" + "," + "'" + y.Revenue + "'" + "," + "'" + y.Fuel + "'" + "," + "'" + y.AddBlueCost + "'" + "," + "'" + y.DriverCost + "'" + "," + "'" + y.FixedCost + "'" + "," + "'" + y.TotalCost + "'" + "," + "'" + y.Profit + "'" + "," + "'" + y.Utilisation + "'" + ")");

                        //myCon.Open();
                        cmd.Connection = myCon;

                        cmd.ExecuteNonQuery();
                        // myCon.Close();

                    }

                    
                }
                myCon.Close();

            }
            else if (report == 3)
            {
                List<UtilisationReportHelperClass> tl = new List<UtilisationReportHelperClass>();
                
                foreach(UtilisationReportHelperClass el1 in listUtilData)
                {
                    
                    foreach(UtilisationReportHelperClass el2 in el1.list)
                    {

                        if (tl.Count == 0)
                        {

                            UtilisationReportHelperClass hl = new UtilisationReportHelperClass();
                            hl.Customer = el2.Customer;
                            hl.Profit = el2.Profit;
                            hl.TotalCost = el2.TotalCost;
                            hl.Counter++;

                            if (el2.Status == 1)
                            {
                                hl.CountOther++;
                            }
                            else if (el2.Status == 2)
                            {
                                hl.CountRefused++;
                            }
                            else if (el2.Status == 3)
                            {
                                hl.CountSuccess++;
                            }
                            else if (el2.Status == 7)
                            {
                                hl.CountCancelled++;
                            }
                            else
                            {
                                hl.CountOther++;
                            }

                            tl.Add(hl);
                            
                        }
                        else
                        {




                            int p = 0;
                            foreach (UtilisationReportHelperClass el3 in tl)
                            {


                                if (p == tl.Count - 1 && !el3.Customer.Equals(el2.Customer))
                                {
                                    UtilisationReportHelperClass hl = new UtilisationReportHelperClass();
                                    hl.Customer = el2.Customer;
                                    hl.Profit = el2.Profit;
                                    hl.TotalCost = el2.TotalCost;
                                    hl.Counter++;

                                    if (el2.Status == 1)
                                    {
                                        hl.CountOther++;
                                    }
                                    else if (el2.Status == 2)
                                    {
                                        hl.CountRefused++;
                                    }
                                    else if (el2.Status == 3)
                                    {
                                        hl.CountSuccess++;
                                    }
                                    else if (el2.Status == 7)
                                    {
                                        hl.CountCancelled++;
                                    }
                                    else
                                    {
                                        hl.CountOther++;
                                    }

                                    tl.Add(hl);

                                    break;
                                }
                                else if (el3.Customer.Equals(el2.Customer))
                                {

                                    el3.TotalCost += el2.TotalCost;
                                    el3.Profit += el2.Profit;
                                    el3.Counter++;

                                    if (el2.Status == 1)
                                    {
                                        el3.CountOther++;
                                    }
                                    else if (el2.Status == 2)
                                    {
                                        el3.CountRefused++;
                                    }
                                    else if (el2.Status == 3)
                                    {
                                        el3.CountSuccess++;
                                    }
                                    else if (el2.Status == 7)
                                    {
                                        el3.CountCancelled++;
                                    }
                                    else
                                    {
                                        el3.CountOther++;
                                    }
                                    break;
                                }
                                p++;
                            }
                        }                                           
                    }
                }

                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2, F3, F4 ) VALUES (" + "'" + "From" + "'" + "," + "'" + from.ToString() + "'" + "," + "'" + "to" + "'" + "," + "'" + to.ToString() + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close();


                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + "Customer" + "'" + "," + "'" + "Profit " + "'" + "," + "'" + "Total Cost" + "'" + "," + "'" + "no of jobs for customer" + "'" + "," + "'" + "% Success" + "'" + "," + "'" + "% Refused" + "'" + "," + "'" + "% Cancelled" + "'" + "," + "'" + "% Other reasons" + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close();

                myCon.Open();
                foreach (UtilisationReportHelperClass hh in tl)
                {
                    float prc_succ = ((float)hh.CountSuccess / (float) hh.Counter) * 100;
                    float prc_ref = ((float)hh.CountRefused / (float)hh.Counter) * 100;
                    float prc_cancel =(float)(hh.CountCancelled / (float)hh.Counter) * 100;
                    float prc_other = ((float)hh.CountOther / (float)hh.Counter) * 100;

                    prc_succ = float.Parse(prc_succ.ToString("0.00"));
                    prc_ref = float.Parse(prc_ref.ToString("0.00"));
                    prc_cancel = float.Parse(prc_cancel.ToString("0.00"));
                    prc_other = float.Parse(prc_other.ToString("0.00"));

                    //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + hh.Customer + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Counter + "'" + "," + "'" + ((hh.CountSuccess/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountRefused/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountCancelled/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountOther/hh.Counter)*100)  + "'" + ")");
                    //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + hh.Customer + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Counter + "'" + "," + "'" + hh.CountSuccess  + "'" + "," + "'" + hh.CountRefused   + "'" + "," + "'" + hh.CountCancelled  + "'" + "," + "'" + hh.CountOther  + "'" + ")");
                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + hh.Customer + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Counter + "'" + "," + "'" + prc_succ  + "'" + "," + "'" + prc_ref  + "'" + "," + "'" + prc_cancel  + "'" + "," + "'" + prc_other  + "'" + ")");

                   // myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    //myCon.Close();


                }
                myCon.Close();




            }
            else if(report == 4)
            {


                List<UtilisationReportHelperClass> tl = new List<UtilisationReportHelperClass>();

                foreach (UtilisationReportHelperClass el1 in listUtilData)
                {

                    foreach (UtilisationReportHelperClass el2 in el1.list)
                    {

                        if (tl.Count == 0)
                        {

                            UtilisationReportHelperClass hl = new UtilisationReportHelperClass();
                            //hl.PostCode = el2.PostCode;
                            string postcode = "";
                            postcode += el2.PostCode[0];
                            postcode += el2.PostCode[1];
                            postcode += el2.PostCode[2];
                            postcode += el2.PostCode[3];
                            hl.PostCode = postcode;
                            

                            hl.Profit = el2.Profit;
                            hl.TotalCost = el2.TotalCost;
                            hl.Counter++;

                            if (el2.Status == 1)
                            {
                                hl.CountOther++;
                            }
                            else if (el2.Status == 2)
                            {
                                hl.CountRefused++;
                            }
                            else if (el2.Status == 3)
                            {
                                hl.CountSuccess++;
                            }
                            else if (el2.Status == 7)
                            {
                                hl.CountCancelled++;
                            }
                            else
                            {
                                hl.CountOther++;
                            }

                            tl.Add(hl);

                        }
                        else
                        {




                            int p = 0;
                            foreach (UtilisationReportHelperClass el3 in tl)
                            {
                                string postcode = "";
                                postcode += el2.PostCode[0];
                                postcode += el2.PostCode[1];
                                postcode += el2.PostCode[2];
                                postcode += el2.PostCode[3];

                                if (p == tl.Count - 1 && !el3.PostCode.Equals(postcode))
                                {
                                    UtilisationReportHelperClass hl = new UtilisationReportHelperClass();
                                    //hl.PostCode = el2.PostCode;
                                    
                                    hl.PostCode = postcode;


                                    hl.Profit = el2.Profit;
                                    hl.TotalCost = el2.TotalCost;
                                    hl.Counter++;

                                    if (el2.Status == 1)
                                    {
                                        hl.CountOther++;
                                    }
                                    else if (el2.Status == 2)
                                    {
                                        hl.CountRefused++;
                                    }
                                    else if (el2.Status == 3)
                                    {
                                        hl.CountSuccess++;
                                    }
                                    else if (el2.Status == 7)
                                    {
                                        hl.CountCancelled++;
                                    }
                                    else
                                    {
                                        hl.CountOther++;
                                    }

                                    tl.Add(hl);

                                    break;
                                }
                                else if (el3.PostCode.Equals(postcode))
                                {

                                    el3.TotalCost += el2.TotalCost;
                                    el3.Profit += el2.Profit;
                                    el3.Counter++;

                                    if (el2.Status == 1)
                                    {
                                        el3.CountOther++;
                                    }
                                    else if (el2.Status == 2)
                                    {
                                        el3.CountRefused++;


                                    }
                                    else if (el2.Status == 3)
                                    {
                                        el3.CountSuccess++;
                                    }
                                    else if (el2.Status == 7)
                                    {
                                        el3.CountCancelled++;
                                    }
                                    else
                                    {
                                        el3.CountOther++;
                                    }

                                    break;




                                }
                                p++;


                            }
                        }
                    }
                }

                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2, F3, F4 ) VALUES (" + "'" + "From" + "'" + "," + "'" + from.ToString() + "'" + "," + "'" + "to" + "'" + "," + "'" + to.ToString() + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close();


                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + "Postcode" + "'" + "," + "'" + "Profit " + "'" + "," + "'" + "Total Cost" + "'" + "," + "'" + "no of jobs to postcode" + "'" + "," + "'" + "% Success" + "'" + "," + "'" + "% Refused" + "'" + "," + "'" + "% Cancelled" + "'" + "," + "'" + "% Other reasons" + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close();

                myCon.Open();
                foreach (UtilisationReportHelperClass hh in tl)
                {
                    float prc_succ = ((float)hh.CountSuccess / (float)hh.Counter) * 100;
                    float prc_ref = ((float)hh.CountRefused / (float)hh.Counter) * 100;
                    float prc_cancel = (float)(hh.CountCancelled / (float)hh.Counter) * 100;
                    float prc_other = ((float)hh.CountOther / (float)hh.Counter) * 100;

                    prc_succ = float.Parse(prc_succ.ToString("0.00"));
                    prc_ref = float.Parse(prc_ref.ToString("0.00"));
                    prc_cancel = float.Parse(prc_cancel.ToString("0.00"));
                    prc_other = float.Parse(prc_other.ToString("0.00"));

                    //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + hh.Customer + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Counter + "'" + "," + "'" + ((hh.CountSuccess/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountRefused/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountCancelled/hh.Counter)*100)  + "'" + "," + "'" + ((hh.CountOther/hh.Counter)*100)  + "'" + ")");
                    //cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + hh.Customer + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Counter + "'" + "," + "'" + hh.CountSuccess  + "'" + "," + "'" + hh.CountRefused   + "'" + "," + "'" + hh.CountCancelled  + "'" + "," + "'" + hh.CountOther  + "'" + ")");
                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + hh.PostCode + "'" + "," + "'" + hh.Profit + "'" + "," + "'" + hh.TotalCost + "'" + "," + "'" + hh.Counter + "'" + "," + "'" + prc_succ + "'" + "," + "'" + prc_ref + "'" + "," + "'" + prc_cancel + "'" + "," + "'" + prc_other + "'" + ")");

                    //myCon.Open();
                    cmd.Connection = myCon;

                    cmd.ExecuteNonQuery();
                    //myCon.Close();
                }
                myCon.Close();
            }
            else if(report==1)
            {

                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2, F3, F4 ) VALUES (" + "'" + "From" + "'" + "," + "'" + from.ToString() + "'" + "," + "'" + "to" + "'" + "," + "'" + to.ToString() + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close();


                cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + "Vehicle code" + "'" + "," + "'" + "Fixed Cost " + "'" + "," + "'" + "Adblue cost " + "'" + "," + "'" + "Fuel Cost" + "'" + "," + "'" + "Driver Cost" + "'" + "," + "'" + "Total Cost" + "'" + "," + "'" + "Revenue" + "'" + "," + "'" + "Profit/Loss" + "'" + ")");

                myCon.Open();
                cmd.Connection = myCon;

                cmd.ExecuteNonQuery();
                myCon.Close();

                myCon.Open();
                
                foreach (UtilisationReportHelperClass elem in listUtilData)
                {
                    cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2,F3,F4,F5,F6,F7,F8 ) VALUES (" + "'" + elem.VehCode + "'" + "," + "'" + elem.FixedCost + "'" + "," + "'" + elem.AddBlueCost  + "'" + "," + "'" + elem.Fuel + "'" + "," + "'" + elem.DriverCost + "'" + "," + "'" + elem.TotalCost + "'" + "," + "'" + elem.Revenue + "'" + "," + "'" + elem.Profit + "'" + ")");
                    cmd.Connection = myCon;
                    cmd.ExecuteNonQuery();
                }
                myCon.Close();





            }




        }

        public void Processing_RunsToPostcodes(string from, string to)
        {
            //description
            List<RunsToPostcodesHelperClass> runs = new List<RunsToPostcodesHelperClass>();

            Utilisation_DataDB utilDataDB = new Utilisation_DataDB();
            utilDataDB.DateFrom = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
            utilDataDB.DateTo = DateTime.ParseExact(to, "dd/MM/yyyy", CultureInfo.InvariantCulture);

            DataTable utilData = utilDataDB.SelectAllRowsBetweenDates();
            DataTable uniqueMan = utilDataDB.SelectUniqueManifestsBetweenDates3();
            int runsCounter = 0;


            if (uniqueMan != null && uniqueMan.Rows.Count > 0)
            {



                for (int i = 0; i <= uniqueMan.Rows.Count - 1; i++)
                {

                    float tot_utilisation = 0;
                    float prop_profit = 0;
                    float tot_profit = 0;
                    float tot_adblue = 0;
                    float fixedCost = 0;
                    float driverCost = 0;

                    utilDataDB.Man_Number = Int32.Parse(uniqueMan.Rows[i]["Man_Number"].ToString());

                    DataTable manJobs = utilDataDB.SelectUniqueRowBetweenDatesByManifest();
                    if (manJobs != null && manJobs.Rows.Count > 0)
                    {

                        int man = Int32.Parse(manJobs.Rows[0]["Man_Number"].ToString());
                        string vehicle = manJobs.Rows[0]["Man_Veh_Code"].ToString();
                        float tot_revenue = float.Parse(manJobs.Rows[0]["Man_Total_Revenue"].ToString());
                        int tot_packs = Int32.Parse(manJobs.Rows[0]["Man_Total_Packs"].ToString());
                        int job_nbr = Int32.Parse(manJobs.Rows[0]["Bkg_Number"].ToString());
                        int veho = Int32.Parse(manJobs.Rows[0]["Veh"].ToString());

                        tot_revenue = float.Parse(tot_revenue.ToString("0.00"));


                        string vehType = "Error";
                        VehicleDb vehDb = new VehicleDb();
                        vehDb.Id = veho;
                        DataTable vehDT = vehDb.SelectVehicleById();

                        string d = manJobs.Rows[0]["Man_Date_Drv"].ToString();
                        DateTime date = DateTime.ParseExact(d, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);



                        Middle_Layer.VTRN_Data fixedDb = new Middle_Layer.VTRN_Data();
                        fixedDb.Veh = veho;
                        fixedDb.Vtrn_Date_Driver = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);

                        DataTable fixedDT = fixedDb.SelectUsingVehAndDate();

                        if (fixedDT != null && fixedDT.Rows.Count > 0)
                        {
                            fixedCost = float.Parse(fixedDT.Rows[0]["Vtrn_Monies"].ToString());
                            fixedCost = float.Parse(fixedCost.ToString("0.00"));
                        }

                        if (vehDT != null)
                        {
                            vehType = vehDT.Rows[0]["Type"].ToString();
                            if (vehType.Contains("Unit"))
                            {
                                tot_utilisation = ((tot_packs / (float)26) * (float)100);
                            }
                            else if (vehType.Contains("18T"))
                            {
                                tot_utilisation = ((tot_packs / (float)14) * (float)100);
                            }
                            else if (vehType.Contains("7.5T"))
                            {
                                tot_utilisation = ((tot_packs / (float)8) * (float)100);
                            }
                        }
                        tot_utilisation = float.Parse(tot_utilisation.ToString("0.00"));


                        DRV_DutyDB drvDt = new DRV_DutyDB();
                        drvDt.Date = date;
                        drvDt.VehCode = vehicle;
                        drvDt.Veh = veho;

                        DataTable drivDuty = drvDt.SelectRowsUsingDateAndVeh();


                        if (drivDuty != null && drivDuty.Rows.Count > 0)
                        {
                            int min = 0;
                            int hr = 0;
                            int minPlusHr = 0;
                            float totalHours = 0;
                            float overtimeHours = 0;
                            float standardHours = 0;
                            float overtimeEarnings = 0;
                            float standardEarnings = 0;

                            for (int a = 0; a <= drivDuty.Rows.Count - 1; a++)
                            {
                                DateTime time = DateTime.ParseExact(drivDuty.Rows[a]["Duty_Time"].ToString(), "HH:mm:ss", CultureInfo.InvariantCulture);
                                min += time.Minute;
                                hr += time.Hour;

                                minPlusHr = hr * 60 + min;
                                totalHours = (float)minPlusHr / 60;


                                if (overtimeHours > 0)
                                {
                                    overtimeEarnings = overtimeHours * float.Parse(drivDuty.Rows[a]["_drivers_Overtime_Rate"].ToString());
                                    standardHours = totalHours - overtimeHours;
                                    standardEarnings = standardHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings + overtimeEarnings;
                                }
                                else
                                {
                                    standardEarnings = totalHours * float.Parse(drivDuty.Rows[a]["_drivers_Standard_Rate"].ToString());
                                    driverCost = standardEarnings;
                                }

                                driverCost = float.Parse(driverCost.ToString("0.00"));

                            }
                        }

                        CostingDB cst = new CostingDB();
                        cst.Veh = veho;
                        //--------------------------------------cst.Date = DateTime.ParseExact(from, "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        cst.Date = DateTime.ParseExact(d, "dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                        DataTable cstRowDT = cst.SelectRowByVehAndDate();

                        float tot_fuel = 0;
                        float tot_fuel_used = 0;
                        float diesel_cost = 0;
                        if (cstRowDT != null && cstRowDT.Rows.Count > 0)
                        {
                            tot_fuel_used = float.Parse(cstRowDT.Rows[0]["Total_Fuel_Used"].ToString());
                            diesel_cost = float.Parse(cstRowDT.Rows[0]["Diesel_Cost_per_l"].ToString());

                            tot_fuel_used = float.Parse(tot_fuel_used.ToString("0.00"));
                            diesel_cost = float.Parse(diesel_cost.ToString("0.00"));


                            tot_fuel = diesel_cost * tot_fuel_used;

                            float adblue_L = 0;
                            float adblue_cost = 0;

                            adblue_L = float.Parse(cstRowDT.Rows[0]["Approximate_Adblue_L"].ToString());
                            adblue_cost = float.Parse(cstRowDT.Rows[0]["Approximate_Addblue_Cost"].ToString());

                            adblue_L = float.Parse(adblue_L.ToString("0.00"));
                            adblue_cost = float.Parse(adblue_cost.ToString("0.00"));

                            tot_adblue = adblue_L * adblue_cost;
                        }

                        tot_profit = tot_revenue - (tot_fuel + tot_adblue + driverCost + fixedCost);

                        tot_profit = float.Parse(tot_profit.ToString("0.00"));

                        DataTable jobs = utilDataDB.SelectUniqueJobsBetweenDatesAndByManifest();
                        if (jobs != null && jobs.Rows.Count > 0)
                        {

                            for (int y = 0; y <= jobs.Rows.Count - 1; y++)
                            {
                                utilDataDB.Bkg_Number = Int32.Parse(jobs.Rows[y]["Bkg_Number"].ToString());
                                DataTable finalrw = utilDataDB.SelectUniqueRowsBetweenDatesAndByJob();

                                int job = Int32.Parse(finalrw.Rows[0]["Bkg_Number"].ToString());
                                string customer = finalrw.Rows[0]["Bkg_Customer_Code"].ToString();
                                int pallets = Int32.Parse(finalrw.Rows[0]["Bkg_Cons_Packs"].ToString());
                                float revenue = float.Parse(finalrw.Rows[0]["Bkg_Cons_Price"].ToString());
                                string postcode = finalrw.Rows[0]["Cons_Delivery_Postcode"].ToString();

                                //int status = -1;
                                int status = Int32.Parse( finalrw.Rows[0]["Bkg_Status"].ToString());

                                if (postcode.Length >= 4)
                                {
                                    string shortpostcode = "";

                                    shortpostcode += postcode[0];
                                    shortpostcode += postcode[1];
                                    shortpostcode += postcode[2];
                                    shortpostcode += postcode[3];

                                    postcode = shortpostcode;

                                }

                                revenue = float.Parse(revenue.ToString("0.00"));

                                float prop_utilisation = 0;
                                if (finalrw != null && finalrw.Rows.Count > 0)
                                {

                                    if (vehType.Contains("Unit"))
                                    {
                                        prop_utilisation = ((pallets / (float)26) * (float)100);
                                    }
                                    else if (vehType.Contains("18T"))
                                    {
                                        prop_utilisation = ((pallets / (float)14) * (float)100);
                                    }
                                    else if (vehType.Contains("7.5T"))
                                    {
                                        prop_utilisation = ((pallets / (float)8) * (float)100);
                                    }
                                    prop_utilisation = float.Parse(prop_utilisation.ToString("0.00"));

                                    float prop_fixedCost = (prop_utilisation / (float)tot_utilisation) * fixedCost;
                                    float prop_driverCost = (prop_utilisation / (float)tot_utilisation) * driverCost;
                                    float prop_fuelCost = (prop_utilisation / (float)tot_utilisation) * tot_fuel;
                                    float prop_adblue = (prop_utilisation / (float)tot_utilisation) * tot_adblue;

                                    prop_profit = revenue - (prop_fuelCost + prop_adblue + prop_driverCost + prop_fixedCost);


                                    revenue = float.Parse(revenue.ToString("0.00"));
                                    prop_profit = float.Parse(prop_profit.ToString("0.00"));
                                    prop_fixedCost = float.Parse(prop_fixedCost.ToString("0.00"));
                                    prop_driverCost = float.Parse(prop_driverCost.ToString("0.00"));
                                    prop_fuelCost = float.Parse(prop_fuelCost.ToString("0.00"));
                                    prop_adblue = float.Parse(prop_adblue.ToString("0.00"));

                                    int c = -1;
                                    bool flag = false;
                                    

                                    if (runs.Count == 0)
                                    {
                                        RunsToPostcodesHelperClass cl = new RunsToPostcodesHelperClass();


                                        runsCounter++;


                                        
                                        cl.CountPostcode = 1;

                                        if (status == 1)
                                        {
                                            cl.CountOther++;
                                        }
                                        else if (status == 2)
                                        {
                                            cl.CountRefused++;


                                        }
                                        else if (status == 3)
                                        {
                                            cl.CountSuccess++;
                                        }
                                        else if (status == 7)
                                        {
                                            cl.CountCancelled++;
                                        }
                                        else
                                        {
                                            cl.CountOther++;
                                        }



                                        cl.Postcode = postcode;
                                        cl.ProfitLoss += float.Parse(prop_profit.ToString("0.00"));
                                        cl.ProfitLoss = float.Parse(cl.ProfitLoss.ToString("0.00"));
                                        runs.Add(cl);

                                    }
                                    else if (runs.Count > 0)
                                    {
                                        foreach (RunsToPostcodesHelperClass elem in runs)
                                        {
                                            
                                            c++;

                                            


                                            if (elem.Postcode.Equals(postcode))
                                            {

                                                if (status == 1)
                                                {
                                                    elem.CountOther++;
                                                }
                                                else if (status == 2)
                                                {
                                                    elem.CountRefused++;


                                                }
                                                else if (status == 3)
                                                {
                                                    elem.CountSuccess++;
                                                }
                                                else if (status == 7)
                                                {
                                                    elem.CountCancelled++;
                                                }
                                                else
                                                {
                                                    elem.CountOther++;
                                                }



                                                runsCounter++;
                                                elem.CountPostcode++;
                                                elem.ProfitLoss += float.Parse(prop_profit.ToString("0.00"));
                                                elem.ProfitLoss = float.Parse(elem.ProfitLoss.ToString("0.00"));
                                                flag = false;
                                                break;
                                            }
                                            else if (c == runs.Count - 1)
                                            {
                                                
                                                flag = true;
                                            }                                          
                                        }

                                        if (flag)
                                        {
                                            
                                            runsCounter++;
                                            RunsToPostcodesHelperClass cl = new RunsToPostcodesHelperClass();

                                            if (status == 1)
                                            {
                                                cl.CountOther++;
                                            }
                                            else if (status == 2)
                                            {
                                                cl.CountRefused++;
                                            }
                                            else if (status == 3)
                                            {
                                                cl.CountSuccess++;
                                            }
                                            else if (status == 7)
                                            {
                                                cl.CountCancelled++;
                                            }
                                            else
                                            {
                                                cl.CountOther++;
                                            }

                                            cl.CountPostcode = 1;
                                            cl.Postcode = postcode;
                                            cl.ProfitLoss += float.Parse(prop_profit.ToString("0.00"));
                                            cl.ProfitLoss = float.Parse(cl.ProfitLoss.ToString("0.00"));

                                            runs.Add(cl);
                                        }




                                    }





                                }
                            }
                        }
                    }
                }
            }

            cmd = new OleDbCommand("INSERT INTO [1$] ( F1, F2) VALUES (" + "'" + "Total jobs sent:" + "'" + "," + "'" + runsCounter + "'" + ")");
            myCon.Open();
            cmd.Connection = myCon;
            cmd.ExecuteNonQuery();
            myCon.Close();

            cmd = new OleDbCommand("INSERT INTO [1$] ( F1,F2, F3, F4, F5, F6, F7, F8 ) VALUES (" + "'" + "Postcode" + "'" + "," + "'" + "ProfitLoss" + "'" + "," + "'" + "no of times at postcode" + "'"  + "," +  "'" + " % jobs to postcode" + "'" + "," +  "'" + " % success" + "'" + "," + "'" + " % refused" + "'" + "," + "'" + " % cancelled" + "'" + "," + "'" + " % other reasons" + "'" + ")");
            myCon.Open();
            cmd.Connection = myCon;
            cmd.ExecuteNonQuery();
            myCon.Close();




            foreach (RunsToPostcodesHelperClass elem in runs)
            {


                myCon.Open();
                cmd = new OleDbCommand("INSERT INTO [1$] ( F1, F2, F3, F4, F5, F6, F7, F8 ) VALUES (" + "'" + elem.Postcode.ToString() + "'" + "," + "'" + elem.ProfitLoss.ToString() + "'" + "," + "'" + elem.CountPostcode + "'"+ "," + "'" + ((elem.CountPostcode/runsCounter)*100) + "'" + "," + "'" + ((elem.CountSuccess/elem.CountPostcode)*100) + "'" + "," + "'" + ((elem.CountRefused / elem.CountPostcode) * 100) + "'"+ "," + "'" + ((elem.CountCancelled / elem.CountPostcode) * 100) + "'" + "," + "'" + ((elem.CountOther / elem.CountPostcode) * 100) + "'" + ")");
                cmd.Connection = myCon;
                cmd.ExecuteNonQuery();
                myCon.Close();
            }

        }



    }


    public class RunsToPostcodesHelperClass
    {

        public string Postcode { set; get; }

        public int CountPostcode { set; get; }


        public int CountCancelled { set; get; }

        public int CountRefused { set; get; }

        public int CountSuccess { set; get; }

        public int CountOther { set; get; }

       

        public float ProfitLoss { set; get; }

        public float TotalCost { set; get; }

        public RunsToPostcodesHelperClass()
        {
            CountSuccess = 0;
            CountOther = 0;
            CountCancelled = 0;
            CountRefused = 0;


        }


        



    }

    public class ClientNames
    {

       

        public string Client { set; get; }

        public float ProfitLoss { set; get; }

        public float TotalCost { set; get; }

        public string Postcode { set; get; }


        public int CountCancelled { set; get; }

        public int CountRefused { set; get; }

        public int CountSuccess { set; get; }

        public int CountOther { set; get; }


        public ClientNames()
        {
            Client = "";
            ProfitLoss = 0;
        }

    }

    public class UtilisationReportHelperClass
    {
        
        public int ManifestNumber { set; get; }
        public string VehCode { set; get; }
        public string Customer { set; get; }
        public string PostCode { set; get; }
        public int PalletCount { set; get; }
        public float Revenue { set; get; }
        public float Fuel { set; get; }
        public float AddBlueCost { set; get; }
        public float DriverCost { set; get; }
        public float FixedCost { set; get; }
        public float TotalCost { set; get; }
        public float Profit { set; get; }
        public float Utilisation { set; get; }

        public int Status { set; get; }

        public int SummedPalletCount { set; get; }

        public int VehicleCapacity { set; get; }

        public float SummedAdblue { set; get; }
        public float SummedDriverCost { set; get; }

        public float SummedProfit { set; get; }

        public float SummedFixed { set; get; }

        public float SummedTotalCost { set; get; }

        public float SummedFuel { set; get; }
        public float SummedRevenue { set; get; }
        public int Counter { set; get; }

        public int CountCancelled { set; get; }

        public int CountRefused { set; get; }

        public int CountSuccess { set; get; }

        public int CountOther { set; get; }

        public List<UtilisationReportHelperClass> list;

        public UtilisationReportHelperClass()
        {
            CountSuccess = 0;
            CountOther = 0;
            CountCancelled = 0;
            CountRefused = 0;
            Counter = 0;
            list = new List<UtilisationReportHelperClass>();
        }

    }


}
