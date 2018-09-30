using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WrkWebApp.Middle_Layer;

namespace WrkWebApp.Presentation_Layer
{
    public partial class UpdateDriversTable : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if (flUpl.HasFile)
            {
                string sr = flUpl.PostedFile.ContentType;
                try
                {

                    if (sr == "application/vnd.ms-excel" || sr == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    {


                        string flName = "Drivers_Report";

                        flUpl.SaveAs(Server.MapPath("~/App_Data/") + flName + Path.GetExtension(flUpl.FileName));
                        lblInfo.Text = "Upload successful.";



                        string path = Server.MapPath("~/App_Data/" + flName);
                        OleDbConnection myCon = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;MAXSCANROWS=0;HDR=NO;IMEX=1\" ");
                        try
                        {
                            OleDbCommand cmd;
                            cmd = new OleDbCommand("SELECT * FROM " + "[1$]", myCon);
                            OleDbDataAdapter oleda = new OleDbDataAdapter();
                            oleda.SelectCommand = cmd;
                            DataSet ds = new DataSet();
                            DataTable dt = new DataTable();
                            oleda.Fill(dt);

                            GridView1.DataSource = dt;
                            GridView1.DataBind();

                            DriversDb dr = new DriversDb();
                            DataTable driversTb = dr.SelectAllDrivers();


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

                                    dr.First_Name = dt.Rows[i]["F2"].ToString();
                                    dr.Second_Name = dt.Rows[i]["F1"].ToString();

                                    

                                    bool exist = true;
                                    for (int a=0; a<driversTb.Rows.Count-1; a++)
                                    {
                                       string first = driversTb.Rows[a]["First_Name"].ToString();
                                       string second = driversTb.Rows[a]["Second_Name"].ToString();

                                        

                                        if (first.Equals(dr.First_Name) && second.Equals(dr.Second_Name))
                                        {
                                            exist = false;
                                            break;
                                           
                                        }

                                        if (a >= driversTb.Rows.Count - 1)
                                        {
                                            exist = true;
                                        }

                                    }

                                    if(exist)
                                    {
                                        dr.Type_Of_Employment = dt.Rows[i]["F4"].ToString();
                                        dr.drv_card = dt.Rows[i]["F7"].ToString();
                                        dr.AddRow();
                                    }
                                    
                               
                                }

                                i++;
                            }
                        }
                        catch (Exception exx)
                        {

                        }
                    }
                    else
                    {
                        lblInfo.Text = "Wrond extension";
                    }
                }
                catch
                {
                    lblInfo.Text = "Error:";
                }
                finally
                {

                }
            }
            else
            {
                lblInfo.Text = "File Upload fail";
            }
            


            


            
        }


        private void Processing_UploadingCostingReport()
        {


        }


        private void Processing_UploadingDriversDutyReport()
        {



        }


        private void Processing_Uploading_RRReportManBcgCons()
        {





        }

        private void Processing_Uploading_RRReportFixedCosts()
        {








        }

        protected void btnUploadClient_Click(object sender, EventArgs e)
        {
            if (flUplClient.HasFile)
            {
                string sr = flUplClient.PostedFile.ContentType;
                try
                {

                    if (sr == "application/vnd.ms-excel" || sr == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    {


                        string flName = "Client_Report";

                        flUpl.SaveAs(Server.MapPath("~/App_Data/") + flName + Path.GetExtension(flUpl.FileName));
                        lblInfo.Text = "Upload successful.";



                        string path = Server.MapPath("~/App_Data/" + flName);
                        OleDbConnection myCon = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";Extended Properties=\"Excel 8.0;MAXSCANROWS=0;HDR=NO;IMEX=1\" ");
                        try
                        {
                            OleDbCommand cmd;
                            cmd = new OleDbCommand("SELECT * FROM " + "[1$]", myCon);
                            OleDbDataAdapter oleda = new OleDbDataAdapter();
                            oleda.SelectCommand = cmd;
                            DataSet ds = new DataSet();
                            DataTable dt = new DataTable();
                            oleda.Fill(dt);

                            GridView1.DataSource = dt;
                            GridView1.DataBind();

                            ClientDb cl = new ClientDb();
                            DataTable clientTb = cl.SelectAllClients();


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

                                    string name = dt.Rows[i]["F1"].ToString();




                                    bool exist = true;

                                    for (int p =0; p<=clientTb.Rows.Count-1; p++)
                                    {

                                        bool result = clientTb.Rows[p]["Name"].ToString().Equals(name);

                                        if (result)
                                        {
                                            exist = false;
                                            break;
                                        }

                                        if (p >= clientTb.Rows.Count - 1)
                                        {
                                            exist = true;
                                        }

                                        if(exist)
                                        {

                                            cl.Name = name;
                                            cl.AddName();


                                        }





                                    }




                                  
                                    


                                }

                                i++;
                            }
                        }
                        catch (Exception exx)
                        {

                        }
                    }
                    else
                    {
                        lblInfo.Text = "Wrond extension";
                    }
                }
                catch
                {
                    lblInfo.Text = "Error:";
                }
                finally
                {

                }
            }
            else
            {
                lblInfo.Text = "File Upload fail";
            }







        }
    }
}