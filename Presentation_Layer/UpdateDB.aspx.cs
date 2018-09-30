using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WrkWebApp.Middle_Layer;

namespace WrkWebApp.Presentation_Layer
{
    public partial class UpdateDB : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

            if (!Page.IsPostBack)
            {

                try
                {

                    DriversDb drv = new DriversDb();
                    DataTable dt = drv.SelectAllDrivers();


                    ddlSelectDriver.DataSource = dt;
                    ddlSelectDriver.DataValueField = "Id";
                    ddlSelectDriver.DataTextField = "Second_Name";

                    ddlSelectDriver.DataBind();


                    DataTable type = drv.SelectDistinctTypes();

                    ddlSelectType.DataSource = type;
                    ddlSelectType.DataValueField = "Type_Of_Employment";

                    ddlSelectType.DataBind();


                }
                catch (Exception ex)
                {

                }
                finally
                {

                }
            }

        }

        protected void ddlStRateFor_SelectedIndexChanged(object sender, EventArgs e)
        {
            ddlStRateFor.AutoPostBack = true;

            if(ddlStRateFor.SelectedIndex == 0)
            {

                lblSelectDriver.Visible = false;
                ddlSelectDriver.Visible = false;
                

            }
            else if (ddlStRateFor.SelectedIndex == 1)
            {
                lblSelectDriver.Visible = true;
                ddlSelectDriver.Visible = true;
                

            }
        }

        protected void btnConfirm_Click(object sender, EventArgs e)
        {
            // txtOvertimeRate.Text == "";
            //  txtStandardRate.Text == "";

            if (ddlStRateFor.SelectedIndex != -1)
            {


                if (ddlStRateFor.SelectedIndex == 0)
                {

                    if (ddlSelectType.SelectedIndex != -1)
                    {



                        try
                        {


                            DriversDb drv = new DriversDb();

                            

                            DataTable dt = drv.SelectAllDrivers();

                            if (dt == null || dt.Rows.Count == 0)
                            {
                                lblInfo.Text = "Table null or empty";
                                //error here
                            }
                            else if (txtOvertimeRate.Text == "" && txtStandardRate.Text == "")
                            {
                                lblInfo.Text = "Both text boxes empty";
                                //error here, extend to validate number
                            }
                            else
                            {
                                lblInfo.Text = "entering last else";

                                //string type = ddlSelectType.Items[ddlSelectType.SelectedIndex].Text.ToString();
                                string type = ddlSelectType.SelectedItem.Value;
                                drv.Type_Of_Employment = type;
                                Response.Write(type);


                                float ovr;
                                bool isOverTimeNumeric = float.TryParse(txtOvertimeRate.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out ovr);

                                float std;
                                bool isStdNumeric = float.TryParse(txtStandardRate.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out std);


                                if (isOverTimeNumeric || isStdNumeric)
                                {
                                    
                                    

                                    if (isOverTimeNumeric)
                                        {

                                            drv.Overtime_Rate = ovr;
                                            drv.UpdateOverTimeRateBasedOnType();
                                        }

                                        if (isStdNumeric)
                                        {

                                            drv.Standard_Rate = std;
                                            drv.UpdateStandardRateBasedOnType();
                                        }


                                }
                                else
                                {
                                    lblInfo.Text = "Please, enter numeric values";

                                }

                            }




                        }
                        catch
                        {

                        }
                        finally
                        {

                        }
                    }
                    else
                    {
                        lblInfo.Text = "Please select value in: " + "Select Type";
                    }
                }
                else if (ddlStRateFor.SelectedIndex==1)
                {

                    if (ddlSelectDriver.SelectedIndex==-1)
                    {
                        lblInfo.Text = "Please select value in: " + "Select Driver";


                    }
                    else
                    {


                        float ovr;
                        bool isOverTimeNumeric = float.TryParse(txtOvertimeRate.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out ovr);

                        float std;
                        bool isStdNumeric = float.TryParse(txtStandardRate.Text, NumberStyles.Any, CultureInfo.InvariantCulture, out std);


                        if (isOverTimeNumeric || isStdNumeric)
                        {
                            DriversDb drv = new DriversDb();

                            drv.Id = Int32.Parse(ddlSelectDriver.Items[ddlSelectDriver.SelectedIndex].Value.ToString());


                            if (isOverTimeNumeric)
                            {

                                drv.Overtime_Rate = ovr;
                                drv.UpdateOverTimeRate();
                            }

                            if (isStdNumeric)
                            {

                                drv.Standard_Rate = std;
                                drv.UpdateStandardRate();
                            }

                        }
                        else
                        {
                            lblInfo.Text = "Please, Enter numeric values";
                        }

                    }


                }

            }
        }
    }
}