using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WrkWebApp.Presentation_Layer
{
    public partial class MainMenu : System.Web.UI.Page
    {
        protected void btnProduceReports2_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Presentation_Layer/ProduceReportsScreen.aspx");
        }

        protected void btnUpdateDb_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Presentation_Layer/UpdateDB.aspx");
        }

        protected void btnUploadReports_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Presentation_Layer/UploadReportsScreen.aspx");
        }

        protected void btnAnalyse2_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Presentation_Layer/Analyse.aspx");
        }

        protected void btnUpdateDriversTable_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Presentation_Layer/UpdateDriversTable.aspx");
        }
    }
}