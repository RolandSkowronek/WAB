using System;


namespace wab2018
{
    public partial class wykazSkarg : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["user_id"] == null)
            {
                Server.Transfer("logowanie.aspx");
            }
            else
            {
                string rola = (string)Session["rola"];
                if (rola!="1")
                {
                    ASPxGridView1.Enabled = true;
                }
                else
                {
                    ASPxGridView1.Enabled = false;
                }
            }
        }

      
    }
}