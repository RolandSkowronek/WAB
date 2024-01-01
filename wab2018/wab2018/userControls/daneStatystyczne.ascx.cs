using System;
using System.Data;

namespace wab2018
{
    public partial class daneStatystyczne : System.Web.UI.UserControl
    {
        private Class2 cl = new Class2();

        protected void Page_Load(object sender, EventArgs e)
        {
            string idBieglego = (string)Session["id_osoby"];
            if (DropDownList3.Items.Count == 0)
            {
                DropDownList3.DataBind();
                if (Session["ddl2"] != null)
                {
                    try
                    {
                        int pozycja = (int)Session["ddl2"];
                        DropDownList3.SelectedIndex = pozycja;
                    }
                    catch
                    {
                    }
                }
                zmienKwerende();
            }
        }

        protected void changeQuerry(object sender, EventArgs e)
        {
            zmienKwerende();
        }

        protected void zmienKwerende()
        {
            Session["ddl2"] = DropDownList3.SelectedIndex;
            //Label5.Text = DropDownList3.SelectedValue.ToString();
            try
            {
                string idBieglego = (string)Session["id_osoby"];
                string querry = DropDownList3.SelectedValue.ToString();
                DataTable dT = cl.tabelaStatystyczna(querry, idBieglego);
                GridView1.DataSource = null;
                GridView1.DataSourceID = null;
                GridView1.AutoGenerateColumns = true;
                GridView1.DataSource = dT;
                GridView1.DataBind();
            }
            catch
            { }
        }
    }
}