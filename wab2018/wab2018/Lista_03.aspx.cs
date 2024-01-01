using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace wab2018
{
    public partial class Lista_03 : System.Web.UI.Page
    {
        public Class2 cl = new Class2();

        protected void Page_Load(object sender, EventArgs e)
        {
            Session["sesja"] = null;
            if (!IsPostBack)
            {
                try
                {
                    GridView1.DataBind();
                    if ((GridView1.Rows.Count >= 0) && (GridView1.SelectedIndex == -1))
                    {
                        GridView1.SelectedIndex = 0;
                    }
                }
                catch
                { }
                if (Session["user_id"] == null)
                {
                    Server.Transfer("default.aspx");
                }
            }
        }

        protected void GridView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Session["id_polecenia"] = "2";
            Panel3.Visible = true;
            LinkButton4.Visible = true;
            try
            {
                string imie = GridView1.SelectedDataKey[7].ToString();
                string nazw = GridView1.SelectedDataKey[8].ToString();
                Session["id_obciazenia"] = GridView1.SelectedDataKey[0].ToString();
                Session["id_osoby"] = GridView1.SelectedDataKey[6].ToString();
                TextBox1.Text = GridView1.SelectedDataKey[1].ToString();
                TextBox2.Text = GridView1.SelectedDataKey[2].ToString();
                string d1 = GridView1.SelectedDataKey[3].ToString();
                string d2 = GridView1.SelectedDataKey[4].ToString();
                string d3 = GridView1.SelectedDataKey[9].ToString();
                string d4 = GridView1.SelectedDataKey[5].ToString();
                cal13.ustaw_date(GridView1.SelectedDataKey[3].ToString());
                cal11.ustaw_date(GridView1.SelectedDataKey[4].ToString());
                cal12.ustaw_date(GridView1.SelectedDataKey[9].ToString());
                Lista_osob1.ustaw_osobe(GridView1.SelectedDataKey[7].ToString() + " " + GridView1.SelectedDataKey[8].ToString());
                try
                {
                    cal1.ustaw_date(GridView1.SelectedDataKey[5].ToString());
                }
                catch
                {
                    //cal1.ustaw_date(GridView1.SelectedDataKey[5].ToString());
                }

            }
            catch
            {

            }

        }

        protected void LinkButton1_Click(object sender, EventArgs e)
        {
            // zapis
            string id_polecenia = string.Empty;
            int err = 0;
            try
            {
                id_polecenia = (string)Session["id_polecenia"];
            }
            catch
            { }

            switch (id_polecenia)
            {
                case "1":
                    {
                        // nowe obcizenie
                        string id_osoby = Lista_osob1.pobierz_osobe();
                        if (string.IsNullOrEmpty(id_osoby)) id_osoby = (string)Session["id_osoby"];
                        try
                        {
                            Int64 id = Int64.Parse(id_osoby);
                            cl.dodaj_obciazenie_osoby(id_osoby, TextBox1.Text.Trim(), TextBox2.Text.Trim(), DateTime.Parse(cal13.pobierz_date()), DateTime.Parse(cal11.pobierz_date()), DateTime.Parse(cal1.pobierz_date()), (string)Session["user_id"], DateTime.Parse(cal12.pobierz_date()));

                            GridView1.DataBind();
                            Panel3.Visible = false;
                        }
                        catch
                        {
                            err = 1;
                        }
                    }
                    break;
                case "2":
                    {
                        string id_osoby = Lista_osob1.pobierz_osobe();
                        if (string.IsNullOrEmpty(id_osoby)) id_osoby = (string)Session["id_osoby"];
                        try
                        {
                            Int64 id = Int64.Parse(id_osoby);
                            string tst_zwrot = cal12.pobierz_date();
                            string tst_wprowadzenie = cal13.pobierz_date();
                            string tst_termin = cal1.pobierz_date();
                            string tst = cal12.pobierz_date();

                            cl.modyfikuj_obciazenie_osoby((string)Session["id_obciazenia"], id_osoby, TextBox1.Text.Trim(), TextBox2.Text.Trim(), DateTime.Parse(cal13.pobierz_date()), DateTime.Parse(cal11.pobierz_date()), DateTime.Parse(cal11.pobierz_date()), (string)Session["user_id"], DateTime.Parse(cal12.pobierz_date()));
                            GridView1.DataBind();
                            Panel3.Visible = false;

                        }
                        catch
                        {
                            err = 1;
                        }

                    }
                    break;
                default:
                    break;

            }
        }

        protected void LinkButton2_Click(object sender, EventArgs e)
        {
            //anulowanie
            Panel3.Visible = false;
        }

        protected void LinkButton3_Click(object sender, EventArgs e)
        {
            //usunięcie
            try
            {
                string id_obciazenia = (string)Session["id_obciazenia"];
                cl.usun_obciazenie_osoby(id_obciazenia, (string)Session["id_osoby"]);
                GridView1.DataBind();
                Panel3.Visible = false;
            }
            catch
            {

            }
        }

        protected void LinkButton1_Click1(object sender, EventArgs e)
        {
            Session["id_polecenia"] = "1";
            Panel3.Visible = true;
            LinkButton4.Visible = false;
            TextBox1.Text = "";
            TextBox2.Text = "";
            try
            {
                Session["id_obciazenia"] = null;
                Session["id_osoby"] = null;

                cal13.ustaw_date(DateTime.Now.Date.ToString("dd.MM.yyyy"));
                cal11.ustaw_date(DateTime.Now.Date.ToString("dd.MM.yyyy"));
                cal12.ustaw_date(DateTime.Now.Date.ToString("dd.MM.yyyy"));
                cal1.ustaw_date(DateTime.Now.Date.ToString("dd.MM.yyyy"));
            }
            catch
            { }
        }
    }
}