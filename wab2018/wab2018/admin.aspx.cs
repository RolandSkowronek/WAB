using System;
using System.Web.UI;

namespace wab2018
{
    public partial class admin : System.Web.UI.Page
    {
        public Class2 cl = new Class2();


        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                Session["sesja"] = null;
                if (Session["user_id"] == null)
                {
                    try
                    {
                        string rola = (string)Session["rola"];
                        if (rola != "2")
                        {
                            Server.Transfer("default.aspx");

                        }
                    }
                    catch
                    {
                        Server.Transfer("default.aspx");
                    }
                    Server.Transfer("default.aspx");
                }



            }
        }

        protected void hide_all()
        {
            Panel2.Visible = false;
            Panel4.Visible = false;
            Panel5.Visible = false;
            Panel6.Visible = false;
            Panel7.Visible = false;
            Panel8.Visible = false;
            TextBox8.Text = "";
            Panel3.Visible = false;
            GridView5.Enabled = true;
        }

        protected void LinkButton1_Click(object sender, EventArgs e)
        {
            hide_all();
            Panel2.Visible = true;
        }

        protected void LinkButton4_Click(object sender, EventArgs e)
        {
            hide_all();
            Panel4.Visible = true;
        }

        protected void LinkButton3_Click(object sender, EventArgs e)
        {
            hide_all();
            Panel5.Visible = true;
        }

        protected void LinkButton6_Click(object sender, EventArgs e)
        {
            // modyfikacje obciążeń
        }

        protected void GridView5_SelectedIndexChanged(object sender, EventArgs e)
        {
            // wybor specjalizacji
            TextBox8.Text = GridView5.SelectedDataKey[1].ToString();
            Session["id_specjalizacji"] = GridView5.SelectedDataKey[0].ToString();
            Session["id_polecenia"] = "1";
            DropDownList3.SelectedIndex = DropDownList3.Items.IndexOf(DropDownList3.Items.FindByValue(GridView5.SelectedDataKey[2].ToString().Trim()));
            //Label20.Text = GridView5.SelectedDataKey[2].ToString().Trim();
            if (cl.sprawdz_ilosc_specjalizacji(GridView5.SelectedDataKey[0].ToString()) != "0")
            {
                LinkButton11.Enabled = false;

            }
            else
            {
                LinkButton11.Enabled = true;
            }
        }

        protected void LinkButton8_Click(object sender, EventArgs e)
        {
            TextBox8.Text = "";
            Session["id_specjalizacji"] = null;
            Session["id_polecenia"] = "2";
            GridView5.Enabled = false;

        }

        protected void LinkButton10_Click(object sender, EventArgs e)
        {
            switch ((string)Session["id_polecenia"])
            {
                case "1":
                    {
                        TextBox8.Text = GridView5.SelectedDataKey[1].ToString();
                        Session["id_specjalizacji"] = GridView5.SelectedDataKey[0].ToString();

                    }
                    break;
                case "2":
                    {
                        TextBox8.Text = "";
                        GridView5.Enabled = true;
                    }
                    break;

                default:
                    break;
            }
        }

        protected void LinkButton2_Click(object sender, EventArgs e)
        {
            hide_all();
            Panel7.Visible = true;
            TextBox8.Text = "";
            GridView5.Enabled = true;
            DropDownList3.DataBind();
            Session["id_specjalizacji"] = null;
            Session["id_polecenia"] = null;
            try
            {
                if (GridView5.Rows.Count > 0)
                {
                    GridView5.SelectedIndex = 0;
                    DropDownList3.SelectedIndex = DropDownList3.Items.IndexOf(DropDownList3.Items.FindByValue(GridView5.SelectedDataKey[2].ToString().Trim()));

                }
            }
            catch
            {

            }

        }

        protected void LinkButton5_Click(object sender, EventArgs e)
        {
            hide_all();
            Panel6.Visible = true;
        }

        protected void LinkButton9_Click(object sender, EventArgs e)
        {

            try
            {

                string ss = (string)Session["id_specjalizacji"];
                switch ((string)Session["id_polecenia"])
                {
                    case "1": // update
                        {
                            string id_ = (string)Session["id_specjalizacji"];
                            cl.update_specjalizacji(id_, TextBox8.Text.Trim(), int.Parse(DropDownList3.SelectedValue.ToString().Trim()));
                        }
                        break;
                    case "2": // update
                        {
                            cl.dodaj_specyfikacje(TextBox8.Text.Trim(), int.Parse(DropDownList3.SelectedValue.ToString().Trim()));
                        }
                        break;
                    default:
                        {
                            cl.dodaj_specyfikacje(TextBox8.Text.Trim(), int.Parse(DropDownList3.SelectedValue.ToString().Trim()));
                        }
                        break;
                }
            }
            catch { }

            GridView5.DataBind();

        } //end of zapisywanie specjalizacji

        protected void GridView6_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                TextBox9.Text = GridView6.SelectedDataKey[1].ToString();
                TextBox10.Text = GridView6.SelectedDataKey[2].ToString();
                TextBox11.Text = GridView6.SelectedDataKey[3].ToString();
                TextBox12.Text = GridView6.SelectedDataKey[4].ToString();
                TextBox13.Text = GridView6.SelectedDataKey[5].ToString();
                Session["id_polecenia"] = "2";
                Session["id_bazy"] = GridView6.SelectedDataKey[0].ToString();

                string txt = GridView6.SelectedDataKey[6].ToString();
                DropDownList4.SelectedIndex = DropDownList4.Items.IndexOf(DropDownList4.Items.FindByValue(GridView6.SelectedDataKey[6].ToString().Trim()));
                TextBox9.Focus();
            }
            catch
            {

            }
        }

        protected void LinkButton16_Click(object sender, EventArgs e)
        {
            Panel14.Visible = false;
            Panel16.Visible = true;
            dbNazwa.Text = "";
            dbOpis.Text = "";
            dbPasswd.Text = "";
            dbServer.Text = "";
            dbUser.Text = "";

            Session["id_polecenia"] = "1";
            /*  TextBox9.Text = "";
              TextBox10.Text = "";
              TextBox11.Text = "";
              TextBox12.Text = "";
              TextBox13.Text = "";
              TextBox9.Focus();*/
        }

        protected void LinkButton13_Click(object sender, EventArgs e)
        {
            try
            {
                cl.update_bazy_danych(TextBox9.Text.Trim(), TextBox10.Text.Trim(), TextBox11.Text.Trim(), TextBox12.Text.Trim(), TextBox13.Text.Trim(), DropDownList4.SelectedValue.ToString().Trim(), (string)Session["id_bazy"]);

                GridView6.DataBind();
            }
            catch
            {
            }
        }

        protected void LinkButton12_Click(object sender, EventArgs e)
        {
            hide_all();
            Panel8.Visible = true;
            TextBox9.Text = "";
            TextBox10.Text = "";
            TextBox11.Text = "";
            TextBox12.Text = "";
            TextBox13.Text = "";
            Session["id_polecenia"] = null;
            Session["id_bazy"] = null;
            if (GridView6.Rows.Count == 0) Session["id_polecenia"] = "1";
            TextBox9.Focus();

        }

        protected void LinkButton15_Click(object sender, EventArgs e)
        {
            if (((string)Session["id_polecenia"] == "2") && (Session["id_bazy"] != null))
            {
                cl.usun_baze_danych((string)Session["id_bazy"]);
                GridView6.DataBind();
                TextBox9.Text = "";
                TextBox10.Text = "";
                TextBox11.Text = "";
                TextBox12.Text = "";
                TextBox13.Text = "";
                Session["id_polecenia"] = null;
                Session["id_bazy"] = null;
                if (GridView6.Rows.Count == 0) Session["id_polecenia"] = "1";
                TextBox9.Focus();

            }
        }

        protected void LinkButton17_Click(object sender, EventArgs e)
        {
            // test połaczenia
            try
            {
                string svr_name = TextBox10.Text.Trim();
                string db_name = TextBox11.Text.Trim();
                string user_ = TextBox12.Text.Trim();
                string paswd = TextBox13.Text.Trim();
                string result = cl.sprawdz_polaczenie(svr_name, db_name, user_, paswd);
                Label16.Text = result;
            }
            catch
            {

            }
        }

        protected void LinkButton7_Click(object sender, EventArgs e)
        {
            // dodanie użytkownika
            Panel10.Visible = true;
            Session["id_polecenia"] = "1";
            TextBox14.Text = "";
            TextBox15.Text = "";
            TextBox16.Text = "";
            try
            {
                DropDownList1.DataBind();
                DropDownList1.SelectedIndex = 0;
            }
            catch
            { }

            TextBox14.Focus();
            LinkButton20.Visible = false;
        }

        protected void ImageButton1_Click(object sender, ImageClickEventArgs e)
        {
            Panel10.Visible = false;
            Session["id_polecenia"] = null;
        }

        protected void GridView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            // edycja użytkownika
            try
            {
                TextBox15.Text = GridView1.SelectedDataKey[2].ToString();
                TextBox16.Text = GridView1.SelectedDataKey[2].ToString();
                TextBox14.Text = GridView1.SelectedDataKey[1].ToString();
                TextBox17.Text = GridView1.SelectedDataKey[3].ToString();
                TextBox18.Text = GridView1.SelectedDataKey[4].ToString();
                DropDownList1.SelectedIndex = DropDownList1.SelectedIndex = DropDownList1.Items.IndexOf(DropDownList1.Items.FindByValue(GridView1.SelectedDataKey[5].ToString()));
                Session["id_polecenia"] = "2";
                Session["id_osoby"] = GridView1.SelectedDataKey[0].ToString();
                Panel10.Visible = true;
                TextBox14.Focus();
                LinkButton20.Visible = true;
            }
            catch
            { }
        }

        protected void LinkButton18_Click(object sender, EventArgs e)
        {
            // zapis lub update uzytkownika
            try
            {
                switch ((string)Session["id_polecenia"])
                {
                    case "1": //dodanie
                        {
                            cl.dodaj_uzytkownika(TextBox14.Text.Trim(), TextBox15.Text.Trim(), TextBox17.Text.Trim(), TextBox18.Text.Trim(), int.Parse(DropDownList1.SelectedValue.ToString()));

                        }
                        break;
                    case "2":
                        {
                            cl.update_uzytkownika(TextBox14.Text.Trim(), TextBox15.Text.Trim(), TextBox17.Text.Trim(), TextBox18.Text.Trim(), int.Parse(DropDownList1.SelectedValue.ToString()), int.Parse((string)Session["id_osoby"]));

                        }
                        break;


                    default:
                        break;
                }
                Panel10.Visible = false;
                GridView1.DataBind();
            }
            catch
            {

            }
        }

        protected void LinkButton20_Click(object sender, EventArgs e)
        {
            try
            {
                cl.usun_uzytkownika(int.Parse((string)Session["id_osoby"]));
                Panel10.Visible = false;
                GridView1.DataBind();
            }
            catch
            {

            }

        }

        protected void LinkButton11_Click(object sender, EventArgs e)
        {
            try
            {
                string ss = (string)Session["id_specjalizacji"];
                cl.usun_specjalizacje(ss);
                TextBox8.Text = "";
                GridView5.DataBind();
                Session["id_specjalizacji"] = null;
                Session["id_polecenia"] = null;
            }
            catch
            {


            }
        }

        protected void GridView7_SelectedIndexChanged(object sender, EventArgs e)
        {
            TextBox19.Text = GridView7.SelectedDataKey[1].ToString().Trim();
            Session["id_grupy"] = GridView7.SelectedDataKey[0].ToString().Trim();
            Session["id_polecenia"] = 1;// edit
            Label19.Text = "";
        }

        protected void LinkButton23_Click(object sender, EventArgs e)
        {
            TextBox19.Text = GridView7.SelectedDataKey[1].ToString().Trim();
            Session["id_grupy"] = GridView7.SelectedDataKey[0].ToString().Trim();
            Session["id_polecenia"] = 1;// edit
        }

        protected void LinkButton25_Click(object sender, EventArgs e)
        {
            hide_all();
            Panel3.Visible = true;
            GridView7.DataBind();
            if (GridView7.Rows.Count >= 0)
            {
                GridView7.SelectedIndex = 0;
                try
                {
                    TextBox19.Text = GridView7.SelectedDataKey[1].ToString().Trim();
                    Session["id_grupy"] = GridView7.SelectedDataKey[0].ToString().Trim();
                    Session["id_polecenia"] = 1;// edit
                }
                catch
                {
                    TextBox19.Text = "";
                    Session["id_polecenia"] = 2;// nowy
                }

            }
        }

        protected void LinkButton21_Click(object sender, EventArgs e)
        {
            try
            {
                if (cl.dodaj_Grupe_specjalizacj1(TextBox19.Text.Trim()) == "0")
                {
                    GridView7.DataBind();
                    if (GridView7.Rows.Count >= 0)
                    {
                        GridView7.SelectedIndex = 0;
                        try
                        {
                            TextBox19.Text = GridView7.SelectedDataKey[1].ToString().Trim();
                            Session["id_grupy"] = GridView7.SelectedDataKey[0].ToString().Trim();
                            Session["id_polecenia"] = 1;// edit
                        }
                        catch
                        {
                            TextBox19.Text = "";
                            Session["id_polecenia"] = 2;// nowy
                        }

                    }
                }
                else
                {
                    Label19.Text = "Już istnieje grupa o takiej nazwie";
                }
            }
            catch
            {

            }

        }

        protected void LinkButton22_Click(object sender, EventArgs e)
        {
            // zapis grupy specjalizacji

            if (cl.zmien_Grupe_specjalizacji(int.Parse(GridView7.SelectedDataKey[0].ToString().Trim()), TextBox19.Text.Trim()) == "0")
            {
                GridView7.DataBind();
                if (GridView7.Rows.Count >= 0)
                {
                    GridView7.SelectedIndex = 0;
                    try
                    {
                        TextBox19.Text = GridView7.SelectedDataKey[1].ToString().Trim();
                        Session["id_grupy"] = GridView7.SelectedDataKey[0].ToString().Trim();
                        Session["id_polecenia"] = 1;// edit
                    }
                    catch
                    {
                        TextBox19.Text = "";
                        Session["id_polecenia"] = 2;// nowy
                    }

                }

            }
            else
            {
                Label19.Text = "Już istnieje grupa o takiej nazwie";

            }
        }

        protected void LinkButton24_Click(object sender, EventArgs e)
        {
            try
            {
                if (int.Parse(cl.ile__Grup_specjalizacji(GridView7.SelectedDataKey[0].ToString().Trim())) > 0)
                {
                    Label19.Text = "Istnieja jeszcze specjalizacje przypisane do tej grupy!";
                }
                else
                {
                    // delete
                    if (cl.usun__Grup_specjalizacji(GridView7.SelectedDataKey[0].ToString().Trim()) == "0")
                    {
                        GridView7.DataBind();
                        if (GridView7.Rows.Count >= 0)
                        {
                            GridView7.SelectedIndex = 0;
                            try
                            {
                                TextBox19.Text = GridView7.SelectedDataKey[1].ToString().Trim();
                                Session["id_grupy"] = GridView7.SelectedDataKey[0].ToString().Trim();
                                Session["id_polecenia"] = 1;// edit
                            }
                            catch
                            {
                                TextBox19.Text = "";
                                Session["id_polecenia"] = 2;// nowy
                            }

                        }
                        else
                        {
                            TextBox19.Text = "";
                            Session["id_grupy"] = null;
                        }

                    }
                }
            }
            catch (Exception)
            {


            }

        }

        protected void LinkButton31_Click(object sender, EventArgs e)
        {
            Panel16.Visible = false;
            Panel14.Visible = true;
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            string opis = dbOpis.Text;
            string nazwaServera = dbServer.Text;
            string nazwabazy = dbNazwa.Text;
            string user = dbUser.Text;
            string haslo = dbPasswd.Text.Trim();
            string ansver = cl.dodaj_baze_danych(opis, nazwaServera, nazwabazy, user, haslo, DropDownList6.SelectedValue.ToString().Trim());
            try
            {
                int i = int.Parse(ansver);
                if (i == 0)
                {
                    Panel16.Visible = false;
                    Panel14.Visible = true;
                    GridView6.DataBind();
                    Label16.Text = "Dodano pomyślnie bazę danych: " + opis;
                }
            }
            catch
            {

                Label16.Text = "niepowodzenie dodania bazy danych: " + opis + "<br/> Kod błedu: " + ansver;
            }
        }
    }// end of class
}