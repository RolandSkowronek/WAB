using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace wab2018
{
    public partial class Add_01 : System.Web.UI.Page
    {
        public Class2 cl = new Class2();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                if (Session["user_id"] == null)
                {
                    Server.Transfer("default.aspx");
                }
                try
                {
                    string sesja = (string)Session["sesja"];
                    if (sesja == null)
                    {
                        string dt = DateTime.Now.DayOfYear.ToString() + DateTime.Now.TimeOfDay.Ticks.ToString();
                        Session["sesja"] = dt;
                        if (string.IsNullOrEmpty ( txDataPoczatku.Text.Trim () ))
                        {
                            txDataPoczatku.Date = DateTime.Now;
                        //    txDataPoczatku.Text = DateTime.Now.ToString("dd.MM.yyyy");
                            string dataKonca = DateTime.Now.AddYears(4).Year.ToString()+".12.31"  ;
                            txDataKonca.Date = DateTime.Parse(dataKonca );

                          //  txDataKonca.Text = "31.12." + DateTime.Now.AddYears(4).Year.ToString();
                        }
                     
                        TextBox1.Focus();
                    }
                }
                catch (Exception ex)
                {
                }


            }
        }

        protected void GridView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int id_ = 0;
            try
            {
                id_ = int.Parse(GridView1.SelectedDataKey[0].ToString());
                if (id_ > 0)
                {
                    if (Session["sesja"] != null)
                    {
                        cl.dodaj_specjalizacje((string)Session["sesja"], id_);
                        GridView2.DataBind();
                    }


                }
            }
            catch
            {


            }
        }

        protected void GridView2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                int id = int.Parse(GridView2.SelectedDataKey[0].ToString());
                cl.usun_specjalizacje(id);
                GridView2.DataBind();
            }
            catch
            {

            }

        }

        protected void LinkButton1_Click(object sender, EventArgs e)
        {
            string imie = TextBox1.Text.Trim();
            string nazwisko = TextBox2.Text.Trim();
            string adres = TextBox3.Text.Trim();
            string kod_poczt = TextBox4.Text.Trim();
            string miejscowosc = TextBox5.Text.Trim();

            DateTime data_pocz = DateTime.Now;
            try
            {
                data_pocz = txDataPoczatku.Date;
            }
            catch
            {
                data_pocz = DateTime.Parse(txDataPoczatku.Text);
            }

            DateTime data_konc = DateTime.Now;
            try
            {
                data_konc = txDataKonca.Date;
            }
            catch
            {
                data_konc = DateTime.Parse(txDataKonca.Text);
            }

            string id_kreatora = (string)Session["user_id"];
            int err = 0;
            // sprawdzenie poprawnosci dat
            if (data_konc <= data_pocz)
            {
                err = 1;
            }
            // sprawdzenie ilosci specjalizacji
            if (GridView2.Rows.Count == 0)
            {
                err = 2;

            }
            switch (err)
            {
                case 0:
                    {
                        string id_osoby = cl.dodaj_osobe(imie, nazwisko, adres, kod_poczt, miejscowosc, data_pocz, data_konc, int.Parse(id_kreatora), TextBox8.Text.Trim(), TextBox9.Text.Trim(),1);
                        cl.modyfikuj_osobe(id_osoby, int.Parse(id_kreatora), imie, nazwisko, adres, kod_poczt, miejscowosc, data_pocz, data_konc, TextBox8.Text, TextBox9.Text, email.Text, "", 0, adrKor.Text, kodKor.Text, cityKor.Text, tel1.Text, tel2.Text,DateTime.Parse ( "1900-01-01"), DateTime.Parse("1900-01-01"), "",1);
                        try
                        {
                            Int64 ident = Int64.Parse(id_osoby);
                            if (ident > 0)
                            {
                                for (int i = 0; i <= GridView2.Rows.Count - 1; i++)
                                {
                                    Int32 specjalizacja = (Int32)GridView2.DataKeys[i][1];
                                    cl.dodaj_specjalizacje_osoby(specjalizacja.ToString(), id_osoby);
                                }
                            }
                        }

                        catch
                        { }
                        cl.usun_specjalizacje_tymczasowe((string)Session["sesja"]);
                        Session["sesja"] = null;
                        cl.update_specjalizacjiWidoku(id_osoby);
                        Server.Transfer("WykazBieglych.aspx");
                    }
                    break;
                case 1:
                    {
                        Page.ClientScript.RegisterStartupScript(this.GetType(), "", "Alert('Data końca powołania jest mniejsza niż data początku okresu powołania!')", true);
                        //alert1.Visible = true;
                        //alert1.alert_txt("Data końca powołania jest mniejsza niż data początku okresu powołania!", 0);
                    }
                    break;
                case 2:
                    {
                        ///alert1.Visible = true;
                        Page.ClientScript.RegisterStartupScript(this.GetType(), "", "Alert('Proszę wybrać co najmniej jedną specjalizację!!')", true);

                        ///alert1.alert_txt("Proszę wybrać co najmniej jedną specjalizację!", 0);
                    }
                    break;
                default:
                    break;
            }

        }

        protected void LinkButton2_Click(object sender, EventArgs e)
        {
            cl.usun_specjalizacje_tymczasowe((string)Session["sesja"]);
            TextBox1.Text = "";
            TextBox2.Text = "";
            TextBox3.Text = "";
            TextBox4.Text = "";
            TextBox5.Text = "";
            txDataPoczatku.Text = DateTime.Now.Date.ToShortDateString();
            txDataKonca.Text = DateTime.Now.AddYears(4).Date.ToShortDateString();

        }

        protected void txDataPoczatku_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int year = DateTime.Parse(txDataPoczatku.Text).Year + 4;
                txDataKonca.Text = "31-12-" + year.ToString();
            }
            catch
            {


            }
        }
    }
}