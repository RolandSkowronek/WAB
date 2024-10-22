using System;
using System.Data;
using DevExpress.Web;
using System.Collections;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Configuration;
namespace wab2018
{
    public partial class Lista_11 : System.Web.UI.Page
    {
        public Class2 cl = new Class2();


        protected void Page_Load(object sender, EventArgs e)
        {
            if (CheckBox4.Checked)
            {
                normalny.Visible = false;
                archiwum.Visible = true;
            }
            else
            {
                archiwum.Visible = false;
                normalny.Visible = true;
            }
            if (!IsPostBack)
            {


                if (Session["user_id"] == null)
                {
                    Server.Transfer("logowanie.aspx");
                }
                else
                {
                    Session["level"] = (string)Session["rola"];
                }
                Specjalizacje_temp.DataBind();
                ustaw_baze();
                ASPxDateEdit1.Text = DateTime.Now.Date.ToShortDateString();
                ASPxDateEdit2.Text = (DateTime.Now.Date.Year + 4).ToString() + "-12-31";
                ASPxGridView5.DataBind();

            }
            Specjalizacje_temp.DataBind();
            ustaw_baze();
            ASPxDateEdit1.Text = DateTime.Now.Date.ToShortDateString();
            ASPxDateEdit2.Text = (DateTime.Now.Date.Year + 4).ToString() + "-12-31";
            ASPxGridView5.DataBind();
            ASPxGridView5.DataBind();
            try
            {
                AppSettingsReader app = new AppSettingsReader();
                int debug = (int)app.GetValue("debug", typeof(int));
                Session["debug"] = debug;
            }
            catch (Exception)
            { }
            try
            {

                int sesj = 0;// (int)Session["skargiEdycja"];
                if (sesj == 1)
                {
                    ASPxGridView7.Visible = false;
                    LinkButton10.Visible = false;
                    Panel11.Visible = true;
                }
            }
            catch 
            {


            }
            SqlDataSkargi2.DataBind();
            ASPxGridView7.DataBind();
            string powP = (string)Session["od"];
            //  poczPowolania.Text = powP;
            //  koniecPowolania.Text = (string)Session["do"];

        }


        protected void wyswietlPoSpecjalizacji(object sender, EventArgs e)
        {
            //wybiera z listy i odświeża widok
            ustaw_baze();
        }

        protected void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            ustaw_baze();
        }

        protected void print_(object sender, EventArgs e)
        {
            //  ASPxGridViewExporter1..GridView = ASPxGridView2;
            ASPxGridViewExporter1.PrintSelectCheckBox = true;
            ASPxGridViewExporter1.Landscape = true;

            ASPxGridViewExporter1.FileName = "Wykaz biegłych";
            ASPxGridViewExporter1.WritePdfToResponse();
        }

        protected void makeExcell(object sender, EventArgs e)
        {
            ASPxGridViewExporter1.FileName = "Wykaz Biegłych";
            ASPxGridViewExporter1.WriteXlsToResponse();
        }


        protected void ustaw_baze()
        {
           
            string kwerenda = string.Empty;


            if (CheckBox4.Checked)
            {
                //archiwum
                SqlDataSource1 = null;
                if (CheckBox1.Checked)
                {
                    //z kategoriami
                    DropDownList1.Enabled = true;
                    string idSspecjalizacji = DropDownList1.SelectedValue.ToString().Trim();
                    kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby WHERE (tbl_specjalizacje_osob.id_specjalizacji = " + DropDownList1.SelectedValue.ToString().Trim() + ") and View_listaBieglych.data_koncowa<getdate() order by nazwisko";

                }
                else
                {
                    //bez kategorii
                    DropDownList1.Enabled = false;
                    kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby where  View_listaBieglych.data_koncowa < getdate() order by nazwisko";

                }
                archiwum1.SelectCommand = kwerenda;
                ASPxGridView8.DataBind();

            }
            else
            {
                archiwum1 = null;
                if (CheckBox1.Checked)
                {
                    //z kategoriami

                    DropDownList1.Enabled = true;
                    string idSspecjalizacji = DropDownList1.SelectedValue.ToString().Trim();
                    kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby where (tbl_specjalizacje_osob.id_specjalizacji = " + DropDownList1.SelectedValue.ToString().Trim() + ") and View_listaBieglych.data_koncowa >= getdate() order by nazwisko";

                }
                else
                {
                    DropDownList1.Enabled = false;
                    //bez kategorii
                    DropDownList1.Enabled = false;
                    kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby where  View_listaBieglych.data_koncowa >= getdate() order by nazwisko";

                }

                SqlDataSource1.SelectCommand = kwerenda;
                SqlDataSource1.DataBind();
                ASPxGridView2.DataBind();

            }
            /*      string kwerenda = string.Empty;
                  switch (bit)
                  {
                      case 1:
                          {
                              //tylko specjakizacje

                              string idSspecjalizacji = DropDownList1.SelectedValue.ToString().Trim();
                                        //SELECT DISTINCT tytul, imie, nazwisko, adres, data_poczatkowa, data_koncowa, ident, adres2, zawieszony, (specjalizacjeWidok +' '+ specjalizacja_opis) as specjalizacjeWidok ,  uwagi, specjalizacja_opis, tel1 as telefon FROM View_listaBieglych where data_koncowa>=getdate() order by nazwisko
                              kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby WHERE (tbl_specjalizacje_osob.id_specjalizacji = " + DropDownList1.SelectedValue.ToString().Trim() + ") and  View_listaBieglych.data_koncowa>=getdate() order by nazwisko";

                          }ASPxGridView2
                          break;
                      case 2:
                          {
                         //     kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby where  View_listaBieglych.data_koncowa < getdate() order by nazwisko";

                              //tylko archiwum
                          }
                          break;
                      case 3:
                          {
                              //archiwum ze specjalizacja
                              string idSspecjalizacji = DropDownList1.SelectedValue.ToString().Trim();
                     //         kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby WHERE (tbl_specjalizacje_osob.id_specjalizacji = " + DropDownList1.SelectedValue.ToString().Trim() + ") and View_listaBieglych.data_koncowa<getdate() order by nazwisko";

                          }
                          break;
                      default:
                          {
                               kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby where  View_listaBieglych.data_koncowa >= getdate() order by nazwisko";

                          }
                          break;
                  }



                  SqlDataSource1.SelectCommand = kwerenda;
                  SqlDataSource1.DataBind();

                  ASPxGridView2.DataBind(); */
            //  maintable = cl.wyciagnijDaneBieglych(kwerenda);



        } // end of ustaw_baze
        protected void callbackPanel_Callback(object source, DevExpress.Web.CallbackEventArgsBase e)
        {
            // tu jest przekazywany parametr

            try
            {
                string rola = (string)Session["rola"];
                int level = int.Parse(rola.Trim());
                if (level > 1)
                {

                    Panel1.Visible = false;
                    Panel2.Visible = true;
                }
                else
                {
                    Panel1.Visible = true;
                    Panel2.Visible = false;
                }
            }
            catch 
            {
                popup.ShowOnPageLoad = false;
            }

            int employeeId = 0;
            try
            {
                employeeId = Convert.ToInt32(e.Parameter);
                Session["employeeId"] = employeeId;
            }
            catch (Exception)
            {
                employeeId = (int)Session["employeeId"];
            }
            string idOsoby = string.Empty;
            try
            {
                idOsoby = (string)Session["id_osoby"];
            }
            catch
            { }
            //      GridView25.DataBind();
            if (employeeId.ToString() != idOsoby)
            {
                string txt = DateTime.Now.Ticks.ToString();
                Session["sesja"] = txt;
                cl.odczytaj_specjalizacje_osoby(employeeId.ToString(), (string)Session["sesja"]);
                // przenosze wszystko tu 
                Session["id_osoby"] = employeeId.ToString().Trim();
                DataTable czlowieczek = cl.dane_osobowe(employeeId.ToString().Trim());
                DataTable korespondencja = cl.dane_korespondencyjne(employeeId.ToString().Trim());
                //tel1 ,tel2 ,email ,adr_kores ,kod_poczt_kor ,miejscowosc_kor 
                SqlDataSource2.DataBind();
                //dodanie tymczasowych specjalizacji

                GridView25.DataBind();
                Specjalizacje_temp.DataBind();
                ASPxGridView6.DataBind();
                SqlDataSkargi2.DataBind();
                ASPxGridView7.DataBind();
                try
                {
                    // imie, nazwisko, ulica, kod_poczt, miejscowosc, data_poczatkowa, data_koncowa, id_kreatora,  pesel, czyus,  tytul,  czy_zaw 
                    DataTable daneBieglego = new DataTable();
                    daneBieglego = Tbiegly();
                    DataRow biegly = daneBieglego.NewRow();

                    TxBTytul.Text = czlowieczek.Rows[0]["tytul"].ToString().Trim();
                    TxbImie.Text = czlowieczek.Rows[0]["imie"].ToString().Trim();
                    TxBNazwisko.Text = czlowieczek.Rows[0]["nazwisko"].ToString().Trim();
                    TxbPesel.Text = czlowieczek.Rows[0]["pesel"].ToString().Trim();
                    ASPxDateEdit3.EditFormatString = "yyyy-MM-dd";
                    ASPxDateEdit4.EditFormatString = "yyyy-MM-dd";
                    string d1 = DateTime.Parse(czlowieczek.Rows[0]["data_poczatkowa"].ToString().Trim()).ToShortDateString();
                    string d2 = DateTime.Parse(czlowieczek.Rows[0]["data_koncowa"].ToString().Trim()).ToShortDateString();
                    ASPxDateEdit3.Date = DateTime.Parse(czlowieczek.Rows[0]["data_poczatkowa"].ToString().Trim());

                    ASPxDateEdit3.Date = DateTime.Parse(czlowieczek.Rows[0]["data_poczatkowa"].ToString().Trim());
                    ASPxDateEdit4.Date = DateTime.Parse(czlowieczek.Rows[0]["data_koncowa"].ToString().Trim());
                    //   poczPowolania.Value = d1;
                    // koniecPowolania.Value = d2;
                    Session["od"] = d1;
                    Session["do"] = d2;
                    //dane kontaktowe
                    TxAdres1.Text = czlowieczek.Rows[0]["ulica"].ToString();
                    Txkod1.Text = czlowieczek.Rows[0]["kod_poczt"].ToString();
                    TxMiejscowosc1.Text = czlowieczek.Rows[0]["miejscowosc"].ToString();

                    TxAdres2.Text = korespondencja.Rows[0]["adr_kores"].ToString();
                    TxKod2.Text = korespondencja.Rows[0]["kod_poczt_kor"].ToString();
                    Txmiejscowosc2.Text = korespondencja.Rows[0]["miejscowosc_kor"].ToString();
                    TxTelefon1.Text = korespondencja.Rows[0]["tel1"].ToString();
                    TxTelefon2.Text = korespondencja.Rows[0]["tel2"].ToString();
                    TxEmail.Text = korespondencja.Rows[0]["email"].ToString();

                    //==========================================
                    string zawieszenie = czlowieczek.Rows[0]["czy_zaw"].ToString().Trim();
                    txspecjalizacja_opis.Text = czlowieczek.Rows[0]["specjalizacja_opis"].ToString().Trim();
                    TxUwagi.Text = czlowieczek.Rows[0]["uwagi"].ToString().Trim();
                    bool zawieszenie1 = false;
                    try
                    {
                        zawieszenie1 = bool.Parse(zawieszenie);
                    }
                    catch
                    {

                    }
                    if (zawieszenie1)
                    {
                        DropDownList4.SelectedIndex = 1;
                        zawieszenieData.Visible = true;
                        zawieszenieData.Date = DateTime.Parse(czlowieczek.Rows[0]["d_zawieszenia"].ToString());
                        zawieszenieData.Text = zawieszenieData.Date.ToShortDateString();
                    }
                    else
                    {
                        DropDownList4.SelectedIndex = 0;
                        zawieszenieData.Visible = false;
                        zawieszenieData.Text = "";

                    }

                    DataTable dt = new DataTable();
                    try
                    {
                        dt = (DataTable)Session["daneBieglego"];
                        TxBTytul.Text = dt.Rows[0]["tytul"].ToString();
                        TxbImie.Text = dt.Rows[0]["imie"].ToString();
                        TxBNazwisko.Text = dt.Rows[0]["nazwisko"].ToString();
                        TxAdres1.Text = dt.Rows[0]["adresZamieszkania"].ToString();
                        Txkod1.Text = dt.Rows[0]["kodPocztowyZamieszkania"].ToString();
                        TxMiejscowosc1.Text = dt.Rows[0]["miejscowoscZamieszkania"].ToString();
                        TxAdres2.Text = dt.Rows[0]["adresKorespondencyjny"].ToString();
                        TxKod2.Text = dt.Rows[0]["kodPocztowyKorespondencyjny"].ToString();
                        Txmiejscowosc2.Text = dt.Rows[0]["miejscowoscKorespondencyjna"].ToString();
                        TxTelefon1.Text = dt.Rows[0]["telefon1"].ToString();
                        TxTelefon2.Text = dt.Rows[0]["telefon2"].ToString();
                        TxEmail.Text = dt.Rows[0]["email"].ToString();
                        txspecjalizacja_opis.Text = dt.Rows[0]["opis"].ToString();
                        TxbPesel.Text = dt.Rows[0]["pesel"].ToString();
                        zawieszenieData.Text = dt.Rows[0]["dataZawieszenia"].ToString();
                        uwagi.Text = dt.Rows[0]["uwagi"].ToString();

                    }
                    catch
                    {

                    }
                    zapamietajDaneBieglego();
                    GridView25.DataBind();
                }
                catch 
                { }

            }
            else
            {
                //zapelnienie pustego
                if (pustePola())
                {
                    DataTable dt = new DataTable();
                    try
                    {
                        dt = (DataTable)Session["daneBieglego"];
                        TxBTytul.Text = dt.Rows[0]["tytul"].ToString();
                        TxbImie.Text = dt.Rows[0]["imie"].ToString();
                        TxBNazwisko.Text = dt.Rows[0]["nazwisko"].ToString();
                        TxAdres1.Text = dt.Rows[0]["adresZamieszkania"].ToString();
                        Txkod1.Text = dt.Rows[0]["kodPocztowyZamieszkania"].ToString();
                        TxMiejscowosc1.Text = dt.Rows[0]["miejscowoscZamieszkania"].ToString();
                        TxAdres2.Text = dt.Rows[0]["adresKorespondencyjny"].ToString();
                        TxKod2.Text = dt.Rows[0]["kodPocztowyKorespondencyjny"].ToString();
                        Txmiejscowosc2.Text = dt.Rows[0]["miejscowoscKorespondencyjna"].ToString();
                        TxTelefon1.Text = dt.Rows[0]["telefon1"].ToString();
                        TxTelefon2.Text = dt.Rows[0]["telefon2"].ToString();
                        TxEmail.Text = dt.Rows[0]["email"].ToString();
                        txspecjalizacja_opis.Text = dt.Rows[0]["opis"].ToString();
                        TxbPesel.Text = dt.Rows[0]["pesel"].ToString();
                        zawieszenieData.Text = dt.Rows[0]["dataZawieszenia"].ToString();
                        uwagi.Text = dt.Rows[0]["uwagi"].ToString();
                        //    poczPowolania.Date = DateTime.Parse(dt.Rows[0]["dataPowolaniaPoczatek"].ToString());
                        //                        koniecPowolania.Date = DateTime.Parse(dt.Rows[0]["dataPowolaniaKoniec"].ToString());


                    }
                    catch
                    { }
                }
            }
        }
        protected void ASPxGridView1_CustomErrorText(object sender, DevExpress.Web.ASPxGridViewCustomErrorTextEventArgs e)
        {

        }

        protected void changeQuerry(object sender, EventArgs e)
        {
            // Label7.Text = DropDownList3.SelectedValue.ToString();
            try
            {
                string idBieglego = (string)Session["id_bieglego"];
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
        #region zawieszenie
        protected void LinkButton5_Click(object sender, EventArgs e)
        {
            //dodaj zawieszenie

            //dodaj do historii powolan
            DateTime begin_ = DateTime.Parse(ASPxDateEdit1.Text.Trim());
            DateTime end_ = DateTime.Parse(ASPxDateEdit2.Text.Trim());
            int err = 0;
            if (end_ < begin_)
            {
                err = 1;
            }
            string user = string.Empty;
            string id_osoby = string.Empty;

            try
            {
                id_osoby = (string)Session["id_osoby"];
                user = (string)Session["user_id"];
            }
            catch
            {
                err = 2;
            }

            switch (err)
            {
                case 0:
                    {
                        Label5.Text = cl.dodajPowolanie(int.Parse(id_osoby), begin_, end_, int.Parse(user));
                        GridView25.DataBind();
                    }
                    break;
                case 1:
                    {
                        //  Label5.Text = "Data początku okresu powołania późniejsza niż końca!";
                    }
                    break;
                case 2:
                    {
                        Label5.Text = "Brak uprawnień!";
                    }
                    break;
                default:
                    break;
            }

            ASPxDateEdit3.Text = (string)Session["do"];
            zapamietajDaneBieglego();

        }

        protected void LinkButton8_Click(object sender, EventArgs e)

        {
            //zmien zawieszenie
            DateTime begin_ = DateTime.Parse(ASPxDateEdit1.Text.Trim());
            DateTime end_ = DateTime.Parse(ASPxDateEdit2.Text.Trim());
            int err = 0;
            if (end_ < begin_)
            {
                err = 1;
            }
            string user = string.Empty;
            string id_osoby = string.Empty;

            try
            {
                id_osoby = (string)Session["id_osoby"];
                user = (string)Session["user_id"];

            }
            catch
            {
                err = 2;
            }
            GridView25.DataBind();
            string idpowolania = string.Empty;
            try
            {
                idpowolania = (string)Session["idPowolania"];
            }
            catch (Exception)
            { }
            if (idpowolania == null)
            {
                err = 3;
            }

            switch (err)
            {
                case 0:
                    {
                        cl.usunPowolanie(int.Parse(idpowolania), user);
                        Label5.Text = cl.dodajPowolanie(int.Parse(id_osoby), begin_, end_, int.Parse(user));
                        GridView25.DataBind();
                    }
                    break;
                case 1:
                    {
                        //  Label5.Text = "Data początku okresu powołania późniejsza niż końca!";
                    }
                    break;
                case 2:
                    {
                        Label5.Text = "Brak uprawnień!";
                    }
                    break;
                default:
                    break;
            }
            string powP = (string)Session["od"];
            //   poczPowolania.Text = powP;
            // koniecPowolania.Text = (string)Session["do"];
            zapamietajDaneBieglego();

        }

        protected void LinkButton9_Click(object sender, EventArgs e)
        {
            //usun powolanie
            int err = 0;

            string user = string.Empty;
            string id_osoby = string.Empty;
            string id_powolania = string.Empty;
            try
            {
                id_osoby = (string)Session["id_osoby"];
                user = (string)Session["user_id"];
                id_powolania = (string)Session["idPowolania"];
            }
            catch
            {
                //err = 2;
            }


            if (string.IsNullOrEmpty(id_powolania) == true)
            {
                err = 1;
            }
            switch (err)
            {
                case 1: { } break;
                case 2: { } break;

                default:
                    {
                        cl.usunPowolanie(int.Parse(id_powolania), user);
                        GridView25.DataBind();
                    }
                    break;
            }
            string powP = (string)Session["od"];
            //  poczPowolania.Text = powP;
            // koniecPowolania.Text = (string)Session["do"];
            zapamietajDaneBieglego();

        }
        #endregion
        protected void GridView25_SelectedIndexChanged(object sender, EventArgs e)
        {
            //zawieszenie do edycji
            try
            {
                GridView25.DataBind();

                ASPxDateEdit1.Date = DateTime.Parse(GridView25.SelectedDataKey[1].ToString().Trim());
                ASPxDateEdit1.Text = ASPxDateEdit1.Date.ToShortDateString();
                ASPxDateEdit2.Date = DateTime.Parse(GridView25.SelectedDataKey[2].ToString().Trim());
                ASPxDateEdit2.Text = ASPxDateEdit2.Date.ToShortDateString();
                Session["idPowolania"] = GridView25.SelectedDataKey[0].ToString().Trim();

            }
            catch (Exception)
            { }
            string txt = GridView25.SelectedDataKey[2].ToString().Trim();
            // string powP = (string)Session["od"];
            // poczPowolania.Text = powP;
            // koniecPowolania.Text = (string)Session["do"];
        }


        protected void Button2_Click(object sender, EventArgs e)
        {

        }

        protected void zapiszDaneBieglego(object sender, EventArgs e)
        {


            string log = string.Empty;
            log = log + "Start =======================" + DateTime.Now.ToString() + "<br/>";
            string idBieglego = (string)Session["id_osoby"];
            log = log + "id biegłego :" + idBieglego + "<br/>";

            string tytul = TxBTytul.Text.Trim();
            log = log + "Tytul: " + tytul + "<br/>";
            string nazwisko = TxBNazwisko.Text.Trim();
            log = log + "Nazwisko: " + nazwisko + "<br/>";
            string imie = TxbImie.Text.Trim();
            string pesel = TxbPesel.Text.Trim();
            string ulica1 = TxAdres1.Text.Trim();
            string kod1 = Txkod1.Text.Trim();
            string miejscowosc1 = TxMiejscowosc1.Text.Trim();
            string ulica2 = TxAdres2.Text.Trim();
            string kod2 = TxKod2.Text.Trim();
            string miejscowosc2 = Txmiejscowosc2.Text.Trim();
            string tel1 = TxTelefon1.Text.Trim();
            string tel2 = TxTelefon2.Text.Trim();
            string emil = TxEmail.Text.Trim();
            string czyZaw = "0";
            string pw1 = ASPxDateEdit3.Date.ToShortDateString();
            string pw2 = ASPxDateEdit4.Date.ToShortDateString();

            string powolanieOd = pw1;// ASPxDateEdit5.Date.ToShortDateString();
            string powolanieDo = pw2;// ASPxDateEdit6.Date.ToShortDateString();
            log = log + "Powolanie od: " + powolanieOd + "<br/>";
            log = log + "Powołanie do" + powolanieDo + "<br/>";
            string uwagi = TxUwagi.Text.Trim();
            string DataZawieszenia = "";
            //==============================================
            //zapisz dane
            // zapis
            int err = 0;
            int er2 = 0;

            try
            {
                czyZaw = DropDownList4.SelectedValue.ToString();
                if (czyZaw == "1")
                {
                    DataZawieszenia = zawieszenieData.Date.ToShortDateString();
                }

                int user_id = int.Parse((string)Session["user_id"]);
                // zapis wart. poczatkowych
                DateTime dat_1 = DateTime.Now;
                DateTime dat_2 = DateTime.Now;
                try
                {
                    dat_1 = ASPxDateEdit3.Date;
                }
                catch 
                {
                    //  dat_1= poczPowolania.Date;
                }

                try
                {
                    dat_2 = DateTime.Parse(powolanieDo.Trim());
                }
                catch 
                {
                    dat_2 = ASPxDateEdit4.Date;
                }

                if (dat_1 >= dat_2) err = 2;
                log = log + "powolanie Od =  " + powolanieOd + "<br/>";
                log = log + "powolanie Do =  " + powolanieDo + "<br/>";
                log = log + "dat_1 =  " + dat_1.ToString() + "<br/>";
                log = log + "dat_2 =  " + dat_2.ToString() + "<br/>";
                log = log + "DataZawieszenia =  " + DataZawieszenia + "<br/>";

                if (err == 0)
                {
                    log = log + "err=  " + err + "<br/>";
                    string txt = "1";// cl.modyfikuj_osobe(idBieglego, user_id, imie, nazwisko, ulica1, kod1, miejscowosc1, DateTime.Parse(powolanieOd), DateTime.Parse(powolanieDo), tytul, pesel, emil, uwagi, int.Parse(czyZaw), ulica2, kod2, miejscowosc2, tel1, tel2,dat DataZawieszenia, DataZawieszenia, txspecjalizacja_opis.Text.Trim(),1);
                    log = log + "komunikat po modyfikacji" + txt + "<br/>";
                    if (txt != "0")
                    {
                        er2 = 1;
                    }
                    ASPxGridView6.DataBind();
                    // popup .ShowOnPageLoad = false;
                    cl.usun_specjalizacje_osoby(int.Parse(idBieglego));
                    DataTable specki = cl.odczytaj_specjalizacje_tymczasowe_osoby_lista(idBieglego, (string)Session["sesja"]);
                    if (specki.Rows.Count > 0)
                    {
                        for (int i = 0; i <= specki.Rows.Count - 1; i++)
                        {
                            Int32 specjalizacja = (Int32)specki.Rows[i][0];
                            cl.dodaj_specjalizacje_osoby(specjalizacja.ToString(), idBieglego);
                        }
                    }
                    Session["sesja"] = null;
                }//
            }
            catch (Exception ex)
            {
                log = log + "Komunikat błedu głównego" + ex.Message + "<br/>";

            }
            Session["reload"] = true;
            //dodaj specjalizacje do widoku
            cl.update_specjalizacjiWidoku(idBieglego);
            cl.usunSpecjalizacjetymczasowe();
            if (CheckBox4.Checked)
            {
                ASPxGridView8.DataBind();
            }
            else
            {
                ASPxGridView2.DataBind();
            }

            int debug = 1;
            try
            {
                debug = (int)Session["debug"];
            }
            catch (Exception)
            { }
            Session["sesja"] = null;
            Session["id_osoby"] = null;
            if (debug == 1)
            {
                Label5.Text = log;
            }

            //zapis na zaś

            zapamietajDaneBieglego();

        }


        protected void ASPxDateEdit6_DateChanged(object sender, EventArgs e)
        {

            DateTime dt = ASPxDateEdit4.Date;

            if (ASPxDateEdit4.Text != "0001.01.01")
            {
                if (ASPxDateEdit4.Text != "0001.01.01")
                {
                    Session["do"] = ASPxDateEdit4.Text;
                }
            }
        }







        //=======================================
        Hashtable copiedValues = null;
        string[] copiedFields = new string[] { "id_" };



        protected void ASPxGridView5_CustomButtonCallback(object sender, ASPxGridViewCustomButtonCallbackEventArgs e)
        {
            if (e.ButtonID != "Clone") return;
            copiedValues = new Hashtable();
            foreach (string fieldName in copiedFields)
            {
                copiedValues[fieldName] = ASPxGridView5.GetRowValues(e.VisibleIndex, fieldName);
            }

            string id = copiedValues["id_"].ToString();

            string sesja = (string)Session["sesja"];
            int specjalizacja = int.Parse(id);
            cl.dodaj_specjalizacje(sesja, specjalizacja);

            Specjalizacje_temp.DataBind();

            ASPxGridView6.DataBind();
            //Label6.Text = sesja;

            string powP = (string)Session["od"];
            //poczPowolania.Text = powP;
            // koniecPowolania.Text = (string)Session["do"];


        }

        protected void ASPxGridView6_CustomButtonCallback(object sender, ASPxGridViewCustomButtonCallbackEventArgs e)
        {

            zapamietajDaneBieglego();

            //usuniecie specjalizacji
            ASPxGridView6.DataBind();
            try
            {
                if (e.ButtonID != "Clone2") return;
                copiedValues = new Hashtable();
                string id = ASPxGridView6.GetRowValues(e.VisibleIndex, "id_").ToString();


                string sesja = (string)Session["sesja"];
                int specjalizacja = int.Parse(id);
                cl.usun_specyfikacjeTmp(id);
                Specjalizacje_temp.DataBind();
                ASPxGridView6.DataBind();
            }
            catch (Exception)
            { }

            // dodanie dat
            string powP = (string)Session["od"];
            //poczPowolania.Text=powP;
            //koniecPowolania .Text= (string)Session["do"];
        }

        protected void ASPxDateEdit7_DateChanged(object sender, EventArgs e)
        {
            Session["dataZawieszenia"] = zawieszenieData.Value.ToString();
        }

        protected void Button3_Click(object sender, EventArgs e)
        {
            //usun biegłego
            string id_osoby = (string)Session["id_osoby"];
            if (id_osoby != null)
            {
                cl.usun_osobe(id_osoby, (string)Session["user_id"]);
            }
            Session["sesja"] = null;
            popup.ShowOnPageLoad = false;
            ASPxGridView2.DataBind();
        }

        protected void ASPxGridView7_CustomButtonCallback(object sender, ASPxGridViewCustomButtonCallbackEventArgs e)
        {
            //wybranie skargi

            ASPxGridView7.DataBind();
            try
            {
                if (e.ButtonID != "wybierz") return;
                copiedValues = new Hashtable();
                string id = ASPxGridView7.GetRowValues(e.VisibleIndex, "ident").ToString();

                Session["idSkargi"] = id;

                Label7.Visible = false;
                ASPxDateZakreslenie.Visible = false;

                Label6.Text = "Dodanie nowej skargi";
                Panel11.Visible = true;
                txNumer.Text = ASPxGridView7.GetRowValues(e.VisibleIndex, "numer").ToString();
                txRok.Text = ASPxGridView7.GetRowValues(e.VisibleIndex, "rok").ToString();
                txSygnatura.Text = ASPxGridView7.GetRowValues(e.VisibleIndex, "Sygnatura").ToString();
                txWizytator.Text = ASPxGridView7.GetRowValues(e.VisibleIndex, "wizytator").ToString();
                uwagi.Text = ASPxGridView7.GetRowValues(e.VisibleIndex, "uwagi").ToString();
                string d1 = ASPxGridView7.GetRowValues(e.VisibleIndex, "dataPisma").ToString();
                string d2 = ASPxGridView7.GetRowValues(e.VisibleIndex, "dataWplywu").ToString();
                ASPxDatePismo.Date = DateTime.Parse(d1);// ASPxGridView7.GetRowValues(e.VisibleIndex, "dataPisma").ToString();
                ASPxDateWplyw.Date = DateTime.Parse(d2);// ASPxGridView7.GetRowValues(e.VisibleIndex, "dataWplywu").ToString();

                string zakreslono = ASPxGridView7.GetRowValues(e.VisibleIndex, "zakreslono").ToString();
                if (zakreslono == "true")
                {
                    CheckBox3.Checked = true;
                    Label7.Visible = true;
                    ASPxDateZakreslenie.Visible = true;
                    ASPxDateZakreslenie.Value = ASPxGridView7.GetRowValues(e.VisibleIndex, "dataZakreslenia").ToString();
                }

                CheckBox3.Checked = false;
                Session["skargiEdycja"] = 1;
                ASPxGridView7.Visible = false;
                LinkButton10.Visible = false;
                Panel11.Visible = true;
                ASPxGridView7.DataBind();

            }
            catch 
            {}

            ASPxGridView7.DataBind();

        }

        protected void zapiszSkarge(object sender, EventArgs e)
        {
            //zapis lub modyfikacja
            string idSkargi = "0";
            try
            {
                idSkargi = (string)Session["idSkargi"];
            }
            catch (Exception)
            { }
            if (string.IsNullOrEmpty(idSkargi))
            {
                idSkargi = "0";
            }



            string idModyfikatora = (string)Session["user_id"];
            string idBieglego = (string)Session["id_osoby"];
            string numer = txNumer.Text.Trim();
            string rok = txRok.Text.Trim();
            string dataWplywu = ASPxDateWplyw.Text;
            string dataPisma = ASPxDatePismo.Text.Trim();
            string sygnatura = txSygnatura.Text.Trim();
            string wizytator = txWizytator.Text.Trim();
            string uwagii = uwagi.Text.Trim();
            bool zakreslono = CheckBox3.Checked;
            string dataZakreslenia = ASPxDateZakreslenie.Text.Trim();

            if (!CheckBox3.Checked)
            {
                dataZakreslenia = "1900-01-01 00:00:00";
            }
            cl.dodajSkarge(int.Parse(idBieglego), numer, rok, sygnatura, dataWplywu, dataPisma, wizytator, uwagii, zakreslono, dataZakreslenia, idModyfikatora, int.Parse(idSkargi));

            Panel11.Visible = false;
            ASPxGridView7.Visible = true;
            LinkButton10.Visible = true;
            ASPxGridView7.DataBind();
            Session["idSkargi"] = "0";

        }

        protected void LinkButton10_Click(object sender, EventArgs e)
        {
            //nowa skarga otwarcie okna
            Label6.Text = "Dodanie nowej skargi";
            Panel11.Visible = true;
            txNumer.Text = "";
            txRok.Text = DateTime.Now.Year.ToString();
            txSygnatura.Text = "";
            txWizytator.Text = "";
            uwagi.Text = "";
            ASPxDatePismo.Text = DateTime.Now.ToShortDateString();
            ASPxDateWplyw.Text = DateTime.Now.ToShortDateString();
            CheckBox3.Checked = false;
            ASPxGridView7.Visible = false;
            LinkButton10.Visible = false;

        }

        protected void anulowanieSkargi(object sender, EventArgs e)
        {
            Panel11.Visible = false;
            ASPxGridView7.Visible = true;
            LinkButton10.Visible = true;
        }

        protected void ASPxGridView7_CustomCallback(object sender, ASPxGridViewCustomCallbackEventArgs e)
        {
            ASPxGridView7.DataBind();
            try
            {
                ASPxGridView7.Visible = false;
                LinkButton10.Visible = false;
                Panel11.Visible = true;
                ASPxGridView7.DataBind();
            }
            catch 
            {
            }


            ASPxGridView7.DataBind();
        }

        protected void ASPxPageControl1_ActiveTabChanged(object source, TabControlEventArgs e)
        {
        }

        protected void pokazDateZ(object sender, EventArgs e)
        {
            if (CheckBox3.Checked)
            {
                Label7.Visible = true;
                ASPxDateZakreslenie.Visible = true;
            }
            else
            {
                Label7.Visible = false;
                ASPxDateZakreslenie.Visible = false;
            }
        }

        protected void LinkButton13_Click(object sender, EventArgs e)
        {
            // usunięcie 
            string idSkargi = "0";
            try
            {
                idSkargi = (string)Session["idSkargi"];
            }
            catch (Exception)
            {

            }
            if (idSkargi != "0")
            {
                string idUsuwajacego = (string)Session["user_id"];
                cl.usunSkarge(idUsuwajacego, int.Parse(idSkargi));
            }
            Session["idSkargi"] = "0";
            ASPxGridView7.Visible = true;
            LinkButton10.Visible = true;
            Panel11.Visible = false;
            ASPxGridView7.DataBind();
        }

        protected void LinkButton14_Click(object sender, EventArgs e)
        {

            if (CheckBox1.Checked)
            {
                //jedna
                robRaportjednejSpecjalizacji(DropDownList1.SelectedItem.Value.Trim());

            }
            else
            {
                robRaportWszystkichSpecjalizacji();
            }

        }

        protected void robRaportjednejSpecjalizacji(String specjalizacja)
        {
            //podliczenie
            DataTable specjalizacjeWyliczenie = new DataTable();
            specjalizacjeWyliczenie.Columns.Add("nr", typeof(string));
            specjalizacjeWyliczenie.Columns.Add("str", typeof(string));
            DataTable specjalizacje = new DataTable();
            specjalizacje = cl.odczytaj_specjalizacjeLista();


            DataTable Biegli = new DataTable();
            string idSpecjalizacji = specjalizacja;
            string nazwaSpecjalizacji = DropDownList1.SelectedItem.Text.Trim();
            int iloscStron = 0;
            Biegli = cl.wyciagnijBieglychZSpecjalizacja(idSpecjalizacji, CheckBox4.Checked);
            if (Biegli.Rows.Count > 0)
            {
                iloscStron = 1;
                PdfPTable tabelaGlowna = new PdfPTable(4);
                int[] tblWidth = { 8, 30, 30, 32 };

                tabelaGlowna = new PdfPTable(4);
                tabelaGlowna = generujCzescRaportu(Biegli, nazwaSpecjalizacji);

                if (tabelaGlowna.Rows.Count > 15)
                {
                    int counter = 0;
                    PdfPTable newTable = new PdfPTable(4);
                    newTable.SetWidths(tblWidth);
                    // podziel tabele

                    foreach (var TableRow in tabelaGlowna.Rows)
                    {
                        counter++;
                        newTable.Rows.Add(TableRow);
                        if (counter == 15)
                        {
                            iloscStron++;
                            counter = 0;
                        }
                    }
                    DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                    wyliczenie[0] = nazwaSpecjalizacji;
                    wyliczenie[1] = iloscStron.ToString();
                    specjalizacjeWyliczenie.Rows.Add(wyliczenie);

                }
                else
                {
                    DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                    wyliczenie[0] = nazwaSpecjalizacji;
                    wyliczenie[1] = iloscStron.ToString();
                    specjalizacjeWyliczenie.Rows.Add(wyliczenie);
                }
                // uttwórz listę osób z taka specjalizacją 
            }



            //==============================================================


            // wyciąfnij listę ludzi z dana specjalizacją 
            var plFont = FontFactory.GetFont(BaseFont.COURIER, BaseFont.CP1257, 10, Font.NORMAL);
            var plFont2 = FontFactory.GetFont(BaseFont.HELVETICA, BaseFont.CP1257, 9, Font.NORMAL);

            var plFont3 = FontFactory.GetFont(BaseFont.COURIER, BaseFont.CP1257, 15, Font.NORMAL);

            Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 0f);
            PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
            pdfDoc.Open();

            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;" + "filename=zestawienie_Specjalizacji.pdf");


            var plFontBIG = FontFactory.GetFont(BaseFont.HELVETICA, BaseFont.CP1257, 35, Font.NORMAL);
            PdfPTable fitst = new PdfPTable(1);
            fitst.DefaultCell.Border = Rectangle.NO_BORDER;
            PdfPCell cell = new PdfPCell(new Paragraph(" ", plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("LISTA", plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;

            cell.Border = Rectangle.NO_BORDER;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("BIEGŁYCH SĄDOWYCH", plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("", plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("SĄDU OKRĘGOWEGO", plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            fitst.AddCell(cell);
            pdfDoc.Add(fitst);
            pdfDoc.NewPage();
            PdfPTable tab = new PdfPTable(3);
            int[] tblWidth2 = { 10, 80, 10 };
            tab.SetWidths(tblWidth2);
            cell = new PdfPCell(new Paragraph("", plFontBIG));
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            tab.AddCell(cell);
            tab.AddCell(cell);
            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Lp.", plFont2));

            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Zakres", plFont2));

            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Numer strony", plFont2));

            tab.AddCell(cell);

            int biezacaStrona = 0;
            int il = 0;
            foreach (DataRow dRow in specjalizacjeWyliczenie.Rows)
            {
                il++;
                tab.AddCell(new Paragraph(il.ToString(), plFont2));
                tab.AddCell(new Paragraph(dRow[0].ToString(), plFont2));
                biezacaStrona = biezacaStrona + int.Parse(dRow[1].ToString());
                tab.AddCell(new Paragraph(biezacaStrona.ToString(), plFont2));
            }

            pdfDoc.Add(tab);

            pdfDoc.NewPage();


            Biegli = new DataTable();
            iloscStron = 0;
            Biegli = cl.wyciagnijBieglychZSpecjalizacja(idSpecjalizacji, CheckBox4.Checked);
            if (Biegli.Rows.Count > 0)
            {
                iloscStron = 1;
                pdfDoc.Add(new Paragraph(" "));
                pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, plFont3)));
                pdfDoc.Add(new Paragraph(" "));
                PdfPTable tabelaGlowna = new PdfPTable(4);
                int[] tblWidth = { 8, 30, 30, 32 };
                tabelaGlowna.SetWidths(tblWidth);
                tabelaGlowna.AddCell(new Paragraph("Lp.", plFont2));
                tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", plFont2));
                tabelaGlowna.AddCell(new Paragraph("Adres- telefom", plFont2));
                tabelaGlowna.AddCell(new Paragraph("Zakres", plFont2));


                pdfDoc.Add(tabelaGlowna);
                tabelaGlowna = new PdfPTable(4);

                tabelaGlowna = generujCzescRaportu(Biegli, nazwaSpecjalizacji);

                if (tabelaGlowna.Rows.Count > 15)
                {
                    int counter = 0;
                    PdfPTable newTable = new PdfPTable(4);
                    newTable.SetWidths(tblWidth);
                    // podziel tabele

                    foreach (var TableRow in tabelaGlowna.Rows)
                    {
                        counter++;
                        newTable.Rows.Add(TableRow);
                        if (counter == 15)
                        {
                            iloscStron++;
                            counter = 0;
                            pdfDoc.Add(newTable);
                            pdfDoc.NewPage();
                            pdfDoc.Add(new Paragraph(" "));
                            pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji + " ciąg dalszy", plFont3)));
                            pdfDoc.Add(new Paragraph(" "));

                            tabelaGlowna = new PdfPTable(4);

                            tabelaGlowna.AddCell(new Paragraph("Lp.", plFont2));
                            tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", plFont2));
                            tabelaGlowna.AddCell(new Paragraph("Adres- telefom", plFont2));

                            tabelaGlowna.AddCell(new Paragraph("Zakres", plFont2));
                            tabelaGlowna.SetWidths(tblWidth);
                        }
                    }
                    DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                    wyliczenie[0] = nazwaSpecjalizacji;
                    wyliczenie[1] = iloscStron.ToString();
                    specjalizacjeWyliczenie.Rows.Add(wyliczenie);
                    pdfDoc.Add(newTable);
                    pdfDoc.NewPage();
                }
                else
                {
                    DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                    wyliczenie[0] = nazwaSpecjalizacji;
                    wyliczenie[1] = iloscStron.ToString();
                    specjalizacjeWyliczenie.Rows.Add(wyliczenie);

                    pdfDoc.Add(tabelaGlowna);
                    pdfDoc.NewPage();

                }
                // uttwórz listę osób z taka specjalizacją 
            }

            Response.Write(pdfDoc);
            pdfDoc.Close();
        }


        protected void robRaportWszystkichSpecjalizacji()
        {
            //podliczenie
            DataTable specjalizacjeWyliczenie = new DataTable();
            specjalizacjeWyliczenie.Columns.Add("nr", typeof(string));
            specjalizacjeWyliczenie.Columns.Add("str", typeof(string));
            DataTable specjalizacje = new DataTable();
            specjalizacje = cl.odczytaj_specjalizacjeLista();

            foreach (DataRow dRow in specjalizacje.Rows)
            {

                DataTable Biegli = new DataTable();
                string idSpecjalizacji = dRow[0].ToString().Trim();
                string nazwaSpecjalizacji = dRow[1].ToString().Trim();
                int iloscStron = 0;

                Biegli = cl.wyciagnijBieglychZSpecjalizacja(idSpecjalizacji, CheckBox4.Checked);
                if (Biegli.Rows.Count > 0)
                {
                    iloscStron = 1;
                    PdfPTable tabelaGlowna = new PdfPTable(4);
                    int[] tblWidth = { 8, 30, 30, 32 };

                    tabelaGlowna = new PdfPTable(4);
                    tabelaGlowna = generujCzescRaportu(Biegli, nazwaSpecjalizacji);

                    if (tabelaGlowna.Rows.Count > 15)
                    {
                        int counter = 0;
                        PdfPTable newTable = new PdfPTable(4);
                        newTable.SetWidths(tblWidth);
                        // podziel tabele

                        foreach (var TableRow in tabelaGlowna.Rows)
                        {
                            counter++;
                            newTable.Rows.Add(TableRow);
                            if (counter == 15)
                            {
                                iloscStron++;
                                counter = 0;
                            }
                        }
                        DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                        wyliczenie[0] = nazwaSpecjalizacji;
                        wyliczenie[1] = iloscStron.ToString();
                        specjalizacjeWyliczenie.Rows.Add(wyliczenie);

                    }
                    else
                    {
                        DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                        wyliczenie[0] = nazwaSpecjalizacji;
                        wyliczenie[1] = iloscStron.ToString();
                        specjalizacjeWyliczenie.Rows.Add(wyliczenie);
                    }
                    // uttwórz listę osób z taka specjalizacją 
                }

            }// end of each




            //==============================================================


            // wyciąfnij listę ludzi z dana specjalizacją 
            var plFont = FontFactory.GetFont(BaseFont.COURIER, BaseFont.CP1257, 10, Font.NORMAL);
            var plFont2 = FontFactory.GetFont(BaseFont.HELVETICA, BaseFont.CP1257, 9, Font.NORMAL);

            var plFont3 = FontFactory.GetFont(BaseFont.COURIER, BaseFont.CP1257, 15, Font.NORMAL);

            Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 0f);
            PdfWriter.GetInstance(pdfDoc, Response.OutputStream);
            pdfDoc.Open();

            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;" + "filename=zestawienie_Specjalizacji.pdf");


            var plFontBIG = FontFactory.GetFont(BaseFont.HELVETICA, BaseFont.CP1257, 35, Font.NORMAL);
            PdfPTable fitst = new PdfPTable(1);
            fitst.DefaultCell.Border = Rectangle.NO_BORDER;
            PdfPCell cell = new PdfPCell(new Paragraph(" ", plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("LISTA", plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;

            cell.Border = Rectangle.NO_BORDER;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("BIEGŁYCH SĄDOWYCH", plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("", plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("SĄDU OKRĘGOWEGO", plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            fitst.AddCell(cell);
            pdfDoc.Add(fitst);
            pdfDoc.NewPage();
            PdfPTable tab = new PdfPTable(3);
            int[] tblWidth2 = { 10, 80, 10 };
            tab.SetWidths(tblWidth2);
            cell = new PdfPCell(new Paragraph("", plFontBIG));
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            tab.AddCell(cell);
            tab.AddCell(cell);
            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Lp.", plFont2));

            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Zakres", plFont2));

            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Numer strony", plFont2));

            tab.AddCell(cell);

            int biezacaStrona = 0;
            int il = 0;
            foreach (DataRow dRow in specjalizacjeWyliczenie.Rows)
            {
                il++;
                tab.AddCell(new Paragraph(il.ToString(), plFont2));
                tab.AddCell(new Paragraph(dRow[0].ToString(), plFont2));
                biezacaStrona = biezacaStrona + int.Parse(dRow[1].ToString());
                tab.AddCell(new Paragraph(biezacaStrona.ToString(), plFont2));
            }

            pdfDoc.Add(tab);

            pdfDoc.NewPage();

            foreach (DataRow dRow in specjalizacje.Rows)
            {

                DataTable Biegli = new DataTable();
                string idSpecjalizacji = dRow[0].ToString().Trim();
                string nazwaSpecjalizacji = dRow[1].ToString().Trim();
                int iloscStron = 0;
                Biegli = cl.wyciagnijBieglychZSpecjalizacja(idSpecjalizacji, CheckBox4.Checked);
                if (Biegli.Rows.Count > 0)
                {
                    iloscStron = 1;
                    pdfDoc.Add(new Paragraph(" "));
                    pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, plFont3)));
                    pdfDoc.Add(new Paragraph(" "));
                    PdfPTable tabelaGlowna = new PdfPTable(4);
                    int[] tblWidth = { 8, 30, 30, 32 };
                    tabelaGlowna.AddCell(new Paragraph("Lp.", plFont2));
                    tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", plFont2));
                    tabelaGlowna.AddCell(new Paragraph("Adres- telefom", plFont2));

                    tabelaGlowna.AddCell(new Paragraph("Zakres", plFont2));
                    tabelaGlowna.SetWidths(tblWidth);


                    pdfDoc.Add(tabelaGlowna);
                    tabelaGlowna = new PdfPTable(5);

                    tabelaGlowna = generujCzescRaportu(Biegli, nazwaSpecjalizacji);

                    if (tabelaGlowna.Rows.Count > 15)
                    {
                        int counter = 0;
                        PdfPTable newTable = new PdfPTable(4);
                        newTable.SetWidths(tblWidth);
                        // podziel tabele

                        foreach (var TableRow in tabelaGlowna.Rows)
                        {
                            counter++;
                            newTable.Rows.Add(TableRow);
                            if (counter == 15)
                            {
                                iloscStron++;
                                counter = 0;
                                pdfDoc.Add(newTable);
                                pdfDoc.NewPage();
                                pdfDoc.Add(new Paragraph(" "));
                                pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji + " ciąg dalszy", plFont3)));
                                pdfDoc.Add(new Paragraph(" "));

                                tabelaGlowna = new PdfPTable(4);
                                tabelaGlowna.SetWidths(tblWidth);
                                tabelaGlowna.AddCell(new Paragraph("Lp.", plFont2));
                                tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", plFont2));
                                tabelaGlowna.AddCell(new Paragraph("Adres- telefom", plFont2));


                                tabelaGlowna.AddCell(new Paragraph("Zakres", plFont2));
                            }
                        }
                        DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                        wyliczenie[0] = nazwaSpecjalizacji;
                        wyliczenie[1] = iloscStron.ToString();
                        specjalizacjeWyliczenie.Rows.Add(wyliczenie);
                        pdfDoc.Add(newTable);
                        pdfDoc.NewPage();
                    }
                    else
                    {
                        DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                        wyliczenie[0] = nazwaSpecjalizacji;
                        wyliczenie[1] = iloscStron.ToString();
                        specjalizacjeWyliczenie.Rows.Add(wyliczenie);

                        pdfDoc.Add(tabelaGlowna);
                        pdfDoc.NewPage();

                    }
                    // uttwórz listę osób z taka specjalizacją 
                }

            }// end of each

            //again

            //strona główna
            Response.Write(pdfDoc);
            pdfDoc.Close();






        }
        protected PdfPTable generujCzescRaportu(DataTable biegli, string specjalizacje)
        {

            int[] tblWidth = { 8, 30, 30, 32 };
            var plFont2 = FontFactory.GetFont(BaseFont.HELVETICA, BaseFont.CP1257, 10, Font.NORMAL);

            PdfPTable tabelaGlowna = new PdfPTable(4);
            tabelaGlowna.SetWidths(tblWidth);
            int iterator = 0;
            foreach (DataRow biegly in biegli.Rows)
            {
                iterator++;
                string Idbieglego = biegly[0].ToString();
                DataTable listaSpecjalizacjiBieglego = new DataTable();
                listaSpecjalizacjiBieglego = cl.odczytaj_specjalizacje_osoby2(Idbieglego);
                //dbo.tbl_osoby.ident, dbo.tbl_osoby.imie, dbo.tbl_osoby.nazwisko, dbo.tbl_osoby.ulica, dbo.tbl_osoby.kod_poczt, dbo.tbl_osoby.miejscowosc,   dbo.tbl_osoby.data_koncowa,  dbo.tbl_osoby.tytul,
                string imie = biegly[1].ToString();
                string nazwisko = biegly[2].ToString();
                string tytul = biegly[7].ToString();
                string dataKonca = string.Empty;
                try
                {
                    dataKonca = DateTime.Parse(biegly[6].ToString()).ToShortDateString();
                }
                catch 
                {

                }

                string innerTable = imie + Environment.NewLine + nazwisko + Environment.NewLine + tytul + Environment.NewLine + "kadencja do dnia: " + dataKonca;
                tabelaGlowna.AddCell(new Paragraph(iterator.ToString(), plFont2));
                tabelaGlowna.AddCell(new Paragraph(innerTable, plFont2));
                string ulica = biegly[3].ToString();
                string kod = biegly[4].ToString();
                string miejscowosc = biegly[5].ToString();
                string tel = biegly[8].ToString();
                string adresTable = ulica + Environment.NewLine + kod + " " + miejscowosc + Environment.NewLine + tel;
                tabelaGlowna.AddCell(new Paragraph(adresTable, plFont2));
                string specki = string.Empty;
                string specjalizacjaOpis = cl.odczytaj_specjalizacje_osobyOpis(Idbieglego);
                // tabelaGlowna.AddCell(new Paragraph(specjalizacjaOpis, plFont2));
                foreach (DataRow specRow in listaSpecjalizacjiBieglego.Rows)
                {
                    specki = specki + specRow[0].ToString().ToLower() + "; ";
                }
                specki = specki + specjalizacjaOpis;
                tabelaGlowna.AddCell(specki);
            }


            return tabelaGlowna;

        }

        protected void CheckBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (CheckBox4.Checked)
            {
                normalny.Visible = false;
                archiwum.Visible = true;
            }
            else
            {
                archiwum.Visible = false;
                normalny.Visible = true;
            }
            ustaw_baze();
        }



        protected void popup_WindowCallback(object source, PopupWindowCallbackArgs e)
        {
            //zamykanie popup
            Session["employeeId"] = null;
            Session["daneBieglego"] = null;
            Session["id_osoby"] = null;
            string txt = e.Parameter.ToString();
        }

        protected void popup_PopupElementResolve(object sender, ControlResolveEventArgs e)
        {
            bool txt = e.ResolvedControl.Visible;
        }

        protected void popup_Unload(object sender, EventArgs e)
        {

        }


        public DataTable Tbiegly()
        {
            DataTable dT = new DataTable();
            dT.Columns.Add("tytul", typeof(string));
            dT.Columns.Add("imie", typeof(string));
            dT.Columns.Add("nazwisko", typeof(string));
            dT.Columns.Add("adresZamieszkania", typeof(string));
            dT.Columns.Add("kodPocztowyZamieszkania", typeof(string));
            dT.Columns.Add("miejscowoscZamieszkania", typeof(string));
            dT.Columns.Add("adresKorespondencyjny", typeof(string));
            dT.Columns.Add("kodPocztowyKorespondencyjny", typeof(string));
            dT.Columns.Add("miejscowoscKorespondencyjna", typeof(string));
            dT.Columns.Add("telefon1", typeof(string));
            dT.Columns.Add("telefon2", typeof(string));
            dT.Columns.Add("email", typeof(string));
            dT.Columns.Add("opis", typeof(string));
            dT.Columns.Add("uwagi", typeof(string));
            dT.Columns.Add("dataPowolaniaPoczatek", typeof(string));
            dT.Columns.Add("dataPowolaniaKoniec", typeof(string));
            dT.Columns.Add("zawieszenie", typeof(string));
            dT.Columns.Add("dataZawieszenia", typeof(string));
            dT.Columns.Add("pesel", typeof(string));
            dT.Columns.Add("specjalizacja_opis", typeof(string));
            return dT;

        }

        public bool pustePola()
        {
            bool result = true;
            if (!string.IsNullOrEmpty(TxbImie.Text.Trim()))
            {
                return false;
            }
            if (!string.IsNullOrEmpty(TxbImie.Text.Trim()))
            {
                return false;
            }
            return result;

        }

        protected void ASPxGridView2_HeaderFilterFillItems(object sender, ASPxGridViewHeaderFilterEventArgs e)
        {
            string ttx = e.Values[1].ToString();
        }

        protected void zapamietajDaneBieglego()
        {
            try
            {
                string pow1 = (string)Session["od"];
                string pow2 = (string)Session["do"];
                DataTable daneBieglego = new DataTable();
                daneBieglego = Tbiegly();
                DataRow biegly = daneBieglego.NewRow();
                biegly["tytul"] = TxBTytul.Text;
                biegly["imie"] = TxbImie.Text;
                biegly["nazwisko"] = TxBNazwisko.Text;
                biegly["adresZamieszkania"] = TxAdres1.Text;
                biegly["kodPocztowyZamieszkania"] = Txkod1.Text;
                biegly["miejscowoscZamieszkania"] = TxMiejscowosc1.Text;
                biegly["adresKorespondencyjny"] = TxAdres2.Text;
                biegly["kodPocztowyKorespondencyjny"] = TxKod2.Text;
                biegly["miejscowoscKorespondencyjna"] = Txmiejscowosc2.Text;
                biegly["telefon1"] = TxTelefon1.Text;
                biegly["telefon2"] = TxTelefon2.Text;
                biegly["email"] = TxEmail.Text;
                biegly["dataPowolaniaPoczatek"] = pow1;
                biegly["dataPowolaniaKoniec"] = pow2;
                biegly["pesel"] = TxbPesel.Text;
                biegly["zawieszenie"] = DropDownList4.SelectedValue.ToString().Trim();
                biegly["dataZawieszenia"] = zawieszenieData.Text;
                biegly["uwagi"] = TxUwagi.Text;
                
                biegly["opis"] = txspecjalizacja_opis.Text.Trim();
                daneBieglego.Rows.Clear();
                daneBieglego.Rows.Add(biegly);
                Session["daneBieglego"] = daneBieglego;
            }
            catch
            {

            }

        }//end of zapamietajDaneBieglego

        protected void zmianaZawieszenia(object sender, EventArgs e)
        {
            zapamietajDaneBieglego();
            switch (DropDownList4.SelectedValue)
            {
                case "1":
                    {
                        zawieszenieData.Visible = true;
                        if (zawieszenieData.Text == "")
                        {
                            zawieszenieData.Date = DateTime.Now;
                            zawieszenieData.Text = zawieszenieData.Date.ToShortDateString();
                        }

                    }
                    break;
                default:
                    {
                        zawieszenieData.Visible = false;
                        zawieszenieData.Text = "";
                    }
                    break;
            }

            // dodanie dat
            //  string powP = (string)Session["od"];
            //  poczPowolania.Text = powP;
            //  koniecPowolania.Text = (string)Session["do"];

        }




        protected void ASPxDateEdit3_DateChanged(object sender, EventArgs e)
        {
            DateTime dT = ASPxDateEdit3.Date;

            // DateTime dt = poczPowolania.Date;
            //if (poczPowolania.Text != "0001.01.01")
            // {//   if (poczPowolania.Text != "")
            // {
            //   Session["od"] = poczPowolania.Text;
            // }
            //}



        }

        protected void carTabPage_ActiveTabChanged(object source, TabControlEventArgs e)
        {

        }

    }
}