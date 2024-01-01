using System;
using System.Data;
using DevExpress.Web;
using System.Collections;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Configuration;
using System.IO;
using DevExpress.XtraPrinting;
using System.Web;
using System.Net;
using System.Net.Mime;

namespace wab2018
{
    public partial class wykazBieglych : System.Web.UI.Page
    {
        public Class2 cl = new Class2();
        private cm Common = new cm();
        Hashtable copiedValues = null;
        string[] copiedFields = new string[] { "id_" };

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
        
            if (ASPxPageControl1.ActiveTabIndex != 5)
            {
                ASPxDateEdit1.Text = DateTime.Now.Date.ToShortDateString();
                ASPxDateEdit2.Text = (DateTime.Now.Date.Year + 4).ToString() + "-12-31";
                ASPxGridView5.DataBind();

            }
            else
            {
                if (string.IsNullOrEmpty(ASPxDateEdit1.Text))
                {
                    ASPxDateEdit1.Date = DateTime.Now.Date;
                    ASPxDateEdit1.Text = ASPxDateEdit1.Date.ToShortDateString();
                }
                if (string.IsNullOrEmpty(ASPxDateEdit2.Text))
                {
                    ASPxDateEdit2.Date = DateTime.Now.Date.AddMonths(1);
                    ASPxDateEdit2.Text = ASPxDateEdit2.Date.ToShortDateString();
                }

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


            }
            Specjalizacje_temp.DataBind();
            ustaw_baze();
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
                    ListaSkarg.Visible = false;
                    LinkButton10.Visible = false;
                    Panel11.Visible = true;
                }
            }
            catch (Exception ex)
            {


            }
            SqlDataSkargi2.DataBind();
            ListaSkarg.DataBind();
            string powP = (string)Session["od"];
            //  poczPowolania.Text = powP;
            //  koniecPowolania.Text = (string)Session["do"];
            var parametr = Request.QueryString["skarga"];
            if (parametr != null)
            {
                int idBieglego = cl.podajIdOsobyPoNumerzeSkargi(int.Parse(parametr));
                string nazwisko= cl.podajNazwiskoOsobyPoNumerzeSkargi(int.Parse(parametr));
                try
                {
                    var kontrolka = listaBieglych.FindFilterRowTemplateControl("Nazwisko");
                    GridViewDataColumn kolumna = (GridViewDataColumn)listaBieglych.Columns["Nazwisko"];
                    listaBieglych.ApplyFilterToColumn (kolumna, DevExpress.Data.Filtering.CriteriaOperator.Parse(nazwisko.Trim()).ToString());
         //           DevExpress.Data.Filtering.CriteriaOperator.Parse("nazwisko=" + nazwisko.Trim()).ToString()
                     
                }
                catch (Exception)
                {

                }
                if (idBieglego > 0)
                {
                    for (int i = 0; i < listaBieglych.VisibleRowCount; i++)
                    {
                        DataRow dr = listaBieglych.GetDataRow(i);
                        {
                            var inNr = dr[9].ToString();
                            if (idBieglego==int.Parse (inNr))
                            {

                            }
                        }
                    }

                }
            }
        }

        protected void obsługaArchiwum(object sender, EventArgs e)
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

        protected void wyswietlPoSpecjalizacji(object sender, EventArgs e)
        {
            ustaw_baze();
        }

        protected void uruchomFiltrowaniePoSpecjalizacji(object sender, EventArgs e)
        {
            if (CheckBox1.Checked)
            {
                DropDownList1.Enabled = true;
            }
            listaBieglych.DataBind();
        }

        protected void print_(object sender, EventArgs e)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                PrintableComponentLink pcl = new PrintableComponentLink(new PrintingSystem());
                listaBieglych.Columns[0].Visible = false;
                ASPxGridViewExporter1.FileName = "Wykaz biegłych";
                pcl.Component = ASPxGridViewExporter1;

                pcl.Margins.Left = pcl.Margins.Right = 50;
                pcl.Landscape = true;
                pcl.CreateDocument(false);
                pcl.PrintingSystem.Document.AutoFitToPagesWidth = 1;
                pcl.ExportToPdf(ms);
                WriteResponse(this.Response, ms.ToArray(), System.Net.Mime.DispositionTypeNames.Inline.ToString());
            }
            //  ASPxGridViewExporter1..GridView = ASPxGridView2;
            /*PxGridViewExporter1.PrintSelectCheckBox = true;
            ASPxGridViewExporter1.Landscape = true;

            
            ASPxGridViewExporter1.Landscape = true;
            
            ASPxGridViewExporter1.WritePdfToResponse();
            */
            listaBieglych.Columns["Column"].Visible = true;
        }
        public static void WriteResponse(HttpResponse response, byte[] filearray, string type)
        {
            response.ClearContent();
            response.Buffer = true;
            response.Cache.SetCacheability(HttpCacheability.Private);
            response.ContentType = "application/pdf";
            ContentDisposition contentDisposition = new ContentDisposition();
            contentDisposition.FileName = "test.pdf";
            contentDisposition.DispositionType = type;
            response.AddHeader("Content-Disposition", contentDisposition.ToString());
            response.BinaryWrite(filearray);
            HttpContext.Current.ApplicationInstance.CompleteRequest();
            try
            {
                response.End();
            }
            catch (System.Threading.ThreadAbortException)
            {
            }

        }
        protected void twórzZestawienie(object sender, EventArgs e)
        {
            internationalPDF();
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

        protected void makeExcell(object sender, EventArgs e)
        {
            ASPxGridViewExporter1.FileName = "Wykaz Biegłych";
            ASPxGridViewExporter1.WriteXlsToResponse();
        }

        protected void usunSpecjalizacje(object sender, DevExpress.Web.ASPxGridViewCustomButtonCallbackEventArgs e)
        {


            zapamietajDaneBieglego();

            //usuniecie specjalizacji
            ASPxGridView6.DataBind();
            try
            {
                if (e.ButtonID != "Clone2") return;
                //   copiedValues = new Hashtable();
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

        protected void dodajSpecjalizacje(object sender, DevExpress.Web.ASPxGridViewCustomButtonCallbackEventArgs e)
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

        protected void pokazDateZ(object sender, EventArgs e)
        {
            ASPxDateZakreslenie.Enabled = CheckBox3.Checked;

        }

        protected void zapiszSkarge(object sender, EventArgs e)
        {
            int user_id = int.Parse((string)Session["user_id"]);

            bool nowaSkarha = false;
            string kwerenda = string.Empty;
            try
            {
                nowaSkarha = (bool)Session["nowaSkarga"];
            }
            catch
            { }
            string idBieglego = "0";
            try
            {
                idBieglego = (string)Session["id_osoby"];

            }
            catch (Exception)
            {

            }
            if (nowaSkarha)
            {

                cl.dodajSkarge(int.Parse(idBieglego), txNumer.Text.Trim(), txRok.Text, txSygnatura.Text, ASPxDateWplyw.Date.ToShortDateString(), ASPxDatePismo.Date.ToShortDateString(), txWizytator.Text, uwagi.Text, CheckBox3.Checked, ASPxDateZakreslenie.Date.ToShortDateString(), user_id.ToString(), 1);
            }
            else
            {
                int idSkargi = int.Parse((string)Session["idSkargi"]);
                cl.usunSkarge(user_id.ToString(), idSkargi);
                cl.dodajSkarge(int.Parse(idBieglego), txNumer.Text.Trim(), txRok.Text, txSygnatura.Text, ASPxDateWplyw.Date.ToShortDateString(), ASPxDatePismo.Date.ToShortDateString(), txWizytator.Text, uwagi.Text, CheckBox3.Checked, ASPxDateZakreslenie.Date.ToShortDateString(), user_id.ToString(), 1);
            }
            Panel11.Visible = false;
            ListaSkarg.DataBind();
            ListaSkarg.Visible = true;

        }

        protected void anulowanieSkargi(object sender, EventArgs e)
        {
            Panel11.Visible = false;
            ListaSkarg.Visible = true;
        }

        protected void usunSkarge(object sender, EventArgs e)
        {
            int user_id = int.Parse((string)Session["user_id"]);

            int idSkargi = int.Parse((string)Session["idSkargi"]);
            cl.usunSkarge(user_id.ToString(), idSkargi);

            Panel11.Visible = false;
            ListaSkarg.DataBind();
            ListaSkarg.Visible = true;
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
            DateTime dataPoczatkuZawieszenia = DateTime.Parse("1900-01-01");
            DateTime dataKoncaZawieszenia = DateTime.Parse("1900-01-01");

            //==============================================
            //zapisz dane
            // zapis
            int err = 0;
            int er2 = 0;
            try
            {
                czyZaw = ddlZawiszenie.SelectedValue.ToString();
                if (czyZaw == "1")
                {
                    dataPoczatkuZawieszenia = poczatekZawieszeniaData.Date;
                    dataKoncaZawieszenia = koniecZawieszeniaData.Date;

                }

                int user_id = int.Parse((string)Session["user_id"]);
                // zapis wart. poczatkowych
                DateTime dat_1 = DateTime.Now;
                DateTime dat_2 = DateTime.Now;
                try
                {
                    dat_1 = ASPxDateEdit3.Date;
                }
                catch (Exception ex)
                {
                    //  dat_1= poczPowolania.Date;
                }

                try
                {
                    dat_2 = DateTime.Parse(powolanieDo.Trim());
                }
                catch (Exception ex)
                {
                    dat_2 = ASPxDateEdit4.Date;
                }

                if (dat_1 >= dat_2) err = 2;
                log = log + "powolanie Od =  " + powolanieOd + "<br/>";
                log = log + "powolanie Do =  " + powolanieDo + "<br/>";
                log = log + "dat_1 =  " + dat_1.ToString() + "<br/>";
                log = log + "dat_2 =  " + dat_2.ToString() + "<br/>";
                log = log + "DataZawieszenia =  " + dataPoczatkuZawieszenia + "<br/>";

                if (err == 0)
                {
                    log = log + "err=  " + err + "<br/>";
                    string txt = cl.modyfikuj_osobe(idBieglego, user_id, imie, nazwisko, ulica1, kod1, miejscowosc1, DateTime.Parse(powolanieOd), DateTime.Parse(powolanieDo), tytul, pesel, emil, uwagi, int.Parse(czyZaw), ulica2, kod2, miejscowosc2, tel1, tel2, dataPoczatkuZawieszenia, dataKoncaZawieszenia, txspecjalizacja_opis.Text.Trim(), 1);
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
                //ASPxGridView8.DataBind();
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
            listaBieglych.DataBind();
        }

        protected void usunBieglego(object sender, EventArgs e)
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

        protected void zmianaZawieszenia(object sender, EventArgs e)
        {
            zmianaWyswietlaniaZawieszenia();
        }



        protected void ASPxDateEdit7_DateChanged(object sender, EventArgs e)
        {

        }

        protected void dodajPowolanie(object sender, EventArgs e)
        {

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

        protected void zmienPowolanie(object sender, EventArgs e)
        {
            DateTime begin_ = ASPxDateEdit1.Date;
            DateTime end_ = ASPxDateEdit2.Date;
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

        protected void usunPowolanie(object sender, EventArgs e)
        {

        }

        protected void wybiezPowolanie(object sender, EventArgs e)
        {
            //powolanie do edycji
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

        }

        protected void nowaSkarga(object sender, EventArgs e)
        {
            for (int i = DateTime.Now.Year - 10; i < DateTime.Now.Year + 2; i++)
            {

                txRok.Items.Add(new System.Web.UI.WebControls.ListItem(i.ToString()));
            }
            try
            {
                txRok.SelectedIndex = txRok.Items.IndexOf(txRok.Items.FindByValue(DateTime.Now.Year.ToString()));

            }
            catch (Exception)
            {


            }
            Panel11.Visible = true;
            ListaSkarg.Visible = false;
            Session["nowaSkarga"] = true;

            txNumer.Text = cl.PodajNumerNowejSkargi(int.Parse(txRok.SelectedValue.ToString()));
            txWizytator.Text = "";
            uwagi.Text = "";
            txSygnatura.Text = "";
            CheckBox3.Checked = false;
            ASPxDateWplyw.Date = DateTime.Now;
            ASPxDatePismo.Date = DateTime.Now;
            ASPxDateZakreslenie.Date = DateTime.Now;
            ASPxDateZakreslenie.Enabled = false;

        }

        protected void ASPxGridView7_CustomButtonCallback(object sender, DevExpress.Web.ASPxGridViewCustomButtonCallbackEventArgs e)
        {
            Session["nowaSkarga"] = false;
            Panel11.Visible = true;
            ListaSkarg.Visible = false;
            var VisibleIndex = e.VisibleIndex;
            for (int i = DateTime.Now.Year - 10; i < DateTime.Now.Year + 2; i++)
            {

                txRok.Items.Add(new System.Web.UI.WebControls.ListItem(i.ToString()));
            }
            var dane = ListaSkarg.GetDataRow(VisibleIndex);

            txNumer.Text = dane[0].ToString();
            try
            {
                txRok.SelectedIndex = txRok.Items.IndexOf(txRok.Items.FindByValue(dane[1].ToString()));
            }
            catch (Exception)
            { }
            txSygnatura.Text = dane[4].ToString();
            txWizytator.Text = dane[5].ToString();

            ASPxDateWplyw.Date = (DateTime)dane[2];
            ASPxDatePismo.Date = (DateTime)dane[3];
            ASPxDateZakreslenie.Date = (DateTime)dane[7];
            CheckBox3.Checked = false;
            if (dane[6].ToString() == "1")
            {
                CheckBox3.Checked = true;
            }
            uwagi.Text = dane[8].ToString();
            Session["idSkargi"] = dane[9].ToString();
        }

        protected void otworzPopup(int idBieglego, short zakladka)
        {

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
            catch (Exception ex)
            {
                popup.ShowOnPageLoad = false;
            }

            string idOsoby = string.Empty;
            try
            {
                idOsoby = (string)Session["id_osoby"];
            }
            catch
            { }

            if (DropDownList4.Items.Count == 0)
            {
                int idStatystyki = 0;
                DropDownList4.DataBind();
                try
                {
                    idStatystyki = (int)Session["idStatystyki"];
                    
                }
                catch
                {
                    
                }
                try
                {
                    string txt = (string)Session["idStatystykitxt"];
                    idStatystyki = int.Parse(txt);
                }
                catch (Exception)
                {
                    
                }
                try
                {
                    DropDownList4.SelectedIndex = idStatystyki;
                }
                catch (Exception)
                {
                    
                }
            }

            try
            {
                string kwerenda = DropDownList4.SelectedValue.ToString();
                pokazStatystyki(kwerenda,idBieglego);
            }
            catch
            {


            }
            if (idBieglego.ToString() != idOsoby)
            {
                Panel11.Visible = false;
                ListaSkarg.Visible = true;
                string txt = DateTime.Now.Ticks.ToString();
                Session["sesja"] = txt;
                cl.odczytaj_specjalizacje_osoby(idBieglego.ToString(), (string)Session["sesja"]);
                // przenosze wszystko tu 
                Session["id_osoby"] = idBieglego.ToString().Trim();
                DataTable czlowieczek = cl.dane_osobowe(idBieglego.ToString().Trim());
                DataTable korespondencja = cl.dane_korespondencyjne(idBieglego.ToString().Trim());
                //tel1 ,tel2 ,email ,adr_kores ,kod_poczt_kor ,miejscowosc_kor 
                SqlDataSource2.DataBind();
                //dodanie tymczasowych specjalizacji
                ASPxPageControl1.TabIndex = zakladka;
                GridView25.DataBind();
                Specjalizacje_temp.DataBind();
                ASPxGridView6.DataBind();
                SqlDataSkargi2.DataBind();
                ListaSkarg.DataBind();
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
                    zmianaWyswietlaniaZawieszenia();
                    if (zawieszenie1)
                    {

                        ddlZawiszenie.SelectedIndex = 1;
                        poczatekZawieszeniaData.Visible = true;
                        poczatekZawieszeniaData.Date = DateTime.Parse(czlowieczek.Rows[0]["d_zawieszenia"].ToString());
                        poczatekZawieszeniaData.Text = poczatekZawieszeniaData.Date.ToShortDateString();
                        koniecZawieszeniaData.Visible = true;
                        koniecZawieszeniaData.Date = DateTime.Parse(czlowieczek.Rows[0]["dataKoncaZawieszenia"].ToString());
                        koniecZawieszeniaData.Text = koniecZawieszeniaData.Date.ToShortDateString();

                    }
                    else
                    {
                        ddlZawiszenie.SelectedIndex = 0;
                        poczatekZawieszeniaData.Visible = false;
                        koniecZawieszeniaData.Visible = false;
                    }


                    zapamietajDaneBieglego();
                    GridView25.DataBind();
                }
                catch (Exception ex)
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
                        poczatekZawieszeniaData.Text = dt.Rows[0]["dataZawieszeniaPoczatek"].ToString();
                        uwagi.Text = dt.Rows[0]["uwagi"].ToString();
                        //    poczPowolania.Date = DateTime.Parse(dt.Rows[0]["dataPowolaniaPoczatek"].ToString());
                        //                        koniecPowolania.Date = DateTime.Parse(dt.Rows[0]["dataPowolaniaKoniec"].ToString());


                    }
                    catch
                    { }
                }
            }

        }
        protected void callbackPanel_Callback(object sender, DevExpress.Web.CallbackEventArgsBase e)
        {
            // tu jest przekazywany parametr


            int employeeId = 0;
            try
            {
                employeeId = Convert.ToInt32(e.Parameter);
                Session["employeeId"] = employeeId;
            }
            catch (Exception)
            {
                if (Session["employeeId"] != null)
                {
                    employeeId = (int)Session["employeeId"];
                }

            }
            otworzPopup(employeeId, 1);



        }

        protected void popup_WindowCallback(object source, DevExpress.Web.PopupWindowCallbackArgs e)
        {

        }

        protected void popup_Unload(object sender, EventArgs e)
        {
            listaBieglych.DataBind();
            listaBieglych0.DataBind();
        }

        protected void popup_PopupElementResolve(object sender, DevExpress.Web.ControlResolveEventArgs e)
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
            dT.Columns.Add("dataZawieszeniaPoczatek", typeof(string));
            dT.Columns.Add("dataZawieszeniaKoniec", typeof(string));
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
                biegly["zawieszenie"] = ddlZawiszenie.SelectedValue.ToString().Trim();
                biegly["dataZawieszeniaPoczatek"] = poczatekZawieszeniaData.Text;
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
        protected void ustaw_baze()
        {

            string kwerenda = string.Empty;
            if (CheckBox4.Checked)
            {
                //archiwum
                daneBieglych = null;
                if (CheckBox1.Checked)
                {
                    //z kategoriami
                    DropDownList1.Enabled = true;
                    string idSspecjalizacji = DropDownList1.SelectedValue.ToString().Trim();
                    kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby WHERE (tbl_specjalizacje_osob.id_specjalizacji = " + DropDownList1.SelectedValue.ToString().Trim() + ") and View_listaBieglych.data_koncowa<=getdate() order by nazwisko";

                }
                else
                {
                    //bez kategorii
                    DropDownList1.Enabled = false;
                    kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby where  View_listaBieglych.data_koncowa <= getdate() order by nazwisko";

                }
                daneBieglychArchiwum.SelectCommand = kwerenda;
                listaBieglych0.DataBind();

            }
            else
            {
                //     archiwum1 = null;
                if (CheckBox1.Checked)
                {
                    //z kategoriami

                    DropDownList1.Enabled = true;
                    string idSspecjalizacji = DropDownList1.SelectedValue.ToString().Trim();
                    kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby where (tbl_specjalizacje_osob.id_specjalizacji = " + DropDownList1.SelectedValue.ToString().Trim() + ") and View_listaBieglych.data_koncowa > getdate() order by nazwisko";

                }
                else
                {
                    DropDownList1.Enabled = false;
                    //bez kategorii
                    DropDownList1.Enabled = false;
                    kwerenda = "SELECT DISTINCT View_listaBieglych.tytul as tytul, View_listaBieglych.imie as imie,  View_listaBieglych.uwagi as uwagi, View_listaBieglych.nazwisko as nazwisko, View_listaBieglych.adres as adres,View_listaBieglych.tel1 as telefon, View_listaBieglych.specjalizacja_opis as opis , View_listaBieglych.data_poczatkowa as data_poczatkowa, View_listaBieglych.data_koncowa as data_koncowa,View_listaBieglych.ident as ident, View_listaBieglych.adres2 as adres2, View_listaBieglych.zawieszony as zawieszony, View_listaBieglych.specjalizacjeWidok as specjalizacjeWidok FROM  View_listaBieglych LEFT OUTER JOIN tbl_specjalizacje_osob ON View_listaBieglych.ident = tbl_specjalizacje_osob.id_osoby where  View_listaBieglych.data_koncowa > getdate() order by nazwisko";

                }

                daneBieglych.SelectCommand = kwerenda;
                daneBieglych.DataBind();
                listaBieglych.DataBind();

            }

            //  maintable = cl.wyciagnijDaneBieglych(kwerenda);



        } // end of ustaw_baze
        private void zmianaWyswietlaniaZawieszenia()
        {
            if (ddlZawiszenie.SelectedIndex == 0)
            {
                poczatekZawieszeniaData.Visible = false;
                koniecZawieszeniaData.Visible = false;
                lblPoczatekZawieszenia.Visible = false;

                lblKoniecZawieszenia.Visible = false;

            }
            else
            {
                poczatekZawieszeniaData.Visible = true;
                koniecZawieszeniaData.Visible = true;
                lblPoczatekZawieszenia.Visible = true;
                lblKoniecZawieszenia.Visible = true;
                if (poczatekZawieszeniaData.Date.Year < 1910)
                {
                    poczatekZawieszeniaData.Date = DateTime.Now;
                    poczatekZawieszeniaData.Text = poczatekZawieszeniaData.Date.ToShortDateString();
                }
                if (koniecZawieszeniaData.Date.Year < 1910)
                {
                    koniecZawieszeniaData.Date = poczatekZawieszeniaData.Date.AddMonths(1);
                    koniecZawieszeniaData.Text = koniecZawieszeniaData.Date.ToShortDateString();
                }
            }

        }


        public void changeQuerry(object sender, EventArgs e)
        {
            Session["ddl2"] = DropDownList3.SelectedIndex;
            Label5.Text = DropDownList3.SelectedValue.ToString();
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
        public void changeQuerry2(object sender, EventArgs e)
        {
            Session["ddl2"] = DropDownList4.SelectedIndex;
            try
            {
                int idBieglego =int.Parse ( (string)Session["id_osoby"]);
                string querry = DropDownList4.SelectedValue.ToString();
                pokazStatystyki(querry, idBieglego);
            }
            catch
            { }

        }
        protected void pokazStatystyki(string kwerenda,int idBieglego)
        {
            
            DataTable dT = cl.tabelaStatystyczna(kwerenda, idBieglego.ToString ());
            GridView26.DataSource = null;
            GridView26.DataSourceID = null;
            GridView26.AutoGenerateColumns = true;
            GridView26.DataSource = dT;
            GridView26.DataBind();

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
         
            string sylfaenpath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\sylfaen.ttf";
            BaseFont sylfaen = BaseFont.CreateFont(sylfaenpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            Font head = new Font(sylfaen, 12f, Font.NORMAL, BaseColor.BLACK);
            Font normal = new Font(sylfaen, 10f, Font.NORMAL, BaseColor.BLACK);
            Font underline = new Font(sylfaen, 10f, Font.UNDERLINE, BaseColor.BLACK);




            //    var cl.plFontBIG = FontFactory.GetFont(BaseFont.HELVETICA, BaseFont.CP1257, 35, Font.NORMAL);



            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10f, 10f, 10f, 0f);
            string path = Server.MapPath(@"~//pdf");// Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string fileName = path + "//zestawienie_Specjalizacji_" + DateTime.Now.ToString().Replace(":", "-") + ".pdf";
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(fileName, FileMode.Create));
            pdfDoc.Open();


            pdfDoc.AddTitle("zestawienie_Specjalizacji");
            pdfDoc.AddCreationDate();


            pdfDoc.Open();

            Response.ContentType = "application/pdf";
            Response.AddHeader("content-disposition", "attachment;" + "filename=zestawienie_Specjalizacji.pdf");


            PdfPTable fitst = new PdfPTable(1);
            fitst.DefaultCell.Border = Rectangle.NO_BORDER;
            PdfPCell cell = new PdfPCell(new Paragraph(" ", head));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("LISTA", head));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;

            cell.Border = Rectangle.NO_BORDER;
            fitst.AddCell(cell);
            string text = "BIEGŁYCH SĄDOWYCH ";

            cell = new PdfPCell(new Paragraph(text, head));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("", head));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("SĄDU OKRĘGOWEGO", head));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            fitst.AddCell(cell);
            pdfDoc.Add(fitst);
            pdfDoc.NewPage();
            PdfPTable tab = new PdfPTable(3);
            int[] tblWidth2 = { 10, 80, 10 };
            tab.SetWidths(tblWidth2);
            cell = new PdfPCell(new Paragraph("", cl.plFontBIG));
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            tab.AddCell(cell);
            tab.AddCell(cell);
            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Lp.", cl.plFont2));

            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Zakres", cl.plFont2));

            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Numer strony", cl.plFont2));

            tab.AddCell(cell);

            int biezacaStrona = 0;
            int il = 0;
            foreach (DataRow dRow in specjalizacjeWyliczenie.Rows)
            {
                il++;
                tab.AddCell(new Paragraph(il.ToString(), cl.plFont2));
                tab.AddCell(new Paragraph(dRow[0].ToString(), cl.plFont2));
                biezacaStrona = biezacaStrona + int.Parse(dRow[1].ToString());
                tab.AddCell(new Paragraph(biezacaStrona.ToString(), cl.plFont2));
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
                pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, cl.plFont1)));
                pdfDoc.Add(new Paragraph(" "));
                PdfPTable tabelaGlowna = new PdfPTable(4);
                int[] tblWidth = { 8, 30, 30, 32 };
                tabelaGlowna.SetWidths(tblWidth);
                tabelaGlowna.AddCell(new Paragraph("Lp.", cl.plFont2));
                tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", cl.plFont2));
                tabelaGlowna.AddCell(new Paragraph("Adres- telefon", cl.plFont2));
                tabelaGlowna.AddCell(new Paragraph("Zakres", cl.plFont2));


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
                            pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji + " ciąg dalszy", cl.plFont1)));
                            pdfDoc.Add(new Paragraph(" "));

                            tabelaGlowna = new PdfPTable(4);

                            tabelaGlowna.AddCell(new Paragraph("Lp.", cl.plFont2));
                            tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", cl.plFont2));
                            tabelaGlowna.AddCell(new Paragraph("Adres- telefon", cl.plFont2));

                            tabelaGlowna.AddCell(new Paragraph("Zakres", cl.plFont2));
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

            pdfDoc.Close();
            pdfDoc.Close();
            WebClient client = new WebClient();
            Byte[] buffer = client.DownloadData(fileName);
            if (buffer != null)
            {
                Response.ContentType = "application/pdf";
                Response.AddHeader("content-lenght", buffer.Length.ToString());
                Response.BinaryWrite(buffer);
            }


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

            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10f, 10f, 10f, 0f);
            //  PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream("C:\\temp\\" + filename, FileMode.Create));

            string path = Server.MapPath(@"~//pdf"); //Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments );
            string fileName = path + "//zestawienie_Specjalizacji_" + DateTime.Now.ToString().Replace(":", "-") + ".pdf";
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(fileName, FileMode.Create));
            pdfDoc.Open();


            pdfDoc.AddTitle("zestawienie_Specjalizacji");
            pdfDoc.AddCreationDate();

            PdfPTable fitst = new PdfPTable(1);
            fitst.DefaultCell.Border = Rectangle.NO_BORDER;
            PdfPCell cell = new PdfPCell(new Paragraph(" ", cl.plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);
            cell = new PdfPCell(new Paragraph("LISTA", cl.plFontBIG));


            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;

            cell.Border = Rectangle.NO_BORDER;
            fitst.AddCell(cell);
            cell = new PdfPCell(new Paragraph("BIEGŁYCH SĄDOWYCH", cl.plFontBIG));

            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("", cl.plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.Border = Rectangle.NO_BORDER;
            cell.FixedHeight = 100;
            fitst.AddCell(cell);

            cell = new PdfPCell(new Paragraph("SĄDU OKRĘGOWEGO", cl.plFontBIG));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            fitst.AddCell(cell);
            pdfDoc.Add(fitst);
            pdfDoc.NewPage();
            PdfPTable tab = new PdfPTable(3);
            int[] tblWidth2 = { 10, 80, 10 };
            tab.SetWidths(tblWidth2);
            cell = new PdfPCell(new Paragraph("", cl.plFontBIG));
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            tab.AddCell(cell);
            tab.AddCell(cell);
            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Lp.", cl.plFont2));

            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Zakres", cl.plFont2));

            tab.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Numer strony", cl.plFont2));

            tab.AddCell(cell);

            int biezacaStrona = 0;
            int il = 0;
            foreach (DataRow dRow in specjalizacjeWyliczenie.Rows)
            {
                il++;
                tab.AddCell(new Paragraph(il.ToString(), cl.plFont2));
                tab.AddCell(new Paragraph(dRow[0].ToString(), cl.plFont2));
                biezacaStrona = biezacaStrona + int.Parse(dRow[1].ToString());
                tab.AddCell(new Paragraph(biezacaStrona.ToString(), cl.plFont2));
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
                    pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, cl.plFont3)));
                    pdfDoc.Add(new Paragraph(" "));
                    PdfPTable tabelaGlowna = new PdfPTable(4);
                    int[] tblWidth = { 8, 30, 30, 32 };
                    tabelaGlowna.AddCell(new Paragraph("Lp.", cl.plFont2));
                    tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", cl.plFont2));
                    tabelaGlowna.AddCell(new Paragraph("Adres- telefon", cl.plFont2));

                    tabelaGlowna.AddCell(new Paragraph("Zakres", cl.plFont2));
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
                                pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji + " ciąg dalszy", cl.plFont3)));
                                pdfDoc.Add(new Paragraph(" "));

                                tabelaGlowna = new PdfPTable(4);
                                tabelaGlowna.SetWidths(tblWidth);
                                tabelaGlowna.AddCell(new Paragraph("Lp.", cl.plFont2));
                                tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", cl.plFont2));
                                tabelaGlowna.AddCell(new Paragraph("Adres- telefon", cl.plFont2));


                                tabelaGlowna.AddCell(new Paragraph("Zakres", cl.plFont2));
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


            pdfDoc.Close();
            WebClient client = new WebClient();
            Byte[] buffer = client.DownloadData(fileName);
            if (buffer != null)
            {
                Response.ContentType = "application/pdf";
                Response.AddHeader("content-lenght", buffer.Length.ToString());
                Response.BinaryWrite(buffer);
            }


        }
        protected PdfPTable generujCzescRaportu(DataTable biegli, string specjalizacje)
        {

            int[] tblWidth = { 8, 30, 30, 32 };


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
                catch (Exception ex)
                {


                }

                string innerTable = imie + Environment.NewLine + nazwisko + Environment.NewLine + tytul + Environment.NewLine + "kadencja do dnia: " + dataKonca;
                tabelaGlowna.AddCell(new Paragraph(iterator.ToString(), cl.plFont1));
                tabelaGlowna.AddCell(new Paragraph(innerTable, cl.plFont1));
                string ulica = biegly[3].ToString();
                string kod = biegly[4].ToString();
                string miejscowosc = biegly[5].ToString();
                string tel = biegly[8].ToString();
                string adresTable = ulica + Environment.NewLine + kod + " " + miejscowosc + Environment.NewLine + tel;
                tabelaGlowna.AddCell(new Paragraph(adresTable, cl.plFont1));
                string specki = string.Empty;
                string specjalizacjaOpis = cl.odczytaj_specjalizacje_osobyOpis(Idbieglego);
                // tabelaGlowna.AddCell(new Paragraph(specjalizacjaOpis, cl.plFont2));
                foreach (DataRow specRow in listaSpecjalizacjiBieglego.Rows)
                {
                    specki = specki + specRow[0].ToString().ToLower() + "; ";
                }
                specki = specki + specjalizacjaOpis;
                tabelaGlowna.AddCell(new Paragraph(specki, cl.plFont1));
            }


            return tabelaGlowna;

        }

        protected void ASPxGridViewExporter1_RenderBrick(object sender, ASPxGridViewExportRenderingEventArgs e)
        {
            GridViewDataColumn dataColumn = e.Column as GridViewDataColumn;
            if (e.RowType == GridViewRowType.Data && dataColumn != null && dataColumn.FieldName == "Z")
            {
                if (e.Text == "Niezaznaczony")
                {
                    e.Text = "";
                }
                else
                {
                    e.Text = "Zawieszony";
                }
            }
        }

    

        static private void internationalPDF()
        {
            try
            {
                iTextSharp.text.Document document = new iTextSharp.text.Document(PageSize.A4, 72, 65, 72, 65);

                string filename = "international" + DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Day.ToString() + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + DateTime.Now.Second.ToString() + ".pdf";
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream("C:\\temp\\" + filename, FileMode.Create));
                document.AddAuthor("RussianPDFtest");
                document.AddTitle("Załączone zdjęcie");
                document.AddCreationDate();

                string sylfaenpath = Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\sylfaen.ttf";
                BaseFont sylfaen = BaseFont.CreateFont(sylfaenpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                Font head = new Font(sylfaen, 12f, Font.NORMAL, BaseColor.BLUE);
                Font normal = new Font(sylfaen, 10f, Font.NORMAL, BaseColor.BLACK);
                Font underline = new Font(sylfaen, 10f, Font.UNDERLINE, BaseColor.BLACK);

                document.Open();
                document.Add(new Paragraph("CZ-Reservation / Załączone zdjęcie          BU-16122011-37", head));
                document.Add(new Paragraph(" ", normal));
                document.Add(new Paragraph(" ", normal));
                document.Add(new Paragraph("Informacje te są przesyłane bezpośrednio do właściciela:", underline));
                document.Add(new Paragraph(" ", normal));
                PdfPTable table1 = new PdfPTable(2);
                table1.TotalWidth = document.PageSize.Width - 72f - 65f;
                table1.LockedWidth = true;
                float[] widths1 = new float[] { 1f, 4f };
                table1.SetWidths(widths1);
                table1.HorizontalAlignment = 0;
                PdfPCell table1cell11 = new PdfPCell(new Phrase("Usunięcie się powiodło:", normal));
                table1cell11.Border = 0;
                table1.AddCell(table1cell11);
                PdfPCell table1cell12 = new PdfPCell(new Phrase("Ferienhaus 'Waldesruh'", normal));
                table1cell12.Border = 0;
                table1.AddCell(table1cell12);
                PdfPCell table1cell21 = new PdfPCell(new Phrase("Menu główne:", normal));
                table1cell21.Border = 0;
                table1.AddCell(table1cell21);
                PdfPCell table1cell22 = new PdfPCell(new Phrase("15344 Strausberg, Am Marienberg 45", normal));
                table1cell22.Border = 0;
                table1.AddCell(table1cell22);
                PdfPCell table1cell31 = new PdfPCell(new Phrase("D'évènement:", normal));
                table1cell31.Border = 0;
                table1.AddCell(table1cell31);
                PdfPCell table1cell32 = new PdfPCell(new Phrase("czr04012012", normal));
                table1cell32.Border = 0;
                table1.AddCell(table1cell32);
                PdfPCell table1cell41 = new PdfPCell(new Phrase("Centres bien-être:", normal));
                table1cell41.Border = 0;
                table1.AddCell(table1cell41);
                PdfPCell table1cell42 = new PdfPCell(new Phrase("12.02.2012", normal));
                table1cell42.Border = 0;
                table1.AddCell(table1cell42);
                PdfPCell table1cell51 = new PdfPCell(new Phrase("Accès handicapés:", normal));
                table1cell51.Border = 0;
                table1.AddCell(table1cell51);
                PdfPCell table1cell52 = new PdfPCell(new Phrase("18.02.2012", normal));
                table1cell52.Border = 0;
                table1.AddCell(table1cell52);
                PdfPCell table1cell61 = new PdfPCell(new Phrase("Août:", normal));
                table1cell61.Border = 0;
                table1.AddCell(table1cell61);
                PdfPCell table1cell62 = new PdfPCell(new Phrase("5", normal));
                table1cell62.Border = 0;
                table1.AddCell(table1cell62);
                document.Add(table1);
                document.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Try-Catch-Fehler!");
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }

        protected void zmienRokSkargi(object sender, EventArgs e)
        {
            string rok = txRok.SelectedValue;
            int rokD = DateTime.Now.Year;
            try
            {
                rokD = int.Parse(rok);
            }
            catch (Exception)
            { }
            txNumer.Text = cl.PodajNumerNowejSkargi(rokD);
        }



        protected void zmienStatystyke(object sender, EventArgs e)
        {

            var cos = sender;
        }

        

        protected void DropDownList4_SelectedIndexChanged(object sender, EventArgs e)
        {
            Session["idStatystyki"] = DropDownList4.SelectedIndex;
            int idBieglego = int.Parse((string)Session["id_osoby"]);
            Session["idStatystykitxt"] = DropDownList4.SelectedIndex.ToString();

            string kwerenda = DropDownList4.SelectedValue.ToString();
            pokazStatystyki(kwerenda, idBieglego);

        }

        protected void listaBieglych_BeforePerformDataSelect(object sender, EventArgs e)
        {

        }

        protected void makeExcellforBIP(object sender, EventArgs e)
        {
           
        }
    }
}