using DevExpress.Web;
using DevExpress.Web.ASPxTreeList;
using DevExpress.Web.Data;
using DevExpress.XtraPrinting;
using DevExpress.XtraPrinting.Native;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Web;
using System.Web.UI.WebControls;

namespace wab2018
{
    public partial class biegliLista : System.Web.UI.Page
    {
        private readonly nowiMediatorzy nm = new nowiMediatorzy();
        private cm Cm = new cm();
        private Class2 cl = new Class2();
        private string pesel = string.Empty;
        protected void Page_Load(object sender, EventArgs e)
        {
            //GridViewFeaturesHelper.SetupGlobalGridViewBehavior(grid);
           
            if (!IsPostBack)
            {
                if (Session["user_id"] == null)
                {
                    Server.Transfer("logowanie.aspx");
                }

                string rola = (string)Session["rola"];
                switch (rola)
                {
                    case "2":
                        {
                            grid.Visible = true;
                            grid0.Visible = false;
                        }
                        break;

                    case "3":
                        {
                            grid.Visible = true;
                            grid0.Visible = false;
                        }
                        break;

                    case "4":
                        {
                            grid.Visible = true;
                            grid0.Visible = false;
                        }
                        break;

                    default:
                        {
                            grid.Visible = false;
                            grid0.Visible = true;
                        }
                        break;
                }
            }

         
         
       

            ustawKwerendeOdczytu();
            var parametr = Request.QueryString["skarga"];
            if (parametr != null)
            {
                string staraSkarhe = (string)Session["skargaId"];
                if (staraSkarhe != parametr)
                {
                    Session["flagaSkarg"] = 0;
                }
                Session["skargaId"] = parametr;
                int flagaSkarg = 0;
                try
                {
                    flagaSkarg = (int)Session["flagaSkarg"];
                }
                catch (Exception)
                {
                }

                if (flagaSkarg == 0)
                {
                    int idBieglego = cl.podajIdOsobyPoNumerzeSkargi(int.Parse(parametr));
                    Session["id_osoby"] = idBieglego;
                    string nazwisko = cl.podajNazwiskoOsobyPoNumerzeSkargi(int.Parse(parametr));
                    int visibleIndex = grid.FindVisibleIndexByKeyValue(idBieglego);

                    //Remove the items
                    grid.Selection.SelectRow(visibleIndex);
                    grid.StartEdit(visibleIndex);
                    try
                    {
                        ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
                        pageControl.ActiveTabIndex = 6;
                        Session["flagaSkarg"] = 1;
                    }
                    catch (Exception)
                    {
                    }
                }
                Session["skargaId"] = parametr;
            }
            try
            {
                AppSettingsReader app = new AppSettingsReader();
                string Theme =(string) app.GetValue("stylTabeli", typeof(string));
                grid.Theme = Theme;
                grid0.Theme = Theme;
            }
            catch (Exception)
            { }
           
         
        }

        private DataTable GenerujListeSpecjalizacji()
        {
            DataTable SpecjalizacjeTemp = cl.Specjalizacje();
            DataTable GrupySpecjalizacjiTemp = cl.GrupySpecjalizacji();
            DataTable Specki = new DataTable();
            Specki.Columns.Add("Nazwa", typeof(String));
            Specki.Columns.Add("nr", typeof(int));
            Specki.Columns.Add("Grupa", typeof(String));
            Specki.Columns.Add("NrGrupy", typeof(int));
            if (SpecjalizacjeTemp != null)
            {
                DataRow pojedynczaSpecka = Specki.NewRow();
                pojedynczaSpecka[0] = "Wszystkie specjalizacje"; 
                pojedynczaSpecka[1] = 0;
                pojedynczaSpecka[2] = 0;
                Specki.Rows.Add(pojedynczaSpecka);



                foreach (DataRow Specjalizacja in SpecjalizacjeTemp.Rows)
                {
                    pojedynczaSpecka = Specki.NewRow();
                    pojedynczaSpecka[0] = Specjalizacja[0];
                    pojedynczaSpecka[1] = Specjalizacja[1];
                    pojedynczaSpecka[2] = Specjalizacja[2];
                    pojedynczaSpecka[3] = Specjalizacja[3];
                    Specki.Rows.Add(pojedynczaSpecka);
                }
            }
            return Specki;
        }
       

        protected void updateMediatora(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {
            zawieszenia zaw = new zawieszenia();
            string txt = mediatorzy.SelectCommand;
            //dane osobowe
            string tytul = nm.controlText("txTytul", grid);
            string imie = nm.controlText("txImie", grid);
            string nazwisko = nm.controlText("txNazwisko", grid);
            string data_poczatkowa = nm.controlTextDate("txPoczatekPowolania", grid);
            string data_koncowa = nm.controlTextDate("txDataKoncaPowolania", grid);
             
            e.NewValues["tytul"] = nm.controlText("txTytul", grid);
            e.NewValues["imie"] = nm.controlText("txImie", grid);
            e.NewValues["nazwisko"] = nm.controlText("txNazwisko", grid);
            e.NewValues["data_poczatkowa"] = nm.controlTextDate("txPoczatekPowolania", grid);
            e.NewValues["data_koncowa"] = nm.controlTextDate("txDataKoncaPowolania", grid);
            bool zawieszenie = nm.controlCheckBox("zawiszeniaCbox", grid);

            e.NewValues["czy_zaw"] = zawieszenie;
            if (zawieszenie)
            {
                e.NewValues["d_zawieszenia"] = nm.controlTextDate("txDataPoczatkuZawieszenia", grid);
                e.NewValues["dataKoncaZawieszenia"] = nm.controlTextDate("txDataKoncaZawieszenia", grid);
            }
            else
            {
                e.NewValues["d_zawieszenia"] = DateTime.Now;
                e.NewValues["dataKoncaZawieszenia"] = DateTime.Now;

            }
            if (nm.controlText("txPESEL", grid) == null)
            {
                e.NewValues["Pesel"] = 0;
            }
            else
            {
                e.NewValues["Pesel"] = nm.controlText("txPESEL", grid);
            }
            //dane adresowe
            e.NewValues["ulica"] = nm.controlText("txAdres", grid);
            e.NewValues["kod_poczt"] = nm.controlText("txKodPocztowy", grid);
            e.NewValues["miejscowosc"] = nm.controlText("txMiejscowosc", grid);
            e.NewValues["tel1"] = nm.controlText("txTelefon1", grid);
            e.NewValues["tel2"] = nm.controlText("txTelefon2", grid);
            e.NewValues["email"] = nm.controlText("txEmail", grid);
            //dane korespondencyjne
            e.NewValues["adr_kores"] = nm.controlText("txAdresKorespondencyjny", grid);
            e.NewValues["kod_poczt_kor"] = nm.controlText("txKodPocztowyKorespondencyjny", grid);
            e.NewValues["miejscowosc_kor"] = nm.controlText("txMiejscowoscKorespondencyjny", grid);
            // uwagi i specjalizacje
            e.NewValues["uwagi"] = nm.controlTextMemo("txUwagi", grid);
            e.NewValues["specjalizacja_opis"] = nm.controlTextMemo("txSpecjalizacjeOpis", grid);
            e.NewValues["instytucja"] = nm.controlText("txInstytucja", grid);
        }

        protected void InsertData(object sender, ASPxDataInitNewRowEventArgs e)
        {
            e.NewValues["data_poczatkowa"] = DateTime.Now.Date;
            DateTime dataKoncz = DateTime.Parse(DateTime.Now.AddYears(5).Year.ToString() + "-12-31");
            e.NewValues["data_koncowa"] = dataKoncz;
            //d_zawieszenia
            e.NewValues["d_zawieszenia"] = DateTime.Now;
            e.NewValues["dataKoncaZawieszenia"] = dataKoncz;
            string userId = (string)Session["user_id"];
            string idOsoby = cl.dodaj_osobe(1, int.Parse(userId));

            Session["idMediatora"] = idOsoby;
            Session["id_osoby"] = idOsoby;
            Session["czy_zaw"] = "0";
        }

        protected void grid_StartRowEditing(object sender, ASPxStartRowEditingEventArgs e)
        {
            // rozpoczecie edycji
            System.Web.UI.Page page = HttpContext.Current.CurrentHandler as System.Web.UI.Page;
       
            string id = e.EditingKeyValue.ToString();
            Session["idMediatora"] = id;
            Session["id_osoby"] = id;
            
           
        }

        protected void grid_RowInserting(object sender, ASPxDataInsertingEventArgs e)
        {
            //dane osobowe
            e.NewValues["tytul"] = nm.controlText("txTytul", grid);
            e.NewValues["imie"] = nm.controlText("txImie", grid);
            e.NewValues["nazwisko"] = nm.controlText("txNazwisko", grid);
            e.NewValues["data_poczatkowa"] = nm.controlTextDate("txPoczatekPowolania", grid);
            e.NewValues["data_koncowa"] = nm.controlTextDate("txDataKoncaPowolania", grid);
            var cos = nm.controlCheckbox("zawiszeniaCbox", grid);

            e.NewValues["czy_zaw"] = false;
            e.NewValues["d_zawieszenia"] = nm.controlTextDate("txDataPoczatkuZawieszenia", grid);
            e.NewValues["dataKoncaZawieszenia"] = nm.controlTextDate("txDataKoncaZawieszenia", grid);



            if (nm.controlText("txPESEL", grid) == null)
            {
                e.NewValues["Pesel"] = 0;
            }
            else
            {
                try
                {
                    e.NewValues["Pesel"] = Int64.Parse(nm.controlText("txPESEL", grid));
                }
                catch
                {
                    {
                        e.NewValues["Pesel"] = 0;
                    }
                }
              
                //dane adresowe
                e.NewValues["ulica"] = nm.controlText("txAdres", grid);
                e.NewValues["kod_poczt"] = nm.controlText("txKodPocztowy", grid);
                e.NewValues["miejscowosc"] = nm.controlText("txMiejscowosc", grid);
                e.NewValues["tel1"] = nm.controlText("txTelefon1", grid);
                e.NewValues["tel2"] = nm.controlText("txTelefon2", grid);
                e.NewValues["email"] = nm.controlText("txEmail", grid);
                //dane korespondencyjne
                e.NewValues["adr_kores"] = nm.controlText("txAdresKorespondencyjny", grid);
                e.NewValues["kod_poczt_kor"] = nm.controlText("txKodPocztowyKorespondencyjny", grid);
                e.NewValues["miejscowosc_kor"] = nm.controlText("txMiejscowoscKorespondencyjny", grid);
                // uwagi i specjalizacje
                e.NewValues["uwagi"] = nm.controlTextMemo("txUwagi", grid);
                e.NewValues["specjalizacja_opis"] = nm.controlTextMemo("txSpecjalizacjeOpis", grid);
                e.NewValues["instytucja"] = nm.controlText("txInstytucja", grid);
            }
        }

        protected void grid_CancelRowEditing(object sender, ASPxStartRowEditingEventArgs e)
        {
            var cos = e.EditingKeyValue;
            if (e.EditingKeyValue == null)
            {
                try
                {
                    int idOsoby = int.Parse((string)Session["id_osoby"]);
                    nm.usunTworzonaOsobe(idOsoby);
                }
                catch 
                {
                }
            }
        } // end of grid_CancelRowEditing

        protected void grid_RowValidating(object sender, ASPxDataValidationEventArgs e)
        {
          
        }

        protected void grid_BeforePerformDataSelect(object sender, EventArgs e)
        {
            ustawKwerendeOdczytu();
            mediatorzy.SelectCommand = (string)Session["kwerenda"];
        }

        protected void poSpecjalizacji(object sender, EventArgs e)
        {
            ustawKwerendeOdczytu();
        }

        protected void ustawKwerendeOdczytu()
        {

            int czyCzynny = 0;
            czyCzynny = int.Parse(DropDownList2.SelectedValue);
            string kwerendabazowa = "";
            string nazwaSpeckajlizacji = string.Empty;
            

            switch (czyCzynny)
            {
                case 2: 
                    {
                        //typ >2
                        if (ASPxCheckBox2.Checked)
                        {
                            string specjalizacja = DropDownList1.SelectedValue;
                            nazwaSpeckajlizacji = NazwaSpecjalizacji(specjalizacja);

                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (data_koncowa >= GETDATE()) and (czyus = 0) and typ < 2 ";
                            kwerendabazowa = kwerendabazowa + "  and (select count(*) from tbl_specjalizacje_osob where id_specjalizacji =" + specjalizacja.Trim() + " and id_osoby=tbl_osoby.ident )=1 ";

                            Session["kwerenda"] = kwerendabazowa;

                        }
                        else
                        {
                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (data_koncowa >= GETDATE()) and (czyus = 0) and typ < 2";
                        }

                    }
                    break;
                case 3:
                    {
                        //typ <2
                        if (ASPxCheckBox2.Checked)
                        {
                            string specjalizacja = DropDownList1.SelectedValue;
                            nazwaSpeckajlizacji = NazwaSpecjalizacji(specjalizacja);

                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (czyus = 0) AND (typ >= 2) AND (data_koncowa >= GETDATE()) ";
                            kwerendabazowa = kwerendabazowa + "  and (select count(*) from tbl_specjalizacje_osob where id_specjalizacji =" + specjalizacja.Trim() + " and id_osoby=tbl_osoby.ident )=1 ";

                            Session["kwerenda"] = kwerendabazowa;

                        }
                        else
                        {
                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (data_koncowa >= GETDATE()) and typ >=2 ";
                        }

                    }
                    break;
                default:
                    {
                        if (ASPxCheckBox2.Checked)
                        {
                            string specjalizacja = DropDownList1.SelectedValue;
                            nazwaSpeckajlizacji = NazwaSpecjalizacji(specjalizacja);

                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (czyus  = 0) AND (data_koncowa >= GETDATE()) ";
                            kwerendabazowa = kwerendabazowa + "  and (select count(*) from tbl_specjalizacje_osob where id_specjalizacji =" + specjalizacja.Trim() + " and id_osoby=tbl_osoby.ident )=1 ";

                            Session["kwerenda"] = kwerendabazowa;

                        }
                        else
                        {
                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (data_koncowa >= GETDATE()) ";
                        }



                    }
                    break;
            }

    
            Session["kwerenda"] = kwerendabazowa;

   
            Session["kwerenda"] = kwerendabazowa + " order by nazwisko" ;
            mediatorzy.SelectCommand = kwerendabazowa;
            mediatorzy.DataBind();
        }

        private string NazwaSpecjalizacji(string specjalizacja)
        {
            cm Cm = new cm();
            DataTable parametry = Cm.makeParameterTable();
            parametry.Rows.Add("@idSpecjalizacji", specjalizacja);
            return Cm.runQuerryWithResult("SELECT nazwa   FROM glo_specjalizacje where id_=@idSpecjalizacji", Cm.con_str, parametry);
        }


        protected void ASPxCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            DropDownList1.Enabled = ASPxCheckBox2.Checked;
            ustawKwerendeOdczytu();
        }

        protected void _print(object sender, EventArgs e)
        {
            if (ASPxCheckBox2.Checked)
            {
                // specjalizacje
                grid.Columns["Specjalizacje"].Visible = false;
                grid.Columns["bezTabeli"].Visible = false;
                grid.Columns["jednaSpecjalizacja"].Visible = true;
            }
            else
            {
                grid.Columns["Specjalizacje"].Visible = false;
                grid.Columns["bezTabeli"].Visible = true;
                grid.Columns["jednaSpecjalizacja"].Visible = false;
            }
            using (MemoryStream ms = new MemoryStream())
            {
                PrintableComponentLink pcl = new PrintableComponentLink(new PrintingSystem());
                //listaBieglych.Columns[0].Visible = false;
                var cosik = ms.ToArray();

                ASPxGridViewExporter1.FileName = "Wykaz biegłych";

                pcl.Component = ASPxGridViewExporter1;

                pcl.Margins.Left = pcl.Margins.Right = 50;
                pcl.Landscape = true;
                pcl.CreateDocument(false);
                pcl.PrintingSystem.Document.AutoFitToPagesWidth = 1;
                pcl.ExportToPdf(ms);
                WriteResponse(this.Response, ms.ToArray(), System.Net.Mime.DispositionTypeNames.Inline.ToString());
            }
            grid.Columns["Specjalizacje"].Visible = true;
            grid.Columns["bezTabeli"].Visible = false;
            grid.Columns["jednaSpecjalizacja"].Visible = false;
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


            if (ASPxCheckBox2.Checked)
            {
                //jedna
                robRaportjednejSpecjalizacji(DropDownList1.SelectedItem, getDataFromGridview());
            }
            else
            {
                // robRaportWszystkichSpecjalizacji(getDataFromGridview());
                robRaportWszystkichSpecjalizacjiNowy(getDataFromGridview());
            }
        }

        private void robRaportWszystkichSpecjalizacjiNowy(DataTable dataTable)
        {

            string kwerenda = "SELECT View_SpecjalizacjeIOsoby.ident, tbl_osoby.imie, tbl_osoby.nazwisko, tbl_osoby.ulica, tbl_osoby.kod_poczt, tbl_osoby.miejscowosc, tbl_osoby.data_poczatkowa,                   tbl_osoby.data_koncowa, tbl_osoby.id_kreatora, tbl_osoby.data_kreacji, tbl_osoby.pesel, tbl_osoby.czyus, typ , tbl_osoby.tytul, tbl_osoby.czy_zaw, tbl_osoby.tel1, tbl_osoby.tel2,                   tbl_osoby.email, tbl_osoby.adr_kores, tbl_osoby.kod_poczt_kor, tbl_osoby.miejscowosc_kor, tbl_osoby.uwagi, tbl_osoby.specjalizacjeWidok, tbl_osoby.specjalizacja_opis,                   tbl_osoby.d_zawieszenia, tbl_osoby.typ, tbl_osoby.dataKoncaZawieszenia, tbl_osoby.instytucja, View_SpecjalizacjeIOsoby.nazwa, View_SpecjalizacjeIOsoby.id_ as identyfikatorSpecjalizacji,                   View_SpecjalizacjeIOsoby.Expr1 AS aktwnaSpecjalizacja FROM     tbl_osoby RIGHT OUTER JOIN                   View_SpecjalizacjeIOsoby ON tbl_osoby.ident = View_SpecjalizacjeIOsoby.ident WHERE (tbl_osoby.nazwisko IS NOT NULL) AND (tbl_osoby.typ < 2) AND (View_SpecjalizacjeIOsoby.Expr1 = 1)";
            DataTable daneBieglych = Cm.getDataTable(kwerenda, Cm.con_str);
            foreach (DataRow wiersz in daneBieglych.Rows)
            {
                string ident = wiersz["ident"].ToString().Trim();
                int numberOfRecords = dataTable.AsEnumerable().Where(x => x["id"].ToString() == ident).ToList().Count;
                if (numberOfRecords == 0)
                {
                    wiersz.Delete();
                }
            }
            daneBieglych.AcceptChanges();

            var IlosciBieglychPoSpecjalizacji = new DataTable();
            IlosciBieglychPoSpecjalizacji.Columns.Add("idSpecjalizacji", typeof(int));
            IlosciBieglychPoSpecjalizacji.Columns.Add("NazwaSpecjalizacji", typeof(string));
            IlosciBieglychPoSpecjalizacji.Columns.Add("ilosc", typeof(int));
            IlosciBieglychPoSpecjalizacji.Columns.Add("iloscStron", typeof(int));

            foreach (DataRow dRow in cl.odczytaj_specjalizacjeLista().Rows)
            {
                int idSpecjalizacji = int.Parse(dRow[0].ToString().Trim());
                string nazwaSpecjalizacji = dRow[1].ToString().Trim();

                int numberOfRecords = daneBieglych.AsEnumerable().Where(x => x["identyfikatorSpecjalizacji"].ToString() == idSpecjalizacji.ToString()).ToList().Count;
                DataRow wierszWyliczen = IlosciBieglychPoSpecjalizacji.NewRow();
                wierszWyliczen["idSpecjalizacji"] = idSpecjalizacji;
                wierszWyliczen["NazwaSpecjalizacji"] = nazwaSpecjalizacji;
                wierszWyliczen["ilosc"] = numberOfRecords;
                int iloscStron = 0;
                if (numberOfRecords > 0)
                {
                    IlosciBieglychPoSpecjalizacji.Rows.Add(wierszWyliczen);

                    float ilStr = (float)numberOfRecords / 15;
                    iloscStron = (int)Math.Ceiling(ilStr);
                    wierszWyliczen["iloscStron"] = iloscStron;
                }
            }

            //nagłówek
            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10f, 10f, 10f, 0f);

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
            string napisDodatkowy = "";
            switch (DropDownList2.SelectedIndex)
            {

                case 0: { napisDodatkowy = "Wszystcy biegli"; } break;
                case 1: { napisDodatkowy = "Biegli czynni"; } break;
                case 2: { napisDodatkowy = "Biegli nie czynni"; } break;

            }

            cell = new PdfPCell(new Paragraph(napisDodatkowy, cl.plFont1));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            fitst.AddCell(cell);
            pdfDoc.Add(fitst);
            pdfDoc.NewPage();

            //podliczenie
            int strona = 1;
            PdfPTable tabelaWyliczenia = new PdfPTable(3);
            int[] tblWidthX = { 10, 70, 20 };
            cell = new PdfPCell(new Paragraph("", cl.plFontBIG));
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            tabelaWyliczenia.AddCell(cell);
            tabelaWyliczenia.AddCell(cell);
            cell = new PdfPCell(new Paragraph("L.p.", cl.plFontBIG));
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Nazwa specjalizacji", cl.plFontBIG));
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Strona", cl.plFontBIG));
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            int iterator = 1;
            foreach (DataRow dRwyliczenie in IlosciBieglychPoSpecjalizacji.Rows)
            {
                cell = new PdfPCell(new Paragraph(iterator.ToString(), cl.plFont2));
                cell.Border = Rectangle.NO_BORDER;
                tabelaWyliczenia.AddCell(cell);
                cell = new PdfPCell(new Paragraph(dRwyliczenie["NazwaSpecjalizacji"].ToString().Trim(), cl.plFont2));
                cell.Border = Rectangle.NO_BORDER;
                tabelaWyliczenia.AddCell(cell);
                cell = new PdfPCell(new Paragraph(strona.ToString(), cl.plFont2));
                cell.Border = Rectangle.NO_BORDER;
                tabelaWyliczenia.AddCell(cell);
                strona = strona + int.Parse(dRwyliczenie["iloscStron"].ToString().Trim());
                iterator++;
            }
            pdfDoc.Add(tabelaWyliczenia);
            pdfDoc.NewPage();
            //end of  po specjalizacjach
            // koniec podliczenia

            
            // po specjalizacjach
            foreach (DataRow dRwyliczenie in IlosciBieglychPoSpecjalizacji.Rows)
            {
                string nazwaSpecjalizacji = dRwyliczenie["NazwaSpecjalizacji"].ToString().Trim();
                string IdSpecjalizacji = dRwyliczenie["idSpecjalizacji"].ToString().Trim();


                PdfPTable tabelaGlowna = new PdfPTable(4);
                int[] tblWidth = { 8, 30, 30, 32 };
                int iloscStron = 0;
                if (daneBieglych.Rows.Count > 0)
                {
            
                    tabelaGlowna = new PdfPTable(4);
                    tabelaGlowna = generujCzescRaportuNew(daneBieglych, IdSpecjalizacji);
                    pdfDoc.Add(new Paragraph(" "));
                    pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, cl.plFont3)));
                    pdfDoc.Add(new Paragraph(" "));

                    if (tabelaGlowna.Rows.Count > 15)
                    {
                        //   int counter = 0;
                        int licznik = 0;
                        PdfPTable newTable = new PdfPTable(4);
                        newTable.SetWidths(tblWidth);
                        // podziel tabele
                        int iter = 0;

                        foreach (var TableRow in tabelaGlowna.Rows)
                        {
                            var cos = TableRow.GetCells();
                            //   newTable.Rows.Add(TableRow);
                            PdfPCell celka01 = (PdfPCell)cos.GetValue(0);
                            PdfPCell celka02 = (PdfPCell)cos.GetValue(1);
                            PdfPCell celka03 = (PdfPCell)cos.GetValue(2);
                            PdfPCell celka04 = (PdfPCell)cos.GetValue(3);
                            string data1 = celka01.Phrase.Chunks.ToString();
                            if (iter > 0)
                            {
                                newTable.AddCell(new PdfPCell(new Phrase(iter.ToString())));
                            }
                            else
                            {
                                newTable.AddCell(celka01);
                            }
                            newTable.AddCell(celka02);
                            newTable.AddCell(celka03);
                            newTable.AddCell(celka04);
                            licznik++;
                            iter++;

                            if (licznik == 15)
                            {
                                iloscStron++;
                                licznik = 0;
                                pdfDoc.Add(newTable);
                                pdfDoc.NewPage();
                                pdfDoc.Add(new Paragraph(" "));
                                pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji + " ciąg dalszy", cl.plFont3)));
                                pdfDoc.Add(new Paragraph(" "));

                                newTable = new PdfPTable(4);
                                newTable.SetWidths(tblWidth);
                                newTable.Rows.Clear();
                            }
                        }

                        pdfDoc.Add(newTable);
                        pdfDoc.NewPage();
                    }
                    else
                    {
                        pdfDoc.Add(tabelaGlowna);
                        pdfDoc.NewPage();
                    }
                    // uttwórz listę osób z taka specjalizacją
                }





            }



            pdfDoc.Close();
            string newFilename = fileName + ".pdf";
            AddPageNumber(fileName, newFilename);
        }

        private void robRaportjednejSpecjalizacji(System.Web.UI.WebControls.ListItem selectedItem, DataTable listaBieglych)
        {

            int idSpecjalizacji = 0;
            string nazwaSpecjalizacji = string.Empty;
            try
            {
                idSpecjalizacji = int.Parse(selectedItem.Value);
                nazwaSpecjalizacji = selectedItem.Text;
            }
            catch
            {
                return;
            }
            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10f, 10f, 10f, 0f);

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
            string napisDodatkowy = "";
            switch (DropDownList2.SelectedIndex)
            {

                case 0: { napisDodatkowy = "Wszystcy biegli"; } break;
                case 1: { napisDodatkowy = "Biegli czynni"; } break;
                case 2: { napisDodatkowy = "Biegli nie czynni"; } break;
                
            }

            cell = new PdfPCell(new Paragraph(napisDodatkowy, cl.plFont1));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            fitst.AddCell(cell);



            pdfDoc.Add(fitst);
            pdfDoc.NewPage();

            //podliczenie
            DataTable specjalizacjeWyliczenie = new DataTable();
            specjalizacjeWyliczenie.Columns.Add("nr", typeof(string));
            specjalizacjeWyliczenie.Columns.Add("str", typeof(string));

            int iloscStron = 0;

            /*
            foreach (DataRow dRow in cl.odczytaj_specjalizacjeLista().Rows)
            {
                Biegli = generujTabeleBieglychDoZestawienia();

                int idSpecjalizacji = int.Parse(dRow[0].ToString().Trim());
                string nazwaSpecjalizacji = dRow[1].ToString().Trim();

                foreach (DataRow bieglyZlisty in listaBieglych.Rows)
                {
                    int idBieglego = int.Parse(bieglyZlisty[0].ToString().Trim());
                    int ilosc = cl.czyJestSpecjalizacjauBieglego(idBieglego, idSpecjalizacji);
                    if (ilosc > 0)
                    {
                        DataRow jedenBiegly = Biegli.NewRow();
                        jedenBiegly = wierszZBieglym(bieglyZlisty, Biegli);
                        Biegli.Rows.Add(jedenBiegly);
                    }
                }// end of foreach

                if (Biegli.Rows.Count > 0)
                {
                    float ilStr = (float)Biegli.Rows.Count / 15;
                    iloscStron = (int)Math.Ceiling(ilStr);
                    DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                    wyliczenie[0] = nazwaSpecjalizacji;
                    wyliczenie[1] = iloscStron.ToString();

                    specjalizacjeWyliczenie.Rows.Add(wyliczenie);
                }
            }

            */// dodaj wyliczenia*/
            PdfPTable tabelaGlowna = new PdfPTable(4);
            int[] tblWidth = { 8, 30, 30, 32 };

            if (listaBieglych.Rows.Count > 0)
            {
                float ilStr = (float)listaBieglych.Rows.Count / 15;
                iloscStron = (int)Math.Ceiling(ilStr);
                DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                wyliczenie[0] = nazwaSpecjalizacji;
                wyliczenie[1] = iloscStron.ToString();

                specjalizacjeWyliczenie.Rows.Add(wyliczenie);
            }
            int strona = 1;
            PdfPTable tabelaWyliczenia = new PdfPTable(3);
            int[] tblWidthX = { 10, 70, 20 };
            cell = new PdfPCell(new Paragraph("", cl.plFontBIG));
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            tabelaWyliczenia.AddCell(cell);
            tabelaWyliczenia.AddCell(cell);
            cell = new PdfPCell(new Paragraph("L.p.", cl.plFontBIG));
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Nazwa specjalizacji", cl.plFontBIG));
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Strona", cl.plFontBIG));
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            int iterator = 1;
            foreach (DataRow dRwyliczenie in specjalizacjeWyliczenie.Rows)
            {
                cell = new PdfPCell(new Paragraph(iterator.ToString(), cl.plFont2));
                cell.Border = Rectangle.NO_BORDER;
                tabelaWyliczenia.AddCell(cell);
                cell = new PdfPCell(new Paragraph(dRwyliczenie[0].ToString().Trim(), cl.plFont2));
                cell.Border = Rectangle.NO_BORDER;
                tabelaWyliczenia.AddCell(cell);
                cell = new PdfPCell(new Paragraph(strona.ToString(), cl.plFont2));
                cell.Border = Rectangle.NO_BORDER;
                tabelaWyliczenia.AddCell(cell);
                strona = strona + int.Parse(dRwyliczenie[1].ToString().Trim());
                iterator++;
            }
            pdfDoc.Add(tabelaWyliczenia);
            pdfDoc.NewPage();
            //end of  po specjalizacjach
            // koniec podliczenia
            /*
            foreach (DataRow dRow in cl.odczytaj_specjalizacjeLista().Rows)
            {
                Biegli = generujTabeleBieglychDoZestawienia();

                int idSpecjalizacji = int.Parse(dRow[0].ToString().Trim());
                string nazwaSpecjalizacji = dRow[1].ToString().Trim();

                foreach (DataRow bieglyZlisty in listaBieglych.Rows)
                {
                    int idBieglego = int.Parse(bieglyZlisty[0].ToString().Trim());
                    int ilosc = cl.czyJestSpecjalizacjauBieglego(idBieglego, idSpecjalizacji);
                    if (ilosc > 0)
                    {
                        DataRow jedenBiegly = Biegli.NewRow();
                        jedenBiegly = wierszZBieglym(bieglyZlisty, Biegli);
                        Biegli.Rows.Add(jedenBiegly);
                    }
                }// end of foreach

                PdfPTable tabelaGlowna = new PdfPTable(4);
                int[] tblWidth = { 8, 30, 30, 32 };

                if (Biegli.Rows.Count > 0)
                {
                    DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                    wyliczenie[0] = nazwaSpecjalizacji;
                    wyliczenie[1] = iloscStron.ToString();

                    specjalizacjeWyliczenie.Rows.Add(wyliczenie);
                    tabelaGlowna = new PdfPTable(4);
                    tabelaGlowna = generujCzescRaportu(Biegli, nazwaSpecjalizacji);
                    pdfDoc.Add(new Paragraph(" "));
                    pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, cl.plFont3)));
                    pdfDoc.Add(new Paragraph(" "));

                    if (tabelaGlowna.Rows.Count > 15)
                    {
                        //   int counter = 0;
                        int licznik = 0;
                        PdfPTable newTable = new PdfPTable(4);
                        newTable.SetWidths(tblWidth);
                        // podziel tabele
                        int iter = 0;

                        foreach (var TableRow in tabelaGlowna.Rows)
                        {
                            var cos = TableRow.GetCells();
                            //   newTable.Rows.Add(TableRow);
                            PdfPCell celka01 = (PdfPCell)cos.GetValue(0);
                            PdfPCell celka02 = (PdfPCell)cos.GetValue(1);
                            PdfPCell celka03 = (PdfPCell)cos.GetValue(2);
                            PdfPCell celka04 = (PdfPCell)cos.GetValue(3);
                            string data1 = celka01.Phrase.Chunks.ToString();
                            if (iter > 0)
                            {
                                newTable.AddCell(new PdfPCell(new Phrase(iter.ToString())));
                            }
                            else
                            {
                                newTable.AddCell(celka01);
                            }
                            newTable.AddCell(celka02);
                            newTable.AddCell(celka03);
                            newTable.AddCell(celka04);
                            licznik++;
                            iter++;

                            if (licznik == 15)
                            {
                                iloscStron++;
                                licznik = 0;
                                pdfDoc.Add(newTable);
                                pdfDoc.NewPage();
                                pdfDoc.Add(new Paragraph(" "));
                                pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji + " ciąg dalszy", cl.plFont3)));
                                pdfDoc.Add(new Paragraph(" "));

                                newTable = new PdfPTable(4);
                                newTable.SetWidths(tblWidth);
                                newTable.Rows.Clear();
                            }
                        }

                        pdfDoc.Add(newTable);
                        pdfDoc.NewPage();
                    }
                    else
                    {
                        pdfDoc.Add(tabelaGlowna);
                        pdfDoc.NewPage();
                    }
                    // uttwórz listę osób z taka specjalizacją
                }
            }
            *///end of  po specjalizacjach
              //==============================================================

            if (listaBieglych.Rows.Count > 0)
            {
                DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                wyliczenie[0] = nazwaSpecjalizacji;
                wyliczenie[1] = iloscStron.ToString();

                specjalizacjeWyliczenie.Rows.Add(wyliczenie);
                tabelaGlowna = new PdfPTable(4);
                tabelaGlowna = generujCzescRaportuOne(listaBieglych, nazwaSpecjalizacji);
                pdfDoc.Add(new Paragraph(" "));
                pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, cl.plFont3)));
                pdfDoc.Add(new Paragraph(" "));

                if (tabelaGlowna.Rows.Count > 15)
                {
                    //   int counter = 0;
                    int licznik = 0;
                    PdfPTable newTable = new PdfPTable(4);
                    newTable.SetWidths(tblWidth);
                    // podziel tabele
                    int iter = 0;

                    foreach (var TableRow in tabelaGlowna.Rows)
                    {
                        var cos = TableRow.GetCells();
                        //   newTable.Rows.Add(TableRow);
                        PdfPCell celka01 = (PdfPCell)cos.GetValue(0);
                        PdfPCell celka02 = (PdfPCell)cos.GetValue(1);
                        PdfPCell celka03 = (PdfPCell)cos.GetValue(2);
                        PdfPCell celka04 = (PdfPCell)cos.GetValue(3);
                        string data1 = celka01.Phrase.Chunks.ToString();
                        if (iter > 0)
                        {
                            newTable.AddCell(new PdfPCell(new Phrase(iter.ToString())));
                        }
                        else
                        {
                            newTable.AddCell(celka01);
                        }
                        newTable.AddCell(celka02);
                        newTable.AddCell(celka03);
                        newTable.AddCell(celka04);
                        licznik++;
                        iter++;

                        if (licznik == 15)
                        {
                            iloscStron++;
                            licznik = 0;
                            pdfDoc.Add(newTable);
                            pdfDoc.NewPage();
                            pdfDoc.Add(new Paragraph(" "));
                            pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji + " ciąg dalszy", cl.plFont3)));
                            pdfDoc.Add(new Paragraph(" "));

                            newTable = new PdfPTable(4);
                            newTable.SetWidths(tblWidth);
                            newTable.Rows.Clear();
                        }
                    }

                    pdfDoc.Add(newTable);
                    pdfDoc.NewPage();
                }
                else
                {
                    pdfDoc.Add(tabelaGlowna);
                    pdfDoc.NewPage();
                }
                // uttwórz listę osób z taka specjalizacją
            }
            pdfDoc.Close();
            string newFilename = fileName + ".pdf";
            AddPageNumber(fileName, newFilename);
        }

        protected DataTable getDataFromGridview()
        {
            DataTable identy = new DataTable();
            identy.Columns.Add(new DataColumn("id", typeof(int)));

            int vrc = grid.VisibleRowCount;
            int vrsi = grid.VisibleStartIndex;

            for (int i = 0; i < vrc; i++)
            {
                int id_ = Convert.ToInt32(grid.GetRowValues(i, grid.KeyFieldName));
                DataRow dR = identy.NewRow();
                dR[0] = id_;
                identy.Rows.Add(dR);
            }
            return identy;
        }

      
        protected DataRow wierszZBieglym(DataRow biegliRow, DataTable Biegli)
        {
            DataRow bieglyZatwierdzony = Biegli.NewRow();
            try
            {
                bieglyZatwierdzony[0] = biegliRow[0];
                bieglyZatwierdzony[1] = biegliRow[1];
                bieglyZatwierdzony[2] = biegliRow[2];
                bieglyZatwierdzony[3] = biegliRow[3];
                bieglyZatwierdzony[4] = biegliRow[4];
                bieglyZatwierdzony[5] = biegliRow[5];
                bieglyZatwierdzony[6] = biegliRow[6];
                bieglyZatwierdzony[7] = biegliRow[7];
                bieglyZatwierdzony[8] = biegliRow[8];
            }
            catch (Exception)
            { }

            return bieglyZatwierdzony;
        }

        protected void robRaportWszystkichSpecjalizacji(DataTable listaBieglych)
        {
            // wyciąfnij listę ludzi z dana specjalizacją

            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10f, 10f, 10f, 0f);

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

            //podliczenie
            DataTable specjalizacjeWyliczenie = new DataTable();
            specjalizacjeWyliczenie.Columns.Add("nr", typeof(string));
            specjalizacjeWyliczenie.Columns.Add("str", typeof(string));
            DataTable Biegli = generujTabeleBieglychDoZestawienia();
            int iloscStron = 0;

            foreach (DataRow dRow in cl.odczytaj_specjalizacjeLista().Rows)
            {
                Biegli = generujTabeleBieglychDoZestawienia();

                int idSpecjalizacji = int.Parse(dRow[0].ToString().Trim());
                string nazwaSpecjalizacji = dRow[1].ToString().Trim();

                foreach (DataRow bieglyZlisty in listaBieglych.Rows)
                {
                    int idBieglego = int.Parse(bieglyZlisty[0].ToString().Trim());
                    int ilosc = cl.czyJestSpecjalizacjauBieglego(idBieglego, idSpecjalizacji);
                    if (ilosc > 0)
                    {
                        DataRow jedenBiegly = Biegli.NewRow();
                        jedenBiegly = wierszZBieglym(bieglyZlisty, Biegli);
                        Biegli.Rows.Add(jedenBiegly);
                    }
                }// end of foreach

                PdfPTable tabelaGlowna = new PdfPTable(4);
                int[] tblWidth = { 8, 30, 30, 32 };

                if (Biegli.Rows.Count > 0)
                {
                    float ilStr = (float)Biegli.Rows.Count / 15;
                    iloscStron = (int)Math.Ceiling(ilStr);
                    DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                    wyliczenie[0] = nazwaSpecjalizacji;
                    wyliczenie[1] = iloscStron.ToString();

                    specjalizacjeWyliczenie.Rows.Add(wyliczenie);
                }
            }
            // dodaj wyliczenia
            int strona = 1;
            PdfPTable tabelaWyliczenia = new PdfPTable(3);
            int[] tblWidthX = { 10, 70, 20 };
            cell = new PdfPCell(new Paragraph("", cl.plFontBIG));
            cell.FixedHeight = 100;
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            tabelaWyliczenia.AddCell(cell);
            tabelaWyliczenia.AddCell(cell);
            cell = new PdfPCell(new Paragraph("L.p.", cl.plFontBIG));
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Nazwa specjalizacji", cl.plFontBIG));
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Strona", cl.plFontBIG));
            cell.Border = Rectangle.NO_BORDER;
            tabelaWyliczenia.AddCell(cell);
            int iterator = 1;
            foreach (DataRow dRwyliczenie in specjalizacjeWyliczenie.Rows)
            {
                cell = new PdfPCell(new Paragraph(iterator.ToString(), cl.plFont2));
                cell.Border = Rectangle.NO_BORDER;
                tabelaWyliczenia.AddCell(cell);
                cell = new PdfPCell(new Paragraph(dRwyliczenie[0].ToString().Trim(), cl.plFont2));
                cell.Border = Rectangle.NO_BORDER;
                tabelaWyliczenia.AddCell(cell);
                cell = new PdfPCell(new Paragraph(strona.ToString(), cl.plFont2));
                cell.Border = Rectangle.NO_BORDER;
                tabelaWyliczenia.AddCell(cell);
                strona = strona + int.Parse(dRwyliczenie[1].ToString().Trim());
                iterator++;
            }
            pdfDoc.Add(tabelaWyliczenia);
            pdfDoc.NewPage();
            //end of  po specjalizacjach
            // koniec podliczenia
            //xxxxxx
            foreach (DataRow dRow in cl.odczytaj_specjalizacjeLista().Rows)
            {
                Biegli = generujTabeleBieglychDoZestawienia();

                int idSpecjalizacji = int.Parse(dRow[0].ToString().Trim());
                string nazwaSpecjalizacji = dRow[1].ToString().Trim();

                foreach (DataRow bieglyZlisty in listaBieglych.Rows)
                {
                    int idBieglego = int.Parse(bieglyZlisty[0].ToString().Trim());
                    int ilosc = cl.czyJestSpecjalizacjauBieglego(idBieglego, idSpecjalizacji);
                    if (ilosc > 0)
                    {
                        DataRow jedenBiegly = Biegli.NewRow();
                        jedenBiegly = wierszZBieglym(bieglyZlisty, Biegli);
                        Biegli.Rows.Add(jedenBiegly);
                    }
                }// end of foreach

                PdfPTable tabelaGlowna = new PdfPTable(4);
                int[] tblWidth = { 8, 30, 30, 32 };

                if (Biegli.Rows.Count > 0)
                {
                    DataRow wyliczenie = specjalizacjeWyliczenie.NewRow();
                    wyliczenie[0] = nazwaSpecjalizacji;
                    wyliczenie[1] = iloscStron.ToString();

                    specjalizacjeWyliczenie.Rows.Add(wyliczenie);
                    tabelaGlowna = new PdfPTable(4);
                    tabelaGlowna = generujCzescRaportu(Biegli, nazwaSpecjalizacji);
                    pdfDoc.Add(new Paragraph(" "));
                    pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, cl.plFont3)));
                    pdfDoc.Add(new Paragraph(" "));

                    if (tabelaGlowna.Rows.Count > 15)
                    {
                        //   int counter = 0;
                        int licznik = 0;
                        PdfPTable newTable = new PdfPTable(4);
                        newTable.SetWidths(tblWidth);
                        // podziel tabele
                        int iter = 0;

                        foreach (var TableRow in tabelaGlowna.Rows)
                        {
                            var cos = TableRow.GetCells();
                            //   newTable.Rows.Add(TableRow);
                            PdfPCell celka01 = (PdfPCell)cos.GetValue(0);
                            PdfPCell celka02 = (PdfPCell)cos.GetValue(1);
                            PdfPCell celka03 = (PdfPCell)cos.GetValue(2);
                            PdfPCell celka04 = (PdfPCell)cos.GetValue(3);
                            string data1 = celka01.Phrase.Chunks.ToString();
                            if (iter > 0)
                            {
                                newTable.AddCell(new PdfPCell(new Phrase(iter.ToString())));
                            }
                            else
                            {
                                newTable.AddCell(celka01);
                            }
                            newTable.AddCell(celka02);
                            newTable.AddCell(celka03);
                            newTable.AddCell(celka04);
                            licznik++;
                            iter++;

                            if (licznik == 15)
                            {
                                iloscStron++;
                                licznik = 0;
                                pdfDoc.Add(newTable);
                                pdfDoc.NewPage();
                                pdfDoc.Add(new Paragraph(" "));
                                pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji + " ciąg dalszy", cl.plFont3)));
                                pdfDoc.Add(new Paragraph(" "));

                                newTable = new PdfPTable(4);
                                newTable.SetWidths(tblWidth);
                                newTable.Rows.Clear();
                            }
                        }

                        pdfDoc.Add(newTable);
                        pdfDoc.NewPage();
                    }
                    else
                    {
                        pdfDoc.Add(tabelaGlowna);
                        pdfDoc.NewPage();
                    }
                    // uttwórz listę osób z taka specjalizacją
                }
            }//end of  po specjalizacjach
             //==============================================================

            pdfDoc.Close();
            string newFilename = fileName + ".pdf";
            AddPageNumber(fileName, newFilename);
        }

        protected DataTable generujTabeleBieglychDoZestawienia()
        {
            DataTable Biegli = new DataTable();
            Biegli.Columns.Add("ident", typeof(int));
            Biegli.Columns.Add("imie", typeof(string));
            Biegli.Columns.Add("nazwisko", typeof(string));
            Biegli.Columns.Add("ulica", typeof(string));
            Biegli.Columns.Add("kod_poczt", typeof(string));
            Biegli.Columns.Add("miejscowosc", typeof(string));
            Biegli.Columns.Add("data_koncowa", typeof(string));
            Biegli.Columns.Add("tytul", typeof(string));
            Biegli.Columns.Add("tel1", typeof(string));
            return Biegli;
        }

        protected PdfPTable generujCzescRaportu(DataTable biegli, string specjalizacje)
        {
            if (biegli.Rows.Count == 0)
            {
                return null;
            }
            int[] tblWidth = { 8, 30, 30, 32 };

            PdfPTable tabelaGlowna = new PdfPTable(4);
            tabelaGlowna.SetWidths(tblWidth);
            int iterator = 0;
            tabelaGlowna.AddCell(new Paragraph("Lp.", cl.plFont2));
            tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", cl.plFont2));
            tabelaGlowna.AddCell(new Paragraph("Adres- telefon", cl.plFont2));
            tabelaGlowna.AddCell(new Paragraph("Zakres", cl.plFont2));
            int iloscBieglych = biegli.Rows.Count;

            foreach (DataRow biegly in biegli.Rows)
            {
                iterator++;
                string Idbieglego = biegly[0].ToString();
                DataTable daneBieglego = cl.wyciagnijBieglegoZSpecjalizacja(int.Parse(Idbieglego));
                if (daneBieglego.Rows.Count == 0)
                {
                    continue;
                }
                DataRow daneJednegoBieglego = daneBieglego.Rows[0];
                DataTable listaSpecjalizacjiBieglego = new DataTable();
                listaSpecjalizacjiBieglego = cl.odczytaj_specjalizacje_osoby2(Idbieglego);
                //tbl_osoby.ident, tbl_osoby.imie, tbl_osoby.nazwisko, tbl_osoby.ulica, tbl_osoby.kod_poczt, tbl_osoby.miejscowosc,   tbl_osoby.data_koncowa,  tbl_osoby.tytul,
                string imie = daneJednegoBieglego[1].ToString();
                string nazwisko = daneJednegoBieglego[2].ToString();
                string tytul = daneJednegoBieglego[7].ToString();
                string telefon = daneJednegoBieglego[8].ToString();
                string email = daneJednegoBieglego[9].ToString();
                string dataKonca = string.Empty;
                try
                {
                    dataKonca = DateTime.Parse(daneJednegoBieglego[6].ToString()).ToShortDateString();
                }
                catch
                { }

                string innerTable = imie + Environment.NewLine + nazwisko + Environment.NewLine + tytul + Environment.NewLine + "kadencja do dnia: " + dataKonca;
                string ulica = daneJednegoBieglego[3].ToString();
                string kod = daneJednegoBieglego[4].ToString();
                string miejscowosc = daneJednegoBieglego[5].ToString();
                string tel = daneJednegoBieglego[8].ToString();
                string adresTable = ulica + Environment.NewLine + kod + " " + miejscowosc + Environment.NewLine + tel + Environment.NewLine + email;
                string specki = string.Empty;
                string specjalizacjaOpis = cl.odczytaj_specjalizacje_osobyOpis(Idbieglego);
                // tabelaGlowna.AddCell(new Paragraph(specjalizacjaOpis, cl.plFont2));
                foreach (DataRow specRow in listaSpecjalizacjiBieglego.Rows)
                {
                    specki = specki + specRow[0].ToString().ToLower() + "; ";
                }
                specki = specki + specjalizacjaOpis;
                tabelaGlowna.AddCell(new Paragraph(iterator.ToString(), cl.plFont1));
                tabelaGlowna.AddCell(new Paragraph(innerTable, cl.plFont1));
                tabelaGlowna.AddCell(new Paragraph(adresTable, cl.plFont1));
                tabelaGlowna.AddCell(new Paragraph(specki, cl.plFont1));
            }

            return tabelaGlowna;
        }

        protected PdfPTable generujCzescRaportuNew(DataTable biegli, string idSpecjalizacji)
        {




            if (biegli.Rows.Count == 0)
            {
                return null;
            }
            int[] tblWidth = { 8, 30, 30, 32 };

            PdfPTable tabelaGlowna = new PdfPTable(4);
            tabelaGlowna.SetWidths(tblWidth);
            int iterator = 0;
            tabelaGlowna.AddCell(new Paragraph("Lp.", cl.plFont2));
            tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", cl.plFont2));
            tabelaGlowna.AddCell(new Paragraph("Adres- telefon", cl.plFont2));
            tabelaGlowna.AddCell(new Paragraph("Zakres", cl.plFont2));
            int iloscBieglych = biegli.Rows.Count;

            var result = biegli
    .AsEnumerable()
    .Where(myRow => myRow.Field<int>("identyfikatorSpecjalizacji") == int.Parse (idSpecjalizacji )).ToArray();


            foreach (DataRow biegly in result)
            {
                iterator++;
                string Idbieglego = biegly["ident"].ToString();
        
                DataTable listaSpecjalizacjiBieglego = new DataTable();
                listaSpecjalizacjiBieglego = cl.odczytaj_specjalizacje_osoby2(Idbieglego);
                //"SELECT View_SpecjalizacjeIOsoby.ident, tbl_osoby.imie, tbl_osoby.nazwisko, tbl_osoby.ulica, tbl_osoby.kod_poczt, tbl_osoby.miejscowosc, tbl_osoby.data_poczatkowa,                   tbl_osoby.data_koncowa, tbl_osoby.id_kreatora, tbl_osoby.data_kreacji, tbl_osoby.pesel, tbl_osoby.czyus, typ , tbl_osoby.tytul, tbl_osoby.czy_zaw, tbl_osoby.tel1, tbl_osoby.tel2,                   tbl_osoby.email, tbl_osoby.adr_kores, tbl_osoby.kod_poczt_kor, tbl_osoby.miejscowosc_kor, tbl_osoby.uwagi, tbl_osoby.specjalizacjeWidok, tbl_osoby.specjalizacja_opis,                   tbl_osoby.d_zawieszenia, tbl_osoby.typ, tbl_osoby.dataKoncaZawieszenia, tbl_osoby.instytucja, View_SpecjalizacjeIOsoby.nazwa, View_SpecjalizacjeIOsoby.id_ as identyfikatorSpecjalizacji,                   View_SpecjalizacjeIOsoby.Expr1 AS aktwnaSpecjalizacja FROM     tbl_osoby RIGHT OUTER JOIN                   View_SpecjalizacjeIOsoby ON tbl_osoby.ident = View_SpecjalizacjeIOsoby.ident WHERE (tbl_osoby.nazwisko IS NOT NULL) AND (tbl_osoby.typ < 2) AND (View_SpecjalizacjeIOsoby.Expr1 = 1)";
                string imie = biegly["imie"].ToString();
                string nazwisko = biegly["nazwisko"].ToString();
                string tytul = biegly["tytul"].ToString();
                string telefon = biegly["tel1"].ToString();
                string email = biegly["email"].ToString();
                string dataKonca = string.Empty;
                try
                {
                    dataKonca = DateTime.Parse(biegly["data_koncowa"].ToString()).ToShortDateString();
                }
                catch
                { }

                string innerTable = imie + Environment.NewLine + nazwisko + Environment.NewLine + tytul + Environment.NewLine + "kadencja do dnia: " + dataKonca;
                string ulica = biegly["ulica"].ToString();
                string kod = biegly["kod_poczt"].ToString();
                string miejscowosc = biegly["miejscowosc"].ToString();
                string tel = biegly["tel2"].ToString();
                string adresTable = ulica + Environment.NewLine + kod + " " + miejscowosc + Environment.NewLine + tel + Environment.NewLine + email;
                string specki = string.Empty;
                string specjalizacjaOpis = cl.odczytaj_specjalizacje_osobyOpis(Idbieglego);
                // tabelaGlowna.AddCell(new Paragraph(specjalizacjaOpis, cl.plFont2));
                foreach (DataRow specRow in listaSpecjalizacjiBieglego.Rows)
                {
                    specki = specki + specRow[0].ToString().ToLower() + "; ";
                }
                specki = specki + specjalizacjaOpis;
                tabelaGlowna.AddCell(new Paragraph(iterator.ToString(), cl.plFont1));
                tabelaGlowna.AddCell(new Paragraph(innerTable, cl.plFont1));
                tabelaGlowna.AddCell(new Paragraph(adresTable, cl.plFont1));
                tabelaGlowna.AddCell(new Paragraph(specki, cl.plFont1));
            }

            return tabelaGlowna;
        }


        protected PdfPTable generujCzescRaportuOne(DataTable biegli, string specjalizacje)
        {
            if (biegli.Rows.Count == 0)
            {
                return null;
            }
            int[] tblWidth = { 8, 30, 30, 32 };

            PdfPTable tabelaGlowna = new PdfPTable(4);
            tabelaGlowna.SetWidths(tblWidth);
            int iterator = 0;
            tabelaGlowna.AddCell(new Paragraph("Lp.", cl.plFont2));
            tabelaGlowna.AddCell(new Paragraph("Nazwisko i imię", cl.plFont2));
            tabelaGlowna.AddCell(new Paragraph("Adres- telefon", cl.plFont2));
            tabelaGlowna.AddCell(new Paragraph("Zakres", cl.plFont2));
            int iloscBieglych = biegli.Rows.Count;

            foreach (DataRow biegly in biegli.Rows)
            {
                DataTable daneBieglego = cl.wyciagnijBieglegoZSpecjalizacja(int.Parse(biegly[0].ToString()));
                if (daneBieglego.Rows.Count == 0)
                {
                    continue;
                }

                iterator++;
                string Idbieglego = daneBieglego.Rows[0][0].ToString();
                //tbl_osoby.ident, tbl_osoby.imie, tbl_osoby.nazwisko, tbl_osoby.ulica, tbl_osoby.kod_poczt, tbl_osoby.miejscowosc,   tbl_osoby.data_koncowa,  tbl_osoby.tytul,
                string imie = daneBieglego.Rows[0][1].ToString();
                string nazwisko = daneBieglego.Rows[0][2].ToString();
                string tytul = daneBieglego.Rows[0][7].ToString();
                string telefon = daneBieglego.Rows[0][8].ToString();
                string email = daneBieglego.Rows[0][9].ToString();
                string dataKonca = string.Empty;
                try
                {
                    dataKonca = DateTime.Parse(daneBieglego.Rows[0][6].ToString()).ToShortDateString();
                }
                catch
                { }

                string innerTable = imie + Environment.NewLine + nazwisko + Environment.NewLine + tytul + Environment.NewLine + "kadencja do dnia: " + dataKonca;
                string ulica = daneBieglego.Rows[0][3].ToString();
                string kod = daneBieglego.Rows[0][4].ToString();
                string miejscowosc = daneBieglego.Rows[0][5].ToString();
                string tel = daneBieglego.Rows[0][8].ToString();
                string specjalizacjaOPisSpecjalizacji = daneBieglego.Rows[0][10].ToString();

                string adresTable = ulica + Environment.NewLine + kod + " " + miejscowosc + Environment.NewLine + tel + Environment.NewLine + email;
                string opisSpecjalizacji = cl.odczytaj_specjalizacje_osobyOpis(Idbieglego, specjalizacje.Trim());

                tabelaGlowna.AddCell(new Paragraph(iterator.ToString(), cl.plFont1));
                tabelaGlowna.AddCell(new Paragraph(innerTable, cl.plFont1));
                tabelaGlowna.AddCell(new Paragraph(adresTable, cl.plFont1));
                tabelaGlowna.AddCell(new Paragraph(specjalizacje.ToUpper() +" "+ opisSpecjalizacji+" ; "+ specjalizacjaOPisSpecjalizacji, cl.plFont1));
            }

            return tabelaGlowna;
        }

        protected void makeExcell(object sender, EventArgs e)
        {
            ASPxGridViewExporter1.FileName = "Wykaz Biegłych";
            ASPxGridViewExporter1.WriteXlsToResponse();
        }

        private void AddPageNumber(string fileIn, string fileOut)
        {
            byte[] bytes = File.ReadAllBytes(fileIn);
            Font blackFont = FontFactory.GetFont("Arial", 12, Font.NORMAL, BaseColor.BLACK);
            using (MemoryStream stream = new MemoryStream())
            {
                PdfReader reader = new PdfReader(bytes);
                using (PdfStamper stamper = new PdfStamper(reader, stream))
                {
                    int pages = reader.NumberOfPages;
                    for (int i = 3; i <= pages; i++)
                    {
                        ColumnText.ShowTextAligned(stamper.GetUnderContent(i), Element.ALIGN_RIGHT, new Phrase((i - 2).ToString(), blackFont), 568f, 15f, 0);
                    }
                }
                bytes = stream.ToArray();
            }
            File.WriteAllBytes(fileOut, bytes);
            WebClient client = new WebClient();
            Byte[] buffer = client.DownloadData(fileOut);
            if (buffer != null)
            {
                Response.ContentType = "application/pdf";
                Response.AddHeader("content-lenght", buffer.Length.ToString());
                Response.BinaryWrite(buffer);
            }
        }

        protected void grid_CustomColumnDisplayText(object sender, ASPxGridViewColumnDisplayTextEventArgs e)
        {
            if (e.Column.Index == 7)
            {
                e.EncodeHtml = false;
            }
        }

        protected void ASPxGridViewExporter1_RenderBrick(object sender, ASPxGridViewExportRenderingEventArgs e)
        {
            if (e.Column.Name.Contains("Specjalizacje"))
            {
                e.Column.Visible = false;
                //e.Column.AdaptivePriority
                //e.TextValue = string.Format("{0:N2}%", (decimal)e.Value);
            }
        }

        protected void ChangeList(object sender, EventArgs e)
        {
            ustawKwerendeOdczytu();
        }

       

        protected void grid_HtmlDataCellPrepared(object sender, ASPxGridViewTableDataCellEventArgs e)
        {
            
            if (e.DataColumn.FieldName != "typ")
                return;
          

        }

        protected void grid_HtmlRowPrepared(object sender, ASPxGridViewTableRowEventArgs e)
        {
            if (e.RowType != GridViewRowType.Data) return;

          
            int typ = Convert.ToInt32(e.GetValue("typ"));
            if (typ >= 2)
                e.Row.BackColor = System.Drawing.Color.LightYellow;
        }
    }

  
}