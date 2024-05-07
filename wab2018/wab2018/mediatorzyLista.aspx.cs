using System;
using DevExpress.Web.Data;
using DevExpress.Web;
using System.Configuration;
using System.IO;
using DevExpress.XtraPrinting;
using System.Web;
using System.Net;
using System.Net.Mime;


namespace wab2018
{
    public partial class mediatorzyLista : System.Web.UI.Page
    {
        nowiMediatorzy nm = new nowiMediatorzy();
        cm Cm = new cm();
        Class2 cl = new Class2();
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
                    case "5":
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
                int idBieglego = cl.podajIdOsobyPoNumerzeSkargi(int.Parse(parametr));
                string nazwisko = cl.podajNazwiskoOsobyPoNumerzeSkargi(int.Parse(parametr));
                int visibleIndex = grid.FindVisibleIndexByKeyValue(idBieglego);

                //Remove the items
                grid.Selection.SelectRow(visibleIndex);
                grid.StartEdit(visibleIndex);
                
                ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
                pageControl.ActiveTabIndex = 2;
            }
            }

        protected void updateMediatora(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {
            string txt=mediatorzy.SelectCommand;
            //dane osobowe
            e.NewValues["tytul"] = nm.controlText("txTytul", grid);
            e.NewValues["imie"] = nm.controlText("txImie", grid);
            e.NewValues["nazwisko"] = nm.controlText("txNazwisko", grid);
            e.NewValues["data_poczatkowa"] = nm.controlTextDate("txPoczatekPowolania", grid);
            e.NewValues["data_koncowa"] = nm.controlTextDate("txDataKoncaPowolania", grid);
            if (nm.controlText("txPESEL", grid) == null)
            {
                e.NewValues["Pesel"] = 0;
            }
            else
            {
                e.NewValues["Pesel"] = nm.controlText("txPESEL", grid);

            }
            //d_zawieszenia

            var cos = (string)Session["czy_zawN"]; //nm.controlCheckbox("cbZawieszenie", grid);
            e.NewValues["czy_zaw"] = false;

            if (cos == "1")
            {
                e.NewValues["czy_zaw"] = true;
                DateTime poczZaw = (DateTime)Session["poczatekZawieszeniaN"];
                DateTime konZaw = (DateTime)Session["koniecZawieszeniaN"];

                e.NewValues["d_zawieszenia"] = poczZaw;
                e.NewValues["dataKoncaZawieszenia"] = konZaw;
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
            DateTime   dataKoncz =  DateTime.Parse ( DateTime.Now.AddYears(5).Year.ToString() +"-"+ DateTime.Now.AddMonths(1).Month .ToString("D2") + "-01").AddDays (-1);
            e.NewValues["data_koncowa"] = dataKoncz;
            //d_zawieszenia
            e.NewValues["d_zawieszenia"] = DateTime.Now;
            e.NewValues["dataKoncaZawieszenia"] = dataKoncz;
            string idOsoby = cl.dodaj_osobe(2,0);

            Session["idMediatora"] = idOsoby;
            Session["id_osoby"] = idOsoby;
        }

        protected void grid_StartRowEditing(object sender, ASPxStartRowEditingEventArgs e)
        {
            //txKoniecZawieszenia.
            string id = e.EditingKeyValue.ToString();
            Session["idMediatora"] = id;
            Session["id_osoby"] = id;
            object zawieszenie = grid.GetRowValuesByKeyValue(e.EditingKeyValue, "czy_zaw");
            var zawieszkax = zawieszenie.ToString();
            if (zawieszkax.ToString() == "True")
            {
                zawieszkax = "1";
            }
            else
            {
                zawieszkax = "0";
            }

                Session["czy_zaw"] = zawieszkax.ToString();
            if (zawieszkax.ToString() == "1")
            {
                object poczatekZawieszenia = grid.GetRowValuesByKeyValue(e.EditingKeyValue, "d_zawieszenia");
                Session["poczatekZawieszenia"] = (DateTime)poczatekZawieszenia;
                object koniecZawieszenia = grid.GetRowValuesByKeyValue(e.EditingKeyValue, "dataKoncaZawieszenia");
                Session["koniecZawieszenia"] = (DateTime)koniecZawieszenia;
            }


        }


        protected void grid_RowInserting(object sender, ASPxDataInsertingEventArgs e)
        {
            //dane osobowe
            e.NewValues["tytul"] = nm.controlText("txTytul", grid);
            e.NewValues["imie"] = nm.controlText("txImie", grid);
            e.NewValues["nazwisko"] = nm.controlText("txNazwisko", grid);
            e.NewValues["data_poczatkowa"] = nm.controlTextDate("txPoczatekPowolania", grid);
            e.NewValues["data_koncowa"] = nm.controlTextDate("txDataKoncaPowolania", grid);
            if (nm.controlText("txPESEL", grid) == null)
            {
                e.NewValues["Pesel"] = 0;
            }
            else
            {
                e.NewValues["Pesel"] = nm.controlText("txPESEL", grid);

            }
            //d_zawieszenia

            var cos = (string)Session["czy_zawN"]; //nm.controlCheckbox("cbZawieszenie", grid);
            e.NewValues["czy_zaw"] = false;

            if (cos == "1")
            {
                e.NewValues["czy_zaw"] = true;
                DateTime poczZaw = (DateTime)Session["poczatekZawieszeniaN"];
                DateTime konZaw = (DateTime)Session["koniecZawieszeniaN"];

                e.NewValues["d_zawieszenia"] = poczZaw;
                e.NewValues["dataKoncaZawieszenia"] = konZaw;
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

        protected void grid_CancelRowEditing(object sender, ASPxStartRowEditingEventArgs e)
        {
            var cos = e.EditingKeyValue;
            if (e.EditingKeyValue==null)
            {
                try
                {
                    int idOsoby = int.Parse((string)Session["id_osoby"]);
                    nm.usunTworzonaOsobe(idOsoby);
                }
                catch (Exception ex)
                {
                    
                }
             

            }

        } // end of grid_CancelRowEditing

      

        protected void grid_RowValidating(object sender, ASPxDataValidationEventArgs e)
        {
            var nazwisko = e.NewValues["nazwisko"];
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
            string kwerendabazowa = "SELECT DISTINCT ulica, kod_poczt, miejscowosc, czy_zaw, tel2, email, d_zawieszenia, dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja FROM tbl_osoby WHERE (czyus = 0) AND (typ = 2) AND (data_koncowa >= GETDATE())";
            Session["kwerenda"] = kwerendabazowa;
            if (!ASPxCheckBox1 .Checked)
            {
                Session["kwerenda"] = kwerendabazowa;
            }
            else
            {
                Session["kwerenda"] = "SELECT DISTINCT ulica, kod_poczt, miejscowosc, czy_zaw, tel2, email, d_zawieszenia, dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja FROM tbl_osoby WHERE (czyus = 0) AND (typ = 2) AND (data_koncowa < GETDATE())";
            }
            // po specjalizacji
            if (DropDownList1.SelectedIndex==-1)
            {
                DropDownList1.SelectedIndex = 0;
            }
            string kwerenda=(string)Session["kwerenda"];
            try
            {
                if (ASPxCheckBox2.Checked)
                {
                    string specjalizacja = DropDownList1.SelectedValue;
                  
                    kwerenda = kwerenda + "  and (select count(*) from tbl_specjalizacje_osob where id_specjalizacji =" + specjalizacja.Trim() + " and id_osoby=tbl_osoby.ident )=1 ";
                }
               
                
            }
            catch (Exception)
            {   }
            Session["kwerenda"] = kwerenda;
            mediatorzy.DataBind();


        }

        protected void zminaArchiwum(object sender, EventArgs e)
        {

            ustawKwerendeOdczytu();
        }

        protected void ASPxCheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            DropDownList1.Enabled = ASPxCheckBox2.Checked;
            ustawKwerendeOdczytu();
        }

        protected void _print(object sender, EventArgs e)
        {
            
            using (MemoryStream ms = new MemoryStream())
            {
                PrintableComponentLink pcl = new PrintableComponentLink(new PrintingSystem());
                //listaBieglych.Columns[0].Visible = false;
                
                ASPxGridViewExporter1.FileName = "Wykaz biegłych";
                pcl.Component = ASPxGridViewExporter1;

                pcl.Margins.Left = pcl.Margins.Right = 50;
                pcl.Landscape = true;
                pcl.CreateDocument(false);
                pcl.PrintingSystem.Document.AutoFitToPagesWidth = 1;
                pcl.ExportToPdf(ms);
                WriteResponse(this.Response, ms.ToArray(), System.Net.Mime.DispositionTypeNames.Inline.ToString());
            }

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

    }

}