using OfficeOpenXml;

using DevExpress.Office.Internal;
using DevExpress.Office.Utils;
using DevExpress.Web;
using DevExpress.Web.ASPxTreeList;
using DevExpress.Web.Data;
using DevExpress.XtraPrinting;

using DevExpress.XtraPrinting.Native;
using DevExpress.XtraPrintingLinks;
using DevExpress.XtraRichEdit.Import.Doc;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mime;
using System.Threading;
using System.Web;
using System.Web.DynamicData;
using System.Web.Services.Description;
using System.Web.UI.WebControls;
using OfficeOpenXml.Style;
using System.Text;



namespace wab2018
{
    public partial class biegliLista : System.Web.UI.Page
    {
        public  tabele tb = new tabele();
        private readonly nowiMediatorzy nm = new nowiMediatorzy();
        private cm Cm = new cm();
        private Class2 cl = new Class2();
        private string pesel = string.Empty;
        protected void Page_Load(object sender, EventArgs e)
        {
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
            var tel1 = nm.controlText("txTelefon1", grid);
            var tel2 = nm.controlText("txTelefon2", grid);
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
                var tel1 = nm.controlText("txTelefon1", grid);
                var tel2 = nm.controlText("txTelefon2", grid);
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
                        //czynni 
                        if (SpecjalizacjeCheckBox.Checked)
                        {
                            string specjalizacja = DropDownList1.SelectedValue;
                            nazwaSpeckajlizacji = NazwaSpecjalizacji(specjalizacja);

                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE  (data_koncowa >= GETDATE()) and (czyus = 0) and typ = 1 ";
                            kwerendabazowa = kwerendabazowa + "  and (select count(*) from tbl_specjalizacje_osob where id_specjalizacji =" + specjalizacja.Trim() + " and id_osoby=tbl_osoby.ident )=1 ";

                            Session["kwerenda"] = kwerendabazowa;
                        }
                        else
                        {
                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (data_koncowa >= GETDATE()) and (czyus = 0) and typ = 1 ";
                        }

                    }
                    break;
                case 3:
                    {
                        //Archiwalni 
                        if (SpecjalizacjeCheckBox.Checked)
                        {
                            string specjalizacja = DropDownList1.SelectedValue;
                            nazwaSpeckajlizacji = NazwaSpecjalizacji(specjalizacja);

                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (czyus = 0) AND (typ >= 2) AND (data_koncowa <= GETDATE()) and typ =1 ";
                            kwerendabazowa = kwerendabazowa + "  and (select count(*) from tbl_specjalizacje_osob where id_specjalizacji =" + specjalizacja.Trim() + " and id_osoby=tbl_osoby.ident )=1 ";

                            Session["kwerenda"] = kwerendabazowa;

                        }
                        else
                        {
                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (data_koncowa <= GETDATE()) and typ = 1 ";
                        }

                    }
                    break;
                default:
                    {
                        if (SpecjalizacjeCheckBox.Checked)
                        {
                            string specjalizacja = DropDownList1.SelectedValue;
                            nazwaSpeckajlizacji = NazwaSpecjalizacji(specjalizacja);

                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (czyus  = 0) And (typ = 1) ";
                            kwerendabazowa = kwerendabazowa + "  and (select count(*) from tbl_specjalizacje_osob where id_specjalizacji =" + specjalizacja.Trim() + " and id_osoby=tbl_osoby.ident )=1 ";

                            Session["kwerenda"] = kwerendabazowa;

                        }
                        else
                        {
                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ  FROM tbl_osoby WHERE (typ = 1) ";
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
            DropDownList1.Enabled = SpecjalizacjeCheckBox.Checked;
            ustawKwerendeOdczytu();
        }


        private string GetSpec(string imie, string nazwisko)
        {
            DataTable parametry = Cm.makeParameterTable();
            parametry.Rows.Add("@imie", imie);
            parametry.Rows.Add("@nazwisko", nazwisko);
          var ident = Cm.runQuerryWithResult("SELECT ident   FROM [tbl_osoby] where imie=@imie and nazwisko=@nazwisko", Cm.con_str, parametry);
            parametry = Cm.makeParameterTable();
            parametry.Rows.Add("@ident", ident);
            
            DataTable specki= Cm.getDataTable("SELECT DISTINCT dbo.glo_specjalizacje.nazwa FROM dbo.tbl_specjalizacje_osob INNER JOIN dbo.glo_specjalizacje ON dbo.tbl_specjalizacje_osob.id_specjalizacji = dbo.glo_specjalizacje.id_ WHERE     (dbo.tbl_specjalizacje_osob.id_osoby = @ident)", Cm.con_str, parametry);

            string SpeckiTxt = string.Empty;

            if (specki.Rows.Count>0)
            {
                foreach (DataRow row in specki.Rows)
                {


                    SpeckiTxt += row[0].ToString() +Environment.NewLine;
                }
            }
            return SpeckiTxt;
        }

        protected void _print(object sender, EventArgs e)

        {


            DataTable dt = new DataTable();
            foreach (GridViewColumn column in grid.VisibleColumns)
            {
                var col = column as GridViewDataColumn;
                if (col != null)
                    dt.Columns.Add(col.FieldName);
            }
            for (int i = 0; i < grid.VisibleRowCount; i++)
            {
                DataRow row = dt.NewRow();
                foreach (GridViewColumn column in grid.VisibleColumns)
                {
                    var col = column as GridViewDataColumn;
                    if (col != null)
                    {
                        var cellValue = grid.GetRowValues(i, col.FieldName);
                        row[col.FieldName] = cellValue;
                    }
                }
                dt.Rows.Add(row);
            }
            IList<DoWydruku> doWydrukuLista = new List<DoWydruku>();
            
            DoWydruku doWydruku = new DoWydruku();
            foreach (DataRow item in dt.Rows)
            {
                
                var zaw = item[5].ToString();
                if (zaw == "0") { zaw = ""; } else { zaw = "zawieszono"; };
              
                doWydruku = new DoWydruku();
                doWydruku.tytul = item[0].ToString();
                doWydruku.nazwisko = item[1].ToString();
                doWydruku.imie = item[3].ToString();
                doWydruku.powolanieOd = item[2].ToString(); 
                doWydruku.zawieszono = zaw;  
                doWydruku.telefon = item[6].ToString();
                doWydruku.uwagi = item[7].ToString();
                doWydruku.spejalizacje = GetSpec(item[3].ToString(), item[1].ToString());
                doWydrukuLista.Add(doWydruku);
            }

            robRaportDoWydruku(doWydrukuLista);

         
        }

        protected void _excell(object sender, EventArgs e)

        {


            DataTable dt = new DataTable();
            foreach (GridViewColumn column in grid.VisibleColumns)
            {
                var col = column as GridViewDataColumn;
                if (col != null)
                    dt.Columns.Add(col.FieldName);
            }
            for (int i = 0; i < grid.VisibleRowCount; i++)
            {
                DataRow row = dt.NewRow();
                foreach (GridViewColumn column in grid.VisibleColumns)
                {
                    var col = column as GridViewDataColumn;
                    if (col != null)
                    {
                        var cellValue = grid.GetRowValues(i, col.FieldName);
                        row[col.FieldName] = cellValue;
                    }
                }
                dt.Rows.Add(row);
            }
            DataTable excelTable = new DataTable();
            excelTable.Columns.Add("tytul",typeof(string));
            excelTable.Columns.Add("imie", typeof(string));
            excelTable.Columns.Add("nazwisko", typeof(string));
            excelTable.Columns.Add("powolanie", typeof(string));
            excelTable.Columns.Add("zawieszono", typeof(string));
            excelTable.Columns.Add("telefon", typeof(string));
            excelTable.Columns.Add("uwagi", typeof(string));
            excelTable.Columns.Add("specjalizacje", typeof(string));
            foreach (DataRow item in dt.Rows)
            {
                DataRow dr = excelTable.NewRow();
                var zaw = item[5].ToString();
                if (zaw == "0") { zaw = ""; } else { zaw = "zawieszono"; };

                dr[0] = item[0].ToString();//tytul
                dr[1] = item[1].ToString();//imie
                dr[2] = item[3].ToString();//nazwisko
                try
                {
                    dr[3] = item[2].ToString().Substring(0,11);//powolanie
                }
                catch 
                {
                    dr[3] ="";

                }

                dr[4] = zaw;// item[5].ToString();
                dr[5] = item[6].ToString();//uwagi
                dr[6] = item[7].ToString();//uwagi
                dr[7] = GetSpec(item[3].ToString(), item[1].ToString());
                excelTable.Rows.Add(dr);
            }





            string tenPlikNazwa = "Zestawienie";
            string path = Server.MapPath("Templates") + "\\" + tenPlikNazwa + ".xlsx";
            FileInfo existingFile = new FileInfo(path);
            if (!existingFile.Exists)
            {
                return;
            }
            string download = Server.MapPath("Templates") + @"\" + tenPlikNazwa + "";

            FileInfo fNewFile = new FileInfo(download + "_.xlsx");

            using (ExcelPackage MyExcel = new ExcelPackage(existingFile))
            {
                ExcelWorksheet MyWorksheet1 = MyExcel.Workbook.Worksheets[1];

              

                MyWorksheet1 = tb.tworzArkuszwExcle(MyExcel.Workbook.Worksheets[1], excelTable, 8, 0, 2, true, true, false, false, false);
                
                try
                {
                    MyExcel.SaveAs(fNewFile);

                    this.Response.Clear();
                    this.Response.ContentType = "application/vnd.ms-excel";
                    this.Response.AddHeader("Content-Disposition", "attachment;filename=" + fNewFile.Name);
                    this.Response.WriteFile(fNewFile.FullName);
                    this.Response.End();
                }
                catch (Exception ex)
                {

                }
            }//end of using


        }

        private void robRaportDoWydruku(IList<DoWydruku> ListaDoWydruku)
        {

            //nagłówek
            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10f, 10f, 10f, 0f);
           
            
            string path = Server.MapPath(@"~//pdf"); //Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments );
            string fileName = path + "//Zestawienie_Specjalizacji_" + DateTime.Now.ToString().Replace(":", "-") + ".pdf";
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(fileName, FileMode.Create));

            pdfDoc.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
            pdfDoc.Open();

            pdfDoc.AddTitle("zestawienie_Specjalizacji");
            pdfDoc.AddCreationDate();
            PdfPTable table = new PdfPTable(8);

            table.AddCell(new PdfPCell(new Phrase("Tytuł", cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Imię", cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Nazwisko" ,cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Powołanie do", cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Zawieszono" ,cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Telefon" ,cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Uwagi", cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Specjalizacje", cl.plFont2)));

            pdfDoc.Add(table);

            foreach (var item in ListaDoWydruku)
            {
                table = new  PdfPTable(8);
                table.AddCell(new PdfPCell(new Phrase(item.tytul.ToString(), cl.plFont2)));
                table.AddCell(new PdfPCell(new Phrase(item.imie.ToString(), cl.plFont2)));
                table.AddCell(new PdfPCell(new Phrase(item.nazwisko.ToString(), cl.plFont2)));
                string data = item.powolanieOd.ToString();
                try
                {
                    data = data.Substring(0, 10);
                }
                catch 
                {
                }
                
                table.AddCell(new PdfPCell(new Phrase(data , cl.plFont2)));
                table.AddCell(new PdfPCell(new Phrase(item.zawieszono.ToString(), cl.plFont2)));
                table.AddCell(new PdfPCell(new Phrase(item.telefon.ToString(), cl.plFont2)));

                table.AddCell(new PdfPCell(new Phrase(item.uwagi.ToString(), cl.plFont2)));
                table.AddCell(new PdfPCell(new Phrase(item.spejalizacje.ToString(), cl.plFont2)));


                pdfDoc.Add(table);
            }

            pdfDoc.Close();

         

            WebClient client = new WebClient();
            Byte[] buffer = client.DownloadData(fileName);
            if (buffer != null)
            {
                Response.ContentType = "application/pdf";
                Response.AddHeader("content-lenght", buffer.Length.ToString());
                Response.BinaryWrite(buffer);
            }

            ProcessStartInfo startInfo = new ProcessStartInfo(fileName);
            Process.Start(startInfo);
        }

       
        protected void twórzZestawienie(object sender, EventArgs e)
        {


            if (SpecjalizacjeCheckBox.Checked)
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

            string kwerenda = "SELECT View_SpecjalizacjeIOsoby.ident, tbl_osoby.imie, tbl_osoby.nazwisko, tbl_osoby.ulica, tbl_osoby.kod_poczt, tbl_osoby.miejscowosc, tbl_osoby.data_poczatkowa, tbl_osoby.data_koncowa, tbl_osoby.id_kreatora, tbl_osoby.data_kreacji, tbl_osoby.pesel, tbl_osoby.czyus, typ , tbl_osoby.tytul, tbl_osoby.czy_zaw, tbl_osoby.tel1 , tbl_osoby.tel2, tbl_osoby.email, tbl_osoby.adr_kores, tbl_osoby.kod_poczt_kor, tbl_osoby.miejscowosc_kor, tbl_osoby.uwagi, tbl_osoby.specjalizacjeWidok, tbl_osoby.specjalizacja_opis,                   tbl_osoby.d_zawieszenia, tbl_osoby.typ, tbl_osoby.dataKoncaZawieszenia, tbl_osoby.instytucja, View_SpecjalizacjeIOsoby.nazwa, View_SpecjalizacjeIOsoby.id_ as identyfikatorSpecjalizacji,                   View_SpecjalizacjeIOsoby.Expr1 AS aktwnaSpecjalizacja FROM     tbl_osoby RIGHT OUTER JOIN                   View_SpecjalizacjeIOsoby ON tbl_osoby.ident = View_SpecjalizacjeIOsoby.ident WHERE (tbl_osoby.nazwisko IS NOT NULL) AND (tbl_osoby.typ < 2) AND (View_SpecjalizacjeIOsoby.Expr1 = 1)";
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
                string tel = biegly["tel1"].ToString();
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


         
            try
            {
                DateTime data_koncowa = Convert.ToDateTime(e.GetValue("data_koncowa"));
              
                int wskaznik = DateTime.Compare(data_koncowa, DateTime.UtcNow);
                
                if (wskaznik < 0)
                    e.Row.BackColor = System.Drawing.Color.LightYellow;

            }
            catch 
            { }
            
            
        }


      

        public class tabele
        {

         //   private common cm = new common();
          //  private dataReaders dr = new dataReaders();

            public TableCell HeaderCell_(string text, int columns, int rows)
            {
                TableCell HeaderCell = new TableCell();
                HeaderCell.Text = text;
                HeaderCell.ColumnSpan = columns;
                HeaderCell.RowSpan = rows;
                return HeaderCell;
            }

            public GridViewRow Grw(object sender)
            {
                GridViewRow HeaderGridRow = null;
                GridView HeaderGrid = (GridView)sender;
                HeaderGridRow = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);
                HeaderGridRow.Font.Size = 7;
                HeaderGridRow.HorizontalAlign = HorizontalAlign.Center;
                HeaderGridRow.VerticalAlign = VerticalAlign.Top;
                return HeaderGridRow;
            }

            public DataTable makeSumRow(DataTable table, int ilKolumn)
            {
                DataTable wyjsciowa = new DataTable();
                for (int i = 0; i < ilKolumn; i++)
                {
                    wyjsciowa.Columns.Add("d_" + i.ToString("D2"), typeof(double));
                }
                DataTable tabelka = tabellaLiczbowa(table);
                DataColumnCollection col = tabelka.Columns;

                object sumObject;
                DataRow wiersz = wyjsciowa.NewRow();
                for (int i = 1; i < ilKolumn; i++)
                {
                    try
                    {
                        string idkolumny = "d_" + (i).ToString("D2");

                        if (col.Contains(idkolumny))
                        {
                            sumObject = tabelka.Compute("Sum(" + idkolumny + ")", "");
                            wiersz[idkolumny] = double.Parse(sumObject.ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        
                    }
                }
                wyjsciowa.Rows.Add(wiersz);
                return wyjsciowa;
            }

            public DataTable makeSumRow(DataTable table, int ilKolumn, int dlugoscLinii)
            {
                //stworzenie tabeli przejściowej
                DataTable tabelaRobocza = new DataTable();
                for (int i = 1; i <= dlugoscLinii; i++)
                {
                    string nazwaKolumny = "d_" + i.ToString("D2");
                    tabelaRobocza.Columns.Add(nazwaKolumny, typeof(double));
                }

                foreach (DataRow jedenWiersz in table.Rows)
                {
                    int index = 1;
                    DataRow nowyWiersz = tabelaRobocza.NewRow();
                    for (int i = 1; i <= ilKolumn; i++)
                    {
                        string nazwaKolumny = "d_" + i.ToString("D2");
                        string nazwaKolumnySumy = "d_" + index.ToString("D2");
                        double wartosc = double.Parse(jedenWiersz[nazwaKolumny].ToString());
                        double wartoscSumowana = 0;
                        try
                        {
                            wartoscSumowana = double.Parse(nowyWiersz[nazwaKolumnySumy].ToString());
                        }
                        catch
                        { }
                        wartoscSumowana = wartoscSumowana + wartosc;
                        nowyWiersz[nazwaKolumnySumy] = wartoscSumowana;
                        if (index == dlugoscLinii)
                        {
                            index = 0;
                        }
                        index++;
                    }
                    tabelaRobocza.Rows.Add(nowyWiersz);
                }

                object sumObject;
                DataRow wiersz = tabelaRobocza.NewRow();
                for (int i = 1; i <= tabelaRobocza.Columns.Count; i++)
                {
                    try
                    {
                        string idkolumny = "d_" + (i).ToString("D2");

                        sumObject = tabelaRobocza.Compute("Sum(" + idkolumny + ")", "");
                        wiersz[idkolumny] = double.Parse(sumObject.ToString());
                    }
                    catch (Exception ex)
                    {
                       
                    }
                }
                tabelaRobocza.Rows.Clear();
                tabelaRobocza.Rows.Add(wiersz);
                return tabelaRobocza;
            }

            public void makeHeader(System.Web.UI.WebControls.GridView GridViewName, DataTable dT, System.Web.UI.WebControls.GridView GridViewX)
            {
                try
                {
                    int row = 0;
                    TableCell HeaderCell = new TableCell();
                    GridViewRow HeaderGridRow = null;
                    string hv = "h";
                    Style stl = new Style();
                    foreach (DataRow dR in dT.Rows)
                    {
                        if (int.Parse(dR[0].ToString().Trim()) > row)
                        {
                            GridView HeaderGrid = (GridView)GridViewName;
                            HeaderGridRow = Grw(GridViewName);
                            row = int.Parse(dR[0].ToString().Trim());
                            try
                            {
                                hv = dR[4].ToString().Trim();
                            }
                            catch
                            { }
                        }
                        if (hv == "v")
                        {
                            stl.CssClass = "verticaltext";
                        }
                        else
                        {
                            stl.CssClass = "horizontaltext";
                        }

                        HeaderCell = new TableCell();
                        HeaderCell.Text = dR[1].ToString().Trim();
                        HeaderCell.Style.Clear();
                        HeaderCell.ApplyStyle(stl);
                        HeaderCell.ColumnSpan = int.Parse(dR[2].ToString().Trim());
                        HeaderCell.RowSpan = int.Parse(dR[3].ToString().Trim());
                        HeaderGridRow.Cells.Add(HeaderCell);
                        GridViewX.Controls[0].Controls.AddAt(0, HeaderGridRow);
                    }
                }
                catch (Exception ex)
                {
                } // end of try
            }

            public void makeHeader(DataTable dT, System.Web.UI.WebControls.GridView GridViewX)
            {
                System.Web.UI.WebControls.GridView sn = new System.Web.UI.WebControls.GridView();
                try
                {
                    int row = 0;
                    TableCell HeaderCell = new TableCell();
                    GridViewRow HeaderGridRow = null;
                    string hv = "h";
                    Style stl = new Style();
                    foreach (DataRow dR in dT.Rows)
                    {
                        if (int.Parse(dR[0].ToString().Trim()) > row)
                        {
                            GridView HeaderGrid = (GridView)sn;
                            HeaderGridRow = Grw(sn);
                            row = int.Parse(dR[0].ToString().Trim());
                            try
                            {
                                hv = dR[4].ToString().Trim();
                            }
                            catch
                            { }
                        }
                        if (hv == "v")
                        {
                            stl.CssClass = "verticaltext";
                        }
                        else
                        {
                            stl.CssClass = "horizontaltext";
                        }

                        HeaderCell = new TableCell();
                        HeaderCell.Text = dR[1].ToString().Trim();
                        HeaderCell.Style.Clear();
                        HeaderCell.ApplyStyle(stl);
                        HeaderCell.ColumnSpan = int.Parse(dR[2].ToString().Trim());
                        HeaderCell.RowSpan = int.Parse(dR[3].ToString().Trim());
                        HeaderGridRow.Cells.Add(HeaderCell);
                        GridViewX.Controls[0].Controls.AddAt(0, HeaderGridRow);
                    }
                }
                catch (Exception ex)
                {
                
                } // end of try
            }

            public void makeHeaderT3(DataTable dT, System.Web.UI.WebControls.GridView GridViewX)
            {
                System.Web.UI.WebControls.GridView sn = new System.Web.UI.WebControls.GridView();
                try
                {
                    int row = 0;
                    TableCell HeaderCell = new TableCell();
                    GridViewRow HeaderGridRow = null;
                    string hv = "h";
                    Style stl = new Style();
                    foreach (DataRow dR in dT.Rows)
                    {
                        if (int.Parse(dR[0].ToString().Trim()) > row)
                        {
                            GridView HeaderGrid = (GridView)sn;
                            HeaderGridRow = Grw(sn);
                            row = int.Parse(dR[0].ToString().Trim());
                            try
                            {
                                hv = dR[4].ToString().Trim();
                            }
                            catch
                            { }
                        }
                        if (hv == "v")
                        {
                            stl.CssClass = "spetialVertical";
                        }
                        else
                        {
                            stl.CssClass = "horizontaltext ";
                        }

                        HeaderCell = new TableCell();
                        HeaderCell.Text = dR[1].ToString().Trim();
                        HeaderCell.Style.Clear();
                        HeaderCell.ApplyStyle(stl);
                        HeaderCell.ColumnSpan = int.Parse(dR[2].ToString().Trim());
                        HeaderCell.RowSpan = int.Parse(dR[3].ToString().Trim());
                        HeaderGridRow.Cells.Add(HeaderCell);
                        GridViewX.Controls[0].Controls.AddAt(0, HeaderGridRow);
                    }
                }
                catch (Exception ex)
                {
                    
                } // end of try
            }

            public DataTable tabellaLiczbowa(DataTable tabelaWejsciowa)
            {
                if (tabelaWejsciowa == null)
                {
                    return null;
                }
                DataTable tabelaRobocza = new DataTable();
                int iloscKolumn = tabelaWejsciowa.Columns.Cast<DataColumn>().Count(c => c.ColumnName.StartsWith("d_"));

                for (int i = 1; i <= iloscKolumn; i++)
                {
                    string nazwaKolumny = "d_" + i.ToString("D2");
                    tabelaRobocza.Columns.Add(nazwaKolumny, typeof(double));
                }
                foreach (DataRow Drow in tabelaWejsciowa.Rows)
                {
                    try
                    {
                        try
                        {
                            if (Drow["nazwisko"].ToString().Trim() == "")
                            {
                                continue;
                            }
                        }
                        catch
                        {
                        }

                        DataRow wierszTymczasowy = tabelaRobocza.NewRow();
                        for (int i = 1; i <= iloscKolumn; i++)
                        {
                            string dana = string.Empty;
                            string nazwaKolumny = "d_" + i.ToString("D2");
                            double liczba = 0;
                            dana = Drow[nazwaKolumny].ToString();
                            if (string.IsNullOrEmpty(dana.Trim()))
                            {
                                dana = "0";
                            }
                            else
                            {
                                try
                                {
                                    dana = dana.Replace(".", ",");
                                    liczba = double.Parse(dana);
                                }
                                catch (Exception ex)
                                {
                                    
                                }
                            }

                            wierszTymczasowy[nazwaKolumny] = liczba;
                        }
                        tabelaRobocza.Rows.Add(wierszTymczasowy);
                    }
                    catch (Exception ex)
                    {
                        
                    }
                }

                return tabelaRobocza;
            }

            public void makeSumRow(DataTable table, GridViewRowEventArgs e, string tenplik)
            {
                makeSumRow(table, e);
            }

            public void makeSumRow(DataTable table, GridViewRowEventArgs e)
            {
                makeSumRow(table, e, 1, "Ogółem");
            }

            public void makeSumRow(DataTable table, GridViewRowEventArgs e, int przesuniecie)
            {
                makeSumRow(table, e, przesuniecie, "Ogółem");
            }

            public void makeSumRow(DataTable table, GridViewRowEventArgs e, int przesuniecie, string razem)
            {
                DataTable tabelka = tabellaLiczbowa(table);
                if (tabelka == null)
                {
                    
                    return;
                }
                object sumObject;
                int ilKolumn = e.Row.Cells.Count;
                e.Row.Cells[0 + przesuniecie].Text = razem;
                for (int i = 1; i < e.Row.Cells.Count; i++)
                {
                    try
                    {
                        string idkolumny = "d_" + (i).ToString("D2");
                        sumObject = tabelka.Compute("Sum(" + idkolumny + ")", "");
                        e.Row.Cells[i + przesuniecie].Text = sumObject.ToString();
                        e.Row.Cells[i + przesuniecie].CssClass = "center normal";
                    }
                    catch (Exception ex)
                    {
                       
                    }
                }
            }
            public void makeSumRow(DataTable table, GridViewRowEventArgs e, int przesuniecie, string razem, bool isGray)
            {
                DataTable tabelka = tabellaLiczbowa(table);
                if (tabelka == null)
                {
                 
                    return;
                }
                object sumObject;
                int ilKolumn = e.Row.Cells.Count;
                e.Row.Cells[0 + przesuniecie].Text = razem;
                for (int i = 1; i < e.Row.Cells.Count; i++)
                {
                    try
                    {
                        string idkolumny = "d_" + (i).ToString("D2");
                        sumObject = tabelka.Compute("Sum(" + idkolumny + ")", "");
                        e.Row.Cells[i + przesuniecie].Text = sumObject.ToString();
                        e.Row.Cells[i + przesuniecie].CssClass = "center normal gray";
                    }
                    catch (Exception ex)
                    {
                       
                    }
                }
            }

            public void makeSumRow(DataTable table, GridViewRowEventArgs e, int przesuniecie, int polaczenie)
            {
                DataTable tabelka = tabellaLiczbowa(table);
                if (tabelka == null)
                {
                    
                    return;
                }
                object sumObject;
                int ilKolumn = e.Row.Cells.Count;
                e.Row.Cells[0].ColumnSpan = polaczenie;
                e.Row.Cells[0].Text = "Ogółem";
                try
                {
                    for (int i = 1; i < polaczenie; i++)
                    {
                        e.Row.Cells.RemoveAt(1);
                    }
                }
                catch
                { }
                for (int i = 1; i < e.Row.Cells.Count; i++)
                {
                    try
                    {
                        string idkolumny = "d_" + (i).ToString("D2");
                        sumObject = tabelka.Compute("Sum(" + idkolumny + ")", "");
                        e.Row.Cells[i].Text = sumObject.ToString();
                        e.Row.Cells[i].CssClass = "center normal";
                    }
                    catch (Exception ex)
                    {
                       
                    }
                }
            }

            public GridViewRow PodsumowanieTabeli(DataTable dane, int iloscKolumn, string cssStyleDlaTabeli)
            {
                DataTable tabelka = tabellaLiczbowa(dane);
                if (tabelka == null)
                {
                    
                    return null;
                }
                object sumObject;
                GridViewRow NewTotalRow = new GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert);
                NewTotalRow.Cells.Add(cela("Razem", 1, 2, cssStyleDlaTabeli));
                for (int i = 1; i < iloscKolumn; i++)
                {
                    try
                    {
                        string idkolumny = "d_" + (i).ToString("D2");
                        
                        sumObject = tabelka.Compute("Sum(" + idkolumny + ")", "");

                        NewTotalRow.Cells.Add(cela(sumObject.ToString(), 1, 1, cssStyleDlaTabeli));
                    }
                    catch (Exception ex)
                    {
                       
                    }
                }
                return NewTotalRow;
            }


            //tabele pod dynamicznymi
            public GridViewRow wierszTabeli(DataTable dane, int iloscKolumn, int idWiersza, string idtabeli, string tekst, int colSpan, int rowSpan, string CssStyleDlaTekstu, string cssStyleDlaTabeli, string drugiText, int colSpanDrugi, int rowSpanDrugi, string cssStyleDrugi)
            {
                // nowy wiersz

                GridViewRow NewTotalRow = new GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert);
                NewTotalRow.Cells.Add(cela(drugiText, colSpanDrugi, rowSpanDrugi, cssStyleDrugi));

                NewTotalRow.Cells.Add(cela(tekst, rowSpan, colSpan, CssStyleDlaTekstu));
                DataRow jedenWiersz = dane.Rows[idWiersza - 1];
                for (int i = 1; i < iloscKolumn; i++)
                {
                    try
                    {
                        string nazwaKolumny = "d_" + i.ToString("D2");
                        NewTotalRow.Cells.Add(cela("<a class='" + CssStyleDlaTekstu + "' href=\"javascript: openPopup('popup.aspx?sesja=" + idWiersza.ToString().Trim() + "!" + idtabeli.ToString().Trim() + "!" + i.ToString().Trim() + "!3')\">" + jedenWiersz[nazwaKolumny].ToString().Trim() + "</a>", 1, 1, cssStyleDlaTabeli));
                    }
                    catch (Exception ex)
                    {
                       
                    }
                }
                return NewTotalRow;
            }




            public GridViewRow wierszTabeli(DataTable dane, int iloscKolumn, int idWiersza, string idtabeli, string tekst, int colSpan, int rowSpan, string CssStyleDlaTekstu, string cssStyleDlaTabeli)
            {
                if (dane == null)
                {
                    return null;
                }
                GridViewRow NewTotalRow = new GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert);
                NewTotalRow.Cells.Add(cela(tekst, rowSpan, colSpan, CssStyleDlaTekstu));
                DataRow jedenWiersz = dane.Rows[idWiersza - 1];
                for (int i = 1; i < iloscKolumn; i++)
                {
                    try
                    {
                        string nazwaKolumny = "d_" + i.ToString("D2");
                        NewTotalRow.Cells.Add(cela("<a class='" + CssStyleDlaTekstu + "' href=\"javascript: openPopup('popup.aspx?sesja=" + idWiersza.ToString().Trim() + "!" + idtabeli.ToString().Trim() + "!" + i.ToString().Trim() + "!3')\">" + jedenWiersz[nazwaKolumny].ToString().Trim() + "</a>", 1, 1, cssStyleDlaTabeli));
                    }
                    catch
                    {
                        try
                        {
                            NewTotalRow.Cells.Add(cela("<a class='" + CssStyleDlaTekstu + "' href=\"javascript: openPopup('popup.aspx?sesja=" + idWiersza.ToString().Trim() + "!" + idtabeli.ToString().Trim() + "!" + i.ToString().Trim() + "!3')\">0</a>", 1, 1, cssStyleDlaTabeli));
                        }
                        catch (Exception ex)
                        {
                           
                        }
                    }
                }
                return NewTotalRow;
            }// end of
            public GridViewRow wierszTabeli(DataTable dane, int iloscKolumn, int idWiersza, string idtabeli, string tekst, int colSpan, int rowSpan, string CssStyleDlaTekstu, string cssStyleDlaTabeli, bool isGray)
            {
                if (dane == null)
                {
                    return null;
                }
                if (isGray)
                {
                    CssStyleDlaTekstu = CssStyleDlaTekstu + " gray";
                    cssStyleDlaTabeli = cssStyleDlaTabeli + " gray";
                }
                GridViewRow NewTotalRow = new GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert);
                NewTotalRow.Cells.Add(cela(tekst, rowSpan, colSpan, CssStyleDlaTekstu));
                DataRow jedenWiersz = dane.Rows[idWiersza - 1];
                for (int i = 1; i < iloscKolumn; i++)
                {
                    try
                    {
                        string nazwaKolumny = "d_" + i.ToString("D2");
                        NewTotalRow.Cells.Add(cela("<a class='" + CssStyleDlaTekstu + "' href=\"javascript: openPopup('popup.aspx?sesja=" + idWiersza.ToString().Trim() + "!" + idtabeli.ToString().Trim() + "!" + i.ToString().Trim() + "!3')\">" + jedenWiersz[nazwaKolumny].ToString().Trim() + "</a>", 1, 1, cssStyleDlaTabeli));
                    }
                    catch
                    {
                        try
                        {
                            NewTotalRow.Cells.Add(cela("<a class='" + CssStyleDlaTekstu + "' href=\"javascript: openPopup('popup.aspx?sesja=" + idWiersza.ToString().Trim() + "!" + idtabeli.ToString().Trim() + "!" + i.ToString().Trim() + "!3')\">0</a>", 1, 1, cssStyleDlaTabeli));
                        }
                        catch (Exception ex)
                        {
                           
                        }
                    }
                }
                return NewTotalRow;
            }// end of
            public GridViewRow wierszTabeli(DataTable dane, int iloscKolumn, int idWiersza, string idtabeli, string tekst, int colSpan, int rowSpan, string CssStyleDlaTekstu, string cssStyleDlaTabeli, bool isGray, bool lastIsGray)
            {
                if (dane == null)
                {
                    return null;
                }
                if (isGray)
                {
                    CssStyleDlaTekstu = CssStyleDlaTekstu + " gray";
                    cssStyleDlaTabeli = cssStyleDlaTabeli + " gray";
                }
                GridViewRow NewTotalRow = new GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert);
                NewTotalRow.Cells.Add(cela(tekst, rowSpan, colSpan, CssStyleDlaTekstu));
                DataRow jedenWiersz = dane.Rows[idWiersza - 1];
                for (int i = 1; i < iloscKolumn; i++)
                {
                    try
                    {
                        string nazwaKolumny = "d_" + i.ToString("D2");

                        if (lastIsGray && (i >= iloscKolumn - 2))
                        {

                            NewTotalRow.Cells.Add(cela("<a class='" + CssStyleDlaTekstu + "' href=\"javascript: openPopup('popup.aspx?sesja=" + idWiersza.ToString().Trim() + "!" + idtabeli.ToString().Trim() + "!" + i.ToString().Trim() + "!3')\">" + jedenWiersz[nazwaKolumny].ToString().Trim() + "</a>", 1, 1, cssStyleDlaTabeli + " gray"));
                        }
                        else
                            NewTotalRow.Cells.Add(cela("<a class='" + CssStyleDlaTekstu + "' href=\"javascript: openPopup('popup.aspx?sesja=" + idWiersza.ToString().Trim() + "!" + idtabeli.ToString().Trim() + "!" + i.ToString().Trim() + "!3')\">" + jedenWiersz[nazwaKolumny].ToString().Trim() + "</a>", 1, 1, cssStyleDlaTabeli));
                    }
                    catch
                    {
                        try
                        {
                            NewTotalRow.Cells.Add(cela("<a class='" + CssStyleDlaTekstu + "' href=\"javascript: openPopup('popup.aspx?sesja=" + idWiersza.ToString().Trim() + "!" + idtabeli.ToString().Trim() + "!" + i.ToString().Trim() + "!3')\">0</a>", 1, 1, cssStyleDlaTabeli));
                        }
                        catch (Exception ex)
                        {
                           
                        }
                    }
                }
                return NewTotalRow;
            }// end of

            public GridViewRow wierszTabeli(string[] lista, int iloscKolumn, int idWiersza, string tekst, int colSpan, int rowSpan, string CssStyleDlaTekstu, string cssStyleDlaTabeli, bool ostatniaEdytowalna)
            {
                GridViewRow NewTotalRow = new GridViewRow(0, 0, DataControlRowType.DataRow, DataControlRowState.Insert);
                NewTotalRow.Cells.Add(cela(tekst, rowSpan, colSpan, CssStyleDlaTekstu));

                for (int i = 1; i < iloscKolumn; i++)
                {
                    try
                    {
                        string nazwaKolumny = "d_" + i.ToString("D2");
                        string textC = lista[i];
                        NewTotalRow.Cells.Add(cela(
                            textC, 1, 1, cssStyleDlaTabeli));
                    }
                    catch (Exception ex)
                    {
                       
                    }
                }
                if (ostatniaEdytowalna)
                {
                    NewTotalRow.Cells.Add(cela("<input id = \"Text1\" type = \"text\" />", 1, 1, "borderTopLeft"));
                }
                return NewTotalRow;
            }// end of


            public TableCell cela(string text, int rowSpan, int colSpan, string cssClass)
            {
                TableCell HeaderCell = new TableCell();
                HeaderCell.Height = 10;
                HeaderCell.HorizontalAlign = HorizontalAlign.Center;
                HeaderCell.ColumnSpan = colSpan;
                HeaderCell.RowSpan = rowSpan;
                HeaderCell.CssClass = cssClass;
                HeaderCell.Text = text;
                return HeaderCell;
            }
            public TableCell cela(string text, int rowSpan, int colSpan, string cssClass, bool isGray)
            {
                TableCell HeaderCell = new TableCell();
                HeaderCell.Height = 10;
                HeaderCell.HorizontalAlign = HorizontalAlign.Center;
                HeaderCell.ColumnSpan = colSpan;
                HeaderCell.RowSpan = rowSpan;
                HeaderCell.CssClass = cssClass + " gray";
                HeaderCell.Text = text;
                return HeaderCell;
            }
            public ExcelWorksheet tworzArkuszwExcle(ExcelWorksheet Arkusz, DataTable daneDoArkusza, int iloscKolumn, int przesunięcieX, int przesuniecieY, bool lp, bool suma, bool stanowisko, bool funkcja, bool nazwiskoiImeieOsobno)
            {
                return tworzArkuszwExcle(Arkusz, daneDoArkusza, iloscKolumn, przesunięcieX, przesuniecieY, lp, suma, stanowisko, funkcja, nazwiskoiImeieOsobno, false);
            }

            public ExcelWorksheet tworzArkuszwExcle(ExcelWorksheet Arkusz, DataTable daneDoArkusza, int iloscKolumn, int przesunięcieX, int przesuniecieY, bool lp, bool suma, bool stanowisko, bool funkcja, bool nazwiskoiImeieOsobno, bool obramowanieOststniej)
            {
                if (daneDoArkusza == null)
                {
                    
                    return Arkusz;
                }
                try
                {
                    int wiersz = przesuniecieY;
                    int dod = 0;
                    foreach (DataRow dR in daneDoArkusza.Rows)
                    {
                        int dodatek = 0;

                        for (int i = 0; i < iloscKolumn; i++)
                        {
                            try
                            {
                            
                                try
                                {
                                    var value = dR[i].ToString().Trim();
                                    Arkusz.Cells[wiersz, i + przesunięcieX + dodatek+1].Value = value;
                                }
                                catch
                                {
                                    Arkusz.Cells[wiersz, i + przesunięcieX + dodatek].Value = (dR[i].ToString().Trim());
                                }
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek + i+1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                Arkusz.Cells[wiersz, i + przesunięcieX + dodatek+1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            }
                            catch (Exception ex)
                            {
                                
                            }
                        }
                      
                        wiersz++;
                        dod = dodatek;
                    }

                  
                }
                catch (Exception ex)
                {
                    
                }

                return Arkusz;
            }

            public ExcelWorksheet tworzArkuszwExcle(ExcelWorksheet Arkusz, DataTable daneDoArkusza, int iloscKolumn, int przesunięcieX, int przesuniecieY, bool lp, bool suma, bool stanowisko, bool funkcja, bool nazwiskoiImeieOsobno, bool obramowanieOststniej, bool pustaKolumnaZaNazwiskiem)
            {
                if (daneDoArkusza == null)
                {
                    
                    return Arkusz;
                }
                try
                {
                    int wiersz = przesuniecieY;
                    int dod = 0;
                    foreach (DataRow dR in daneDoArkusza.Rows)
                    {
                        int dodatek = 0;
                        if (lp)
                        {
                            try
                            {
                                dodatek++;

                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.ShrinkToFit = true;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = wiersz - przesuniecieY + 1;
                            }
                            catch (Exception ex)
                            {
                               
                            }
                        }
                        if (stanowisko)
                        {
                            try
                            {
                                dodatek++;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.ShrinkToFit = true;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                                string value = (dR["stanowisko"].ToString().Trim());
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = value;
                            }
                            catch (Exception ex)
                            {
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = "";
                                
                            }
                        }
                        if (funkcja)
                        {
                            try
                            {
                                dodatek++;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.ShrinkToFit = true;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Green);

                                string value = (dR["funkcja"].ToString().Trim());
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = value;
                            }
                            catch (Exception ex)
                            {
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = "";
                               
                            }
                        }
                        if (nazwiskoiImeieOsobno)
                        {
                            dodatek++;
                            try
                            {
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.ShrinkToFit = true;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Green);
                                string value = (dR["nazwisko"].ToString().Trim());
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = value;
                                dodatek++;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.ShrinkToFit = true;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Green);
                                value = (dR["imie"].ToString().Trim());
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = value;
                            }
                            catch (Exception ex)
                            {
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = "";
                               
                            }
                        }
                        else
                        {
                            try
                            {
                                dodatek++;
                                string value = dR["imie"].ToString().Trim() + " " + dR["nazwisko"].ToString().Trim();
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = value;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.ShrinkToFit = true;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            }
                            catch (Exception ex)
                            {
                               
                            }
                        }

                        if (pustaKolumnaZaNazwiskiem)
                        {
                            przesunięcieX = przesunięcieX + 1;
                        }

                        for (int i = 1; i < iloscKolumn; i++)
                        {
                            try
                            {
                                string colunmName = "d_" + (i).ToString("D2");
                                try
                                {
                                    double value = double.Parse(dR[colunmName].ToString().Trim());
                                    Arkusz.Cells[wiersz, i + przesunięcieX + dodatek].Value = value;
                                }
                                catch
                                {
                                    Arkusz.Cells[wiersz, i + przesunięcieX + dodatek].Value = (dR[colunmName].ToString().Trim());
                                }
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek + i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                Arkusz.Cells[wiersz, i + przesunięcieX + dodatek].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                            }
                            catch (Exception ex)
                            {
                                
                            }
                        }
                        if (obramowanieOststniej)
                        {
                            Arkusz.Cells[wiersz, przesunięcieX + dodatek + iloscKolumn].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            Arkusz.Cells[wiersz, iloscKolumn + przesunięcieX + dodatek].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        }
                        wiersz++;
                        dod = dodatek;
                    }

                    if (suma)
                    {
                        DataTable tabelka = tabellaLiczbowa(daneDoArkusza);
                        object sumObject;

                        Arkusz.Cells[wiersz, przesunięcieX + dod].Value = "Razem";
                        Arkusz.Cells[wiersz, przesunięcieX + dod].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        Arkusz.Cells[wiersz, przesunięcieX + dod].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        for (int i = 1; i < iloscKolumn; i++)
                        {
                            try
                            {
                                string idkolumny = "d_" + (i).ToString("D2");
                                sumObject = tabelka.Compute("Sum(" + idkolumny + ")", "");
                                Arkusz.Cells[wiersz, i + przesunięcieX + dod].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                                Arkusz.Cells[wiersz, i + przesunięcieX + dod].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                Arkusz.Cells[wiersz, i + przesunięcieX + dod].Value = (sumObject.ToString());
                            }
                            catch (Exception ecx)
                            {
                                string mes = ecx.Message;
                                Arkusz.Cells[wiersz, i + przesunięcieX + dod].Value = mes;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    
                }

                return Arkusz;
            }

            public ExcelWorksheet tworzArkuszwExcle(ExcelWorksheet Arkusz, DataTable daneDoArkusza, int iloscKolumn, int przesunięcieX, int przesuniecieY, bool lp, bool suma, bool stanowisko, bool funkcja, bool nazwiskoiImeieOsobno, bool obramowanieOststniej, bool przesuniecieiteracji, bool mp)
            {
                if (daneDoArkusza == null)
                {
                    
                    return Arkusz;
                }
                try
                {
                    int wiersz = przesuniecieY;
                    int dod = 0;
                    foreach (DataRow dR in daneDoArkusza.Rows)
                    {
                        if (dR["nazwisko"].ToString().Trim() == "")
                        {
                            continue;
                        }
                        int dodatek = 0;
                        if (lp)
                        {
                            try
                            {
                                dodatek++;

                                if (!przesuniecieiteracji)
                                {
                                    komorkaExcela(Arkusz, wiersz + 1, przesunięcieX + dodatek, (wiersz - przesuniecieY + 1).ToString(), false, 0, 0, true, false);
                                }
                                else
                                {
                                    komorkaExcela(Arkusz, wiersz + 1, przesunięcieX + dodatek, (wiersz - przesuniecieY + 1).ToString(), false, 0, 0, true, false);
                                }
                            }
                            catch (Exception ex)
                            {
                               
                            }
                        }
                        if (stanowisko)
                        {
                            try
                            {
                                dodatek++;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.ShrinkToFit = true;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);

                                string value = (dR["stanowisko"].ToString().Trim());
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = value;
                            }
                            catch (Exception ex)
                            {
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = "";
                                
                            }
                        }
                        if (funkcja)
                        {
                            try
                            {
                                dodatek++;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.ShrinkToFit = true;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Green);

                                string value = (dR["funkcja"].ToString().Trim());
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = value;
                            }
                            catch (Exception ex)
                            {
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = "";
                               
                            }
                        }
                        if (nazwiskoiImeieOsobno)
                        {
                            dodatek++;
                            try
                            {
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.ShrinkToFit = true;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Green);
                                string value = (dR["nazwisko"].ToString().Trim());
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Value = value;
                                dodatek++;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.ShrinkToFit = true;
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Green);
                                value = (dR["imie"].ToString().Trim());
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek + 1].Value = value;
                            }
                            catch (Exception ex)
                            {
                                Arkusz.Cells[wiersz, przesunięcieX + dodatek + 1].Value = "";
                               
                            }
                        }
                        else
                        {
                            try
                            {
                                dodatek++;
                                string value = dR["imie"].ToString().Trim() + " " + dR["nazwisko"].ToString().Trim();
                                if (!przesuniecieiteracji)
                                {
                                    komorkaExcela(Arkusz, wiersz, przesunięcieX + dodatek, value, false, 0, 0, false, false);
                                }
                                else
                                {
                                    komorkaExcela(Arkusz, wiersz + 1, przesunięcieX + dodatek, value, false, 0, 0, false, false);
                                }
                            }
                            catch (Exception ex)
                            {
                                
                            }
                        }

                        for (int i = 1; i < iloscKolumn; i++)
                        {
                            try
                            {
                                string colunmName = "d_" + (i).ToString("D2");
                                try
                                {
                                    komorkaExcela(Arkusz, wiersz + 1, i + przesunięcieX + dodatek, (dR[colunmName].ToString().Trim()), false, 0, 0, true, false);
                                }
                                catch
                                {
                                    Arkusz.Cells[wiersz + 1, i + przesunięcieX + dodatek].Value = "0";
                                }
                            }
                            catch
                            {
                                Arkusz.Cells[wiersz + 1, i + przesunięcieX + dodatek].Value = "0";
                            }
                        }
                        if (obramowanieOststniej)
                        {
                            Arkusz.Cells[wiersz, przesunięcieX + dodatek + iloscKolumn].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            Arkusz.Cells[wiersz, iloscKolumn + przesunięcieX + dodatek].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        }
                        wiersz++;
                        dod = dodatek;
                    }

                    if (suma)
                    {
                        DataTable tabelka = tabellaLiczbowa(daneDoArkusza);
                        object sumObject;
                        if (mp)
                        {
                            wiersz = wiersz + 1;
                        }
                        komorkaExcela(Arkusz, wiersz, przesunięcieX + dod, "Razem", false, 0, 0);
                        for (int i = 1; i <= iloscKolumn + 1; i++)
                        {
                            try
                            {
                                string idkolumny = "d_" + (i).ToString("D2");
                                sumObject = tabelka.Compute("Sum(" + idkolumny + ")", "");
                                komorkaExcela(Arkusz, wiersz, i + przesunięcieX + dod, sumObject.ToString(), false, 0, 0, true, false);
                            }
                            catch (Exception ecx)
                            {
                               
                            }
                        }
                    }
                    if (mp)
                    {
                        for (int i = 1; i < iloscKolumn; i++)
                        {
                            komorkaExcela(Arkusz, przesuniecieY, przesunięcieX + 2 + i, i.ToString(), false, 0, 0, true, false);
                        }
                    }
                }
                catch (Exception ex)
                {
                    
                }

                return Arkusz;
            }

            public ExcelWorksheet tworzArkuszwExcle(ExcelWorksheet Arkusz, DataTable daneDoArkusza, int iloscKolumn, int przesunięcieX, int przesuniecieY, bool lp, bool suma, bool stanowisko, bool funkcja, bool nazwiskoiImeieOsobno, bool obramowanieOststniej, string Linia01, string Linia02, string Linia03)
            {
                Arkusz = tworzArkuszwExcle(Arkusz, daneDoArkusza, iloscKolumn, przesunięcieX, przesuniecieY, lp, suma, stanowisko, funkcja, nazwiskoiImeieOsobno, obramowanieOststniej);

                Arkusz.Cells[1, 1].Value = Linia01; ;
                Arkusz.Cells[2, 1].Value = Linia02; ;
                Arkusz.Cells[3, 1].Value = Linia03; ;

                return Arkusz;
            }

             public ExcelWorksheet tworznaglowki(ExcelWorksheet Arkusz, DataTable daneDoArkusza, int iloscwierszy, int przesunięcieX, int przesuniecieY, string tekstNadTabela)
            {
                komorkaExcela(Arkusz, 1, 2, tekstNadTabela, false, 0, 0);
                int wiersz = przesuniecieY;
                for (int i = 1; i < iloscwierszy + 1; i++)
                {
                    try
                    {
                        komorkaExcela(Arkusz, przesuniecieY, i + przesunięcieX, daneDoArkusza.Rows[i - 1][1].ToString().Trim(), false, 0, 0, true, false);
                    }
                    catch (Exception ex)
                    {
                        //cm.log.Error("KP tworznaglowki " + ex.Message);
                    }
                }
                return Arkusz;
            }

            public void komorkaExcela(ExcelWorksheet Arkusz, int wiersz, int kolumna, string tekst, bool zlaczenie, int rowSpan, int colSpan)
            {
                komorkaExcela(Arkusz, wiersz, kolumna, tekst, zlaczenie, rowSpan, colSpan, false, false);
            }

            public void komorkaExcela(ExcelWorksheet Arkusz, int wiersz, int kolumna, string tekst, bool zlaczenie, int rowSpan, int colSpan, bool wycentrowanie, bool wyszarzenie)
            {
                if (zlaczenie)
                {
                    try
                    {
                        Arkusz.Cells[wiersz, kolumna, wiersz + rowSpan, kolumna + colSpan].Merge = true;
                        Arkusz.Cells[wiersz, kolumna, wiersz + rowSpan, kolumna + colSpan].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                        Arkusz.Cells[wiersz, kolumna, wiersz + rowSpan, kolumna + colSpan].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    catch (Exception ex)
                    {
                        //cm.log.Error("komorkaExcela merge " + ex.Message);
                    }
                }
                try
                {
                    Arkusz.Cells[wiersz, kolumna].Style.ShrinkToFit = true;
                    Arkusz.Cells[wiersz, kolumna].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                    if (wycentrowanie)
                    {
                        Arkusz.Cells[wiersz, kolumna].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    }
                    Arkusz.Cells[wiersz, kolumna].Value = tekst;
                    if (wyszarzenie)
                    {
                        Arkusz.Cells[wiersz, kolumna].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        Arkusz.Cells[wiersz, kolumna].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Gray);
                    }
                }
                catch (Exception ex)
                {
                    //cm.log.Error("komorkaExcela merge " + ex.Message);
                }

            }

            public DataTable naglowek(string plik, int numerArkusza)
            {
                if (string.IsNullOrEmpty(plik.Trim()))
                {
                    return null;
                }
                IList<string> komorki = new List<string>();

                DataTable schematNaglowka = new DataTable();
                schematNaglowka.Columns.Add("wiersz", typeof(int));
                schematNaglowka.Columns.Add("kolumna", typeof(int));
                schematNaglowka.Columns.Add("text", typeof(string));
                schematNaglowka.Columns.Add("rowSpan", typeof(int));
                schematNaglowka.Columns.Add("colSpan", typeof(int));

                var package = new ExcelPackage(new FileInfo(plik));
                using (package)
                {
                    int iloscZakladek = package.Workbook.Worksheets.Count;
                    if (iloscZakladek >= numerArkusza)
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[numerArkusza];

                        int rows = worksheet.Dimension.End.Row;
                        int columns = worksheet.Dimension.End.Column;

                        for (int i = 1; i <= rows; i++)
                        {
                            for (int j = 1; j <= columns; j++)
                            {
                                object baseE = worksheet.Cells[i, j];
                                ExcelCellBase celka = (ExcelCellBase)baseE;

                                bool polaczony = (bool)celka.GetType().GetProperty("Merge").GetValue(celka, null);
                                var kolumny = celka.GetType().GetProperty("Columns").GetValue(celka, null);
                                var wiersze = celka.GetType().GetProperty("Rows").GetValue(celka, null);
                                var text = celka.GetType().GetProperty("Value").GetValue(celka, null);

                                DataRow komorka = schematNaglowka.NewRow();
                                if (polaczony && text != null)
                                {
                                    IList<int> lista = okreslKomorke(i, j, rows, columns, worksheet);

                                    komorka["wiersz"] = i;
                                    komorka["kolumna"] = j;
                                    komorka["text"] = text;
                                    komorka["colSpan"] = lista[0].ToString();
                                    komorka["rowSpan"] = lista[1].ToString();

                                    schematNaglowka.Rows.Add(komorka);
                                    int k = lista[1];
                                    if (k > 1)
                                    {
                                        j = (j + k) - 1;
                                    }
                                }
                                else
                                {
                                    komorka["wiersz"] = i;
                                    komorka["kolumna"] = j;
                                    komorka["text"] = text;
                                    komorka["colSpan"] = 1;
                                    komorka["rowSpan"] = 1;
                                    if (text != null)
                                    {
                                        schematNaglowka.Rows.Add(komorka);
                                    }
                                }
                            }
                        }
                    }
                }

                DataTable dT_01 = new DataTable();
                dT_01.Columns.Clear();
                dT_01.Columns.Add("Column1", typeof(string));
                dT_01.Columns.Add("Column2", typeof(string));
                dT_01.Columns.Add("Column3", typeof(string));
                dT_01.Columns.Add("Column4", typeof(string));
                dT_01.Columns.Add("Column5", typeof(string));

                // max ilosc wierszy
                var max = schematNaglowka.Rows.OfType<DataRow>().Select(row => row["wiersz"]).Max();

                if (max != null)
                {
                    int wiersz = 0;
                    for (int i = (int)max; i >= 0; i--)
                    {
                        wiersz++;
                        //wyciągnij dane tylko z wierszem
                        string selectString = "wiersz=" + i.ToString();
                        DataRow[] jedenWiersz = schematNaglowka.Select(selectString);
                        foreach (var komorka in jedenWiersz)
                        {
                            dT_01.Rows.Add(new Object[] { wiersz.ToString(), komorka["text"], komorka["rowSpan"], komorka["colSpan"], "h" });
                        }
                    }
                }

                return dT_01;
            }

            protected IList<int> okreslKomorke(int wierszPoczatkowy, int kolumnaPoczatkowa, int iloscWierszy, int iloscKolumn, ExcelWorksheet worksheet)
            {
                IList<int> wyniki = new List<int>();
                int rowSpan = 0;
                int colSpan = 0;

                bool mergedY = false;

                for (int i = wierszPoczatkowy; i <= iloscWierszy + 1; i++)
                {
                    object baseE = worksheet.Cells[i, kolumnaPoczatkowa];

                    ExcelCellBase celka = (ExcelCellBase)baseE;
                    bool polaczony = (bool)celka.GetType().GetProperty("Merge").GetValue(celka, null);
                    var text = celka.GetType().GetProperty("Value").GetValue(celka, null);
                    if (!polaczony)
                    {
                        break;
                    }
                    else
                    {
                        if (mergedY)
                        {
                            if (text != null)
                            {
                                break;
                            }
                        }
                        mergedY = true;
                    }
                    rowSpan++;
                }
                bool mergedX = false;
                for (int j = kolumnaPoczatkowa; j <= iloscKolumn + 1; j++)
                {
                    object baseE = worksheet.Cells[wierszPoczatkowy, j];

                    ExcelCellBase celka = (ExcelCellBase)baseE;
                    bool polaczony = (bool)celka.GetType().GetProperty("Merge").GetValue(celka, null);
                    var text = celka.GetType().GetProperty("Value").GetValue(celka, null);
                    if (!polaczony)
                    {
                        break;
                    }
                    else
                    {
                        if (mergedX)
                        {
                            if (text != null)
                            {
                                break;
                            }
                        }
                        mergedX = true;
                    }
                    colSpan++;
                }
                wyniki.Add(rowSpan);
                wyniki.Add(colSpan);
                return wyniki;
            }

            public DataTable SchematTabelinaglowkowej()
            {
                DataTable tabelaNaglowkowa = new DataTable();
                tabelaNaglowkowa.Columns.Clear();
                tabelaNaglowkowa.Columns.Add("wiersz", typeof(string));
                tabelaNaglowkowa.Columns.Add("text", typeof(string));
                tabelaNaglowkowa.Columns.Add("Column3", typeof(string));
                tabelaNaglowkowa.Columns.Add("Column4", typeof(string));
                tabelaNaglowkowa.Columns.Add("Column5", typeof(string));
                tabelaNaglowkowa.Columns.Add("Column6", typeof(string));
                return tabelaNaglowkowa;
            }

            public string komorkaHTML(string text, int colspan, int rowspan, string style)
            {
                StringBuilder builder = new StringBuilder();
                builder.Append("<td ");
                if (!string.IsNullOrEmpty(style.Trim()))
                {
                    builder.Append(" class='" + style + "' ");
                }
                if (rowspan > 0)
                {
                    builder.Append(" rowspan='" + rowspan + "' ");
                }
                if (colspan > 0)
                {
                    builder.Append(" colspan='" + colspan + "' ");
                }
                builder.AppendLine(">");
                builder.AppendLine("<p>" + text + "</p>");
                builder.AppendLine("</td>");
                return builder.ToString();
            }

            public string komorkaHTMLbezP(string text, int colspan, int rowspan, string style)
            {
                StringBuilder builder = new StringBuilder();
                builder.Append("<td ");
                if (!string.IsNullOrEmpty(style.Trim()))
                {
                    builder.Append(" class='" + style + "' ");
                }
                if (rowspan > 0)
                {
                    builder.Append(" rowspan='" + rowspan + "' ");
                }
                if (colspan > 0)
                {
                    builder.Append(" colspan='" + colspan + "' ");
                }
                builder.AppendLine(">");

                builder.AppendLine(text);

                builder.AppendLine("</td>");
                return builder.ToString();
            }
        }
    }

    public class DoWydruku
    { 
        public string tytul { get; set; }

        public string nazwisko { get; set; }
        public string imie { get; set; }
        public string powolanieOd { get; set; }

        public string zawieszono { get; set; }

        public string telefon { get; set; }

        public string uwagi { get; set; }
        public string spejalizacje { get; set; }

    }
}