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
using NPOI.SS.Formula;
using DevExpress.Utils.OAuth;
using OfficeOpenXml.Table;
using DevExpress.XtraReports.Parameters;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using static NPOI.HSSF.Util.HSSFColor;




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
            bool  zawieszenie = false;
            if (HiddenField.Value=="true")
            {
                zawieszenie = true;
            }
          
            e.NewValues["czy_zaw"] = zawieszenie;
            if (zawieszenie)
            {
                e.NewValues["d_zawieszenia"] = nm.controlTextDate("txDataPoczatkuZawieszenia", grid);
                e.NewValues["dataKoncaZawieszenia"] = nm.controlTextDate("txDataKoncaZawieszenia", grid);
                e.NewValues["rodzaj_zawieszenia"] = nm.controlDropDownList("powZawDropDownList", grid);
                e.NewValues["Informacje_o_wstrzymaniu"] = nm.controlText("txInformacjeowstrzymaniu", grid);

            }
            else
            {
                e.NewValues["d_zawieszenia"] = DateTime.Now;
                e.NewValues["dataKoncaZawieszenia"] = DateTime.Now;
                e.NewValues["rodzaj_zawieszenia "]="";
                e.NewValues["Informacje_o_wstrzymaniu"] = "";
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
            e.NewValues["uwagiBIP"] = nm.controlTextMemo("txUwagiBIP", grid);
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
                e.NewValues["uwagiBIP"] = nm.controlTextMemo("txUwagiBIP", grid);
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

                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, uwagiBIP,  specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ, rodzaj_zawieszenia, Informacje_o_wstrzymaniu  FROM tbl_osoby WHERE  (data_koncowa >= GETDATE()) and (czyus = 0) and typ = 1 ";
                            kwerendabazowa = kwerendabazowa + "  and (select count(*) from tbl_specjalizacje_osob where id_specjalizacji =" + specjalizacja.Trim() + " and id_osoby=tbl_osoby.ident )=1 ";

                            Session["kwerenda"] = kwerendabazowa;
                        }
                        else
                        {
                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, uwagiBIP,  specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ,rodzaj_zawieszenia, Informacje_o_wstrzymaniu  FROM tbl_osoby WHERE (data_koncowa >= GETDATE()) and (czyus = 0) and typ = 1 ";
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

                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, uwagiBIP,  specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ,rodzaj_zawieszenia, Informacje_o_wstrzymaniu  FROM tbl_osoby WHERE (czyus = 0) AND (typ >= 2) AND (data_koncowa <= GETDATE()) and typ =1 ";
                            kwerendabazowa = kwerendabazowa + "  and (select count(*) from tbl_specjalizacje_osob where id_specjalizacji =" + specjalizacja.Trim() + " and id_osoby=tbl_osoby.ident )=1 ";

                            Session["kwerenda"] = kwerendabazowa;

                        }
                        else
                        {
                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, uwagiBIP,  specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ ,rodzaj_zawieszenia, Informacje_o_wstrzymaniu FROM tbl_osoby WHERE (data_koncowa <= GETDATE()) and typ = 1 ";
                        }

                    }
                    break;
                default:
                    {
                        if (SpecjalizacjeCheckBox.Checked)
                        {
                            string specjalizacja = DropDownList1.SelectedValue;
                            nazwaSpeckajlizacji = NazwaSpecjalizacji(specjalizacja);

                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, uwagiBIP,  specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ,rodzaj_zawieszenia, Informacje_o_wstrzymaniu  FROM tbl_osoby WHERE (czyus  = 0) And (typ = 1) ";
                            kwerendabazowa = kwerendabazowa + "  and (select count(*) from tbl_specjalizacje_osob where id_specjalizacji =" + specjalizacja.Trim() + " and id_osoby=tbl_osoby.ident )=1 ";

                            Session["kwerenda"] = kwerendabazowa;

                        }
                        else
                        {
                            kwerendabazowa = "SELECT ulica, kod_poczt, miejscowosc, COALESCE (czy_zaw, 0) AS czy_zaw, tel2, email, COALESCE (d_zawieszenia, '1900-01-01') AS d_zawieszenia, COALESCE (dataKoncaZawieszenia, '1900-01-01') AS dataKoncaZawieszenia, GETDATE() AS now, tytul, uwagi, uwagiBIP,  specjalizacja_opis, dbo.specjalizacjeLista(ident) AS specjalizacjeWidok, miejscowosc_kor, kod_poczt_kor, adr_kores, imie, ident, data_poczatkowa, data_koncowa, pesel, tel1, typ, nazwisko, instytucja, REPLACE(REPLACE(REPLACE(specjalizacjeWidok, '<table>', ''), '<br>', ''), '<br/>', '') AS bezTabeli, '" + nazwaSpeckajlizacji + "' as jednaSpecjalizacja, czyus, typ,rodzaj_zawieszenia, Informacje_o_wstrzymaniu  FROM tbl_osoby WHERE (typ = 1) ";
                        }
                    }
                    break;
            }

    
            Session["kwerenda"] = kwerendabazowa;

   
            Session["kwerenda"] = kwerendabazowa + " order by nazwisko" ;
            mediatorzy.SelectCommand = kwerendabazowa;
            mediatorzy.DataBind();
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
            excelTable.Columns.Add("tytul", typeof(string));
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
                int Ident = getId(item[2].ToString(), item[1].ToString());
                dr[0] = item[0].ToString();//tytul
                dr[1] = item[1].ToString();//imie
                dr[2] = item[2].ToString();//nazwisko
                try
                {
                    dr[3] = item[3].ToString().Substring(0, 11);//powolanie
                }
                catch
                {
                    dr[3] = "";

                }

                dr[4] = zaw;// item[5].ToString();
                dr[5] = item[6].ToString();//uwagi
                dr[6] = item[7].ToString();//uwagi
                dr[7] = GetSpec(Ident);
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
                catch
                {

                }
            }//end of using

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

        private int getId(string imie, string nazwisko)
        {
            DataTable parametry = Cm.makeParameterTable();
            parametry.Rows.Add("@imie", imie);
            parametry.Rows.Add("@nazwisko", nazwisko);
            string ident = Cm.runQuerryWithResult("SELECT ident   FROM [tbl_osoby] where imie=@imie and nazwisko=@nazwisko", Cm.con_str, parametry);
            int identInt = 0;
            try
            {
                identInt = int.Parse(ident);
            }
            catch  {}
            return identInt;

        }

        private string GetOpisSpecjalizacji(int ident)
        {
            DataTable parametry = Cm.makeParameterTable();

            parametry = Cm.makeParameterTable();
            parametry.Rows.Add("@ident", ident);

            DataTable specki = Cm.getDataTable("SELECT opis FROM tbl_specjalizacje_osob WHERE id_osoby = @ident", Cm.con_str, parametry);

            string SpeckiTxt = string.Empty;

            if (specki.Rows.Count > 0)
            {
                foreach (DataRow row in specki.Rows)
                {


                    SpeckiTxt += row[0].ToString() + Environment.NewLine;
                }
            }
            return SpeckiTxt;

        }

        private string GetSpec(int ident)
        {
            DataTable parametry = Cm.makeParameterTable();
            
            parametry = Cm.makeParameterTable();
            parametry.Rows.Add("@ident", ident);
            
            DataTable specki= Cm.getDataTable("SELECT DISTINCT glo_specjalizacje.nazwa FROM tbl_specjalizacje_osob INNER JOIN glo_specjalizacje ON tbl_specjalizacje_osob.id_specjalizacji = glo_specjalizacje.id_ WHERE     (tbl_specjalizacje_osob.id_osoby = @ident)", Cm.con_str, parametry);

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
 
        private IList<DoWydruku> ListaDoDalszejObrobki()
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

                string zaw = string.Empty ;
               
                int Ident = getId(item[2].ToString(), item[1].ToString());
                if (Ident==0)
                {
                    break;
                }
                doWydruku = new DoWydruku();
                doWydruku.ident = Ident;
                doWydruku.tytul = item[0].ToString();
                doWydruku.nazwisko = item[1].ToString();
                doWydruku.imie = item[2].ToString();
                doWydruku.powolanieOd = item[3].ToString();
                doWydruku.powolanieDo = item[4].ToString();
                
                doWydruku.telefon = item[6].ToString();
                doWydruku.uwagi = item[8].ToString();
                doWydruku.Kadencja = item[4].ToString();
                doWydruku.spejalizacje = GetOpisSpecjalizacji(Ident);
                doWydruku.spejalnosc = GetSpec(Ident);
                DataTable zawieszenia = DaneZawieszen(Ident);
                string PoczatekZawieszeni =string.Empty;
                string KoniecZawieszenia = string.Empty;    
                int iloscWierszy = 0;
                try
                {
                    iloscWierszy = zawieszenia.Rows.Count;
                    if (iloscWierszy > 0)
                    {
                        var czyZaw =(bool) zawieszenia.Rows[0][0];
                        if (czyZaw) 
                        {
                            zaw = "zawieszono";
                            PoczatekZawieszeni = zawieszenia.Rows[0][1].ToString();
                            KoniecZawieszenia = zawieszenia.Rows[0][2].ToString();

                        }
                    
                    }
                }
                catch 
                { }
                doWydruku.zawieszono = zaw;
                doWydruku.PoczatekZawieszeni = PoczatekZawieszeni;// GetPoczatekZawieszenia(Ident);
                doWydruku.KoniecZawieszenia = KoniecZawieszenia;//GetKoniecZawieszenia(Ident);
                doWydruku.uwagiBIP = GetUwagiBIP(Ident);

                doWydrukuLista.Add(doWydruku);
                
            }

            return doWydrukuLista;

        }

        private string GetUwagiBIP(int ident)
        {
            DataTable parametry = Cm.makeParameterTable();

            parametry = Cm.makeParameterTable();
            parametry.Rows.Add("@ident", ident);

            string uwagiBip = Cm.runQuerryWithResult("SELECT uwagiBIP FROM tbl_osoby WHERE ident = @ident", Cm.con_str, parametry);

            
            return uwagiBip;
        }

        private DataTable DaneZawieszen(int ident)
        {
            
                DataTable parametry = Cm.makeParameterTable();

                parametry = Cm.makeParameterTable();
                parametry.Rows.Add("@ident", ident);

                DataTable zawieszenia = Cm.getDataTable("SELECT czy_zaw, d_zawieszenia, dataKoncaZawieszenia  FROM tbl_osoby WHERE ident = @ident", Cm.con_str, parametry);

                
                return zawieszenia;

            


        }
      
        private void robRaportDoWydruku(IList<DoWydruku> ListaDoWydruku)
        {

            //nagłówek
            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10f, 10f, 10f, 0f);
            
            string path = Server.MapPath(@"~//pdf");
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
                robRaportjednejSpecjalizacji(DropDownList1.SelectedItem, getDataFromGridview());
            }
            else
            {
                robRaportWszystkichSpecjalizacjiNowy(getDataFromGridview());
            }
        }
       
       
        protected void makeExcell(object sender, EventArgs e)
        {
            ASPxGridViewExporter1.FileName = "Wykaz Biegłych";
            ASPxGridViewExporter1.WriteXlsToResponse();
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
        protected void _print(object sender, EventArgs e)
        {
            IList<DoWydruku> TaListaDoDalszejObrobki = ListaDoDalszejObrobki();
            robRaportDoWydruku(TaListaDoDalszejObrobki);
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

        protected void twórzZestawienieBIP(object sender, EventArgs e)
        {

            // do likwidacji
            IList<DoWydruku> TaListaDoDalszejObrobki = ListaDoDalszejObrobki();
            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4, 10f, 10f, 10f, 0f);


            string path = Server.MapPath(@"~//pdf"); //Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments );
            string fileName = path + "//Zestawienie_Specjalizacji_BIP_" + DateTime.Now.ToString().Replace(":", "-") + ".pdf";
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(fileName, FileMode.Create));

            pdfDoc.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());
            pdfDoc.Open();

            pdfDoc.AddTitle("zestawienie_Specjalizacji");
            pdfDoc.AddCreationDate();

            PdfPTable table = new PdfPTable(1);
            
            PdfPCell headerCell = new PdfPCell(new Phrase("LISTA BIEGŁYCH SĄDOWYCH PRZY SĄDZIE OKRĘGOWYM na dzień " + DateTime.Now.ToShortDateString(), cl.plFont3));

            headerCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            headerCell.Border = Rectangle.NO_BORDER;
            
            table.AddCell(headerCell);
            pdfDoc.Add(table);
            table = new PdfPTable(7);

            int[] tblWidth = { 20, 30, 10, 10,10,10,10 };


            table.SetWidths(tblWidth);


            table.AddCell(new PdfPCell(new Phrase("Specjalność", cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Specjalizacje", cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Nazwisko", cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Imię", cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Tytuł", cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Kadencja", cl.plFont2)));
            table.AddCell(new PdfPCell(new Phrase("Uwagi", cl.plFont2)));

            pdfDoc.Add(table);

            foreach (var item in TaListaDoDalszejObrobki)
            {
                table = new PdfPTable(7);
                table.SetWidths(tblWidth);
                table.AddCell(new PdfPCell(new Phrase(item.spejalnosc.ToString(), cl.plFont2)));
                table.AddCell(new PdfPCell(new Phrase(item.spejalizacje.ToString(), cl.plFont2)));
                table.AddCell(new PdfPCell(new Phrase(item.nazwisko.ToString(), cl.plFont2)));
                
                table.AddCell(new PdfPCell(new Phrase(item.imie.ToString(), cl.plFont2)));
                table.AddCell(new PdfPCell(new Phrase(item.tytul.ToString(), cl.plFont2)));
                string data = item.powolanieDo.ToString();
                try
                {
                    data = data.Substring(0, 10);
                }
                catch
                {
                }

                table.AddCell(new PdfPCell(new Phrase(data, cl.plFont2)));

                // sprawdzenie zawieszenia 
                var zawieszenie = item.zawieszono.ToString();
                if (zawieszenie == "zawieszono")
                {
                    string poleUwagi= string.Empty;
                    poleUwagi = item.uwagiBIP.ToString()+Environment.NewLine;
                    poleUwagi += "przerwa w opiniowaniu od " + item.PoczatekZawieszeni.Substring(0, 10) + " do " + item.KoniecZawieszenia.Substring(0, 10);
                    table.AddCell(new PdfPCell(new Phrase(poleUwagi, cl.plFont2)));

                }
                else
                {
                    table.AddCell(new PdfPCell(new Phrase(item.uwagiBIP.ToString(), cl.plFont2)));
                }
             //aaaa


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

        protected void tworzZestawienieBIPexcel(object sender, EventArgs e)
        {

            IList<DoWydruku> TaListaDoDalszejObrobki = ListaDoDalszejObrobki();
            IList<ListaIlosciSpecjalizacji> ListaIlosciSpecjalizacji = new List<ListaIlosciSpecjalizacji>();
            foreach (var item in TaListaDoDalszejObrobki)
            {
                int iloscSpecjalizacji = PodajIloscSpecjalizacji(item.ident);
                if (iloscSpecjalizacji > 0)
                {
                    ListaIlosciSpecjalizacji elementListyIlosciSpecjalizacji = new ListaIlosciSpecjalizacji
                    {
                        UserId = item.ident,
                        IloscSpecjalizacji = iloscSpecjalizacji
                    };

                    ListaIlosciSpecjalizacji.Add(elementListyIlosciSpecjalizacji);
                }
            }




            DataTable excelTableBIP = new DataTable();
            excelTableBIP.Columns.Add("specjalność", typeof(string));
            excelTableBIP.Columns.Add("specjalizacje", typeof(string));
            excelTableBIP.Columns.Add("nazwisko", typeof(string));
            excelTableBIP.Columns.Add("imie", typeof(string));
            excelTableBIP.Columns.Add("tytuł", typeof(string));
            excelTableBIP.Columns.Add("kadencka", typeof(string));
            excelTableBIP.Columns.Add("uwagi", typeof(string));

            int i = 0;

            
                foreach (var item in TaListaDoDalszejObrobki)
                {
                    try
                    {
                        var ilosc = ListaIlosciSpecjalizacji.First(x => x.UserId == item.ident).IloscSpecjalizacji;
                        if (ilosc > 1)
                        {
                            DataTable listaSpecjalizacjiBieglego = ListaIlosciSpecjalizacjiBieglego(item.ident);
                            foreach (DataRow dRow in listaSpecjalizacjiBieglego.Rows)
                            {
                                DataRow dr = excelTableBIP.NewRow();


                                dr[0] = dRow[0].ToString();//spejalnosc
                                dr[1] = dRow[1].ToString();//spejalizacje
                                dr[2] = item.nazwisko.ToString();//nazwisko
                                dr[3] = item.imie.ToString();//imie
                                dr[4] = item.tytul.ToString();//tytul

                                try
                                {
                                    dr[5] = item.powolanieDo.ToString().Substring(0, 11);//powolanie
                                }
                                catch
                                {
                                    dr[53] = "";

                                }


                                dr[6] = item.uwagi.ToString();//uwagi

                                excelTableBIP.Rows.Add(dr);






                            }




                        }
                        else
                        {
                            DataRow dr = excelTableBIP.NewRow();


                            dr[0] = item.spejalnosc.ToString();//spejalnosc
                            dr[1] = item.spejalizacje.ToString();//spejalizacje
                            dr[2] = item.nazwisko.ToString();//nazwisko
                            dr[3] = item.imie.ToString();//imie
                            dr[4] = item.tytul.ToString();//tytul

                            try
                            {
                                dr[5] = item.powolanieDo.ToString().Substring(0, 11);//powolanie
                            }
                            catch
                            {
                                dr[53] = "";

                            }


                            dr[6] = item.uwagi.ToString();//uwagi

                            excelTableBIP.Rows.Add(dr);
                        }
                    }
                    catch 
                    {

                       
                    }
                   
                }
            

                string tenPlikNazwa = "ZestawienieBIP";
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



                    MyWorksheet1 = tb.tworzArkuszwExcle(MyExcel.Workbook.Worksheets[1], excelTableBIP, 8, 0, 2, true, true, false, false, false);

                    try
                    {
                        MyExcel.SaveAs(fNewFile);

                        this.Response.Clear();
                        this.Response.ContentType = "application/vnd.ms-excel";
                        this.Response.AddHeader("Content-Disposition", "attachment;filename=" + fNewFile.Name);
                        this.Response.WriteFile(fNewFile.FullName);
                        this.Response.End();
                    }
                    catch
                    {

                    }
                }//end of using
           

           

        }

        private PdfPTable NaglowekZestawienia()
        {

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

            return fitst;



        }

        private PdfPTable tabelaWyliczenia(DataTable IlosciBieglychPoSpecjalizacji)
        {

            int strona = 1;
            PdfPTable tabelaWyliczenia = new PdfPTable(3);
            int[] tblWidthX = { 10, 70, 20 };

            PdfPCell cell = new PdfPCell(new Paragraph("", cl.plFontBIG));
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

            return tabelaWyliczenia;

        }

        private void robRaportWszystkichSpecjalizacjiNowy(DataTable dataTable)
        {

            string kwerenda = "SELECT View_SpecjalizacjeIOsoby.ident, tbl_osoby.imie, tbl_osoby.nazwisko, tbl_osoby.ulica, tbl_osoby.kod_poczt, tbl_osoby.miejscowosc, tbl_osoby.data_poczatkowa, tbl_osoby.data_koncowa, tbl_osoby.id_kreatora, tbl_osoby.data_kreacji, tbl_osoby.pesel, tbl_osoby.czyus, typ , tbl_osoby.tytul, tbl_osoby.czy_zaw, tbl_osoby.tel1 , tbl_osoby.tel2, tbl_osoby.email, tbl_osoby.adr_kores, tbl_osoby.kod_poczt_kor, tbl_osoby.miejscowosc_kor, tbl_osoby.uwagi, uwagiBIP,  tbl_osoby.specjalizacjeWidok, tbl_osoby.specjalizacja_opis,                   tbl_osoby.d_zawieszenia, tbl_osoby.typ, tbl_osoby.dataKoncaZawieszenia, tbl_osoby.instytucja, View_SpecjalizacjeIOsoby.nazwa, View_SpecjalizacjeIOsoby.id_ as identyfikatorSpecjalizacji,                   View_SpecjalizacjeIOsoby.Expr1 AS aktwnaSpecjalizacja FROM     tbl_osoby RIGHT OUTER JOIN                   View_SpecjalizacjeIOsoby ON tbl_osoby.ident = View_SpecjalizacjeIOsoby.ident WHERE (tbl_osoby.nazwisko IS NOT NULL) AND (tbl_osoby.typ < 2) AND (View_SpecjalizacjeIOsoby.Expr1 = 1)";
            DataTable daneBieglych = Cm.getDataTable(kwerenda, Cm.con_str, Cm.makeParameterTable());
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

            // koniec naglowka

            pdfDoc.Add(NaglowekZestawienia());
            pdfDoc.NewPage();

            //podliczenie

            pdfDoc.Add(tabelaWyliczenia(IlosciBieglychPoSpecjalizacji));
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

                            if (licznik == 13)
                            {
                                iloscStron++;
                                licznik = 0;


                                pdfDoc.Add(newTable);
                                pdfDoc.NewPage();
                                pdfDoc.Add(new Paragraph(" "));
                                pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, cl.plFont3)));
                                pdfDoc.Add(new Paragraph(" "));

                                newTable = new PdfPTable(4);
                                newTable.SetWidths(tblWidth);

                                newTable.Rows.Clear();
                                newTable.AddCell(new Paragraph("Lp.", cl.plFont2));
                                newTable.AddCell(new Paragraph("Nazwisko i imię", cl.plFont2));
                                newTable.AddCell(new Paragraph("Adres- telefon", cl.plFont2));
                                newTable.AddCell(new Paragraph("Zakres", cl.plFont2));
                                pdfDoc.Add(newTable);
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

            //podliczenie

            pdfDoc.NewPage();

            //==============================================================
            int iloscStron = 0;
            if (listaBieglych.Rows.Count > 0)
            {

                PdfPTable tabelaGlowna = new PdfPTable(4);
                tabelaGlowna = generujCzescRaportuOne(listaBieglych, nazwaSpecjalizacji, idSpecjalizacji);
                // tabelaGlowna = generujCzescRaportuNew(listaBieglych, idSpecjalizacji.ToString());


                pdfDoc.Add(new Paragraph(" "));
                pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, cl.plFont3)));
                pdfDoc.Add(new Paragraph(" "));
                int[] tblWidth = { 8, 30, 30, 32 };
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
                            pdfDoc.Add(new Paragraph(new Paragraph("        " + nazwaSpecjalizacji, cl.plFont3)));
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

        private DataTable getDataFromGridview()
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

        private string getInformacjeOwstrzymaniu(string idbieglego)
        {
            return cl.wyciagnijInformacjeOWsrzymaniuBieglego(idbieglego);
        }
        private PdfPTable generujCzescRaportuNew(DataTable biegli, string idSpecjalizacji)
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
    .Where(myRow => myRow.Field<int>("identyfikatorSpecjalizacji") == int.Parse(idSpecjalizacji)).ToArray();


            foreach (DataRow biegly in result)
            {
                iterator++;
                tabelaGlowna.AddCell(new Paragraph(iterator.ToString(), cl.plFont1));
                tabelaGlowna.AddCell(getBieglyInfo(biegly));
                tabelaGlowna.AddCell(getBieglyInfoAdres(biegly));
                tabelaGlowna.AddCell(new Paragraph(cl.PobierzOpisSpecjalizacji(biegly["ident"].ToString(), idSpecjalizacji), cl.plFont1));
            }

            return tabelaGlowna;
        }

        private PdfPTable generujCzescRaportuOne(DataTable biegli, string specjalizacje, int idSpecjalizacji)
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
            var result = biegli.AsEnumerable().ToArray();

            foreach (DataRow biegly in result)
            {
                if (biegly == null)
                {
                    continue;
                }
                DataTable daneBieglego = cl.wyciagnijBieglegoZSpecjalizacja(int.Parse(biegly[0].ToString()));
                if (daneBieglego.Rows.Count == 0)
                {
                    continue;
                }
                if (daneBieglego == null)
                {
                    continue;
                }
                iterator++;
                DataRow tenBiegly = daneBieglego.Rows[0];
                string Idbieglego = daneBieglego.Rows[0][0].ToString();
                tabelaGlowna.AddCell(new Paragraph(iterator.ToString(), cl.plFont1));
                tabelaGlowna.AddCell(getBieglyInfo(tenBiegly));
                tabelaGlowna.AddCell(getBieglyInfoAdres(tenBiegly));
                tabelaGlowna.AddCell(new Paragraph(cl.PobierzOpisSpecjalizacji(Idbieglego, idSpecjalizacji.ToString()), cl.plFont1));

            }

            return tabelaGlowna;
        }

        private Paragraph getBieglyInfo(DataRow biegly)
        {

            Paragraph result = new Paragraph();
            string Idbieglego = biegly["ident"].ToString();
            string imie = biegly["imie"].ToString();
            string nazwisko = biegly["nazwisko"].ToString();
            string tytul = biegly["tytul"].ToString();
            string czyZaw = biegly["czy_zaw"].ToString();
            string InformacjeOwstrzymaniu = getInformacjeOwstrzymaniu(Idbieglego);
            string dataKonca = getDateFromString(biegly, "data_koncowa");
            string poczatekWstrzymania = getDateFromString(biegly, "d_zawieszenia");
            string koniecWstrzymania = getDateFromString(biegly, "dataKoncaZawieszenia");

            string innerTable = string.Empty;
            innerTable = czyZaw == "True"
                ? imie + Environment.NewLine + nazwisko + Environment.NewLine + tytul + Environment.NewLine + "kadencja do dnia: " + dataKonca + Environment.NewLine + InformacjeOwstrzymaniu + Environment.NewLine + "Przerwa w opiniowaniu od " + poczatekWstrzymania + " do " + koniecWstrzymania
                : imie + Environment.NewLine + nazwisko + Environment.NewLine + tytul + Environment.NewLine + "kadencja do dnia: " + dataKonca + Environment.NewLine + InformacjeOwstrzymaniu;
            result = new Paragraph(innerTable, cl.plFont1);
            return result;

        }
        private Paragraph getBieglyInfoAdres(DataRow tenBiegly)
        {
            Paragraph result = new Paragraph();
            string Instytucja = tenBiegly["instytucja"].ToString();
            string email = tenBiegly["email"].ToString();
            string ulica = tenBiegly["ulica"].ToString();
            string kod = tenBiegly["kod_poczt"].ToString();
            string miejscowosc = tenBiegly["miejscowosc"].ToString();
            string tel = tenBiegly["tel1"].ToString();
            string adresTable = Instytucja + Environment.NewLine + ulica + Environment.NewLine + kod + " " + miejscowosc + Environment.NewLine + tel + Environment.NewLine + email;
            result = new Paragraph(adresTable, cl.plFont1);
            return result;
        }

        private DataTable ListaIlosciSpecjalizacjiBieglego(int ident)
        {
            return cl.ListaIlosciSpecjalizacjiBieglego(ident);
        }

        private int PodajIloscSpecjalizacji(int ident)
        {
            return  cl.PobierzIloscSpecjalizacji(ident);

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

        private string getDateFromString(DataRow biegly, string NazwaPola)
        {
            string result = string.Empty;
            try
            {
                DateTime dateTime = DateTime.Parse(biegly[NazwaPola].ToString());
                if (dateTime.Year==1900) 
                {
                    return string.Empty;
                }
                result = DateTime.Parse(biegly[NazwaPola].ToString()).ToShortDateString();
                
            }
            catch
            { }
            return result;
        }

    }

    public class tabele
    {

      

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
                                Arkusz.Cells[wiersz, i + przesunięcieX + dodatek + 1].Value = value;
                            }
                            catch
                            {
                                Arkusz.Cells[wiersz, i + przesunięcieX + dodatek].Value = (dR[i].ToString().Trim());
                            }
                            Arkusz.Cells[wiersz, przesunięcieX + dodatek + i + 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.Black);
                            Arkusz.Cells[wiersz, i + przesunięcieX + dodatek + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        }
                        catch
                        {

                        }
                    }

                    wiersz++;
                    dod = dodatek;
                }


            }
            catch
            {

            }

            return Arkusz;
        }

    }

    public class DoWydruku
    {
        public int ident { get; set; }
        public string tytul { get; set; }

        public string nazwisko { get; set; }
        public string imie { get; set; }
        public string powolanieOd { get; set; }
        public string powolanieDo { get; set; }
        public string zawieszono { get; set; }

        public string telefon { get; set; }

        public string uwagi { get; set; }

        public string uwagiBIP { get; set; }
        public string spejalizacje { get; set; }

        public string spejalnosc { get; set; }
        public string Kadencja { get; set; }
        public string PoczatekZawieszeni { get; set; }
        public string KoniecZawieszenia { get; set; }
    }

    public class ListaIlosciSpecjalizacji
    {
        public int UserId { get; set; }
        public int IloscSpecjalizacji { get; set; }
    }
}