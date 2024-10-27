using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web.UI.WebControls;

namespace wab2018
{
    public class Class2
    {
        public string con_str = ConfigurationManager.ConnectionStrings["wap"].ConnectionString;
        private cm Common = new cm();

        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static
            BaseFont NewFont = BaseFont.CreateFont(Environment.GetEnvironmentVariable("SystemRoot") + "\\fonts\\sylfaen.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);

        //  static BaseFont NewFont = BaseFont.CreateFont(BaseFont.HELVETICA,BaseFont.CP1257 , BaseFont.EMBEDDED);
        public Font fontPL1 = new Font(NewFont, 10f, Font.NORMAL, BaseColor.BLACK);

        public Font plFont1 = new Font(NewFont, 10f, Font.NORMAL, BaseColor.BLACK);
        public Font plFont = new Font(NewFont, 10f, Font.NORMAL, BaseColor.BLACK);
        public Font plFont2 = new Font(NewFont, 10f, Font.NORMAL, BaseColor.BLACK);
        public Font plFontBIG = new Font(NewFont, 15, Font.NORMAL, BaseColor.BLACK);
        public Font plFont3 = new Font(NewFont, 15, Font.NORMAL, BaseColor.BLACK);

        //==========================================================================================================================================
        //==========================================================================================================================================
        //==========================================================================================================================================
        public string podajNumerSkargiwRoku(string idOsoby, int rok)
        {
            log.Info("Start funkcji usun_Osobe");
            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@rok", rok);
            parameters.Rows.Add("@idBieglego", idOsoby);
            string odp = Common.runQuerryWithResult("select (max(numer)+1) from tbl_skargi where rok=@rok and idBieglego=@idBieglego and czyus=0", con_str, parameters);
            if (string.IsNullOrEmpty(odp))
            {
                odp = "1";
            }
            return odp;
        }// end of podajNumerSkargiwRoku

        public string podajNumerSkargiwRoku(int rok)
        {
            log.Info("Start funkcji usun_Osobe");
            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@rok", rok);

            string odp = Common.runQuerryWithResult("select (max(numer)+1) from tbl_skargi where rok=@rok  and czyus=0", con_str, parameters);
            if (string.IsNullOrEmpty(odp))
            {
                odp = "1";
            }
            return odp;
        }// end of podajNumerSkargiwRoku

        public string usun_osobe(string id, string deleter)
        {
            log.Info("Start funkcji usun_Osobe");
            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@id_", id);
            parameters.Rows.Add("@usuwajacy", deleter);
            Common.runQuerry("UPDATE tbl_osoby SET czyus =1, d_usuniecia =getdate(), id_usuwajacego =@usuwajacy,   przyczyna_usuniecia =''  where ident=@id_", con_str, parameters);
            return null;
        }// end ofusun_osobe

        public string usun_specjalizacje(int id)
        {
            log.Info("Start funkcji usun_specjalizacje");
            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@id_", id);
            Common.runQuerry("delete from tbl_sprcjalizacje_temp where id_=@id_", con_str, parameters);
            return null;
        }// end of usun_specjalizacje

        public string usun_specjalizacje_osoby(int id)
        {
            log.Info("Start funkcji usun_specjalizacje_osoby");

            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@id_", id);
            Common.runQuerry("delete from  tbl_specjalizacje_osob  where id_osoby=@id_", con_str, parameters);
            return null;
        }// end of usun_specjalizacje_osoby

        public string usun_specjalizacje_tymczasowe(string sesja)
        {
            log.Info("Start funkcji usun_specjalizacje_tymczasowe");

            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@id_", sesja);
            Common.runQuerry("delete from  tbl_sprcjalizacje_temp  where nr_sesji=@id_", con_str, parameters);
            return null;
        }// end of usun_specjalizacje_tymczasowe

        public string dodaj_specjalizacje(string sesja, int id)
        {
            log.Info("Start funkcji dodaj_specjalizacje");

            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@nr_sesji", sesja);
            parameters.Rows.Add("@id_spec", id);
            Common.runQuerry("insert into  tbl_sprcjalizacje_temp(nr_sesji, id_spec)  values(@nr_sesji, @id_spec) ", con_str, parameters);
            return null;
        }// end of dodaj_specjalizacje

        public string usunSpecjalizacjetymczasowe()
        {
            log.Info("Start funkcji dodaj_specjalizacje");

            DataTable parameters = Common.makeParameterTable();

            Common.runQuerry("delete from  tbl_sprcjalizacje_temp", con_str, parameters);
            return "0";
        }// end of dodaj_specjalizacje

        public string dodaj_Grupe_specjalizacj1(string nazwa)
        {
            log.Info("Start funkcji dodaj_Grupe_specjalizacj1");

            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@nazwa", nazwa);
            Common.runQuerry("insert into glo_grupy_specjalizacji (nazwa)  values (@nazwa)", con_str, parameters);
            return "0";
        }// end of dodaj_specjalizacje

        public string ile__Grup_specjalizacji(string id_)
        {
            var conn = new SqlConnection(con_str);
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("SELECT COUNT(*)  FROM glo_specjalizacje where grupa=@grupa", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@grupa", id_);
                    string odp = sqlCmd.ExecuteScalar().ToString().Trim();
                    conn.Close();
                    return odp;
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_specjalizacje

        public string usun__Grup_specjalizacji(string id_)
        {
            var conn = new SqlConnection(con_str);
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("delete from  glo_grupy_specjalizacji where ident=@id_", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@id_", id_);
                    sqlCmd.ExecuteScalar();
                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_specjalizacje

        public string zmien_Grupe_specjalizacji(int id_, string nazwa)
        {
            var conn = new SqlConnection(con_str);
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("update  glo_grupy_specjalizacji set nazwa= @nazwa where ident=@id_ ", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@nazwa", nazwa);
                    sqlCmd.Parameters.AddWithValue("@id_", id_);
                    sqlCmd.ExecuteScalar();
                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_specjalizacje

        public string dodaj_specjalizacje_osoby(string specjalizacja, string id)
        {
            var conn = new SqlConnection(con_str);
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("INSERT     INTO  tbl_specjalizacje_osob (id_osoby, id_specjalizacji) VALUES (@id_osoby, @id_specjalizacji)", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@id_osoby", id);
                    sqlCmd.Parameters.AddWithValue("@id_specjalizacji", specjalizacja);
                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_specjalizacje

        public string sprawdz_ilosc_specjalizacji(string specjalizacja)
        {
            var conn = new SqlConnection(con_str);
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("select count (*) from   tbl_specjalizacje_osob where id_specjalizacji= @id_specjalizacji", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@id_specjalizacji", specjalizacja);
                    string odp = sqlCmd.ExecuteScalar().ToString();
                    conn.Close();
                    return odp;
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_specjalizacje

        public string usunPowolanie(int id_powolania, string user)
        {
            var conn = new SqlConnection(con_str);
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("update tbl_powolania set czyus=1,modyfikator=@modyfikator, d_modyfikacji=Getdate() where ident=@id", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@id", id_powolania);
                    sqlCmd.Parameters.AddWithValue("@modyfikator", user);
                    sqlCmd.ExecuteScalar();
                    conn.Close();
                    return "1";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_specjalizacje

        public DataTable odczytaj_specjalizacjeLista()
        {
            SqlConnection conn = new SqlConnection(con_str);
            DataSet lista = new DataSet();
            DataTable specjalizacje = new DataTable();
            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT DISTINCT  id_, nazwa FROM         glo_specjalizacje ORDER BY nazwa", conn);

                daMenu.Fill(lista);
                conn.Close();
                specjalizacje = lista.Tables[0];
            }
            catch 
            {
                conn.Close();
            }
            return specjalizacje;
        }// end of dodaj_specjalizacje

        public string odczytaj_specjalizacje_osobyOpis(string id)
        {
            string result = string.Empty;
            SqlCommand sqlCmd;
            SqlConnection conn = new SqlConnection(con_str);
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("SELECT     specjalizacja_opis FROM         [tbl_osoby] WHERE     (ident = @id_) ORDER BY nazwisko", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@id_", id);

                    result = sqlCmd.ExecuteScalar().ToString();
                    conn.Close();
                }
                catch (Exception ex)
                {
                    conn.Close();
                    return ex.Message;
                }
            }
            return result;
        }// end of dodaj_specjalizacje
        public string PobierzOpisSpecjalizacji(string UserId, string IdSpecjalizacji)
        {
            string result = string.Empty;
            SqlCommand sqlCmd;
            SqlConnection conn = new SqlConnection(con_str);
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("SELECT distinct [opis]  FROM [tbl_specjalizacje_osob] where id_osoby=@idOsoby and id_specjalizacji=@idSpecjalizacji", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@idOsoby", UserId);
                    sqlCmd.Parameters.AddWithValue("@idSpecjalizacji", IdSpecjalizacji);

                    result = sqlCmd.ExecuteScalar().ToString();
                    conn.Close();
                }
                catch (Exception ex)
                {
                    conn.Close();
                    return ex.Message;
                }
            }
            return result;
        }// end of dodaj_specjalizacje

        public int PobierzIloscSpecjalizacji(int UserId)
        {
            int result = 0;
            SqlCommand sqlCmd;
            SqlConnection conn = new SqlConnection(con_str);
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("SELECT count(*)   FROM [tbl_specjalizacje_osob] where id_osoby=@idOsoby ", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@idOsoby", UserId);


                    result = int.Parse(sqlCmd.ExecuteScalar().ToString());
                    conn.Close();
                }
                catch
                { }
            }
            return result;
        }// end of dodaj_specjalizacje



        public DataTable ListaIlosciSpecjalizacjiBieglego(int UserId)
        {

            SqlConnection conn = new SqlConnection(con_str);
            DataSet lista = new DataSet();
            try
            {
                conn.Open();
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter();
                sqlDataAdapter.SelectCommand = new SqlCommand("SELECT glo_specjalizacje.nazwa, tbl_specjalizacje_osob.opis, tbl_specjalizacje_osob.id_osoby, tbl_specjalizacje_osob.id_specjalizacji FROM tbl_specjalizacje_osob LEFT OUTER JOIN  glo_specjalizacje ON tbl_specjalizacje_osob.id_specjalizacji = glo_specjalizacje.id_ WHERE        (tbl_specjalizacje_osob.id_osoby = @ident) ", conn);
                sqlDataAdapter.SelectCommand.Parameters.AddWithValue("@ident", UserId);
                sqlDataAdapter.Fill(lista);
                conn.Close();
            }
            catch
            { }
            return lista.Tables[0];


        }// end of ListaIlosciSpecjalizacjiBieglego
        public string odczytaj_specjalizacje_osobyOpis(string id,string nazwa)
        {
            string result = string.Empty;
            SqlCommand sqlCmd;
            SqlConnection conn = new SqlConnection(con_str);
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("SELECT tbl_specjalizacje_osob.opis FROM  tbl_specjalizacje_osob LEFT OUTER JOIN glo_specjalizacje ON tbl_specjalizacje_osob.id_specjalizacji = glo_specjalizacje.id_ WHERE(tbl_specjalizacje_osob.opis IS NOT NULL) AND(glo_specjalizacje.nazwa = @nazwa) AND(tbl_specjalizacje_osob.id_osoby = @id_osoby)", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@id_osoby", id);
                    sqlCmd.Parameters.AddWithValue("@nazwa", nazwa);

                    result = sqlCmd.ExecuteScalar().ToString();
                    conn.Close();
                }
                catch
                {
                    conn.Close();
                   
                }
            }
            return result;
        }// end of dodaj_specjalizacje

        public string odczytaj_specjalizacje_osoby(string id, string sesja)
        {
            SqlConnection conn = new SqlConnection(con_str);
            DataSet lista = new DataSet();
            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT     id_specjalizacji FROM         tbl_specjalizacje_osob WHERE     (id_osoby = @id_) ORDER BY id_", conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@id_", id);
                daMenu.Fill(lista);
                // zapis

                int il = 0;
                try
                {
                    il = lista.Tables[0].Rows.Count;
                    if (il >= 0)
                    {
                        foreach (DataRow dr in lista.Tables[0].Rows)
                        {
                            string val_ = dr[0].ToString();
                            dodaj_specjalizacje(sesja, int.Parse(val_));
                        }
                    }
                }
                catch
                { }
                conn.Close();
            }
            catch 
            { }
            return "lista";
        }// end of dodaj_specjalizacje
      
        public DataTable odczytaj_specjalizacje_tymczasowe_osoby_lista(string id, string sesja)
        {
            SqlConnection conn = new SqlConnection(con_str);
            DataSet lista = new DataSet();
            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT     id_spec FROM          tbl_sprcjalizacje_temp WHERE     nr_sesji=@sesja ORDER BY id_", conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@sesja", sesja);
                daMenu.Fill(lista);
                // zapis

                conn.Close();
            }
            catch 
            { }
            return lista.Tables[0];
        }// end of dodaj_specjalizacje

        public DataTable wyciagnijBieglychZSpecjalizacja(string idSpecjalizacji, bool archiwum)
        {
            DataTable dT = new DataTable();
            SqlConnection conn = new SqlConnection(con_str);
            DataSet lista = new DataSet();
            string kwerenda = string.Empty;
            if (archiwum)
            {
                kwerenda = "SELECT DISTINCT dbo.tbl_osoby.ident, dbo.tbl_osoby.imie, dbo.tbl_osoby.nazwisko, dbo.tbl_osoby.ulica, dbo.tbl_osoby.kod_poczt, dbo.tbl_osoby.miejscowosc,   dbo.tbl_osoby.data_koncowa,  dbo.tbl_osoby.tytul, dbo.tbl_osoby.tel1 FROM  dbo.tbl_specjalizacje_osob LEFT OUTER JOIN                          dbo.tbl_osoby ON dbo.tbl_specjalizacje_osob.id_osoby = dbo.tbl_osoby.ident WHERE    (dbo.tbl_osoby.typ=1) and (dbo.tbl_osoby.czyus=0) and   dbo.tbl_osoby.data_koncowa<getdate() and  (dbo.tbl_specjalizacje_osob.id_specjalizacji = @idSpecjalizacji)";
            }
            else
            {
                kwerenda = "SELECT DISTINCT dbo.tbl_osoby.ident, dbo.tbl_osoby.imie, dbo.tbl_osoby.nazwisko, dbo.tbl_osoby.ulica, dbo.tbl_osoby.kod_poczt, dbo.tbl_osoby.miejscowosc,   dbo.tbl_osoby.data_koncowa,  dbo.tbl_osoby.tytul, dbo.tbl_osoby.tel1 FROM  dbo.tbl_specjalizacje_osob LEFT OUTER JOIN                          dbo.tbl_osoby ON dbo.tbl_specjalizacje_osob.id_osoby = dbo.tbl_osoby.ident WHERE  (dbo.tbl_osoby.typ=1) and  (dbo.tbl_osoby.czyus=0) and   dbo.tbl_osoby.data_koncowa>getdate() and  (dbo.tbl_specjalizacje_osob.id_specjalizacji = @idSpecjalizacji)";
            }
            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand(kwerenda, conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@idSpecjalizacji", idSpecjalizacji);
                daMenu.Fill(lista);
                // zapis
                dT = lista.Tables[0];
                conn.Close();
            }
            catch 
            { }
            return dT;
        }// end of wyciagnijBieglychZSpecjalizacja
      
        public string wyciagnijInformacjeOWsrzymaniuBieglego(string idBieglego)
        {
           
           string info=string.Empty;
            
            string kwerenda = string.Empty;

            kwerenda = "SELECT DISTINCT [Informacje_o_wstrzymaniu] FROM  [tbl_osoby] WHERE  [ident]=@ident " ;
            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@ident", idBieglego);
            info =  Common.runQuerryWithResult(kwerenda, con_str, parameters);
      

            return info;
        }// end of wyciagnijInformacjeOWsrzymaniuBieglego

        public DataTable wyciagnijBieglegoZSpecjalizacja(int idBieglego)
        {
            DataTable dT = new DataTable();
            SqlConnection conn = new SqlConnection(con_str);
            string querryExtention = " tbl_osoby.ident =" + idBieglego.ToString();
            string kwerenda = string.Empty;

            kwerenda = "SELECT DISTINCT tbl_osoby.ident, tbl_osoby.imie, tbl_osoby.nazwisko, tbl_osoby.ulica, tbl_osoby.kod_poczt, tbl_osoby.miejscowosc,  tbl_osoby.data_koncowa,  tbl_osoby.tytul, tbl_osoby.tel1, tbl_osoby.email ,tbl_osoby.specjalizacja_opis  ,tbl_osoby.instytucja ,tbl_osoby.instytucja ,tbl_osoby.instytucja ,tbl_osoby.Informacje_o_wstrzymaniu,  tbl_osoby.czy_zaw, tbl_osoby.d_zawieszenia, tbl_osoby.dataKoncaZawieszenia  FROM  tbl_specjalizacje_osob LEFT OUTER JOIN tbl_osoby ON tbl_specjalizacje_osob.id_osoby = tbl_osoby.ident WHERE   " + querryExtention;
            try
            {
                DataSet lista = new DataSet();
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand(kwerenda, conn);
                //  daMenu.SelectCommand.Parameters.AddWithValue("@ident", ident);
                daMenu.Fill(lista);
                // zapis

                dT = lista.Tables[0];
                conn.Close();
            }
            catch
            { }

            return dT;
        }// end of wyciagnijBieglychZSpecjalizacja
        /*
        public DataTable wyciagnijMediatorowZSpecjalizacja(string idSpecjalizacji, bool archiwum)
        {
            DataTable dT = new DataTable();
            SqlConnection conn = new SqlConnection(con_str);
            DataSet lista = new DataSet();
            string kwerenda = string.Empty;
            if (archiwum)
            {
                kwerenda = "SELECT DISTINCT dbo.tbl_osoby.ident, dbo.tbl_osoby.imie, dbo.tbl_osoby.nazwisko, dbo.tbl_osoby.ulica, dbo.tbl_osoby.kod_poczt, dbo.tbl_osoby.miejscowosc,   dbo.tbl_osoby.data_koncowa,  dbo.tbl_osoby.tytul, dbo.tbl_osoby.tel1 FROM  dbo.tbl_specjalizacje_osob ,dbo.tbl_osoby.instytucja  LEFT OUTER JOIN                          dbo.tbl_osoby ON dbo.tbl_specjalizacje_osob.id_osoby = dbo.tbl_osoby.ident WHERE    (dbo.tbl_osoby.typ=2) and (dbo.tbl_osoby.czyus=0) and   dbo.tbl_osoby.data_koncowa<getdate() and  (dbo.tbl_specjalizacje_osob.id_specjalizacji = @idSpecjalizacji)";
            }
            else
            {
                kwerenda = "SELECT DISTINCT dbo.tbl_osoby.ident, dbo.tbl_osoby.imie, dbo.tbl_osoby.nazwisko, dbo.tbl_osoby.ulica, dbo.tbl_osoby.kod_poczt, dbo.tbl_osoby.miejscowosc,   dbo.tbl_osoby.data_koncowa,  dbo.tbl_osoby.tytul, dbo.tbl_osoby.tel1 FROM  dbo.tbl_specjalizacje_osob ,dbo.tbl_osoby.instytucja  LEFT OUTER JOIN                          dbo.tbl_osoby ON dbo.tbl_specjalizacje_osob.id_osoby = dbo.tbl_osoby.ident WHERE  (dbo.tbl_osoby.typ=2) and  (dbo.tbl_osoby.czyus=0) and   dbo.tbl_osoby.data_koncowa>getdate() and  (dbo.tbl_specjalizacje_osob.id_specjalizacji = @idSpecjalizacji)";
            }
            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand(kwerenda, conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@idSpecjalizacji", idSpecjalizacji);
                daMenu.Fill(lista);
                // zapis
                dT = lista.Tables[0];
                conn.Close();
            }
            catch (Exception ec)
            { }
            return dT;
        }// end of dodaj_specjalizacje

        public DataTable wyciagnijDaneBieglych(string kwerenda)
        {
            DataTable dT = new DataTable();
            SqlConnection conn = new SqlConnection(con_str);
            DataSet lista = new DataSet();
            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand(kwerenda, conn);

                daMenu.Fill(lista);
                // zapis
                dT = lista.Tables[0];
                conn.Close();
            }
            catch (Exception ec)
            { }
            return dT;
        }// end of dodaj_specjalizacje
        */
        public DataTable odczytaj_specjalizacje_osoby2(string id)
        {
            DataTable dT = new DataTable();
            SqlConnection conn = new SqlConnection(con_str);
            DataSet lista = new DataSet();
            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT [nazwa]  FROM [View_pozycje_specjalizacji] where   (id_osoby = @id_) ORDER BY id_", conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@id_", id);
                daMenu.Fill(lista);
                // zapis
                dT = lista.Tables[0];
                conn.Close();
            }
            catch 
            { }
            return dT;
        }// end of dodaj_specjalizacje

        public int czyJestSpecjalizacjauBieglego(int idBieglego, int idSpecjalizacji)
        {
            return int.Parse(Common.runQuerryWithResult("SELECT count(*) FROM View_pozycje_specjalizacji where   (id_osoby = " + idBieglego.ToString() + ") and id_specjalizacji =" + idSpecjalizacji.ToString(), Common.con_str));
        }

       
        public string PobierzDane(string klucz)
        {
            log.Info("Start funkcji PobierzDane");

            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@klucz", klucz);
            return Common.runQuerryWithResult("SELECT        distinct  wartosc FROM            tbl_konfiguracja where klucz = @klucz", con_str, parameters);
        }
     
        public string modyfikuj_osobe(string id_, int id_modyfikatora, string imie, string nazwisko, string ulica, string kod_poczt, string miejscowosc, DateTime data_poczatkowa, DateTime data_koncowa, string tytul, string pesel, string email, string uwagi, int czy_zaw, string adr_kor, string kod_kor, string miejsc_kor, string tel1, string tel2, DateTime dataPoczatkuZawieszenia, DateTime dataKoncaZawieszenia, string specjalizacja_opis, int typ)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("INSERT INTO tbl_modyfikacje (id_osoby, imie_org, nazwisko_org, adres_org, kod_poczt_org, miejscowosc_org, data_pocz_org, data_konc_org) SELECT     ident, imie, nazwisko, ulica, kod_poczt, miejscowosc, data_poczatkowa, data_koncowa FROM         tbl_osoby WHERE     (ident = @id) select @@identity", conn);
                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@id", id_);
                    string odp = sqlCmd.ExecuteScalar().ToString();
                    // nowe wartosci tabela modyfikacj
                    conn.Close();
                    conn.Open();
                    sqlCmd = new SqlCommand("update tbl_modyfikacje  set  id_modyfikujacego=@id_mod, data_modyfikacji=getdate(),imie_mod=@imie, nazwisko_mod=@nazwisko, adres_mod=@adres ,kod_poczt_mod=@kod_poczt, miejscowosc_mod=@miejscowosc,data_pocz_mod=@d_zawieszeniaPocz,dataKoncaZawieszenia =@dataKoncaZawieszenia ,data_konc_mod=@data_konc where id_=@id", conn);
                    sqlCmd.Parameters.AddWithValue("@imie", imie);
                    sqlCmd.Parameters.AddWithValue("@nazwisko", nazwisko);
                    sqlCmd.Parameters.AddWithValue("@adres", ulica);
                    sqlCmd.Parameters.AddWithValue("@kod_poczt", kod_poczt);
                    sqlCmd.Parameters.AddWithValue("@miejscowosc", miejscowosc);
                    sqlCmd.Parameters.AddWithValue("@data_pocz", data_poczatkowa);
                    sqlCmd.Parameters.AddWithValue("@data_konc", data_koncowa);
                    sqlCmd.Parameters.AddWithValue("@id_mod", id_modyfikatora);
                    sqlCmd.Parameters.AddWithValue("@d_zawieszeniaPocz", dataPoczatkuZawieszenia);
                    sqlCmd.Parameters.AddWithValue("@dataKoncaZawieszenia", dataKoncaZawieszenia);

                    //dataKoncaZawieszenia
                    sqlCmd.Parameters.AddWithValue("@id", odp);
                    sqlCmd.Parameters.AddWithValue("@typ", typ);
                    sqlCmd.ExecuteScalar();
                    conn.Close();
                    conn.Open();
                    // nowe wartosci tabela osoby
                    sqlCmd = new SqlCommand("update tbl_osoby set imie=@imie, nazwisko=@nazwisko, ulica=@adres ,kod_poczt=@kod_poczt, miejscowosc=@miejscowosc,data_poczatkowa=@data_pocz,data_koncowa=@data_konc,tytul=@tytul, pesel=@pesel, czy_zaw=@czy_zaw, tel1=@tel1, tel2=@tel2,email=@email,adr_kores=@adr_kores, kod_poczt_kor=@kod_kores, miejscowosc_kor=@miejscowosc_kores, uwagi=@uwagi, d_zawieszenia=@d_zawieszenia, dataKoncaZawieszenia =@dataKoncaZawieszenia, specjalizacja_opis=@specjalizacja_opis where ident=@id", conn);
                    sqlCmd.Parameters.AddWithValue("@imie", imie);
                    sqlCmd.Parameters.AddWithValue("@nazwisko", nazwisko);
                    sqlCmd.Parameters.AddWithValue("@adres", ulica);
                    sqlCmd.Parameters.AddWithValue("@kod_poczt", kod_poczt);
                    sqlCmd.Parameters.AddWithValue("@miejscowosc", miejscowosc);
                    sqlCmd.Parameters.AddWithValue("@data_pocz", data_poczatkowa);
                    sqlCmd.Parameters.AddWithValue("@data_konc", data_koncowa);
                    sqlCmd.Parameters.AddWithValue("@id_mod", id_modyfikatora);
                    sqlCmd.Parameters.AddWithValue("@tytul", tytul);
                    sqlCmd.Parameters.AddWithValue("@id", id_);
                    sqlCmd.Parameters.AddWithValue("@pesel", pesel);
                    //nowe
                    sqlCmd.Parameters.AddWithValue("@email", email);
                    sqlCmd.Parameters.AddWithValue("@adr_kores", adr_kor);
                    sqlCmd.Parameters.AddWithValue("@kod_kores", kod_kor);
                    sqlCmd.Parameters.AddWithValue("@miejscowosc_kores", miejsc_kor);
                    sqlCmd.Parameters.AddWithValue("@Uwagi", uwagi);
                    sqlCmd.Parameters.AddWithValue("@czy_zaw", czy_zaw);
                    sqlCmd.Parameters.AddWithValue("@tel1", tel1);
                    sqlCmd.Parameters.AddWithValue("@tel2", tel2);
                    sqlCmd.Parameters.AddWithValue("@d_zawieszenia", dataPoczatkuZawieszenia);
                    sqlCmd.Parameters.AddWithValue("@dataKoncaZawieszenia", dataKoncaZawieszenia);

                    //[specjalizacja_opis]
                    sqlCmd.Parameters.AddWithValue("@specjalizacja_opis", specjalizacja_opis);
                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    conn.Close();
                    return ex.Message;
                }
            }
        }// end of modyfikuj_osobe

        public string usun_obciazenie_osoby(string id_, string id_kreatora)
        {
            var conn = new SqlConnection(con_str);
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("delete from   tbl_obciazenia where id_=@id_", conn);

                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@id_", id_);
                    sqlCmd.ExecuteScalar();
                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_obciazenie_osoby

        public string update_specjalizacji(string id_, string nazwa, int grupa)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            // wyciągnij tabele specjalizacji
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("update glo_specjalizacje set nazwa=@nazwa, grupa=@grupa where id_=@id", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@id", id_);
                    sqlCmd.Parameters.AddWithValue("@nazwa", nazwa);
                    sqlCmd.Parameters.AddWithValue("@grupa", grupa);
                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of update_specjalizacji

        public string update_specjalizacjiWidoku(string id_)
        {
           

            SqlConnection conn = new SqlConnection(con_str);

            DataSet specjalizacje = new DataSet();

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT        dbo.glo_specjalizacje.nazwa FROM            dbo.tbl_specjalizacje_osob LEFT OUTER JOIN     dbo.glo_specjalizacje ON dbo.tbl_specjalizacje_osob.id_specjalizacji = dbo.glo_specjalizacje.id_  WHERE        (dbo.tbl_specjalizacje_osob.id_osoby = @idOsoby)", conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@idOsoby", id_);

                daMenu.Fill(specjalizacje);
                conn.Close();
            }
            catch 
            {
            }

            string lista = string.Empty;
            try
            {
                DataTable spec = specjalizacje.Tables[0];

                foreach (DataRow dr in spec.Rows)
                {
                    lista = lista + dr[0].ToString() + "; ";
                }
            }
            catch (Exception)
            {
            }

            SqlCommand sqlCmd;
            // wyciągnij tabele specjalizacji
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand(" update tbl_osoby set specjalizacjeWidok =@lista where ident=@idOsoby", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@idOsoby", id_);
                    sqlCmd.Parameters.AddWithValue("@lista", lista);
                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of update_specjalizacji

        public string usun_specjalizacje(string id_)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("delete from glo_specjalizacje where  id_=@id", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@id", id_);

                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of update_specjalizacji

        public string dodaj_specyfikacje(string nazwa, int grupa)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("insert into glo_specjalizacje (nazwa,grupa) values (@nazwa,@grupa)", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@nazwa", nazwa);
                    sqlCmd.Parameters.AddWithValue("@grupa", grupa);
                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_specyfikacje

        public string usun_specyfikacjeTmp(string id_)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("delete from  tbl_sprcjalizacje_temp where id_spec =@id", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@id", id_);
                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_specyfikacje

        public string dodaj_uzytkownika(string nazwa, string haslo, string imie, string nazwisko, int rola)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("INSERT    INTO      tbl_users_(user_, pasword, rola, imie, nazwisko) VALUES     ( @user_, @pasword, @rola, @imie, @nazwisko)", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@user_", nazwa);
                    sqlCmd.Parameters.AddWithValue("@pasword", haslo);
                    sqlCmd.Parameters.AddWithValue("@rola", rola);
                    sqlCmd.Parameters.AddWithValue("@imie", imie);
                    sqlCmd.Parameters.AddWithValue("@nazwisko", nazwisko);

                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_uzytkownika

        public string update_uzytkownika(string nazwa, string haslo, string imie, string nazwisko, int rola, int id_)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("UPDATE    tbl_users_ SET   user_ = @user_, pasword = @pasword, rola = @rola, imie = @imie, nazwisko = @nazwisko where id_=@id_", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@user_", nazwa);
                    sqlCmd.Parameters.AddWithValue("@pasword", haslo);
                    sqlCmd.Parameters.AddWithValue("@rola", rola);
                    sqlCmd.Parameters.AddWithValue("@imie", imie);
                    sqlCmd.Parameters.AddWithValue("@nazwisko", nazwisko);
                    sqlCmd.Parameters.AddWithValue("@id_", id_);
                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of update_uzytkownika

        public string usun_uzytkownika(int id_)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("delete from tbl_users_ where id_=@id_", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@id_", id_);
                    sqlCmd.ExecuteScalar();
                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of update_uzytkownika

        public string sprawdz_polaczenie(string db_name, string db_catalog, string user_, string password)
        {
            string result = string.Empty; string connection_string = "Data Source=" + db_name + ";Initial Catalog=" + db_catalog + ";Persist Security Info=True;User ID=" + user_ + ";Password=" + password;
            var conn = new SqlConnection(connection_string);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                try
                {
                    conn.Open();
                    result = "Uzyskano połaczenie z bazą danych!";
                    conn.Close();
                }
                catch (Exception ex)
                {
                    result = "Błąd połaczenia z bazą danych! Komunikat błędu: " + ex.Message;
                }
            }
            return result;
        }// end of dodaj_baze_danych

        public string dodaj_baze_danych(string nazwa, string db_name, string db_catalog, string user_, string password, string id_kwerendy)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("INSERT    INTO tbl_bazy_danych (nazwa, db_name, db_user, db_paswd, db_catalog,id_kwerendy) VALUES    (@nazwa, @db_name, @db_user, @db_paswd, @db_catalog,@id_kwerendy) ", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@nazwa", nazwa);
                    sqlCmd.Parameters.AddWithValue("@db_name", db_name);
                    sqlCmd.Parameters.AddWithValue("@db_user", user_);
                    sqlCmd.Parameters.AddWithValue("@db_paswd", password);
                    sqlCmd.Parameters.AddWithValue("@db_catalog", db_catalog);
                    sqlCmd.Parameters.AddWithValue("@id_kwerendy", id_kwerendy);
                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_baze_danych

        public string update_bazy_danych(string nazwa, string db_name, string db_catalog, string user_, string password, string id_kwerendy, string id_)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("UPDATE    tbl_bazy_danych  SET  nazwa = @nazwa, db_name = @db_name, db_user =@db_user, db_paswd = @db_paswd, db_catalog = @db_catalog,id_kwerendy=@id_kwerendy where id_ =@id_", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@nazwa", nazwa);
                    sqlCmd.Parameters.AddWithValue("@db_name", db_name);
                    sqlCmd.Parameters.AddWithValue("@db_user", user_);
                    sqlCmd.Parameters.AddWithValue("@db_paswd", password);
                    sqlCmd.Parameters.AddWithValue("@db_catalog", db_catalog);
                    sqlCmd.Parameters.AddWithValue("@id_kwerendy", id_kwerendy);
                    sqlCmd.Parameters.AddWithValue("@id_", id_);
                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_baze_danych

        public string usun_baze_danych(string id_)
        {
            var conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("delete from    tbl_bazy_danych where id_=@id_", conn);

                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@id_", id_);
                    sqlCmd.ExecuteScalar();
                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_baze_danych

        public string dodaj_obciazenie_osoby(string id_osoby, string instytucja, string sygnatura, DateTime data_wprowadzenia, DateTime data_otrzymania, DateTime termin, string id_kreatora, DateTime data_zwrotu)
        {
            var conn = new SqlConnection(con_str);
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("INSERT  INTO  tbl_obciazenia(id_osoby, instytucja, sygnatura, data_wprowadzenia, data_otrzymania, termin, data_zapisu, id_kreatora,data_zwrotu)VALUES     (@id_osoby, @instytucja, @sygnatura, @data_wprowadzenia, @data_otrzymania, @termin,getdate(), @id_kreatora,@data_zwrotu)", conn);

                try
                {
                    conn.Open();
                    sqlCmd.Parameters.AddWithValue("@id_osoby", id_osoby);
                    sqlCmd.Parameters.AddWithValue("@instytucja", instytucja);
                    sqlCmd.Parameters.AddWithValue("@sygnatura", sygnatura);
                    sqlCmd.Parameters.AddWithValue("@data_wprowadzenia", data_wprowadzenia);
                    sqlCmd.Parameters.AddWithValue("@data_otrzymania", data_otrzymania);
                    sqlCmd.Parameters.AddWithValue("@termin", termin);
                    sqlCmd.Parameters.AddWithValue("@id_kreatora", id_kreatora);
                    sqlCmd.Parameters.AddWithValue("@data_zwrotu", data_zwrotu);
                    sqlCmd.ExecuteScalar();
                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_obciazenie_osoby

        public string modyfikuj_obciazenie_osoby(string id_, string id_osoby, string instytucja, string sygnatura, DateTime data_wprowadzenia, DateTime data_otrzymania, DateTime termin, string id_kreatora, DateTime data_zwrotu)
        {
            var conn = new SqlConnection(con_str);
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("update   tbl_obciazenia set id_osoby=@id_osoby, instytucja=@instytucja, sygnatura=@sygnatura, data_wprowadzenia=@data_wprowadzenia, data_otrzymania=@data_otrzymania, termin=@termin,data_zwrotu=@data_zwrotu where id_=@id_", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@id_", id_);
                    sqlCmd.Parameters.AddWithValue("@id_osoby", id_osoby);
                    sqlCmd.Parameters.AddWithValue("@instytucja", instytucja);
                    sqlCmd.Parameters.AddWithValue("@sygnatura", sygnatura);
                    sqlCmd.Parameters.AddWithValue("@data_wprowadzenia", data_wprowadzenia);
                    sqlCmd.Parameters.AddWithValue("@data_otrzymania", data_otrzymania);
                    sqlCmd.Parameters.AddWithValue("@data_zwrotu", data_zwrotu);
                    sqlCmd.Parameters.AddWithValue("@termin", termin);

                    sqlCmd.ExecuteScalar();

                    conn.Close();
                    return "0";
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of modyfikuj_obciazenie_osoby

        public string dodaj_osobe(string imie, string nazwisko, string ulica, string kod_poczt, string miejscowosc, DateTime data_poczatkowa, DateTime data_koncowa, int id_kreatora, string tytul, string pesel, int typ)
        {
            var conn = new SqlConnection(con_str);
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("insert into  tbl_osoby (imie,nazwisko,ulica,kod_poczt,miejscowosc,data_poczatkowa,data_koncowa,id_kreatora,data_kreacji,pesel,tytul,typ)  values (@imie,@nazwisko,@ulica,@kod_poczt,@miejscowosc,@data_poczatkowa,@data_koncowa,@id_kreatora,getdate(),@pesel,@tyt,@typ) select @@identity", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@imie", imie);
                    sqlCmd.Parameters.AddWithValue("@nazwisko", nazwisko);
                    sqlCmd.Parameters.AddWithValue("@ulica", ulica);
                    sqlCmd.Parameters.AddWithValue("@kod_poczt", kod_poczt);
                    sqlCmd.Parameters.AddWithValue("@miejscowosc", miejscowosc);
                    sqlCmd.Parameters.AddWithValue("@data_poczatkowa", data_poczatkowa);
                    sqlCmd.Parameters.AddWithValue("@data_koncowa", data_koncowa);
                    sqlCmd.Parameters.AddWithValue("@id_kreatora", id_kreatora);
                    sqlCmd.Parameters.AddWithValue("@tyt", tytul);
                    sqlCmd.Parameters.AddWithValue("@pesel", pesel);
                    sqlCmd.Parameters.AddWithValue("@typ", typ);

                    string odp = sqlCmd.ExecuteScalar().ToString();

                    //  sqlCmd = new SqlCommand("insert into  tbl_osoby (imie,nazwisko,ulica,kod_poczt,miejscowosc,data_poczatkowa,data_koncowa,id_kreatora,data_kreacji,pesel)  values (@imie,@nazwisko,@ulica,@kod_poczt,@miejscowosc,@data_poczatkowa,@data_koncowa,@id_kreatora,getdate(),0) select @@identity", conn);

                    conn.Close();
                    return odp;
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
        }// end of dodaj_osobe

        public string dodaj_osobe(int typ, int idKreatora)
        {
            log.Info("Start funkcji dodaj_osobe w wersji skróconej");

            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@typ", typ);
            parameters.Rows.Add("@IdKreatora", idKreatora);
            string odp = Common.runQuerryWithResult("insert into  tbl_osoby (typ,id_kreatora,data_kreacji,pesel,czy_zaw)  values (@typ,@IdKreatora,Getdate(),0,0) select @@identity", con_str, parameters);
            return odp;
        }// end of dodaj_osobe

        //================================================================================
        public string dodajSkarge(int idBieglego, string numer, string rok, string sygnatura, string dataWplywu, string dataPisma, string wizytator, string uwagi, bool zakreslono, string dataZakreslenia, string idmodyfikatora, int idSkargi)
        {
            var conn = new SqlConnection(con_str);
            string result = string.Empty; ;
            SqlCommand sqlCmd;

            if (idSkargi != 0)
            {
                //modyfikacja
                using (sqlCmd = new SqlCommand())
                {
                    sqlCmd = new SqlCommand("update tbl_skargi set czyUS=1, dataUsuniecia=getdate(),idUsuwajacego=@idUsuwajacego where ident=@ident", conn);

                    try
                    {
                        conn.Open();

                        sqlCmd.Parameters.AddWithValue("@idUsuwajacego", idmodyfikatora);
                        sqlCmd.Parameters.AddWithValue("@ident", idSkargi);
                        sqlCmd.ExecuteScalar();
                        conn.Close();
                    }
                    catch
                    {
                        conn.Close();
                    }
                }
            }

            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("insert into tbl_skargi (idBieglego,numer,rok,dataWplywu,dataPisma,sygnatura,wizytator,zakreslono,dataZakreslenia,uwagi,czyus) values  (@idBieglego,@numer,@rok,@dataWplywu,@dataPisma,@sygnatura,@wizytator,@zakreslono,@dataZakreslenia,@uwagi,0)", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@idBieglego", idBieglego);
                    sqlCmd.Parameters.AddWithValue("@idKreatora", idmodyfikatora);
                    sqlCmd.Parameters.AddWithValue("@numer", numer);
                    sqlCmd.Parameters.AddWithValue("@rok", rok);
                    sqlCmd.Parameters.AddWithValue("@dataWplywu", dataWplywu);
                    sqlCmd.Parameters.AddWithValue("@dataPisma", dataPisma);
                    sqlCmd.Parameters.AddWithValue("@sygnatura", sygnatura);
                    sqlCmd.Parameters.AddWithValue("@wizytator", wizytator);
                    sqlCmd.Parameters.AddWithValue("@zakreslono", zakreslono);
                    sqlCmd.Parameters.AddWithValue("@datazakreslenia", dataZakreslenia);
                    sqlCmd.Parameters.AddWithValue("@uwagi", uwagi);

                    sqlCmd.ExecuteScalar();
                    conn.Close();
                }
                catch (Exception ex)
                {
                    log.Error("Skargi: " + ex.Message);
                    conn.Close();
                }
            }
            return result;
        }// end of czy zawieszony

        public string usunSkarge(string idmodyfikatora, int idSkargi)
        {
            var conn = new SqlConnection(con_str);
            string result = string.Empty; ;
            SqlCommand sqlCmd;

            if (idSkargi != 0)
            {
                //modyfikacja
                using (sqlCmd = new SqlCommand())
                {
                    sqlCmd = new SqlCommand("update tbl_skargi set czyUS=2 where ident=@ident", conn);

                    try
                    {
                        conn.Open();

                        sqlCmd.Parameters.AddWithValue("@idUsuwajacego", idmodyfikatora);
                        sqlCmd.Parameters.AddWithValue("@ident", idSkargi);
                        sqlCmd.ExecuteScalar();
                        conn.Close();
                    }
                    catch (Exception)
                    {
                        conn.Close();
                    }
                }
            }

            return result;
        }// end of usunskarge

        public string PodajNumerNowejSkargi(int rok)
        {
            var conn = new SqlConnection(con_str);
            string result = string.Empty; ;
            SqlCommand sqlCmd;

            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("select max (numer) from tbl_skargi where rok=@rok and czyus=0", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@rok", rok);
                    string odp = sqlCmd.ExecuteScalar().ToString();
                    int numer = int.Parse(odp) + 1;
                    result = numer.ToString();
                    conn.Close();
                }
                catch (Exception)
                {
                    conn.Close();
                    result = "1";
                }
            }

            return result;
        }// end of usunskarge

        //================================================================================

        public string dodajPowolanie(int idBieglego, DateTime poczatek, DateTime koniec, int UserId)
        {
            var conn = new SqlConnection(con_str);
            string result = string.Empty; ;
            SqlCommand sqlCmd;
            int id_powolania = 0;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("select max(id_powolania) from  tbl_powolania where id_bieglego=@id_bieglego", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@id_bieglego", idBieglego);
                    id_powolania = int.Parse(sqlCmd.ExecuteScalar().ToString());
                    conn.Close();
                }
                catch (Exception)
                {
                    conn.Close();
                }
                id_powolania++;
                sqlCmd = new SqlCommand("insert into       tbl_powolania(id_bieglego, data_od, data_do, d_kreacji, kreator,id_powolania,czyus) VALUES(@id_bieglego,@begin,@end,getdate(),@userId,@id_powolania,0) ", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@id_bieglego", idBieglego);
                    sqlCmd.Parameters.AddWithValue("@begin", poczatek);
                    sqlCmd.Parameters.AddWithValue("@end", koniec);
                    sqlCmd.Parameters.AddWithValue("@userId", UserId);
                    sqlCmd.Parameters.AddWithValue("@id_powolania", id_powolania);
                    sqlCmd.ExecuteScalar();

                    result = "OK";
                    conn.Close();
                }
                catch (Exception ex)
                {
                    return ex.Message;
                }
            }
            return result;
        }// end of czy zawieszony
        /*
        public bool modyfikujPowolanie(int idBieglego, int idPowolania, DateTime poczatek, DateTime koniec, int UserId)
        {
            var conn = new SqlConnection(con_str);
            bool result = false;
            SqlCommand sqlCmd;
            int id_powolania = 0;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("insert into       tbl_powolania(id_bieglego, data_od, data_do, d_modyfikacji, modyfikator,id_powolania,czyus) VALUES(@id_bieglego,@begin,@end,getdate(),@userId,@id_powolania,0) ", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@id_bieglego", idBieglego);
                    sqlCmd.Parameters.AddWithValue("@begin", poczatek);
                    sqlCmd.Parameters.AddWithValue("@end", koniec);
                    sqlCmd.Parameters.AddWithValue("@userId", UserId);
                    sqlCmd.Parameters.AddWithValue("@id_powolania", idPowolania);
                    sqlCmd.ExecuteScalar();

                    result = true;
                    conn.Close();
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            return result;
        }// end of czy zawieszony

        public bool czy_zawieszony(int id)
        {
            var conn = new SqlConnection(con_str);
            bool result = false;
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("select czy_zaw from  tbl_osoby where ident=@ident ", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@ident", id);
                    string odp = sqlCmd.ExecuteScalar().ToString();
                    if (odp == "True")
                    {
                        result = true;
                    }

                    conn.Close();
                }
                catch (Exception ex)
                {
                    return false;
                }
            }
            return result;
        }// end of czy zawieszony

        public int acces_level(int id)
        {
            var conn = new SqlConnection(con_str);
            int result = 0;
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("SELECT         level_ FROM            uzytkownik where ident=@ident", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@ident", id);
                    string odp = sqlCmd.ExecuteScalar().ToString();
                    result = int.Parse(odp);

                    conn.Close();
                }
                catch 
                {
                    return 0;
                }
            }
            return result;
        }// end of czy zawieszony

        public string data_Poczatku_zawieszenia(int id)
        {
            var conn = new SqlConnection(con_str);
            string result = "";
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("SELECT DISTINCT TOP (1) dataPoczZaw FROM            tbl_Zawieszenia where  idbieglego=@ident ", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@ident", id);
                    string odp = sqlCmd.ExecuteScalar().ToString();
                    result = odp;
                    conn.Close();
                }
                catch 
                {
                    return "";
                }
            }
            return result;
        }// end of data zawieszenia
        */
        /*
        public string data_Konca_zawieszenia(int id)
        {
            var conn = new SqlConnection(con_str);
            string result = "";
            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("SELECT DISTINCT TOP (1) dataKoncaZaw FROM            tbl_Zawieszenia where  idbieglego=@ident ", conn);

                try
                {
                    conn.Open();

                    sqlCmd.Parameters.AddWithValue("@ident", id);
                    string odp = sqlCmd.ExecuteScalar().ToString();
                    result = odp;
                    conn.Close();
                }
                catch (Exception ex)
                {
                    return "";
                }
            }
            return result;
        }// end of data zawieszenia
        */
        public DataSet loguj(string user_, string haslo)
        {
            SqlConnection conn = new SqlConnection(con_str);

            DataSet dsLogin = new DataSet();

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT id_,rola,imie,nazwisko  FROM tbl_users_ where user_=@user and pasword=@haslo", conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@user", user_);
                daMenu.SelectCommand.Parameters.AddWithValue("@haslo", haslo);
                daMenu.Fill(dsLogin);
                conn.Close();
            }
            catch (Exception ec)
            {
                log.Error("logowanie " + ec.Message);
            }
            return dsLogin;
        } // end of loguj
        /*
        public DataTable tabelaStatystycznaNew(string querry)
        {
            SqlConnection conn = new SqlConnection(con_str);
            DataTable dT = new DataTable();
            DataSet dsLogin = new DataSet();

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand(querry, conn);

                daMenu.Fill(dsLogin);
                conn.Close();
                dT = dsLogin.Tables[0];
            }
            catch (Exception ec)
            {
                return null;
            }
            return dT;
        } // 
        */
        public DataTable tabelaStatystyczna(string querry, string idBieglego)
        {
            SqlConnection conn = new SqlConnection(con_str);
            DataTable dT = new DataTable();
            DataSet dsLogin = new DataSet();

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand(querry, conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@biegly", idBieglego);
                daMenu.Fill(dsLogin);
                conn.Close();
                dT = dsLogin.Tables[0];
            }
            catch 
            {
                return null;
            }
            return dT;
        } // end of tabelaStatystyczna

        public DataTable dane_korespondencyjne(string id_)
        {
            DataTable result = null;
            SqlConnection conn = new SqlConnection(con_str);
            DataSet dsLogin = new DataSet();

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECt tel1 ,tel2 ,email ,adr_kores ,kod_poczt_kor ,miejscowosc_kor ,uwagi  FROM tbl_osoby where ident=@id", conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@id", id_);
                daMenu.Fill(dsLogin);
                conn.Close();
                result = dsLogin.Tables[0];
            }
            catch
            {
            }
            return result;
        } // end of loguj

        public DataTable dane_osobowe(string id_)
        {
            DataTable result = null;
            SqlConnection conn = new SqlConnection(con_str);
            DataSet dsLogin = new DataSet();

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT    imie, nazwisko, ulica, kod_poczt, miejscowosc, data_poczatkowa, data_koncowa, id_kreatora,  pesel, czyus,  tytul,  czy_zaw, specjalizacja_opis,d_zawieszenia,uwagi, dataKoncaZawieszenia   FROM            tbl_osoby where ident=@id", conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@id", id_);
                daMenu.Fill(dsLogin);
                conn.Close();
                result = dsLogin.Tables[0];
            }
            catch (Exception ex)
            {
                DataTable ddTT = new DataTable();
                ddTT.Columns.Add("info", typeof(String));
                DataRow dr = ddTT.NewRow();
                dr[0] = ex.Message;
                ddTT.Rows.Add(dr);
                result = ddTT;
            }
            return result;
        } // end of loguj
        /*
        public DataTable statystykiBiegłego(int id_)
        {
            DataTable result = null;
            SqlConnection conn = new SqlConnection(con_str);
            DataSet dsLogin = new DataSet();

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT      [kwerendaStatystyczna],[opis]  FROM kwerendaStaystyczna where biegly=@biegly order by ident", conn);
                daMenu.SelectCommand.Parameters.AddWithValue("@biegly", id_);
                daMenu.Fill(dsLogin);
                conn.Close();
                result = dsLogin.Tables[0];
            }
            catch
            {
            }
            return result;
        } // end of staystykiBiegłego
        */
        public int podajIdOsobyPoNumerzeSkargi(int idSkargi)
        {
            log.Info("Start funkcji podajIdOsobyPoNumerzeSkargi");
            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@idSakargi", idSkargi);
            try
            {
                return int.Parse(Common.runQuerryWithResult("SELECT idBieglego  FROM tbl_skargi where ident=@idSakargi", con_str, parameters));
            }
            catch (Exception)
            {
            }

            return 0;
        }

        public string podajNazwiskoOsobyPoNumerzeSkargi(int idSkargi)
        {
            log.Info("Start funkcji podajIdOsobyPoNumerzeSkargi");
            DataTable parameters = Common.makeParameterTable();
            parameters.Rows.Add("@idSakargi", idSkargi);
            try
            {
                return Common.runQuerryWithResult("select distinct tbl_osoby.nazwisko from tbl_osoby join tbl_skargi on tbl_osoby .ident=tbl_skargi.idBieglego where tbl_skargi.ident=@idSakargi", con_str, parameters);
            }
            catch (Exception)
            {
            }

            return string.Empty;
        }
        /*
        public DataTable Specjalizacje()
        {
            DataTable result = null;
            SqlConnection conn = new SqlConnection(con_str);
            DataSet dsSpecjalizacje = new DataSet();

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT glo_specjalizacje.nazwa,glo_specjalizacje.id_, glo_grupy_specjalizacji.Nazwa AS NazwaGrupy, glo_grupy_specjalizacji.ident AS IdSpecjalizacji FROM glo_grupy_specjalizacji RIGHT OUTER JOIN glo_specjalizacje ON glo_grupy_specjalizacji.ident = glo_specjalizacje.grupa WHERE (glo_grupy_specjalizacji.Nazwa IS NOT NULL)", conn);
             
                daMenu.Fill(dsSpecjalizacje);
                conn.Close();
                result = dsSpecjalizacje.Tables[0];
            }
            catch
            {
            }
            return result;
        } // end of Specjalizacje
        public DataTable GrupySpecjalizacji()
        {
            DataTable result = null;
            SqlConnection conn = new SqlConnection(con_str);
            DataSet dsSpecjalizacje = new DataSet();

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT DISTINCT Nazwa, ident FROM glo_grupy_specjalizacji ORDER BY Nazwa", conn);

                daMenu.Fill(dsSpecjalizacje);
                conn.Close();
                result = dsSpecjalizacje.Tables[0];
            }
            catch
            {
            }
            return result;
        } // end of Specjalizacje*/
    }
}