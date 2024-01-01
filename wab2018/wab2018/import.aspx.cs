using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Configuration;
using System.IO;
using System.Text;


namespace wab2018
{
    public partial class import : System.Web.UI.Page
    {

          private static readonly log4net.ILog log
   = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public string con_str = string.Empty;
        public string con_str_wcyw = string.Empty;

        private cm common = new cm();
        public Class2 Cl = new Class2();
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                con_str = ConfigurationManager.ConnectionStrings["wap"].ConnectionString;
                con_str_wcyw = ConfigurationManager.ConnectionStrings["wcyw"].ConnectionString;


                Session["sesja"] = null;
                if (!IsPostBack)
                {

                    if (Session["user_id"] == null)
                    {
                        try
                        {
                            string rola = (string)Session["rola"];
                            if (rola != "3")
                            {
                                //     Server.Transfer("default.aspx");

                            }
                        }
                        catch
                        {
                            // Server.Transfer("default.aspx");
                        }
                        //    Server.Transfer("default.aspx");
                    }


                }
            }
            catch (Exception)
            { }


        }


        protected void laduj_dane(int id_kwerendy, string db_server, string db_name, string db_user, string db_password, string idBazy)
        {

            SqlConnection connx = new SqlConnection(con_str);
            int result = 0;

            string ReadQuery = string.Empty;
            string kwerenda = string.Empty;
            //odczyt kwerendy



            string query = string.Empty;
            SqlCommand sqlCmd;

            TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Próba podłączenia do bazy z kwerendami" + Environment.NewLine);
            log.Debug("Próba podłączenia do bazy z kwerendami");
            try
            {
                using (sqlCmd = new SqlCommand())
                {
                    // sqlCmd = new SqlCommand("SELECT     TOP (1) kwerenda FROM         glo_kwerenda where id_=1", connx);
                    try
                    {
                        connx.Open();
                        sqlCmd.CommandText = "SELECT        kwerenda FROM            glo_kwerenda where id_=@id";
                        sqlCmd.Connection = connx;
                        sqlCmd.Parameters.AddWithValue("@id", id_kwerendy);
                        query = sqlCmd.ExecuteScalar().ToString();
                        connx.Close();
                    }
                    catch (Exception ex)
                    {
                        log.Error(ex.Message);
                        connx.Close();
                        Label2.Text = ex.Message;
                        TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Błąd nr 1: nie pobrano kwerendy!!!! " + Environment.NewLine);
                        TextBox1.Text = TextBox1.Text + (ex.Message);
                    }
                    string connection_string = "Data Source=" + db_server + ";Initial Catalog=" + db_name + ";Persist Security Info=True;User ID=" + db_user + ";Password=" + db_password;

                    TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " #############################################" + Environment.NewLine);
                    TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Connection string " + connection_string + Environment.NewLine);
                    TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " ############################################# " + Environment.NewLine);
                    // string connection_string = con_str;

                    int err = 0;
                    TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Odczyt danych z baz zewnętrznych " + Environment.NewLine);
                    log.Debug(" Odczyt danych z baz zewnętrznych ");
                    TextBox1.Text = TextBox1.Text + ("====================================" + Environment.NewLine);
                    DataSet Dds = new DataSet();
                    SqlConnection conn = new SqlConnection(connection_string);
                    TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Połaczenie do bazy " + Environment.NewLine);
                    log.Debug("połaczenie do bazy");
                    if (query.Length != 0)
                    {
                        try
                        {
                            conn.Open();
                            TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Połaczono z bazą nr : " + idBazy + " ! " + Environment.NewLine);
                            log.Debug(" Połaczono z bazą nr : " + idBazy + " ! ");
                            SqlDataAdapter daMenu = new SqlDataAdapter();
                            daMenu.SelectCommand = new SqlCommand(query, conn);
                            daMenu.SelectCommand.Parameters.AddWithValue("@zrodlo", idBazy);
                            daMenu.Fill(Dds);
                            TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Dane zaimportowano z bazy nr: " + idBazy + Environment.NewLine);
                            log.Debug(" Dane zaimportowano z bazy nr: " + idBazy);
                            conn.Close();
                            DataTable nowa = new DataTable();

                            try
                            {
                                nowa = Dds.Tables[0];
                            }
                            catch (Exception ec)
                            {
                                log.Error(ec.Message);
                                Label2.Text = "Import danych " + ec.Message;
                                err = 2;
                            }
                            log.Debug(" Próba połaczenia do bazy docelowej ");
                            TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Próba połaczenia do bazy docelowej " + Environment.NewLine);
                            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con_str))
                            {
                                log.Debug(" Pobieranie klucza timeout do pompowania danych");
                                TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Pobieranie klucza timeout do pompowania danych" + Environment.NewLine);
                                string TimeOut = Cl.PobierzDane("BulkCopyTimeout");
                                int timeOut = 0;
                                try
                                {
                                    timeOut = int.Parse(TimeOut);
                                }
                                catch (Exception eex)
                                {
                                    log.Error("Bład pobrania klucza timeout " + eex.Message);
                                }

                                log.Debug("Timeout ustawiony na " + TimeOut + " sekund");
                                TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Klucza timeout = " + timeOut.ToString() + Environment.NewLine);

                                bulkCopy.BulkCopyTimeout = timeOut;
                                TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Połaczono z bazą docelową " + Environment.NewLine);
                                log.Debug("Połaczono z bazą docelową ");
                                bulkCopy.DestinationTableName = "tbl_main";
                                try
                                {
                                    log.Debug(" Wgranie do bazy docelowej ");
                                    TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Wgranie do bazy docelowej " + Environment.NewLine);
                                    bulkCopy.BulkCopyTimeout = timeOut;
                                    bulkCopy.WriteToServer(nowa);
                                    log.Debug(" Zakończono wgranie do bazy docelowej ");
                                    TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Zakończono wgranie do bazy docelowej " + Environment.NewLine);

                                }
                                catch (Exception ex)
                                {
                                    log.Error(" Błąd wgrania  do bazy docelowej " + ex.Message);
                                    TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Błąd wgrania  do bazy docelowej " + Environment.NewLine);
                                    TextBox1.Text = TextBox1.Text + ("Bład: " + ex.Message + Environment.NewLine);
                                    Label3.Text = ex.Message;
                                    err = 3;
                                }
                            }
                        }
                        catch (Exception ec)
                        {
                            log.Error(" Błąd importu danych z bazy nr: " + idBazy + " " + ec.Message);
                            Label1.Text = ec.Message;
                            TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Błąd importu danych z bazy nr: " + idBazy + Environment.NewLine);
                            TextBox1.Text = TextBox1.Text + ("Bład: " + ec.Message + Environment.NewLine);
                            err = 1;
                        }


                    }
                    else
                    {
                        Label1.Text = "Brak kwerendy";
                        log.Error("Brak kwerendy");
                    }
                    result = err;

                }
            }
            catch (Exception ex)
            {
                log.Error(" Błąd wgrania  do bazy docelowej " + ex.Message);
                TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Błąd wgrania  do bazy docelowej " + Environment.NewLine);
                TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + ex.Message + Environment.NewLine);
            }

            TextBox1.Text = TextBox1.Text + (DateTime.Now.Date.ToLocalTime() + " Koniec!!!!!!!!!!!!!!!!" + Environment.NewLine);
            log.Debug(" Koniec!!!!!!!!!!!!!!!!");
            // return result;
            SqlDataSource2.DataBind();
            GridView1.DataSource = SqlDataSource2;
            GridView1.DataBind();

        }
        /*
        protected int quick()
        {

            
            int err = 0;

            DataSet Dds = new DataSet();
            SqlConnection conn = new SqlConnection(con_str_a);

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT   rtrim( wab.ident) as ident, rtrim(inna.nazwisko)  as nazwisko ,rtrim( inna.imie) as imie, rtrim(inna.ulica) as ulica, rtrim(inna.numer) as numer, inna.kod,rtrim( inna.miejscowosc) as miejscowosc,wab.d_wyslania, wab.d_zwrotu, wab.d_termin, cast(rtrim(wab.czyopinia) as smallint) as czyopinia ,cast( rtrim(wab.czyus)as smallint) as czyus, wab.d_modyfikacji, rtrim((select symbol from repertorium where numer=spr.repertorium)) as rep, rtrim(spr.numer) as nr_sprawy, spr.rok,'1' as zrodlo  FROM inna_strona as inna, ekspertyza as ek, wyslanie as wy, wab as wab, sprawa as spr where inna.ident=ek.id_innej and wy.id_ekspertyzy=ek.ident and  wab.ident=wy.id_wab and wab.id_ekspertyzy=spr.ident and ek.czyus=0 and wy.czyus=0 and wab.czyus=0 ", conn);
                daMenu.Fill(Dds);
                conn.Close();
            }
            catch (Exception ec)
            {
                err = 1;
            }

            DataTable nowa = new DataTable();

            try
            {
                nowa = Dds.Tables[0];
            }
            catch
            {
                err = 2;
            }

            using (SqlBulkCopy bulkCopy = new SqlBulkCopy(con_str))
            {
                bulkCopy.DestinationTableName = "tbl_main";
                try
                {
                    bulkCopy.WriteToServer(nowa);
                }
                catch (Exception ex)
                {
                    err = 3;
                }
            }

            return err;

        }

        */
        protected void Button1_Click(object sender, EventArgs e)
        {


            DataSet Dds = new DataSet();
            SqlConnection conn = new SqlConnection(con_str);

            try
            {
                conn.Open();
                SqlDataAdapter daMenu = new SqlDataAdapter();
                daMenu.SelectCommand = new SqlCommand("SELECT  [id_] ,[nazwa],[db_name],[db_catalog],[db_user],[db_paswd],id_kwerendy  FROM [tbl_bazy_danych]", conn);
                daMenu.Fill(Dds);
                conn.Close();
            }
            catch (Exception ec)
            {
            }

            DataTable table = new DataTable();
            try
            {
                table = Dds.Tables[0];
            }
            catch
            {

            }
            conn = new SqlConnection(con_str);

            SqlCommand sqlCmd;
            using (sqlCmd = new SqlCommand())
            {
                sqlCmd = new SqlCommand("delete from  tbl_main", conn);
                try
                {
                    conn.Open();


                    sqlCmd.ExecuteScalar();

                    conn.Close();

                }
                catch (Exception ex)
                {
                }
            }

            foreach (DataRow row in table.Rows) // Loop over the rows.
            {

                laduj_dane(int.Parse(row[6].ToString()), row[2].ToString(), row[3].ToString(), row[4].ToString(), row[5].ToString(), row[0].ToString());

            }

            GridView1.DataBind();
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            //quick();
        }
    }
}
