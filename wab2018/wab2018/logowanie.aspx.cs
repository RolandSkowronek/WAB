using System;
using System.Data;
using log4net;
namespace wab2018
{
    public partial class logowanie : System.Web.UI.Page
    {
        public Class2 cl = new Class2();

        public log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        protected void Page_Load(object sender, EventArgs e)
        {
            Session["sesja"] = null;
            try
            {
                if (Session["User_id"] == null)
                {
                    TextBox1.Text = "";
                    TextBox1.Focus();
                }
            }
            catch
            { }
        }

        protected void LinkButton1_Click(object sender, EventArgs e)
        {
            zaloguj();
        }

        protected void TextBox2_TextChanged(object sender, EventArgs e)
        {
            zaloguj();
        }

        protected void zaloguj()
        {
            int err = 0;

            Label3.Text = "";
            DataSet login = new DataSet();
            log.Info("Rozpoczecie logowania uzytkownika " +TextBox1.Text.Trim());
            login = cl.loguj(TextBox1.Text.Trim(), TextBox2.Text.Trim());
            try
            {

                DataRow dr = login.Tables[0].Rows[0];
                string id = dr[0].ToString();
                string imie = dr[0].ToString();
                string nazwisko = dr[0].ToString();
                string rola = dr[1].ToString();
                if (rola != "0")
                {
                    Session["user_id"] = id;
                    Session["rola"] = rola;
                }
                else
                {
                    err = 2;
                }
            }
            catch (Exception ex)
            {
                log.Error("logowanie : " + ex.Message);
                err = 1;
            }

            switch (err)
            {
                case 0:
                    {
                        Server.Transfer("default.aspx");
                        TextBox1.Text = "";
                        TextBox2.Text = "";
                        TextBox1.Focus();

                    }
                    break;
                case 1:
                    {
                        Label3.Text = "Bład logowania";
                        TextBox1.Text = "";
                        TextBox2.Text = "";
                        TextBox1.Focus();
                    }
                    break;
                case 2:
                    {
                        Label3.Text = "Nie masz praw do logowania!";
                        log.Error("logowanie : " + "Nie masz praw do logowania!");
                    }
                    break;
                default:
                    break;
            }


        }//zaloguj 
    }
}