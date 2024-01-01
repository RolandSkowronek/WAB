using System;
using System.Web.UI.WebControls;

namespace wab2018
{
    public partial class Master : System.Web.UI.MasterPage
    {
        public log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        protected void Page_Load(object sender, EventArgs e)
        {
            var fileContents = System.IO.File.ReadAllText(Server.MapPath(@"~//version.txt"));    // file read with version
            this.Page.Title = "Portal Biegłych i Mediatorów Sądowych  " + fileContents.ToString().Trim();

            string User_id = string.Empty;

            try
            {

                string id = Request.QueryString["logout"];
                if (id != null)
                {
                    Session["user_id"] = null;
                    Menu1.Items.Clear();
                    MenuItem mn = new MenuItem();
                    mn.NavigateUrl = "logowanie.aspx";
                    mn.Text = "Logowanie";
                    Menu1.Items.Add(mn);
                }
            }
            catch
            { }

            try
            {
                try
                {
                    User_id = (string)Session["user_id"];
                    string rola = (string)Session["rola"];
                }
                catch
                {
                    User_id = "0";
                    Session["user_id"] = User_id;

                }

             
                if (User_id.Length > 0)
                {

                    switch ((string)Session["rola"])
                    {

                        case "0":
                            {
                                Menu1.Items.Clear();
                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "Lista_01.aspx";
                                mn.Text = "Lista osób";
                                Menu1.Items.Add(mn);
                            }
                            break;
                        case "1":
                            {
                                Menu1.Items.Clear();

                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx";
                                mn.Text = "Obciążenie biegłych";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "biegliLista.aspx";
                                mn.Text = "Wykaz biegłych";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "mediatorzyLista.aspx";
                                mn.Text = "Wykaz mediatorów";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "wykazSkarg.aspx";
                                mn.Text = "Wykaz skarg";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx?logout=1";
                                mn.Text = "Wyloguj";
                                Menu1.Items.Add(mn);
                            }
                            break;
                        case "2":
                            {
                                Menu1.Items.Clear();

                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx";
                                mn.Text = "Obciążenie biegłych";
                                Menu1.Items.Add(mn);


                                mn = new MenuItem();
                                mn.NavigateUrl = "biegliLista.aspx";
                                mn.Text = "Wykaz biegłych";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "mediatorzyLista.aspx";
                                mn.Text = "Wykaz mediatorów";
                                Menu1.Items.Add(mn);

                                
                                mn = new MenuItem();
                                mn.NavigateUrl = "Lista_03.aspx";
                                mn.Text = "Obciążenia";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "wykazSkarg.aspx";
                                mn.Text = "Wykaz skarg";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx?logout=1";
                                mn.Text = "Wyloguj";
                                Menu1.Items.Add(mn);
                            }
                            break;
                        case "3":
                            {
                                Menu1.Items.Clear();

                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx";
                                mn.Text = "Obciążenie biegłych";
                                Menu1.Items.Add(mn);


                                mn = new MenuItem();
                                mn.NavigateUrl = "biegliLista.aspx";
                                mn.Text = "Wykaz biegłych";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "mediatorzyLista.aspx";
                                mn.Text = "Wykaz mediatorów";
                                Menu1.Items.Add(mn);

                                
                                mn = new MenuItem();
                                mn.NavigateUrl = "Lista_03.aspx";
                                mn.Text = "Obciążenia";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "wykazSkarg.aspx";
                                mn.Text = "Wykaz skarg";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "admin.aspx";
                                mn.Text = "Administracja";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "import.aspx";
                                mn.Text = "Import danych";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx?logout=1";
                                mn.Text = "Wyloguj";
                                Menu1.Items.Add(mn);

                            }
                            break;
                        case "4":
                            {
                                Menu1.Items.Clear();

                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx";
                                mn.Text = "Obciążenie biegłych";
                                Menu1.Items.Add(mn);


                                mn = new MenuItem();
                                mn.NavigateUrl = "biegliLista.aspx";
                                mn.Text = "Wykaz biegłych";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "mediatorzyLista.aspx";
                                mn.Text = "Wykaz mediatorów";
                                Menu1.Items.Add(mn);


                                mn = new MenuItem();
                                mn.NavigateUrl = "Lista_03.aspx";
                                mn.Text = "Obciążenia";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "wykazSkarg.aspx";
                                mn.Text = "Wykaz skarg";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx?logout=1";
                                mn.Text = "Wyloguj";
                                Menu1.Items.Add(mn);
                            }
                            break;
                        case "5":
                            {
                                Menu1.Items.Clear();

                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx";
                                mn.Text = "Obciążenie biegłych";
                                Menu1.Items.Add(mn);


                                mn = new MenuItem();
                                mn.NavigateUrl = "biegliLista.aspx";
                                mn.Text = "Wykaz biegłych";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "mediatorzyLista.aspx";
                                mn.Text = "Wykaz mediatorów";
                                Menu1.Items.Add(mn);


                                mn = new MenuItem();
                                mn.NavigateUrl = "Lista_03.aspx";
                                mn.Text = "Obciążenia";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "wykazSkarg.aspx";
                                mn.Text = "Wykaz skarg";
                                Menu1.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx?logout=1";
                                mn.Text = "Wyloguj";
                                Menu1.Items.Add(mn);
                            }
                            break;
                        default:
                            {
                                Menu1.Items.Clear();
                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "logowanie.aspx";
                                mn.Text = "Logowanie";
                                Menu1.Items.Add(mn);
                                //  mn = new MenuItem();
                                //   mn.NavigateUrl = "login.aspx?logout=1";
                                //   mn.Text = "Wyloguj";
                                //   Menu1.Items.Add(mn);
                            }
                            break;
                    }


                }
                else
                {
                    
                    Menu1.Items.Clear();
                    MenuItem mn = new MenuItem();
                    mn.NavigateUrl = "logowanie.aspx";
                    mn.Text = "Logowanie";
                    Menu1.Items.Add(mn);
                }


            }
            catch (Exception ex)
            {
                log.Error("master: " + ex.Message);
                User_id = "0";
                Session["user_id"] = User_id;
                Menu1.Items.Clear();
                MenuItem mn = new MenuItem();
                mn.NavigateUrl = "logowanie.aspx";
                mn.Text = "Logowanie";
                Menu1.Items.Add(mn);

            }

        }


        protected void Menu1_MenuItemClick(object sender, MenuEventArgs e)
        {
            try
            {
                if (e.Item.Text == "Wyloguj")
                {
                    Menu1.Items.Clear();
                    MenuItem mn = new MenuItem();
                    mn.NavigateUrl = "login.aspx";
                    mn.Text = "Logowanie";
                    Menu1.Items.Add(mn);
                    Server.Transfer("login.aspx");
                    Session["user_id"] = null;
                }
            }
            catch (Exception)
            {
            }
        }
    }
}