using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace wab2018
{
    public partial class Site1 : System.Web.UI.MasterPage
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string User_id = string.Empty;

            try
            {

                string id = Request.QueryString["logout"];
                if (id != null)
                {
                    Session["user_id"] = null;
                    NavigationMenu.Items.Clear();
                    MenuItem mn = new MenuItem();
                    mn.NavigateUrl = "logowanie.aspx";
                    mn.Text = "Logowanie";
                    NavigationMenu.Items.Add(mn);
                }
            }
            catch
            { }

            try
            {
                try
                {
                    User_id = (string)Session["user_id"];
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
                                NavigationMenu.Items.Clear();
                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "Lista_01.aspx";
                                mn.Text = "Lista osób";
                                NavigationMenu.Items.Add(mn);
                            }
                            break;
                        case "1":
                            {
                                NavigationMenu.Items.Clear();

                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx";
                                mn.Text = "Obciążenie biegłych";
                                NavigationMenu.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "Lista_11.aspx";
                                mn.Text = "Wykaz biegłych";
                                NavigationMenu.Items.Add(mn);


                                mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx?logout=1";
                                mn.Text = "Wyloguj";
                                NavigationMenu.Items.Add(mn);
                            }
                            break;
                        case "2":
                            {
                                NavigationMenu.Items.Clear();

                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx";
                                mn.Text = "Obciążenie biegłych";
                                NavigationMenu.Items.Add(mn);


                                mn = new MenuItem();
                                mn.NavigateUrl = "Lista_11.aspx";
                                mn.Text = "Wykaz biegłych";
                                NavigationMenu.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "Add_01.aspx";
                                mn.Text = "Nowy Biegły";
                                NavigationMenu.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "Lista_03.aspx";
                                mn.Text = "Obciążenia";
                                NavigationMenu.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx?logout=1";
                                mn.Text = "Wyloguj";
                                NavigationMenu.Items.Add(mn);
                            }
                            break;
                        case "3":
                            {
                                NavigationMenu.Items.Clear();

                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx";
                                mn.Text = "Obciążenie biegłych";
                                NavigationMenu.Items.Add(mn);


                                mn = new MenuItem();
                                mn.NavigateUrl = "Lista_11.aspx";
                                mn.Text = "Wykaz biegłych";
                                NavigationMenu.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "wykazMediatorow.aspx";
                                mn.Text = "Wykaz mediatorów";
                                NavigationMenu.Items.Add(mn);


                                mn = new MenuItem();
                                mn.NavigateUrl = "Add_01.aspx";
                                mn.Text = "Nowy Biegły";
                                NavigationMenu.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "nowyMediator.aspx";
                                mn.Text = "Nowy mediator";
                                NavigationMenu.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "Lista_03.aspx";
                                mn.Text = "Obciążenia";
                                NavigationMenu.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "admin.aspx";
                                mn.Text = "Administracja";
                                NavigationMenu.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "import.aspx";
                                mn.Text = "Import danych";
                                NavigationMenu.Items.Add(mn);

                                mn = new MenuItem();
                                mn.NavigateUrl = "default.aspx?logout=1";
                                mn.Text = "Wyloguj";
                                NavigationMenu.Items.Add(mn);
                                
                            }
                            break;
                        default:
                            {
                                NavigationMenu.Items.Clear();
                                MenuItem mn = new MenuItem();
                                mn.NavigateUrl = "logowanie.aspx";
                                mn.Text = "Logowanie";
                                NavigationMenu.Items.Add(mn);
                                //  mn = new MenuItem();
                                //   mn.NavigateUrl = "login.aspx?logout=1";
                                //   mn.Text = "Wyloguj";
                                //   NavigationMenu.Items.Add(mn);
                            }
                            break;
                    }


                }
                else
                {
                    NavigationMenu.Items.Clear();
                    MenuItem mn = new MenuItem();
                    mn.NavigateUrl = "logowanie.aspx";
                    mn.Text = "Logowanie";
                    NavigationMenu.Items.Add(mn);
                }
            }
            catch
            {
                User_id = "0";
                Session["user_id"] = User_id;
                NavigationMenu.Items.Clear();
                MenuItem mn = new MenuItem();
                mn.NavigateUrl = "logowanie.aspx";
                mn.Text = "Logowanie";
                NavigationMenu.Items.Add(mn);

            }

        }

        protected void NavigationMenu_MenuItemClick(object sender, MenuEventArgs e)
        {
            try
            {
                if (e.Item.Text == "Wyloguj")
                {
                    NavigationMenu.Items.Clear();
                    MenuItem mn = new MenuItem();
                    mn.NavigateUrl = "login.aspx";
                    mn.Text = "Logowanie";
                    NavigationMenu.Items.Add(mn);
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