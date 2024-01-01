using System;
using System.Collections.Generic;

using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace wab
{
    public partial class Lista_osob : System.Web.UI.UserControl
    {

        public string pobierz_osobe()
        {
            return HiddenField2.Value.ToString();
        }

        public void ustaw_osobe(string s)
        {
            try
            {
                TextBox1.Text = s;
                
            }
            catch
            {

            }

        }



        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void GridView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //wybor osoby
            HiddenField2.Value = GridView1.SelectedDataKey[0].ToString();
            TextBox1.Text = GridView1.SelectedDataKey[1].ToString()+" "+GridView1.SelectedDataKey[2].ToString();
            Panel3.Visible = false;
        }

        protected void LinkButton1_Click(object sender, EventArgs e)
        {
            Panel3.Visible = true;
        }

        protected void ImageButton1_Click(object sender, ImageClickEventArgs e)
        {
            Panel3.Visible = false;
        }

        protected void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {

            
            lista_osob_x.SelectCommand = "SELECT [imie], [nazwisko], [ulica], [kod_poczt], [miejscowosc], [data_poczatkowa], [data_koncowa], [ident] FROM [View_lista_osob_aktywnych]";
            if (CheckBox1.Checked)
            {
                DropDownList1.Enabled = true;
                DropDownList1.SelectedIndex = 0;
                try
                {
                    Session["id_spec"] = DropDownList1.SelectedValue.ToString();
                    lista_osob_x.SelectCommand = "SELECT [imie], [nazwisko], [ulica], [kod_poczt], [miejscowosc], [data_poczatkowa], [data_koncowa], [ident] FROM View_lista_osob_z_kategoriami where id_specjalizacji=" + DropDownList1.SelectedValue.ToString().Trim();
                }
                catch
                { }
            }
            else
            {
                lista_osob_x.SelectCommand = "SELECT [imie], [nazwisko], [ulica], [kod_poczt], [miejscowosc], [data_poczatkowa], [data_koncowa], [ident] FROM [View_lista_osob_aktywnych]";
                DropDownList1.Enabled = false;
                DropDownList1.SelectedIndex = -1;
            }
            GridView1.DataBind();
        }

        protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                Session["id_spec"] = DropDownList1.SelectedValue.ToString();

                lista_osob_x.SelectCommand = "SELECT [imie], [nazwisko], [ulica], [kod_poczt], [miejscowosc], [data_poczatkowa], [data_koncowa], [ident] FROM View_lista_osob_z_kategoriami where id_specjalizacji=" + DropDownList1.SelectedValue.ToString().Trim();

            }
            catch
            {

            }
            GridView1.DataBind();
        }
    }
}