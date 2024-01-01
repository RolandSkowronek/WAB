using System;
using System.Collections.Generic;

using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace wab
{
    public partial class cal1 : System.Web.UI.UserControl
    {

        public string pobierz_date()
        {
            return HiddenField1.Value.ToString();
        }

        public void ustaw_date(string s)
        {
            try
            {
                TextBox1.Text = DateTime.Parse(s).ToString("dd.MM.yyyy");
                HiddenField1.Value = DateTime.Parse(s).ToShortDateString ();
            }
            catch
            {
                HiddenField1.Value = "";
                TextBox1.Text = "";
            }

        }


        protected void Page_Load(object sender, EventArgs e)
        {
            if (TextBox1.Text.Trim().Length == 0)
            {
                TextBox1.Text = DateTime.Now.Date.ToString("dd.MM.yyyy");
                HiddenField1.Value = DateTime.Now.Date.ToShortDateString();
            }
        }

        protected void ImageButton1_Click(object sender, ImageClickEventArgs e)
        {
            Panel1.Visible = true;
            Calendar1.VisibleDate = DateTime.Parse(HiddenField1.Value.ToString() );
            Calendar1.SelectedDate = DateTime.Parse(HiddenField1.Value.ToString());
        }

        protected void ImageButton2_Click(object sender, ImageClickEventArgs e)
        {
            Panel1.Visible = false;
        }

        protected void Calendar1_SelectionChanged(object sender, EventArgs e)
        {
            TextBox1.Text = Calendar1.SelectedDate.ToString("dd.MM.yyyy");
            HiddenField1.Value = Calendar1.SelectedDate.ToShortDateString();
            Panel1.Visible = false;
        }
    }
}