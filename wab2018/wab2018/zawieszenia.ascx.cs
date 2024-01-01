using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace wab2018
{
    public partial class zawieszenia : System.Web.UI.UserControl
    {
       
     
        public bool czyZawieszenie
            {
                get
            {
                return cbZawieszenie.Checked ;
            }
            set
            {
                cbZawieszenie.Checked = czyZawieszenie;
            }
            
            
        }
        public DateTime pocztek
        {
            get
            {
                return txPoczatekZawiszenia.Date ;
            }
            set
            {
                txPoczatekZawiszenia.Date = pocztek;
            }

        }

        public DateTime koniec
        {

            get
            {
                return txKoniecZawieszenia.Date; 
            }
            set
            {
                txKoniecZawieszenia.Date = koniec;
            }
        }
        protected void Page_Load(object sender, EventArgs e)
        {
          //  Panel1.Visible = cbZawieszenie.Checked;
        }

      

        protected void zawieszeni(object sender, EventArgs e)
        {

            zawieszaj( cbZawieszenie.Checked);
         
        }
        public void zawieszaj(bool zawieszenie)
        {
            Panel1.Visible = zawieszenie;
            cbZawieszenie.Checked = zawieszenie;
            if (txPoczatekZawiszenia.Text == "")
            {
                txPoczatekZawiszenia.Date = DateTime.Now.Date;
            }
            if (txKoniecZawieszenia.Text == "")
            {
                txKoniecZawieszenia.Date = txPoczatekZawiszenia.Date.AddYears(5);
            }

        }
    }
}