using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace wab2018.userControls
{
    public partial class WebUserControl1 : System.Web.UI.UserControl
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            DateTime data1 = ASPxDateEdit1.Date;
            if (ASPxDateEdit1.Date.Year==1)
            {
                ASPxDateEdit1.Date = DateTime.Now;
            }
            if (ASPxDateEdit2.Date.Year == 1)
            {
                ASPxDateEdit2.Date = DateTime.Now;
            }

        }
    }
}