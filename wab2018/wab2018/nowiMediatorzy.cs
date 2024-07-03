using System;
using DevExpress.Web.Data;
using DevExpress.Web;
using System.Data;
using System.Collections;
using System.Web.UI.HtmlControls;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace wab2018
{
    
    class nowiMediatorzy
    {
        public cm _cm = new cm();

        

            public string controlPanel(string control, ASPxGridView grid)
        {
            ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
            Panel Panel1 = pageControl.FindControl(control) as Panel;
            string pvalue = Panel1.Attributes["class"];
            var cosik = Panel1.Style;
            return pvalue;
        }
        public bool controlCheckBox(string control, ASPxGridView grid)
        {
            ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
            CheckBox CB = pageControl.FindControl(control) as CheckBox;

            return CB.Checked;
        }


        public bool controlCheckbox(string control,ASPxGridView grid)
        {
            ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
            ASPxCheckBox txt = pageControl.FindControl(control) as ASPxCheckBox;
            if (txt==null)
            {
                return false;
            }
            return txt.Checked;
        }
        public string controlDropDownList(string control, ASPxGridView grid)
        {
            ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
            try
            {
                DropDownList dropList= pageControl.FindControl(control) as DropDownList;
                if (dropList.SelectedIndex>-1)
                
                {
                    var aaa= dropList.SelectedValue;
                    return dropList.SelectedValue;
                }

            }
            catch (Exception)
            {

               
            }
            
            return "zawieszono";
        }
        public string controlText(string control, ASPxGridView grid)
        {
            ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
            ASPxTextBox txt = pageControl.FindControl(control) as ASPxTextBox;
            if (txt == null)
            {
                return "";
            }
            return txt.Text;
        }
        public string controlTextDate(string control, ASPxGridView grid)
        {
            ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
            ASPxDateEdit txt = pageControl.FindControl(control) as ASPxDateEdit;
            if (txt == null)
            {
                return "";
            }
            
            return txt.Date.ToString("yyyy-MM-dd");
        }
     

        public DateTime controlKoniecZawieszeniafromWUC( ASPxGridView grid)
        {
            ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
            zawieszenia txt = pageControl.FindControl("zawieszenia1") as zawieszenia;
            return txt.koniec.Date;
        }
        
        public string controlTextMemo(string control, ASPxGridView grid)
        {
            ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
            ASPxMemo txt = pageControl.FindControl(control) as ASPxMemo;
            if (txt == null)
            {
                return string.Empty ;
            }
            return txt.Text ;
        }

        public void usunTworzonaOsobe(int idOsoby)
        {
            DataTable parametry = _cm.makeParameterTable();
            parametry.Rows.Add("@idosoby", idOsoby);
            _cm.runQuerry("delete from tbl_osoby where ident=@idOsoby", parametry);
            _cm.runQuerry("delete from tbl_osoby where id_osoby=@idOsoby", parametry);
        }

        public void controlTextVisibility(string control, ASPxGridView grid,bool visible)
        {
            ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
            ASPxTextBox txt = pageControl.FindControl(control) as ASPxTextBox;
            if (txt != null)
            {
                txt.Visible = visible;
            }
        }
        public void controlVisibility(string control, ASPxGridView grid)
        {
            ASPxPageControl pageControl = grid.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
           
            HtmlControl  txt = pageControl.FindControl(control) as HtmlControl;
            if (txt != null)
            {
                
                txt.Style.Clear();
                txt.Style.Add("display", "block");
                txt.Style.Add(HtmlTextWriterStyle.Display, "block;");
            }
        }
    }
}
