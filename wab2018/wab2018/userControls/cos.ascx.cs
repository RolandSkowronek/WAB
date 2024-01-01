using System;
using System.Web.UI;
using DevExpress.Web;

namespace wab2018
{
    public partial class cos : System.Web.UI.UserControl
    {
        Class2 cl = new Class2();
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string[] names = Enum.GetNames(typeof(GridViewEditingMode));

            }
        }

        protected void ASPxGridView1_RowInserted(object sender, DevExpress.Web.Data.ASPxDataInsertedEventArgs e)
        {

            var a= e.AffectedRecords;
           
        }

        protected void ASPxGridView1_InitNewRow(object sender, DevExpress.Web.Data.ASPxDataInitNewRowEventArgs e)
        {
            string biegly = (string)Session["idMediatora"];
            e.NewValues["idBieglego"] = biegly;
        
          
            e.NewValues["czyus"] = "0";

            string idBieglego = (string)Session["id_osoby"];
            string numer = cl.podajNumerSkargiwRoku(idBieglego, DateTime.Now.Year);
            e.NewValues["rok"] = DateTime.Now.Year;
            e.NewValues["numer"] = numer;
          
        }

        protected void ASPxGridView1_CustomCallback(object sender, ASPxGridViewCustomCallbackEventArgs e)
        {
            int index = 0;
           
                ASPxGridView1.SettingsEditing.Mode = (GridViewEditingMode)index;
        }

      
        protected void ASPxGridView1_CellEditorInitialize(object sender, ASPxGridViewEditorEventArgs e)
        {
            e.Editor.ReadOnly = false;
            
            if (e.Column.FieldName == "numer")
            {
                e.Editor.Value = cl.podajNumerSkargiwRoku( DateTime.Now.Year); ;
            }
            if (e.Column.FieldName == "rok")
            {
                e.Editor.Value = DateTime.Now.Year;

            }
            if (e.Column.FieldName == "dataWplywu")
            {
                e.Editor.Value = DateTime.Now.Date;

            }
            if (e.Column.FieldName == "dataPisma")
            {
                e.Editor.Value = DateTime.Now.Date;

            }
           
        }

        protected void ASPxGridView1_RowUpdating(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {
           
            var ident = e.NewValues["ident"];
            var zakreslenie = e.NewValues["zakreslono"];
            if (zakreslenie == null)
            {
                e.NewValues["zakreslono"] = false;
            }
        }

       

        protected void ASPxGridView1_RowInserting(object sender, DevExpress.Web.Data.ASPxDataInsertingEventArgs e)
        {
            string biegly = (string)Session["idMediatora"];
         
            e.NewValues["idBieglego"] = biegly;
        }

       
    }
}