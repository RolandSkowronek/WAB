using System;
using System.Web.UI;
using DevExpress.Web;

namespace wab2018
{
    public partial class skargiOdczyt : System.Web.UI.UserControl
    {
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
            e.NewValues["data_od"] = DateTime.Now;
            e.NewValues["data_do"] = DateTime.Now.AddYears(5);
            string biegly = (string)Session["idMediatora"];
            e.NewValues["id_bieglego"] = biegly;
            e.NewValues["d_kreacji"] = DateTime.Now.Date;
            e.NewValues["kreator"] = (string)Session["user_id"];
            e.NewValues["czyus"] = "0";
          

        }

        protected void ASPxGridView1_CustomCallback(object sender, ASPxGridViewCustomCallbackEventArgs e)
        {
            int index = 0;
           
                ASPxGridView1.SettingsEditing.Mode = (GridViewEditingMode)index;
        }

      
        protected void ASPxGridView1_CellEditorInitialize(object sender, ASPxGridViewEditorEventArgs e)
        {
            e.Editor.ReadOnly = false;
        }

        protected void ASPxGridView1_RowUpdating(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {
           
            var ident = e.NewValues["ident"];
        }

       

        protected void ASPxGridView1_RowInserting(object sender, DevExpress.Web.Data.ASPxDataInsertingEventArgs e)
        {
            string biegly = (string)Session["idMediatora"];
            e.NewValues["id_powolania"] = "1";
            e.NewValues["id_bieglego"] = biegly;
        }

       
    }
}