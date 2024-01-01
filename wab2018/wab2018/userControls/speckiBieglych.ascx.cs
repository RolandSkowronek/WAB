using System;
using System.Collections.Generic;
using System.Linq;
using DevExpress.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace wab2018.userControls
{
    public partial class speckiBieglych : System.Web.UI.UserControl
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void ASPxGridView1_RowInserting(object sender, DevExpress.Web.Data.ASPxDataInsertingEventArgs e)
        {
            string biegly = (string)Session["idMediatora"];

            e.NewValues["id_bieglego"] = biegly;
        }



        protected void ASPxGridView1_RowUpdating(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {
            string biegly = (string)Session["idMediatora"];

            e.NewValues["id_bieglego"] = biegly;
            ASPxPageControl pageControl = ASPxGridView1.FindEditFormTemplateControl("ASPxPageControl1") as ASPxPageControl;
            // ASPxComboBox combobox = pageControl.FindControl("ASPxComboBox1") as ASPxComboBox;
            ASPxComboBox combobox = ASPxGridView1.FindEditRowCellTemplateControl((GridViewDataColumn)ASPxGridView1.Columns["nazwa"], "ASPxComboBox1") as ASPxComboBox;
            var id = e.OldValues["id_"];
            try
            {
                var idspecjalizacji = combobox.SelectedItem.Value;
                e.NewValues["id_specjalizacji"] = idspecjalizacji;
                Int64 te = (Int64)Session["key"];
                e.NewValues["id"] = te;
                e.NewValues["id_"] = te;
            }
            catch (Exception ex)
            { }

            Session["flagaSkarg"] = 1;
        }


        protected void ASPxGridView1_StartRowEditing(object sender, DevExpress.Web.Data.ASPxStartRowEditingEventArgs e)
        {
            Session["key"] = e.EditingKeyValue;
            int rowIndex = ASPxGridView1.FindVisibleIndexByKeyValue(e.EditingKeyValue);
            object  row = ASPxGridView1.GetRow(rowIndex);
            System.Data.DataRowView nazwa = (System.Data.DataRowView)row;
            ASPxComboBox combobox = ASPxGridView1.FindEditRowCellTemplateControl((GridViewDataColumn)ASPxGridView1.Columns["Nazwa specjalizacji"], "ASPxComboBox1") as ASPxComboBox;
           int index= combobox.Items.IndexOfValue(nazwa[3].ToString());
           combobox.SelectedIndex = index;
        }

        protected void ASPxGridView1_CellEditorInitialize(object sender, ASPxGridViewEditorEventArgs e)
        {
            object obj = e.Value;
            System.Data.DataRowView nazwa = (System.Data.DataRowView)obj;
            string aaa = nazwa[3].ToString();

            ASPxComboBox cmb = e.Editor as ASPxComboBox;
            cmb.SelectedIndex = 0;
        }

        protected void ASPxGridView1_CustomCallback(object sender, ASPxGridViewCustomCallbackEventArgs e)
        {
            string cos = e.ToString(); 
        }

        protected void initNewRow(object sender, DevExpress.Web.Data.ASPxDataInitNewRowEventArgs e)
        {
            ASPxComboBox combobox = ASPxGridView1.FindEditRowCellTemplateControl((GridViewDataColumn)ASPxGridView1.Columns["Nazwa specjalizacji"], "ASPxComboBox1") as ASPxComboBox;
            combobox.Visible = true;
            ASPxTextBox ASPxTextBox1 = ASPxGridView1.FindEditRowCellTemplateControl((GridViewDataColumn)ASPxGridView1.Columns["Nazwa specjalizacji"], "ASPxTextBox1") as ASPxTextBox;
            ASPxTextBox1.Visible = false;
            //ASPxTextBox1
        }
    }

      
}