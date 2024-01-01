using System;
using System.Data;


namespace wab2018.userControls
{
    public partial class specjalizacjeBiegli : System.Web.UI.UserControl
    {
        private cm Common = new cm();
        protected void Page_Load(object sender, EventArgs e)
        {
            try
            {
                string txt= (string)Session["id_osoby"];
                ASPxGridView1.DataBind();
            }
            catch 
            {      }
           
        }

        protected void ASPxGridView1_RowUpdating(object sender, DevExpress.Web.Data.ASPxDataUpdatingEventArgs e)
        {
           var klucze= e.Keys[0];
           Session ["id"] = e.NewValues["idOsoby"];
            var stan = e.NewValues["stab"];
            bool stanb = (bool)stan;
            string osoba = (string)Session["id_osoby"];
            var specjalizacja = e.NewValues["idSpecjalizacji"];
            DataTable parametry = Common.makeParameterTable();
            parametry.Rows.Add("@idOsoby", osoba.ToString());
           
            parametry.Rows.Add("@idSpecjalizacji", klucze.ToString());
            if (stanb)
            {
                Common.runQuerry(" insert into tbl_specjalizacje_osob (id_osoby,id_specjalizacji) values (@idOsoby,@idSpecjalizacji)", parametry);
            }
            else
            {
                Common.runQuerry(" delete from  tbl_specjalizacje_osob where id_osoby=@idOsoby and  id_specjalizacji=@idSpecjalizacji", parametry);
            }
            
        }
    }
}