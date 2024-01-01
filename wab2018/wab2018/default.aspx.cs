using System;
using System.Web.UI.WebControls;

namespace wab2018
{
    public partial class _default : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            Session["sesja"] = null;
            if (!IsPostBack)
            {
                zdecyduj();
            }
        }

        protected void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            zdecyduj();
        }

        protected void zdecyduj()
        {
            int flag = 0;
            // wyświetlaanie po kategorii

            if (CheckBox1.Checked)
            {
                DropDownList2.Enabled = CheckBox1.Checked;
                flag = flag + 1;
                if (DropDownList2.Items.Count == 0)
                {
                    try
                    {
                        DropDownList2.DataBind();
                    }
                    catch (Exception ex)
                    {
                    }
                }
                Session["grupa"] = DropDownList2.SelectedValue.ToString().Trim();
            }
            else
            {
                DropDownList2.Enabled = CheckBox1.Checked;
            }
            if (CheckBox2.Checked)
            {
                flag = flag + 2;
                DropDownList1.Enabled = true;
                try
                {
                    Session["last"] = DropDownList1.SelectedValue.ToString();
                }
                catch
                {
                }

                if (CheckBox1.Checked)
                {
                    specjalizacje.SelectCommand = "SELECT DISTINCT id_, nazwa, grupa FROM glo_specjalizacje WHERE (grupa = @grupa) ORDER BY nazwa";
                    DropDownList1.DataBind();
                }
                else
                {
                    specjalizacje.SelectCommand = "SELECT DISTINCT id_, nazwa, grupa FROM glo_specjalizacje ORDER BY nazwa";
                    DropDownList1.DataBind();
                }
                if (DropDownList1.Items.Count == 0)
                {
                    DropDownList1.DataBind();
                }
                try
                {
                    string txt1 = (string)Session["last"];
                    string ii = DropDownList1.Items.FindByValue((string)Session["last"]).Value;
                    int il = DropDownList1.Items.IndexOf(DropDownList1.Items.FindByValue((string)Session["last"]));
                    DropDownList1.SelectedIndex = DropDownList1.Items.IndexOf(DropDownList1.Items.FindByValue((string)Session["last"]));
                    GridView1.DataBind();
                }
                catch
                {
                }
            }
            else
            {
                DropDownList1.Enabled = false;
            }

            wyswietl(flag);
        } // end of zdecyduj

        protected void wyswietl(int flag)
        {
            try
            {
                switch (flag)
                {
                    case 0:
                        {
                            widok_glowny.SelectCommand = "SELECT DISTINCT nazwisko ,nazwisko + ' ' + imie AS biegly, stan, stan_inne, zwrocone, zwrocone_inne, opinie, przeterminowane, grzywna, data_poczatkowa, data_koncowa FROM obciazenia ORDER BY stan,stan_inne, zwrocone";
                        }
                        break;

                    case 1:
                        {
                            widok_glowny.SelectCommand = "SELECT DISTINCT nazwisko ,nazwisko + ' ' + imie AS biegly, stan, stan_inne, zwrocone, zwrocone_inne, opinie, przeterminowane, grzywna, data_poczatkowa, data_koncowa FROM obciazenia  where grupa=" + DropDownList2.SelectedValue.ToString().Trim() + "  order by stan,stan_inne, zwrocone";
                        }
                        break;

                    case 3:
                        {
                            widok_glowny.SelectCommand = "SELECT DISTINCT nazwisko ,nazwisko + ' ' + imie AS biegly, stan, stan_inne, zwrocone, zwrocone_inne, opinie, przeterminowane, grzywna, data_poczatkowa, data_koncowa FROM obciazenia  where  id_specjalizacji=" + DropDownList1.SelectedValue.ToString().Trim() + " and grupa=" + DropDownList2.SelectedValue.ToString().Trim() + "  order by stan,stan_inne, zwrocone";
                        }
                        break;

                    case 2:
                        {
                            widok_glowny.SelectCommand = "SELECT DISTINCT nazwisko ,nazwisko + ' ' + imie AS biegly, stan, stan_inne, zwrocone, zwrocone_inne, opinie, przeterminowane, grzywna, data_poczatkowa, data_koncowa FROM obciazenia  where  id_specjalizacji=" + DropDownList1.SelectedValue.ToString().Trim() + "  order by stan,stan_inne, zwrocone ;";
                        }
                        break;

                    default:
                        break;
                }

                GridView1.DataBind();
            }
            catch
            {
            }
        }

        protected void DropDownList1_SelectedIndexChanged(object sender, EventArgs e)
        {
            zdecyduj();
        }

        protected void GridView1_Sorted(object sender, EventArgs e)
        {
            zdecyduj();
        }

        protected void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            zdecyduj();
        }

        protected void DropDownList2_SelectedIndexChanged(object sender, EventArgs e)
        {
            zdecyduj();
        }

        protected void GridView1_RowCreated1(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.Header)
            {
                // najnizszy
                GridView HeaderGrid = (GridView)sender;
                GridViewRow HeaderGridRow = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Insert);
                HeaderGridRow.Font.Size = 7;
                HeaderGridRow.HorizontalAlign = HorizontalAlign.Center;
                HeaderGridRow.VerticalAlign = VerticalAlign.Top;
                //#4b6c9e
                HeaderGridRow.ForeColor = System.Drawing.Color.White;
                string hexValue = "#4b6c9e"; // You do need the hash
                System.Drawing.Color colour = System.Drawing.ColorTranslator.FromHtml(hexValue); // Yippee
                HeaderGridRow.BackColor = colour;
                TableCell HeaderCell = new TableCell();

                HeaderCell = new TableCell();

                HeaderCell.Text = "Sądowe";

                HeaderCell.RowSpan = 1;
                HeaderGridRow.Cells.Add(HeaderCell);

                HeaderCell = new TableCell();
                HeaderCell.Text = "Inne";
                HeaderCell.ColumnSpan = 1;
                HeaderGridRow.Cells.Add(HeaderCell);

                HeaderCell = new TableCell();

                HeaderCell.Text = "Sądowe";

                HeaderCell.RowSpan = 1;
                HeaderGridRow.Cells.Add(HeaderCell);

                HeaderCell = new TableCell();
                HeaderCell.Text = "Inne";
                HeaderCell.ColumnSpan = 1;
                HeaderGridRow.Cells.Add(HeaderCell);

                GridView1.Controls[0].Controls.AddAt(0, HeaderGridRow);

                ////  drugi wiersz

                //W tym      nieusprawiedliwione

                HeaderGrid = (GridView)sender;
                HeaderGridRow = new GridViewRow(0, 0, DataControlRowType.Header, DataControlRowState.Selected);
                HeaderGridRow.HorizontalAlign = HorizontalAlign.Center;
                HeaderGridRow.VerticalAlign = VerticalAlign.Top;
                HeaderGridRow.ForeColor = System.Drawing.Color.White;
                HeaderGridRow.BackColor = colour;
                HeaderGridRow.Font.Size = 10;
                HeaderCell = new TableCell();

                HeaderCell.Text = "Biegły";
                HeaderCell.RowSpan = 2;
                HeaderCell.ColumnSpan = 1;
                HeaderGridRow.Cells.Add(HeaderCell);
                GridView1.Controls[0].Controls.AddAt(0, HeaderGridRow);

                HeaderCell = new TableCell();
                HeaderCell.Text = "Stan akt";
                HeaderCell.ColumnSpan = 2;
                HeaderGridRow.Cells.Add(HeaderCell);
                GridView1.Controls[0].Controls.AddAt(0, HeaderGridRow);

                HeaderCell = new TableCell();
                HeaderCell.Text = "Ilość akt zwróconych";
                HeaderCell.ColumnSpan = 2;
                HeaderGridRow.Cells.Add(HeaderCell);
                GridView1.Controls[0].Controls.AddAt(0, HeaderGridRow);

                HeaderCell = new TableCell();
                HeaderCell.Text = "W tym z opinią";
                HeaderCell.ColumnSpan = 1;
                HeaderCell.RowSpan = 2;
                HeaderGridRow.Cells.Add(HeaderCell);
                GridView1.Controls[0].Controls.AddAt(0, HeaderGridRow);

                HeaderCell = new TableCell();
                HeaderCell.Text = "W tym po terminie";
                HeaderCell.ColumnSpan = 1;
                HeaderCell.RowSpan = 2;
                HeaderGridRow.Cells.Add(HeaderCell);
                GridView1.Controls[0].Controls.AddAt(0, HeaderGridRow);

                HeaderCell = new TableCell();
                HeaderCell.Text = "Ilość grzywien";
                HeaderCell.ColumnSpan = 1;
                HeaderCell.RowSpan = 2;
                HeaderGridRow.Cells.Add(HeaderCell);
                GridView1.Controls[0].Controls.AddAt(0, HeaderGridRow);

                HeaderCell = new TableCell();
                HeaderCell.Text = "Powołany";
                HeaderCell.ColumnSpan = 1;
                HeaderCell.RowSpan = 2;
                HeaderGridRow.Cells.Add(HeaderCell);

                GridView1.Controls[0].Controls.AddAt(0, HeaderGridRow);
            }
        }
    }
}