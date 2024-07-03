using DevExpress.Web;
using iTextSharp.text.pdf;
using System;
using System.Data;
using System.IO;
using iTextSharp.text;

using System.Diagnostics;
using System.Net;
namespace wab2018
{
    
    public partial class wykazSkarg : System.Web.UI.Page
    {
        private Class2 cl = new Class2();

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Session["user_id"] == null)
            {
                Server.Transfer("logowanie.aspx");
            }
            else
            {
                string rola = (string)Session["rola"];
                if (rola!="1")
                {
                    ASPxGridView1.Enabled = true;
                }
                else
                {
                    ASPxGridView1.Enabled = false;
                }
            }
        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            //pdf
            DataTable dt = new DataTable();
            foreach (GridViewColumn column in ASPxGridView1.VisibleColumns)
            {
                var col = column as GridViewDataColumn;
                if (col != null)
                    if (col.FieldName!="")
                    {
                        dt.Columns.Add(col.FieldName);
                    }
                    
            }
            for (int i = 0; i < ASPxGridView1.VisibleRowCount; i++)
            {
                DataRow row = dt.NewRow();
                foreach (GridViewColumn column in ASPxGridView1.VisibleColumns)
                {
                    var col = column as GridViewDataColumn;
                    if (col != null)
                    {
                        if (col.FieldName != "")
                        {
                            var cellValue = ASPxGridView1.GetRowValues(i, col.FieldName);
                            row[col.FieldName] = cellValue;
                        }
                    }
                }
                dt.Rows.Add(row);
            }
            // make pdf

            iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document(PageSize.A4.Rotate(), 0f, 0f, 10f, 0f);


            string path = Server.MapPath(@"~//pdf"); //Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments );
            string fileName = path + "//Zestawienie_Skarg_" + DateTime.Now.ToString().Replace(":", "-") + ".pdf";
            PdfWriter writer = PdfWriter.GetInstance(pdfDoc, new FileStream(fileName, FileMode.Create));

         
            pdfDoc.Open();
            
            pdfDoc.AddTitle("Zestawienie Skarg");
            pdfDoc.AddCreationDate();
            PdfPTable table = new PdfPTable(11);
            table.WidthPercentage = 90f;

            PdfPCell cell = new PdfPCell(new Paragraph("Biegły", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            table.AddCell(cell);
            
            cell = new PdfPCell(new Paragraph("Rok", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            table.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Data wpływu", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            table.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Data wpływu pisma", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            table.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Sygnatura", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            table.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Sędzie wizytator", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            table.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Zakreślono", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
       
            table.AddCell(cell);

            cell = new PdfPCell(new Paragraph("Data zakreślenia", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
       
            table.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Rodzaj załatwienia", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
     
            table.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Składający skargę", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
 
            table.AddCell(cell);
            cell = new PdfPCell(new Paragraph("Uwagi", cl.plFont2));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
     
            table.AddCell(cell);

            foreach (DataRow item in dt.Rows)
            {
                var biegły = item[2].ToString();
                var rok = item[1].ToString();
                var DataWplywu  = item[3].ToString();
                var dataWplywuPisma = item[4].ToString();
                var sygnatura = item[5].ToString(); 
                var sedzieWizytator = item[6].ToString();
                var zakreslono = item[7].ToString();
                var dataZakreslenia = item[8].ToString();
                var rodzajZalatwienia = item[9].ToString();
                var skladajacySkarge    = item[10].ToString();
                var uwagi = ""; item[11].ToString();
                table.AddCell(defaultPDFCell(biegły));
                table.AddCell(defaultPDFCell(rok));
                table.AddCell(defaultPDFCell(CutDate(DataWplywu)));
                table.AddCell(defaultPDFCell(CutDate(dataWplywuPisma)));
                table.AddCell(defaultPDFCell(sygnatura));
                table.AddCell(defaultPDFCell(sedzieWizytator));
                if (zakreslono == "True")
                {
                    table.AddCell( new PdfPCell(new Paragraph("zakreślono", cl.fontPL1)));
                }
                else
                {
                    table.AddCell(new PdfPCell(new Paragraph("", cl.fontPL1)));
                }
                
                table.AddCell(defaultPDFCell(CutDate(dataZakreslenia)));
                table.AddCell(defaultPDFCell(rodzajZalatwienia));
                table.AddCell(defaultPDFCell(skladajacySkarge));
                table.AddCell(defaultPDFCell(uwagi));

            }
            pdfDoc.Add(table);
            pdfDoc.Close();

            WebClient client = new WebClient();
            Byte[] buffer = client.DownloadData(fileName);
            if (buffer != null)
            {
                Response.ContentType = "application/pdf";
                Response.AddHeader("content-lenght", buffer.Length.ToString());
                Response.BinaryWrite(buffer);
            }


            ProcessStartInfo startInfo = new ProcessStartInfo(fileName);
            Process.Start(startInfo);
        }
        private string CutDate(string Date)
        {
            if (Date.Length>10)
            {
                return Date.Substring(0, 10);
            }
            return Date;
        }
        private PdfPCell defaultPDFCell(string Tekst)
        {
            PdfPCell cell = new PdfPCell(new Paragraph(Tekst, cl.fontPL1));
            cell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
            return cell;
        }
    }
   
}