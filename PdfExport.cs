using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq; 

namespace TelevendFilter
{
    class PdfExport
    {
        static public void SaveBigAudit(string pathDB, string pathSaveFile, IEnumerable<ListDisplay.ListItem> data, 
                                        DateTime? dateFrom, DateTime? dateTo, string sumPaid)
        {
            Document document = new Document();
            document.Info.Title = "Duży audyt"; 
           
            Section section = document.AddSection();
            section.PageSetup.LeftMargin = "1cm";

            Paragraph paragraph;
            Table table;
            Column column;
            Row row;

            paragraph = section.AddParagraph("Wyciąg z audytu wygenerowany " + DateTime.Now);
            paragraph.Format.Font.Size = 6;

            if (dateFrom == null)
            {
                paragraph = section.AddParagraph("Data od: " + DateTime.Parse(data.First().ItemDate).ToShortDateString());
                paragraph.Format.Font.Size = 6;
            }
            else
            {
                paragraph = section.AddParagraph("Data od: " + dateFrom.Value.ToShortDateString());
                paragraph.Format.Font.Size = 6;
            }

            if (dateTo == null)
            {
                paragraph = section.AddParagraph("Data do: " + DateTime.Parse(data.Last().ItemDate).ToShortDateString());
                paragraph.Format.Font.Size = 6;
            }
            else
            {
                paragraph = section.AddParagraph("Data do: " + dateTo.Value.ToShortDateString());
                paragraph.Format.Font.Size = 6;
            }
            paragraph = section.AddParagraph("\n\n");

            //Table section
            table = section.AddTable(); 
            table.Style = "Table";
            table.Borders.Color = Color.Parse("Black"); 
            table.Borders.Width = 0.25; 
            table.Borders.Left.Width = 0.5; 
            table.Borders.Right.Width = 0.5;
            table.Format.Alignment = ParagraphAlignment.Left; 
            table.Rows.LeftIndent = 0;
             
            column = table.AddColumn("2,5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5;
            column = table.AddColumn("2,5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5;
            column = table.AddColumn("2,5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5;
            column = table.AddColumn("3cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5;

            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Format.Font.Size = 6;
            row.Cells[0].AddParagraph("Kod pracownika");
            row.Cells[1].AddParagraph("ID");
            row.Cells[2].AddParagraph("Televend ID");
            row.Cells[3].AddParagraph("Data");
            row.Cells[4].AddParagraph("Cena");
            row.Cells[5].AddParagraph("Produkt");
            row.Cells[6].AddParagraph("Maszyna");

            for (int i = 0; i < data.Count(); i++)
            {
                row = table.AddRow();
                row.Cells[0].AddParagraph(data.ToList()[i].ItemWorkerID);
                row.Cells[2].AddParagraph(data.ToList()[i].ItemID);
                row.Cells[1].AddParagraph(data.ToList()[i].ItemStickerID);
                row.Cells[3].AddParagraph(data.ToList()[i].ItemDate);
                row.Cells[4].AddParagraph(data.ToList()[i].ItemPurchase);
                row.Cells[5].AddParagraph(data.ToList()[i].ItemProduct);
                row.Cells[6].AddParagraph(data.ToList()[i].ItemMachine); 
            }

            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Format.Font.Size = 7;
            row.Cells[0].AddParagraph("Suma");
            row.Cells[4].AddParagraph(sumPaid); 
             
            //Final PDF generation
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true, PdfSharp.Pdf.PdfFontEmbedding.Always); 
            renderer.Document = document;
            renderer.RenderDocument();

            if (!pathSaveFile.EndsWith(".pdf"))
            {
                pathSaveFile += ".pdf";
            }

            renderer.PdfDocument.Save(pathSaveFile); 
            Process.Start(pathSaveFile);
        }

        static public void SaveSmallAudit(string path, string pathSaveFile, IEnumerable<ListDisplay.ListItem> data,
                                          DateTime? dateFrom, DateTime? dateTo, string sumPaid)
        {
            Document document = new Document();
            document.Info.Title = "Mały audyt";

            Section section = document.AddSection();
            section.PageSetup.LeftMargin = "1cm";
            Paragraph paragraph;
            Table table;
            Column column;
            Row row;

            paragraph = section.AddParagraph("Wyciąg z audytu wygenerowany " + DateTime.Now);
            paragraph.Format.Font.Size = 6;

            if (dateFrom == null)
            {
                paragraph = section.AddParagraph("Data od: " + DateTime.Parse(data.First().ItemDate).ToShortDateString());
                paragraph.Format.Font.Size = 6;
            }
            else
            {
                paragraph = section.AddParagraph("Data od: " + dateFrom.Value.ToShortDateString());
                paragraph.Format.Font.Size = 6;
            }

            if (dateTo == null)
            {
                paragraph = section.AddParagraph("Data do: " + DateTime.Parse(data.Last().ItemDate).ToShortDateString());
                paragraph.Format.Font.Size = 6;
            }
            else
            {
                paragraph = section.AddParagraph("Data do: " + dateTo.Value.ToShortDateString());
                paragraph.Format.Font.Size = 6;
            }

            paragraph = section.AddParagraph("\n\n");

            //Table section
            table = section.AddTable();
            table.Style = "Table";
            table.Borders.Color = Color.Parse("Black");
            table.Borders.Width = 0.25;
            table.Borders.Left.Width = 0.5;
            table.Borders.Right.Width = 0.5;
            table.Format.Alignment = ParagraphAlignment.Center;
            table.Rows.LeftIndent = 0;

            column = table.AddColumn("2,5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5;
            column = table.AddColumn("2,5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5;
            column = table.AddColumn("2,5cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5;
            column = table.AddColumn("2cm");
            column.Format.Alignment = ParagraphAlignment.Center;
            column.Format.Font.Size = 5; 

            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Format.Font.Size = 6;
            row.Cells[0].AddParagraph("Kod pracownika");
            row.Cells[1].AddParagraph("ID");
            row.Cells[2].AddParagraph("Televend ID");
            row.Cells[3].AddParagraph("Suma zakupów");                   

            List<string> IDs = DatabaseHandling.GetCardIDs(path);

            int i;
            IEnumerable<ListDisplay.ListItem> currentIDItem;
            for (i = 0; i < IDs.Count(); i++)
            {
                currentIDItem = (data as IEnumerable<ListDisplay.ListItem>).Where(item => item.ItemID == IDs[i]);
                if(currentIDItem.FirstOrDefault().ItemID == null)
                {
                    continue;
                }
                row = table.AddRow();
                row.Cells[0].AddParagraph(currentIDItem.FirstOrDefault().ItemWorkerID);
                row.Cells[2].AddParagraph(IDs[i]);
                row.Cells[1].AddParagraph(currentIDItem.FirstOrDefault().ItemStickerID);
                row.Cells[3].AddParagraph(currentIDItem.Sum(item => Convert.ToDecimal(item.ItemPurchase)).ToString()); 
            }

            row = table.AddRow();
            row.Format.Font.Bold = true;
            row.Format.Font.Size = 7;
            row.Cells[0].AddParagraph("Suma");
            row.Cells[3].AddParagraph(sumPaid); 

            //Final PDF generation
            PdfDocumentRenderer renderer = new PdfDocumentRenderer(true, PdfSharp.Pdf.PdfFontEmbedding.Always);
            renderer.Document = document;
            renderer.RenderDocument();

            if (!pathSaveFile.EndsWith(".pdf"))
            {
                pathSaveFile += ".pdf";
            }

            renderer.PdfDocument.Save(pathSaveFile);
            Process.Start(pathSaveFile);
        }
    }
}
