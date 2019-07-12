using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq; 

namespace TelevendFilter
{
    class ExcelExport
    {
        static public void SaveBigAudit(string pathDB, string pathSaveFile, IEnumerable<ListDisplay.ListItem> data,
                                        DateTime? dateFrom, DateTime? dateTo, string sumPaid)
        { 
            XSSFWorkbook xlWorkBook = new XSSFWorkbook();
            ISheet sheet = xlWorkBook.CreateSheet("1");
            IRow firstRow, row;

            firstRow = sheet.CreateRow(0);
            row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue("Data od");
            row.CreateCell(2).SetCellValue("Data do"); 

            if (dateTo == null)
            {
                row.CreateCell(3).SetCellValue(DateTime.Parse(data.Last().ItemDate).ToShortDateString());
            }
            else
            {
                row.CreateCell(3).SetCellValue(dateTo.Value.ToShortDateString());
            }

            if (dateFrom == null)
            {
                row.CreateCell(1).SetCellValue(DateTime.Parse(data.First().ItemDate).ToShortDateString());
            }
            else
            {
                row.CreateCell(1).SetCellValue(dateFrom.Value.ToShortDateString());
            }

            row = sheet.CreateRow(2);
            row.CreateCell(0).SetCellValue("Kod pracownika");
            row.CreateCell(1).SetCellValue("ID");
            row.CreateCell(2).SetCellValue("Televend ID");
            row.CreateCell(3).SetCellValue("Data");
            row.CreateCell(4).SetCellValue("Cena");
            row.CreateCell(5).SetCellValue("Produkt");
            row.CreateCell(6).SetCellValue("Maszyna");

            int i;
            for (i = 0; i < data.Count(); i++)
            {
                row = sheet.CreateRow(i + 3);
                row.CreateCell(0).SetCellValue(data.ToList()[i].ItemWorkerID);
                row.CreateCell(2).SetCellValue(data.ToList()[i].ItemID);
                row.CreateCell(1).SetCellValue(data.ToList()[i].ItemStickerID);
                row.CreateCell(3).SetCellValue(data.ToList()[i].ItemDate);
                row.CreateCell(4).SetCellValue(data.ToList()[i].ItemPurchase);
                row.CreateCell(5).SetCellValue(data.ToList()[i].ItemProduct);
                row.CreateCell(6).SetCellValue(data.ToList()[i].ItemMachine); 
            }

            row = sheet.CreateRow(i + 3);
            row.CreateCell(0).SetCellValue("Suma");
            row.CreateCell(4).SetCellValue(sumPaid);  

            for(int j = 0; j < 7; j++)
            {
                sheet.AutoSizeColumn(j);
            }

            //Add after auto sizing to avoid enormous first column
            firstRow.CreateCell(0).SetCellValue("Wyciąg z audytu wygenerowany " + DateTime.Now);

            if (!pathSaveFile.EndsWith(".xlsx"))
            {
                pathSaveFile += ".xlsx";
            }

            using (var fs = new FileStream(pathSaveFile, FileMode.Create, FileAccess.Write))
            {
                xlWorkBook.Write(fs);
            }
        }

        static public void SaveSmallAudit(string path, string pathSaveFile, IEnumerable<ListDisplay.ListItem> data,
                                          DateTime? dateFrom, DateTime? dateTo, string sumPaid)
        {
            XSSFWorkbook xlWorkBook = new XSSFWorkbook();
            ISheet sheet = xlWorkBook.CreateSheet("1");
            IRow firstRow, row;

            firstRow = sheet.CreateRow(0);
            row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue("Data od");
            row.CreateCell(2).SetCellValue("Data do");

            if (dateTo == null)
            {
                row.CreateCell(3).SetCellValue(DateTime.Parse(data.Last().ItemDate).ToShortDateString());
            }
            else
            {
                row.CreateCell(3).SetCellValue(dateTo.Value.ToShortDateString());
            }

            if (dateFrom == null)
            {
                row.CreateCell(1).SetCellValue(DateTime.Parse(data.First().ItemDate).ToShortDateString());
            }
            else
            {
                row.CreateCell(1).SetCellValue(dateFrom.Value.ToShortDateString());
            }

            row = sheet.CreateRow(2);
            row.CreateCell(0).SetCellValue("Kod pracownika");
            row.CreateCell(1).SetCellValue("ID");
            row.CreateCell(2).SetCellValue("Televend ID"); 
            row.CreateCell(3).SetCellValue("Suma zakupów"); 

            List<string> IDs = DatabaseHandling.GetCardIDs(path);

            int i;
            int rowCounter = 3;
            IEnumerable<ListDisplay.ListItem> currentIDItem;
            for (i = 0; i < IDs.Count(); i++)
            {
                currentIDItem = (data as IEnumerable<ListDisplay.ListItem>).Where(item => item.ItemID == IDs[i]);
                if (currentIDItem.FirstOrDefault().ItemID == null)
                {
                    continue;
                }

                string currentSum = currentIDItem.Sum(item => Convert.ToDecimal(item.ItemPurchase)).ToString();

                if (currentSum == "0")
                {
                    continue;
                }

                row = sheet.CreateRow(rowCounter);
                row.CreateCell(0).SetCellValue(currentIDItem.FirstOrDefault().ItemWorkerID);
                row.CreateCell(2).SetCellValue(IDs[i]);
                row.CreateCell(1).SetCellValue(currentIDItem.FirstOrDefault().ItemStickerID); 
                row.CreateCell(3).SetCellValue(currentSum); 
                rowCounter++;
            }
            row = sheet.CreateRow(i + 3);
            row.CreateCell(0).SetCellValue("Suma");
            row.CreateCell(3).SetCellValue(sumPaid); 

            for (int j = 0; j < 4; j++)
            {
                sheet.AutoSizeColumn(j);
            }

            //Add after auto sizing to avoid enormous first column
            firstRow.CreateCell(0).SetCellValue("Wyciąg z audytu wygenerowany " + DateTime.Now);

            if (!pathSaveFile.EndsWith(".xlsx"))
            {
                pathSaveFile += ".xlsx";
            }

            using (var fs = new FileStream(pathSaveFile, FileMode.Create, FileAccess.Write))
            {
                xlWorkBook.Write(fs);
            }
        }
    }
}
