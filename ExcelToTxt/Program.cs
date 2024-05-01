using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelToText
{
    class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = @"C:\Users\...\x.xlsx";
            string textFilePath = @"C:\Users\...\x.txt";

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(1); // 2. sayfa
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                SharedStringTablePart sstpart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                using (StreamWriter writer = new StreamWriter(textFilePath))
                {
                    foreach (Row row in sheetData.Elements<Row>())
                    {
                        string cellValue = GetCellValue(row, "A", sstpart);
                        

                        if (!string.IsNullOrWhiteSpace(cellValue)) // Boş satırları atlamak için kontrol
                        {
                            writer.WriteLine("A column value: " +cellValue );
                            
                        }
                    }
                }
            }

            Console.WriteLine("Dosya dönüştürme işlemi tamamlandı.");
        }

        private static string GetCellValue(Row row, string columnName, SharedStringTablePart sstpart)
        {
            Cell cell = row.Elements<Cell>().FirstOrDefault(c => string.Compare(c.CellReference.Value, columnName + row.RowIndex, true) == 0);

            if (cell != null && cell.CellValue != null)
            {
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    int ssid = int.Parse(cell.CellValue.Text);
                    return sstpart.SharedStringTable.ElementAt(ssid).InnerText;
                }
                else
                {
                    return cell.CellValue.Text;
                }
            }
            return string.Empty;
        }
    }
}
