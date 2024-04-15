using OfficeOpenXml;
using System.IO;
using System;
using ClassLibrary;

namespace Host
{
    internal class Program
    {
        static void Main(string[] args)
        {     
            FileInfo existingFile = new FileInfo("C:\\Users\\User\\Desktop\\ТЗиБ\\K6. Info v1.20.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[11];
                Console.WriteLine("Value:" + worksheet.Cells[46, 3].Value.ToString().Trim() + "\n");
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                for (int i = 5; i < colCount; i++)
                {
                    var armatureName = worksheet.Cells[13, 3].Value.ToString().Trim();
                    var cellValue = worksheet.Cells[13, i].Value;
                    if (cellValue != null)
                    {
                        if (cellValue.ToString().Split('/')[0].Trim() == "ЗапО" ||
                            cellValue.ToString().Split('/')[0].Trim() == "ЗапЗ")
                        {
                            var blcap = "Запрет";
                            Console.WriteLine("true");
                            Txt.WriteTxt("C:\\Users\\User\\Desktop\\ТЗиБ\\" + armatureName + "_B1.db", blcap);
                        }
                    }
                }              
            }
        }
    }
}
