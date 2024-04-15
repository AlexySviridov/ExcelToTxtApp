using OfficeOpenXml;
using System.IO;
using System;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;

namespace Host
{
    internal class Program
    {
        static void Main(string[] args)
        {
            FileInfo existingFile = new FileInfo(@"C:\Users\User\Desktop\ТЗиБ\K6. Info v1.20.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[11];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count

                //for (int row = 1; row <= rowCount; row++)
                //{
                //    for (int col = 1; col <= colCount; col++)
                //    {
                //        var cellValue = worksheet.Cells[row, col].Value;
                //        if (cellValue != null)
                //        {
                //            Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + cellValue.ToString().Trim());
                //        }
                //    }
                //}

                Console.WriteLine(" Row:" + 13 + " column:" + 3 + " Value:" + worksheet.Cells[13, 3].Value.ToString().Trim());
            }
        }
    }
}
