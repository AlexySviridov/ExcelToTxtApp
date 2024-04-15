using OfficeOpenXml;
using System.IO;
using System;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using ClassLibrary;

namespace Host
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string pathTxt = @"C:\Users\User\Desktop\ТЗиБ\MyTest.txt";

            FileInfo existingFile = new FileInfo(@"C:\Users\User\Desktop\ТЗиБ\K6. Info v1.20.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[11];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count

                Console.WriteLine(" Row:" + 13 + " column:" + 3 + " Value:" + worksheet.Cells[13, 3].Value.ToString().Trim());

                Txt.CreateTxt(pathTxt);
            }
        }
    }
}
