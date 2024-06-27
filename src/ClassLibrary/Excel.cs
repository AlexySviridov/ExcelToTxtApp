using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;


namespace ClassLibrary
{
    public class Excel
    {
        readonly string[] commandsArray = new[] { "Закр", "Откр", "Вкл", "Откл" };
        readonly string[] bansArray = new[] { "ЗапО", "ЗапЗ" };
        public static void Read(string pathToExcel, int numberWorksheets, int[] ignoredRowsArray, int firstArmatureRow, int firstArmatureColumn)
        {
            FileInfo existingFile = new FileInfo(pathToExcel);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[numberWorksheets];
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                for (int i = firstArmatureRow; i <= rowCount; i++)
                {
                    if (ignoredRowsArray.Contains(i)) continue;
                    for (int j = firstArmatureColumn + 2; j <= colCount; j++)
                    {

                    }
                }
            }
        }
    }
}
