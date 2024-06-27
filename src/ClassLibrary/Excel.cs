using OfficeOpenXml;
using System.IO;


namespace ClassLibrary
{
    public class Excel
    {
        readonly string[] commandsArray = new[] { "Закр", "Откр", "Вкл", "Откл" };
        readonly string[] bansArray = new[] { "ЗапО", "ЗапЗ" };
        public static void Read(string pathToExcel, int numberWorksheets, int[] ignoredRowsArray)
        {
            FileInfo existingFile = new FileInfo(pathToExcel);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[numberWorksheets];
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;
            }
        }
    }
}
