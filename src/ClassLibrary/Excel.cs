using OfficeOpenXml;
using System.IO;


namespace ClassLibrary
{
    public class Excel
    {
        public static void ReadExcel(string pathToExcel, int numberWorksheets)
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
