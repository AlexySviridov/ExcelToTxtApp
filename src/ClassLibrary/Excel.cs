using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;


namespace ClassLibrary
{
    public class Excel
    {
        readonly string[] commandsArray = new[] { "Закр", "Откр", "Вкл", "Откл" };
        readonly string[] bansArray = new[] { "ЗапО", "ЗапЗ" };
        public static void Read(string pathToExcel, int numberWorksheets, int[] ignoredRowsArray, int firstArmatureRow, int ArmatureNameColumn, int firstAlgorithmColumn)
        {
            FileInfo existingFile = new FileInfo(pathToExcel);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[numberWorksheets];
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                for (int i = firstArmatureRow; i <= rowCount; i++)
                {
                    Armature armature = new Armature(worksheet.Cells[i, ArmatureNameColumn].Value.ToString().Trim(), i, new List<string> { });
                    if (ignoredRowsArray.Contains(i)) continue;
                    for (int j = firstAlgorithmColumn; j <= colCount; j++)
                    {
                        if (worksheet.Cells[i, j].Value == null) continue;
                        armature.values.Add(worksheet.Cells[i, j].Value.ToString().Trim());
                    }
                    //foreach (string value in armature.values) Console.Write(value + " ");
                    //Console.WriteLine();
                }
            }
        }
    }
}
