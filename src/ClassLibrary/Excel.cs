using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;


namespace ClassLibrary
{
    public class Excel
    {
        readonly static string[] bansArray = new[] { "ЗапО", "ЗапЗ" };
        readonly static string[] commandsArray = new[] { "Закр", "Откр", "Вкл", "Откл" };        
        enum TypeArmature
        {
            BansNotExists,
            CommandsNotExist, 
            BansAndCommandsExistInFirstField, 
            CommandsExistInSecondField,
            UnidentifiedType
        }
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
                    Console.WriteLine(typeArmature(armature));
                }
            }
        }

        private static TypeArmature typeArmature(Armature armature)
        {
            var bansInFirstField = false;
            var commandsInFirstField = false;
            var commandsInSecondField = false;

            foreach (string value in armature.values)
            {
                var firstFieldValue = value.Split('/')[0];
                if (value.Split('/').Length == 1) 
                {                    
                    if (bansArray.Contains(firstFieldValue)) bansInFirstField = true;
                    else if (commandsArray.Contains(firstFieldValue)) commandsInFirstField = true;
                    else throw new Exception("Неопознанный запрет или команда: " + firstFieldValue);
                }
                else if (value.Split('/').Length > 1)
                {
                    var secondFieldValue = value.Split('/')[1];
                    if (bansArray.Contains(firstFieldValue)) bansInFirstField = true;
                    if (commandsArray.Contains(secondFieldValue)) commandsInSecondField = true;
                    else if (secondFieldValue != "Руч") throw new Exception("Неопознанный запрет или команда: " + secondFieldValue);
                }                
            }

            if (!bansInFirstField) return TypeArmature.BansNotExists;
            else if (!commandsInFirstField && !commandsInSecondField) return TypeArmature.CommandsNotExist;
            else if (bansInFirstField && commandsInFirstField) return TypeArmature.BansAndCommandsExistInFirstField;
            else if (bansInFirstField && commandsInSecondField) return TypeArmature.CommandsExistInSecondField;
            else return TypeArmature.UnidentifiedType;
        }
    }
}
