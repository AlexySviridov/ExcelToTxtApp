using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;


namespace ClassLibrary
{
    public class ExcelTZiB
    {
        readonly static string[] bansArray = new[] { "ЗапО", "ЗапЗ" };
        readonly static string[] commandsArray = new[] { "Закр", "Откр", "Вкл", "Откл" };        

        public static void ReadDbTZiB (string pathToExcel, int numberWorksheet, int[] ignoredRowsArray, int firstArmatureRow, int ArmatureNameColumn, int firstAlgorithmColumn, 
            string pathDirectoryToSave)
        {
            FileInfo existingFile = new FileInfo(pathToExcel);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[numberWorksheet];
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                for (int i = firstArmatureRow; i <= rowCount; i++)
                {
                    Armature armature = new Armature(worksheet.Cells[i, ArmatureNameColumn].Value.ToString().Trim(), new List<string> { }, new List<int> { });
                    if (ignoredRowsArray.Contains(i)) continue;
                    for (int j = firstAlgorithmColumn; j <= colCount; j++)
                    {
                        if (worksheet.Cells[i, j].Value == null) continue;
                        armature.values.Add(worksheet.Cells[i, j].Value.ToString().Trim());
                        armature.valuesColumn.Add(j);
                    }
                    CreateB1B2(armature.name, typeArmature(armature), pathDirectoryToSave);
                    RecordAlgorithm(armature, typeArmature(armature), pathDirectoryToSave, worksheet);
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

        private static void CreateB1B2(string armatureName, TypeArmature typeArmature, string pathDirectoryToSave)
        {
            var pathB1 = Path.Combine(pathDirectoryToSave + armatureName + "_B1.db");
            var pathB2 = Path.Combine(pathDirectoryToSave + armatureName + "_B2.db");

            switch (typeArmature)
            {
                case TypeArmature.UnidentifiedType:
                    throw new Exception("Обработать логику данной арматуры (" + armatureName + ") не представляется возможным для текущей версии программы :(");                    
                case TypeArmature.BansNotExists:
                    Txt.CreateTxt(pathB1, "Команда", false);
                    break;
                case TypeArmature.CommandsNotExist:
                    Txt.CreateTxt(pathB1, "Запрет", false);
                    break;
                default:
                    Txt.CreateTxt(pathB1, "Запрет", true);
                    Txt.CreateTxt(pathB2, "Команда", true);
                    break;
            }
        }

        private static void RecordAlgorithm(Armature armature, TypeArmature typeArmature, string pathDirectoryToSave, ExcelWorksheet worksheet)
        {
            var pathB1 = Path.Combine(pathDirectoryToSave + armature.name+ "_B1.db");
            var pathB2 = Path.Combine(pathDirectoryToSave + armature.name + "_B2.db");

            for (int i = 0; i <= armature.values.Count() - 1; i++)
            {
                var algorithmColumn = armature.valuesColumn[i];
                var string2 = worksheet.Cells[2, algorithmColumn].Value.ToString().Trim() + "||" + worksheet.Cells[3, algorithmColumn].Value.ToString().Trim() +
                                    "||" + worksheet.Cells[4, algorithmColumn].Value.ToString().Trim();
                var string3 = worksheet.Cells[5, algorithmColumn].Value.ToString().Trim() + "||" + worksheet.Cells[6, algorithmColumn].Value.ToString();
                var string4 = worksheet.Cells[7, algorithmColumn].Value.ToString().Trim() + "||" + worksheet.Cells[8, algorithmColumn].Value.ToString().Trim() + "||";

                switch (typeArmature)
                {
                    case TypeArmature.UnidentifiedType:
                        throw new Exception("Обработать логику данной арматуры (" + armature.name + ") не представляется возможным для текущей версии программы :(");
                    case TypeArmature.BansNotExists:
                        string4 += armature.values[i];
                        Txt.WriteTxt(pathB1, string2, string3, string4);
                        break;
                    case TypeArmature.CommandsNotExist:
                        string4 += armature.values[i];
                        Txt.WriteTxt(pathB1, string2, string3, string4);
                        break;
                    case TypeArmature.CommandsExistInSecondField:

                        break;
                    default:
                        //Txt.CreateTxt(pathB1, "Запрет", true);
                        //Txt.CreateTxt(pathB2, "Команда", true);
                        break;
                }
            }

            
            

            //                        if (!unnormalRowsArray.Contains(i))
            //                        {
            //                            var string4 = nakladka + "||" + outputReley + "||" + cellValue.Split('/')[0];
            //                            Txt.WriteTxt(pathTxt, string2, string3, string4);
            //                            if (b2Exists)
            //                            {
            //                                var string4B2 = nakladka + "||" + outputReley + "||" + cellValue.Split('/')[1];
            //                                Txt.WriteTxt(pathTxtB2, string2, string3, string4B2);
            //                            }
            //                        }
            //                        else
            //                        {
            //                            var string4 = nakladka + "||" + outputReley + "||" + cellValue.Split('/')[0];
            //                            if (bansArray.Contains(cellValue.Split('/')[0]))
            //                            {                                        
            //                                Txt.WriteTxt(pathTxt, string2, string3, string4);
            //                            }
            //                            else if (commandsArray.Contains(cellValue.Split('/')[0]))
            //                            {
            //                                Txt.WriteTxt(pathTxtB2, string2, string3, string4);
            //                            }
            //                            else throw new Exception("Неопознанная команда или запрет: " + cellValue.Split('/')[0] + " в ячейке по адресу - строка " + i + " столбец " + j);
        }
    }
}
