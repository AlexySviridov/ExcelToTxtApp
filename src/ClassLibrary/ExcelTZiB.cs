﻿using OfficeOpenXml;
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

        public static void DoAllWork (string pathToExcel, int numberWorksheet, int[] ignoredRowsArray, int firstArmatureRow, int ArmatureNameColumn, int firstAlgorithmColumn, 
            string pathDirectoryToSave)
        {
            FileInfo existingFile = new FileInfo(pathToExcel);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[numberWorksheet];
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                CreateAlgorithmList(worksheet, firstAlgorithmColumn, colCount, out List<Algorithm> algorithmList);
                
                for (int i = firstArmatureRow; i <= rowCount; i++)
                {
                    var armatureName = worksheet.Cells[i, ArmatureNameColumn].Value.ToString().Trim();
                    Armature armature = new Armature(armatureName, new List<string> { }, new List<int> { }, pathDirectoryToSave);

                    if (ignoredRowsArray.Contains(i)) continue;
                    for (int j = firstAlgorithmColumn; j <= colCount; j++)
                    {
                        if (worksheet.Cells[i, j].Value == null) continue;
                        armature.values.Add(worksheet.Cells[i, j].Value.ToString().Trim());
                        armature.valuesColumn.Add(j);
                    }
                    CreateB1B2Txt(armature, typeArmature(armature));
                    RecordAlgorithmToTxt(armature, typeArmature(armature), algorithmList, firstAlgorithmColumn);
                }
            }
        }

        private static void CreateAlgorithmList(ExcelWorksheet worksheet, int firstAlgorithmColumn, int colCount, out List<Algorithm> algorithmList)
        {
            algorithmList = new List<Algorithm>();

            for (int j = firstAlgorithmColumn; j <= colCount; j++)
            {
                var algorithmColumn = j;
                var signalBefore = worksheet.Cells[2, j].Value.ToString().Trim();
                var conditionAnimation = worksheet.Cells[3, j].Value.ToString().Trim();
                var mnenonicDiagram = worksheet.Cells[4, j].Value.ToString().Trim();
                var algorithmPosition = worksheet.Cells[5, j].Value.ToString().Trim();
                var algorithmName = worksheet.Cells[6, j].Value.ToString();
                var overlay = worksheet.Cells[7, j].Value.ToString().Trim();
                var outputRelay = worksheet.Cells[8, j].Value.ToString().Trim();

                if (algorithmColumn == firstAlgorithmColumn)
                {
                    algorithmPosition = "";
                    overlay = "";
                    outputRelay = "";
                }

                var string2 = signalBefore + "||" + conditionAnimation + "||" + mnenonicDiagram;
                var string3 = algorithmPosition + "||" + algorithmName;
                var string4 = overlay + "||" + outputRelay + "||";

                Algorithm algorithm = new Algorithm(algorithmColumn, string2, string3, string4);
                algorithmList.Add(algorithm);
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

        private static void CreateB1B2Txt(Armature armature, TypeArmature typeArmature)
        {
            switch (typeArmature)
            {
                case TypeArmature.UnidentifiedType:
                    throw new Exception("Обработать логику данной арматуры (" + armature.name + ") не представляется возможным для текущей версии программы :(");                    
                case TypeArmature.BansNotExists:
                    Txt.CreateTxt(armature.pathB1, "Команда", false);
                    break;
                case TypeArmature.CommandsNotExist:
                    Txt.CreateTxt(armature.pathB1, "Запрет", false);
                    break;
                default:
                    Txt.CreateTxt(armature.pathB1, "Запрет", true);
                    Txt.CreateTxt(armature.pathB2, "Команда", true);
                    break;
            }
        }

        private static void RecordAlgorithmToTxt(Armature armature, TypeArmature typeArmature, List<Algorithm> algorithmsList, int firstAlgorithmColumn)
        {
            for (int i = 0; i <= armature.values.Count() - 1; i++)
            {
                var algorithmColumnNumber = armature.valuesColumn[i] - firstAlgorithmColumn;
                var string2 = algorithmsList[algorithmColumnNumber].string2;
                var string3 = algorithmsList[algorithmColumnNumber].string3;
                var string4 = algorithmsList[algorithmColumnNumber].string4;

                switch (typeArmature)
                {
                    case TypeArmature.UnidentifiedType:
                        throw new Exception("Обработать логику данной арматуры (" + armature.name + ") не представляется возможным для текущей версии программы O_o");
                    case TypeArmature.BansNotExists:
                        Txt.WriteTxt(armature.pathB1, string2, string3, string4 + armature.values[i]);
                        break;
                    case TypeArmature.CommandsNotExist:
                        Txt.WriteTxt(armature.pathB1, string2, string3, string4 + armature.values[i]);
                        break;
                    case TypeArmature.CommandsExistInSecondField:
                        Txt.WriteTxt(armature.pathB1, string2, string3, string4 + armature.values[i].Split('/')[0]);
                        if (armature.values[i].Split('/').Length > 1) Txt.WriteTxt(armature.pathB2, string2, string3, string4 + armature.values[i].Split('/')[1]);
                        break;
                    case TypeArmature.BansAndCommandsExistInFirstField:
                        if (bansArray.Contains(armature.values[i].Split('/')[0])) 
                            Txt.WriteTxt(armature.pathB1, string2, string3, string4 + armature.values[i].Split('/')[0]);
                        else Txt.WriteTxt(armature.pathB2, string2, string3, string4 + armature.values[i].Split('/')[0]);
                        break;
                }
            }
        }
    }
}
