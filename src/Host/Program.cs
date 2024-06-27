using System.IO;
using System;
using ClassLibrary;
using System.Linq;

namespace Host
{
    internal class Program
    {
        static void Main()
        {
            var pathToExcel = "C:\\Users\\User\\Desktop\\Илья\\K6. Info v1.35.xlsx";
        }

        //static void Main(string[] args)
        //{     
        //    FileInfo existingFile = new FileInfo("C:\\Users\\User\\Desktop\\Илья\\K6. Info v1.35.xlsx");
        //    using (ExcelPackage package = new ExcelPackage(existingFile))
        //    {
        //        ExcelWorksheet worksheet = package.Workbook.Worksheets[12];
        //        int colCount = worksheet.Dimension.End.Column;
        //        int rowCount = worksheet.Dimension.End.Row;

        //        var ignoredRowsArray = new[] { 15, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 87, 88, 90 };
        //        var unnormalRowsArray = new[] { 63, 64, 65, 67, 68, 69, 71, 72, 73, 75, 76, 77, 79, 80, 81, 83, 84, 85};
        //        var commandsArray = new[] { "Закр", "Откр", "Вкл", "Откл"};
        //        var bansArray = new[] { "ЗапО", "ЗапЗ" };

        //        for (int i = 13; i <= rowCount; i++)
        //        {
        //            if (!ignoredRowsArray.Contains(i))
        //            {
        //                for (int j = 5; j <= colCount; j++)
        //                {
        //                    if (worksheet.Cells[i, j].Value != null)
        //                    {
        //                        var cellValue = worksheet.Cells[i, j].Value.ToString().Trim();
        //                        var armatureName = worksheet.Cells[i, 3].Value.ToString().Trim();                                                  
        //                        var pathTxt = "C:\\Users\\User\\Desktop\\ТЗиБ_v2\\" + armatureName + "_B1.db";
        //                        var pathTxtB2 = ("C:\\Users\\User\\Desktop\\ТЗиБ_v2\\" + armatureName + "_B2.db");
        //                        bool b2Exists = B2Exists(cellValue, commandsArray, i, j);

        //                        if (!unnormalRowsArray.Contains(i))
        //                        {
        //                            CreateB1B2TxtFiles(pathTxt, pathTxtB2, cellValue, commandsArray, bansArray, i, j, b2Exists);
        //                        }
        //                        else CreateB1B2TxtFiles2(pathTxt, pathTxtB2);

        //                        var numberPosition = worksheet.Cells[5, j].Value.ToString().Trim();
        //                        var nakladka = worksheet.Cells[7, j].Value.ToString().Trim();
        //                        var outputReley = worksheet.Cells[8, j].Value.ToString().Trim();

        //                        if (j == 5)
        //                        {
        //                            numberPosition = "";
        //                            nakladka = "";
        //                            outputReley = "";
        //                        }

        //                        var string2 = worksheet.Cells[2, j].Value.ToString().Trim() + "||" + worksheet.Cells[3, j].Value.ToString().Trim() +
        //                            "||" + worksheet.Cells[4, j].Value.ToString().Trim();
        //                        var string3 = numberPosition + "||" + worksheet.Cells[6, j].Value.ToString();

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
        //                        }
        //                    }
        //                }
        //            }
        //        }                
        //    }
        //}

        //static private void CreateB1B2TxtFiles2(string pathTxt, string pathTxtB2)
        //{
        //    var B2Exists = true;
        //    if (!File.Exists(pathTxt))
        //    {
        //        Txt.CreateTxt(pathTxt, "Запрет", B2Exists);
        //    }
        //    if (!File.Exists(pathTxtB2))
        //    {
        //        Txt.CreateTxt(pathTxtB2, "Команда", B2Exists);
        //    }
        //}

        //static private void CreateB1B2TxtFiles(string pathTxt, string pathTxtB2, string cellValue, string[] commandsArray, string[] bansArray, int i, int j, bool B2Exists)
        //{
        //    if (!File.Exists(pathTxt))
        //    {
        //        Txt.CreateTxt(pathTxt, TypeBLCAP(cellValue, commandsArray, bansArray, i, j), B2Exists);
        //    }
            
        //    if (B2Exists)
        //    {
        //        if (!File.Exists(pathTxtB2))
        //        {
        //            Txt.CreateTxt(pathTxtB2, "Команда", B2Exists);
        //        }
        //    }
        //}

        //static private bool B2Exists(string cellValue, string[] commandsArray, int i, int j)
        //{
        //    if (cellValue.Split('/').Length > 1 && commandsArray.Contains(cellValue.Split('/')[1]))
        //    {
        //        return true;
        //    }
        //    else if (cellValue.Split('/').Length > 1 && cellValue.Split('/')[1] != "Руч")
        //    {
        //        throw new Exception("Неопознанная команда или запрет: " + cellValue.Split('/')[1] + " в ячейке по адресу - строка " + i + " столбец " + j);
        //    }
        //    else return false;
        //}

        //static private string TypeBLCAP(string cellValue, string[] commandsArray, string[] bansArray, int row, int column)
        //{
        //    if (bansArray.Contains(cellValue.Split('/')[0]))
        //    {
        //        return "Запрет";
        //    }
        //    else if (commandsArray.Contains(cellValue.Split('/')[0]))
        //    {
        //        return "Команда";
        //    }
        //    else throw new Exception("Неопознанная команда или запрет: " + cellValue.Split('/')[0] + " в ячейке по адресу - строка " + row + " столбец " + column);            
        //}
    }
}
