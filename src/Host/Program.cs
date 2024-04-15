using OfficeOpenXml;
using System.IO;
using System;
using ClassLibrary;
using System.Linq;

namespace Host
{
    internal class Program
    {
        static void Main(string[] args)
        {     
            FileInfo existingFile = new FileInfo("C:\\Users\\User\\Desktop\\Илья\\K6. Info v1.20.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[11];
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;
                var ignoreRowArray = new[] { 15, 34, 35, 36, 37, 38, 39, 87, 88, 90, 93, 94, 95 };

                for (int j = 13; j < rowCount; j++)
                {
                    if (!ignoreRowArray.Contains(j))
                    {
                        for (int i = 6; i < colCount; i++)
                        {
                            var excelRow = j;
                            bool B2Exists = false;
                            var armatureName = worksheet.Cells[excelRow, 4].Value.ToString().Trim();
                            var cellValue = worksheet.Cells[excelRow, i].Value;
                            var pathTxt = "C:\\Users\\User\\Desktop\\ТЗиБ\\" + armatureName + "_B1.db";
                            var pathTxtB2 = ("C:\\Users\\User\\Desktop\\ТЗиБ\\" + armatureName + "_B2.db");

                            if (cellValue != null)
                            {
                                if (!File.Exists(pathTxt))
                                {
                                    Txt.CreateTxt(pathTxt, TypeBLCAP(cellValue.ToString().Trim(), j, i));
                                }

                                if (cellValue.ToString().Split('/').Length > 1)
                                {
                                    B2Exists = true;
                                    if (!File.Exists(pathTxtB2))
                                    {
                                        Txt.CreateTxt(pathTxtB2, "Команда");
                                    }
                                }

                                var numberPosition = worksheet.Cells[5, i].Value.ToString().Trim();
                                var nakladka = worksheet.Cells[7, i].Value.ToString().Trim();
                                var outputReley = worksheet.Cells[8, i].Value.ToString().Trim();

                                if (i == 6)
                                {
                                    numberPosition = "";
                                    nakladka = "";
                                    outputReley = "";
                                }

                                var string2 = worksheet.Cells[2, i].Value.ToString().Trim() + "||" + worksheet.Cells[3, i].Value.ToString().Trim() +
                                    "||" + worksheet.Cells[4, i].Value.ToString().Trim();
                                var string3 = numberPosition + "||" + worksheet.Cells[6, i].Value.ToString();
                                var string4 = nakladka + "||" + outputReley + "||" + cellValue.ToString().Split('/')[0].Trim();

                                Txt.WriteTxt(pathTxt, "--SIDESC--");
                                Txt.WriteTxt(pathTxt, string2);
                                Txt.WriteTxt(pathTxt, string3);
                                Txt.WriteTxt(pathTxt, string4);

                                if (B2Exists)
                                {
                                    var string4B2 = nakladka + "||" + outputReley + "||" + cellValue.ToString().Split('/')[1].Trim();
                                    Txt.WriteTxt(pathTxtB2, "--SIDESC--");
                                    Txt.WriteTxt(pathTxtB2, string2);
                                    Txt.WriteTxt(pathTxtB2, string3);
                                    Txt.WriteTxt(pathTxtB2, string4B2);
                                }
                            }
                        }
                    }
                }                
            }
        }

        static private string TypeBLCAP(string cellValue, int j, int i)
        {
            if (cellValue.ToString().Split('/')[0].Trim() == "ЗапО" ||
                            cellValue.ToString().Split('/')[0].Trim() == "ЗапЗ")
            {
                return "Запрет";
            }
            else if (cellValue.ToString().Split('/')[0].Trim() == "Закр" ||
                            cellValue.ToString().Split('/')[0].Trim() == "Откр" ||
                                cellValue.ToString().Split('/')[0].Trim() == "Вкл" ||
                                    cellValue.ToString().Split('/')[0].Trim() == "Откл")
            {
                return "Команда";
            }
            else
            {
                Console.WriteLine("Строка " + j + "Столбец " + i);
                Console.WriteLine(cellValue.ToString().Split('/')[0].Trim());
                throw new Exception("Неопознанная команда или запрет!!");                
            }
        }
    }
}
