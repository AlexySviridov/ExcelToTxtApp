using OfficeOpenXml;
using System.IO;
using System;
using ClassLibrary;

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
                //Console.WriteLine("Value:" + worksheet.Cells[6, 6].Value.ToString().Trim() + "\n");
                int colCount = worksheet.Dimension.End.Column;
                int rowCount = worksheet.Dimension.End.Row;

                for (int i = 6; i < colCount; i++)
                {
                    var armatureName = worksheet.Cells[15, 4].Value.ToString().Trim();
                    var cellValue = worksheet.Cells[15, i].Value;
                    var pathTxt = "C:\\Users\\User\\Desktop\\ТЗиБ\\" + armatureName + "_B1.db";

                    if (cellValue != null)
                    {
                        if (cellValue.ToString().Split('/')[0].Trim() == "ЗапО" ||
                            cellValue.ToString().Split('/')[0].Trim() == "ЗапЗ")
                        {
                            var blcap = "Запрет";
                            //Console.WriteLine("true");
                            if (!File.Exists(pathTxt))
                            {
                                Txt.CreateTxt(pathTxt, blcap);
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
                        }

                        //if (cellValue.ToString().Split('/')[0].Trim() == "Закр" || cellValue.ToString().Split('/')[0].Trim() == "Откр"
                        //    || cellValue.ToString().Split('/')[0].Trim() == "Вкл" || cellValue.ToString().Split('/')[0].Trim() == "Откл")
                        //{
                        //    var blcap = "Команда";
                        //    if (!File.Exists(pathTxt))
                        //    {
                        //        Txt.CreateTxt(pathTxt, blcap);
                        //    }
                        //}
                    }
                }
            }
        }

        static private void TypeBLCAP()
        {

        }
    }
}
