using ClassLibrary;

namespace Host
{
    internal class Program
    {
        static void Main()
        {
            var pathToExcel = "C:\\Users\\User\\Desktop\\Илья\\K6. Info v1.35.xlsx";
            var pathDirectoryToSave = "C:\\Users\\User\\Desktop\\ТЗиБ\\";
            var numberWorksheet = 12;            
            var firstArmatureRow = 13;
            var ArmatureNameColumn = 3;
            var firstAlgorithmColumn = 5;
            var ignoredRowsArray = new[] { 15, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 49, 87, 88, 90 };            

            ExcelTZiB.DoAllJob (pathToExcel, numberWorksheet, ignoredRowsArray, firstArmatureRow, ArmatureNameColumn, firstAlgorithmColumn, pathDirectoryToSave);
        }
    }
}
