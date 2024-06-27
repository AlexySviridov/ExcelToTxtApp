using System.IO;


namespace ClassLibrary
{
    internal class Excel
    {
        public static void ReadExcel(string pathToExcel)
        {
            FileInfo existingFile = new FileInfo(pathToExcel);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {

            }
        }
    }
}
