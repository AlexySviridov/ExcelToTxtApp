using System.IO;
using System.Text;
using System;

namespace ClassLibrary
{
    public class Txt
    {
        public static void WriteTxt(string path, string line)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(path, true, Encoding.GetEncoding(1251)))
                {
                    sw.WriteLine(line);
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public static void CreateTxt(string path, string blcap)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(path, true, Encoding.GetEncoding(1251)))
                {
                    sw.WriteLine("--TABS--");
                    sw.WriteLine("Открыть");
                    sw.WriteLine("Закрыть");
                    sw.WriteLine("--BLCAP--");
                    sw.WriteLine(blcap);
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public static void ReadTxt(string path)
        {
            try
            {
                using (StreamReader sr = File.OpenText(path))
                {
                    string s = "";
                    while ((s = sr.ReadLine()) != null)
                    {
                        Console.WriteLine(s);
                    }
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}
