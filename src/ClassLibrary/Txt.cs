using System.IO;
using System.Text;
using System;

namespace ClassLibrary
{
    public class Txt
    {
        public static void WriteTxt(string path, string string2, string string3, string string4)
        {
            try
            {
                using (StreamWriter sw = new StreamWriter(path, true, Encoding.GetEncoding(1251)))
                {
                    sw.WriteLine("--SIDESC--");
                    sw.WriteLine(string2);
                    sw.WriteLine(string3);
                    sw.WriteLine(string4);
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
    }
}
