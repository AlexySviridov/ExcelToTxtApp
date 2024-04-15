using System.IO;
using System.Text;
using System;

namespace ClassLibrary
{
    public class Txt
    {
        public static void WriteTxt(string path, string blcap)
        {
            try
            {                
                using (FileStream fs = File.Create(path))
                {
                    byte[] info = new UTF8Encoding(true).GetBytes("--TABS--\nОткрыть\nЗакрыть\n--BLCAP--\n" + blcap);
                    fs.Write(info, 0, info.Length);                 
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
