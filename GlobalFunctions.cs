using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace SprayingCabineService
{
    internal class GlobalFunctions
    {


        public static Encoding FileDetectEncoding(string FileName)
        {
            string enc = "";
            if (File.Exists(FileName))
            {
                FileStream filein = new FileStream(FileName, FileMode.Open, FileAccess.Read);
                if ((filein.CanSeek))
                {
                    byte[] bom = new byte[5];
                    filein.Read(bom, 0, 4);
                    // EF BB BF = utf-8 FF FE = ucs-2le, ucs-4le, and ucs-16le FE FF = utf-16 and
                    // ucs-2 00 00 FE FF = ucs-4
                    if ((((bom[0] == 0xEF) && (bom[1] == 0xBB) && (bom[2] == 0xBF)) || ((bom[0] == 0xFF) && (bom[1] == 0xFE)) || ((bom[0] == 0xFE) && (bom[1] == 0xFF)) || ((bom[0] == 0x0) && (bom[1] == 0x0) && (bom[2] == 0xFE) && (bom[3] == 0xFF))))
                        enc = "Unicode";
                    else
                        enc = "ASCII";
                    // Position the file cursor back to the start of the file
                    filein.Seek(0, SeekOrigin.Begin);
                }
                filein.Close();
            }
            if (enc == "Unicode")
                return Encoding.UTF8;
            else
                return Encoding.Default;
        }


        public static string UnicodeToUTF8(string strFrom)
        {
            byte[] bytes = Encoding.Default.GetBytes(strFrom);
            return Encoding.UTF8.GetString(bytes);
        }


    }
}
