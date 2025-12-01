using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace WinformUI.CADHelper
{
    public class TXT
    {
        public static void TXT建立(string Path, string content)
        {
            FileStream fs = new FileStream(Path, FileMode.OpenOrCreate);
            byte[] data = Encoding.Default.GetBytes(content);
            fs.Write(data, 0, data.Length);
            fs.Flush();
            fs.Close();
        }
    }
}
