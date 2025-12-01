using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinformUI.CADHelper
{
    public static class CommonTools
    {
        public static int GetIndexOfLetters(string letter)
        {
            List<string> letters = new List<string> { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            return letters.IndexOf(letter.ToUpper());
        }
        public static int GetNumberInStr(string str)
        {
            string result = System.Text.RegularExpressions.Regex.Replace(str, @"[^0-9]+", "");
            int num = 0;
            int.TryParse(result, out num);
            return num;
        }
    }
}
