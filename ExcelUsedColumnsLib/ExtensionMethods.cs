using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelUsedColumnsLib
{
    public static class ExtensionMethods
    {
        public static string ExcelColumnName(this int Index)
        {
            var chars = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };

            Index -= 1;
            string columnName = null;
            var quotient = Index / 26;
            if (quotient > 0)
            {
                columnName = ExcelColumnName(quotient) + chars[Index % 26];
            }
            else
            {
                columnName = chars[Index % 26].ToString();
            }
            return columnName;
        }

    }
}
