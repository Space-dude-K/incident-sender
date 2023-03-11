using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Incidents
{
    public static class Extensions
    {
        public static string FormatPaddingId(this string str)
        {
            string formatedStr = string.Empty;

            if (str.Length != 19)
            {
                formatedStr = str.PadRight(19);
            }
            else
            {
                formatedStr = str;
            }

            return formatedStr;
        }
    }
}
