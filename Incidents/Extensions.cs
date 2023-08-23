using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Incidents
{
    public static class Extensions
    {
        public static string GetEnumDescription(this Enum value)
        {
            FieldInfo fi = value.GetType().GetField(value.ToString());

            DescriptionAttribute[] attributes =
                (DescriptionAttribute[])fi.GetCustomAttributes(typeof(DescriptionAttribute), false);

            if (attributes != null && attributes.Length > 0)
                return attributes[0].Description;
            else
                return value.ToString();
        }
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
