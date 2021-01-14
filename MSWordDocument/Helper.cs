using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Genexus.Word
{
    public static class Helper
    {
        public static string ToString(string value, string defaultValue)
        {
            return (!string.IsNullOrEmpty(value)) ? value : defaultValue;
        }

        public static string ToString(double value, string defaultValue)
        {
            return (value > 0) ? value.ToString() : defaultValue;
        }

        public static string ToPtUnit(double value, double defaultValue)
        {
            return (value > 0) ? String.Format("{0}pt", value.ToString()) : String.Format("{0}pt", defaultValue.ToString());
        }


        public static string ToColorHex(string colorName, string defaultColorHex = null)
        {
            if (string.IsNullOrEmpty(colorName))
            {
                return defaultColorHex;
            }

            if (colorName.StartsWith("#"))
            {
                return colorName;
            }

            int ColorValue = Color.FromName(colorName).ToArgb();
            return string.Format("{0:x6}", ColorValue);
        }
    }
}
