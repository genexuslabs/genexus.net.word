using DocumentFormat.OpenXml;
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


        /// <summary>
        /// Get properties from string properties
        /// </summary>
        /// <param name="properties"></param>
        /// <returns></returns>
        public static IEnumerable<OpenXmlElement> GetProperties(List<string> properties)
        {
            foreach (string prop in properties)
            {
                if (StyleProperties.Exists(prop))
                {
                    yield return StyleProperties.RunFunctionProperty(prop);
                }
            }
        }


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


        public static string ToRGBHexColor(string colorName, string defaultColorHex = null)
        {
            if (string.IsNullOrEmpty(colorName))
            {
                return defaultColorHex;
            }

            if (colorName.StartsWith("#"))
            {
                return colorName.Substring(1);
            }

            Color color = Color.FromName(colorName);
            return String.Format("{0:X2}{1:X2}{2:X2}", color.R, color.G, color.B).ToUpper();
        }
    }
}
