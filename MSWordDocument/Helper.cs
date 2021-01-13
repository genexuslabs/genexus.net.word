using System;
using System.Collections.Generic;
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
    }
}
