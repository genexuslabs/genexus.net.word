using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Genexus.Word.Helpers
{
    public static class MathOpenXml
    {
        private const int ONEEMUINCENTIMETERS = 360142;

        public static double CentimetersToEMU(double cm)
        {
            return Math.Round(cm * ONEEMUINCENTIMETERS);
        }
    }
}
