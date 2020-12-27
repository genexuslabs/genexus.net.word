using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Genexus.Word
{
    internal class ExceptionManager
    {
        public static void HandleException(Exception ex)
        {
            // see to integrate log4net here
        }

        internal static void LogException(string v)
        {
            throw new NotImplementedException();
        }
    }
}
