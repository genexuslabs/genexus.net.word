using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Genexus.Word

{
    public class OutputCode
    {
        public readonly static int OK = 0;
        public readonly static int FAIL_OPEN = 10;
        public readonly static int FAIL_CREATE = 11;
        public readonly static int FILE_ALREADY_EXISTS = 7;
        public readonly static int FILE_NOT_FOUND = 6;
        public readonly static int INVALID_OPERATION = 1;
    }

    public class AddImageOutputCode
    {
        public readonly static int IMAGE_NOT_FOUND = 12;
    }
}
