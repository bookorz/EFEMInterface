﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFEMInterface.MessageInterface
{
    public class ErrorCategory
    {
        public class ErrorType
        {
            public const string COMMAND = "COMMAND";
        }

        public class ErrorDetail
        {
            public const string DUPLICATE = "DUPLICATE";
            
        }

    }
}
