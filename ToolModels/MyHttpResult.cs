using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ToolModels
{
    public class MyHttpResult
    {
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public string Result { get; set; }
        public string ErrorMessage { get; set; }
        public bool Success { get; set; }
    }
}
