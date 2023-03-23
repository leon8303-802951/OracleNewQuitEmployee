using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class MiddleReturn
    {
        public string StatusCode { get; set; }
        public string StatusDescription { get; set; }
        public string ReturnData { get; set; }
        public string ErrorMessage { get; set; }
        public double UsingSecs { get; set; }

    }
}
