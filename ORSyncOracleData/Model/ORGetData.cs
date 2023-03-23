using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class ORGetData
    {
        public ORGetData()
        {
        }
        public string FunctionName { get; set; }
        public NameValueCollection Pars { get; set; }
    }
}
