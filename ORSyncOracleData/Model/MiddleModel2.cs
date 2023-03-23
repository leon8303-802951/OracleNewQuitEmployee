using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class MiddleModel2
    {
        public MiddleModel2()
        {
            AddHeaders = new Dictionary<string, string>();
        }
        public string URL { get; set; }
        public string Proxy { get; set; }
        public int ProxyPort { get; set; }
        public string SendingData { get; set; }
        public string Method { get; set; }
        public string ContnetType { get; set; }
        public Dictionary<string, string> AddHeaders;
        public string UserName { get; set; }
        public string Password { get; set; }
        //public CredentialCache Cred { get; set; }
        public int? Timeout { get; set; }
    }

}
