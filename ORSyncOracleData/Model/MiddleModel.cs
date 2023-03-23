using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Threading.Tasks;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class MiddleModel
    {
        public string URL { get; set; }
        public string SendingData { get; set; }
        public string Method { get; set; }
        public string ContnetType { get; set; }
        public Dictionary<string, string> AddHeaders;
        public string UserName { get; set; }
        public string Password { get; set; }
        public CredentialCache Cred { get; set; }
        public int? Timeout { get; set; }
    }
}
