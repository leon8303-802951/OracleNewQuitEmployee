using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Text;

namespace OracleNewQuitEmployee.ToolModels
{
    public class MiddleModel3
    {
        public string URL { get; set; }
        public string SendingData { get; set; }
        public string Method { get; set; }
        public string ContnetType { get; set; }
        public NameValueCollection AddHeaders;
        public string UserName { get; set; }
        public string Password { get; set; }
        public CredentialCache Cred { get; set; }
        public int? Timeout { get; set; }
    }
}
