using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class OracleUser
    {
        public string UserId { get; set; }
        public string Username { get; set; }
        public string SuspendedFlag { get; set; }
        public string PersonId { get; set; }
        public string PersonNumber { get; set; }
        public string CredentialsEmailSentFlag { get; set; }
        public string GUID { get; set; }
        public string CreatedBy { get; set; }
        public string CreationDate { get; set; }
        public string LastUpdatedBy { get; set; }
        public string LastUpdateDate { get; set; }
        public List<OracleApiReturnObjLink> Links { get; set; }

    }

}
