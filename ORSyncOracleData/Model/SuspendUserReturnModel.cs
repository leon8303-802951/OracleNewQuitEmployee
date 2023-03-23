using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{

    using Newtonsoft.Json;

    public  class SuspendUserReturnModel
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("meta")]
        public UserReturnMeta Meta { get; set; }

        [JsonProperty("schemas")]
        public string[] Schemas { get; set; }

        [JsonProperty("userName")]
        public string UserName { get; set; }

        [JsonProperty("name")]
        public OracleUsersName Name { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("preferredLanguage")]
        public string PreferredLanguage { get; set; }

        [JsonProperty("urn:scim:schemas:extension:fa:2.0:faUser")]
        public UrnScimSchemasExtensionFa20FaUser UrnScimSchemasExtensionFa20FaUser { get; set; }

        [JsonProperty("active")]
        public string Active { get; set; }
    }

    public  class UserReturnMeta
    {
        [JsonProperty("location")]
        public string Location { get; set; }

        [JsonProperty("resourceType")]
        public string ResourceType { get; set; }

        [JsonProperty("created")]
        public  string  Created { get; set; }

        [JsonProperty("lastModified")]
        public  string  LastModified { get; set; }
    }

    public  class OracleUsersName
    {
        [JsonProperty("familyName")]
        public string FamilyName { get; set; }
    }

    public  class UrnScimSchemasExtensionFa20FaUser
    {
        [JsonProperty("userCategory")]
        public string UserCategory { get; set; }

        [JsonProperty("accountType")]
        public string AccountType { get; set; }

        [JsonProperty("workerInformation")]
        public WorkerInformation WorkerInformation { get; set; }
    }

    public  class WorkerInformation
    {
        [JsonProperty("personNumber")]
        public string PersonNumber { get; set; }

        [JsonProperty("manager")]
        public string Manager { get; set; }

        [JsonProperty("businessUnit")]
        public string BusinessUnit { get; set; }

        [JsonProperty("department")]
        public string Department { get; set; }
    }






}
