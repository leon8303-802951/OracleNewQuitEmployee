using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model.SCIMUser
{

    using Newtonsoft.Json;

    public class OracleAllUsersBySCIM
    {
        [JsonProperty("itemsPerPage")]
        public string ItemsPerPage { get; set; }

        [JsonProperty("startIndex")]
        public string StartIndex { get; set; }

        [JsonProperty("Resources")]
        public OracleAllUsersBySCIM_Resource[] Resources { get; set; }
    }

    public class OracleAllUsersBySCIM_Resource
    {
        /// <summary>
        /// GUID
        /// </summary>
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("meta")]
        public OracleAllUsersBySCIM_Meta Meta { get; set; }

        [JsonProperty("schemas")]
        public string[] Schemas { get; set; }

        [JsonProperty("userName")]
        public string UserName { get; set; }

        [JsonProperty("name")]
        public OracleAllUsersBySCIM_Name Name { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("preferredLanguage")]
        public string PreferredLanguage { get; set; }

        [JsonProperty("urn:scim:schemas:extension:fa:2.0:faUser")]
        public UrnScimSchemasExtensionFa20FaUser UrnScimSchemasExtensionFa20FaUser { get; set; }

        [JsonProperty("active")]
        public bool Active { get; set; }
    }

    public class OracleAllUsersBySCIM_Meta
    {
        [JsonProperty("location")]
        public Uri Location { get; set; }

        [JsonProperty("resourceType")]
        public string ResourceType { get; set; }

        [JsonProperty("created")]
        public DateTimeOffset Created { get; set; }

        [JsonProperty("lastModified")]
        public DateTimeOffset LastModified { get; set; }
    }

    public class OracleAllUsersBySCIM_Name
    {
        [JsonProperty("familyName")]
        public string FamilyName { get; set; }
    }

    public class UrnScimSchemasExtensionFa20FaUser
    {
        [JsonProperty("userCategory")]
        public string UserCategory { get; set; }

        [JsonProperty("accountType")]
        public string AccountType { get; set; }

        [JsonProperty("workerInformation")]
        public WorkerInformation WorkerInformation { get; set; }
    }

    public class WorkerInformation
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
