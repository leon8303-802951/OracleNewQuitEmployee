using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model.Model1
{

    using System.Globalization;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public class OracleUserResultModel
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("meta")]
        public Meta Meta { get; set; }

        [JsonProperty("schemas")]
        public string[] Schemas { get; set; }

        [JsonProperty("userName")] 
        public string UserName { get; set; }

        [JsonProperty("name")]
        public Name Name { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("preferredLanguage")]
        public string PreferredLanguage { get; set; }

        [JsonProperty("emails")]
        public Email[] Emails { get; set; }

        [JsonProperty("roles")]
        public Role[] Roles { get; set; }

        [JsonProperty("urn:scim:schemas:extension:fa:2.0:faUser")]
        public UrnScimSchemasExtensionFa20FaUser UrnScimSchemasExtensionFa20FaUser { get; set; }

        [JsonProperty("active")]
        public bool Active { get; set; }
    }

    public class Email
    {
        [JsonProperty("value")]
        public string Value { get; set; }

        [JsonProperty("type")]
        public string Type { get; set; }

        [JsonProperty("primary")]
        public string Primary { get; set; }
    }

    public class Meta
    {
        [JsonProperty("location")]
        public string Location { get; set; }

        [JsonProperty("resourceType")]
        public string ResourceType { get; set; }

        [JsonProperty("created")]
        public string Created { get; set; }

        [JsonProperty("lastModified")]
        public string LastModified { get; set; }
    }

    public class Name
    {
        [JsonProperty("familyName")] 
        public string FamilyName { get; set; }

        [JsonProperty("givenName")]
        public string GivenName { get; set; }
    }

    public class Role
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("value")]
        public string Value { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }
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

        [JsonProperty("job")]
        public string Job { get; set; }

        [JsonProperty("businessUnit")]
        public string BusinessUnit { get; set; }

        [JsonProperty("department")]
        public string Department { get; set; }
    }
}
