using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{

        using System;
        using Newtonsoft.Json;

        public class OracleJobIDReturnModel
        {
            [JsonProperty("items")]
            public OracleJobID[] Items { get; set; }

            [JsonProperty("count")]
            public int Count { get; set; }

            [JsonProperty("hasMore")]
            public bool HasMore { get; set; }

            [JsonProperty("limit")]
            public int Limit { get; set; }

            [JsonProperty("offset")]
            public int Offset { get; set; }

            [JsonProperty("links")]
            public OracleJobIDLink[] Links { get; set; }
        }

        public class OracleJobID
        {
            [JsonProperty("JobId")]
            public string JobId { get; set; }

            [JsonProperty("JobCode")]
            public string JobCode { get; set; }

            [JsonProperty("JobFamilyId")]
            public string JobFamilyId { get; set; }

            [JsonProperty("ActiveStatus")]
            public string ActiveStatus { get; set; }

            [JsonProperty("FullPartTime")]
            public string FullPartTime { get; set; }

            [JsonProperty("JobFunctionCode")]
            public string JobFunctionCode { get; set; }

            [JsonProperty("ManagerLevel")]
            public string ManagerLevel { get; set; }

            [JsonProperty("MedicalCheckupRequired")]
            public string MedicalCheckupRequired { get; set; }

            [JsonProperty("RegularTemporary")]
            public string RegularTemporary { get; set; }

            [JsonProperty("SetId")]
            public string SetId { get; set; }

            [JsonProperty("EffectiveStartDate")]
            public string EffectiveStartDate { get; set; }

            [JsonProperty("EffectiveEndDate")]
            public string EffectiveEndDate { get; set; }

            [JsonProperty("Name")]
            public string Name { get; set; }

            [JsonProperty("ApprovalAuthority")]
            public string ApprovalAuthority { get; set; }

            [JsonProperty("GradeLadderId")]
            public string GradeLadderId { get; set; }

            [JsonProperty("CreationDate")]
            public string CreationDate { get; set; }

            [JsonProperty("LastUpdateDate")]
            public string LastUpdateDate { get; set; }

            [JsonProperty("links")]
            public OracleJobIDLink[] Links { get; set; }
        }

        public partial class OracleJobIDLink
        {
            [JsonProperty("rel")]
            public string Rel { get; set; }

            [JsonProperty("href")]
            public Uri Href { get; set; }

            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("kind")]
            public string Kind { get; set; }

            [JsonProperty("properties", NullValueHandling = NullValueHandling.Ignore)]
            public OracleJobIDProperties Properties { get; set; }
        }

        public partial class OracleJobIDProperties
        {
            [JsonProperty("changeIndicator")]
            public string ChangeIndicator { get; set; }
        }


    



}
