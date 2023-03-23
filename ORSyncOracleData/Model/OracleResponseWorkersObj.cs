using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    using System;
    using System.Collections.Generic;

    using System.Globalization;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public class OracleResponseWorkersObj
    {
        [JsonProperty("items")]
        public Worker[] Items { get; set; }

        [JsonProperty("count")]
        public long Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public long Limit { get; set; }

        [JsonProperty("offset")]
        public long Offset { get; set; }

        [JsonProperty("links")]
        public WorkerLink[] Links { get; set; }
    }

    public class Worker
    {
        [JsonProperty("PersonId")]
        public long PersonId { get; set; }

        [JsonProperty("PersonNumber")]
        public string PersonNumber { get; set; }

        [JsonProperty("CorrespondenceLanguage")]
        public string CorrespondenceLanguage { get; set; }

        [JsonProperty("BloodType")]
        public object BloodType { get; set; }

        [JsonProperty("DateOfBirth")]
        public string DateOfBirth { get; set; }

        [JsonProperty("DateOfDeath")]
        public object DateOfDeath { get; set; }

        [JsonProperty("CountryOfBirth")]
        public object CountryOfBirth { get; set; }

        [JsonProperty("RegionOfBirth")]
        public object RegionOfBirth { get; set; }

        [JsonProperty("TownOfBirth")]
        public object TownOfBirth { get; set; }

        [JsonProperty("ApplicantNumber")]
        public object ApplicantNumber { get; set; }

        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        [JsonProperty("CreationDate")]
        public string CreationDate { get; set; }

        [JsonProperty("LastUpdatedBy")]
        public string LastUpdatedBy { get; set; }

        [JsonProperty("LastUpdateDate")]
        public string LastUpdateDate { get; set; }

        [JsonProperty("links")]
        public WorkerLink[] Links { get; set; }
    }

    public class WorkerLink
    {
        [JsonProperty("rel")]
        public string Rel { get; set; }

        [JsonProperty("href")]
        public string Href { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("kind")]
        public string Kind { get; set; }

        [JsonProperty("properties", NullValueHandling = NullValueHandling.Ignore)]
        public WorkerProperties Properties { get; set; }
    }

    public class WorkerProperties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }

}
