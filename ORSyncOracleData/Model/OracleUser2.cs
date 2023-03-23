using Newtonsoft.Json;


namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class TmpResult20230119
    {
        [JsonProperty("items")]
        public OracleUser2[] Items { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public int Limit { get; set; }

        [JsonProperty("offset")]
        public int Offset { get; set; }

        [JsonProperty("links")]
        public Link[] Links { get; set; }
    }

    public class OracleUser2
    {
        [JsonProperty("UserId")]
        public string UserId { get; set; }

        [JsonProperty("Username")]
        public string Username { get; set; }

        [JsonProperty("SuspendedFlag")]
        public bool SuspendedFlag { get; set; }

        [JsonProperty("PersonId")]
        public string PersonId { get; set; }

        [JsonProperty("PersonNumber")]
        public string PersonNumber { get; set; }

        [JsonProperty("CredentialsEmailSentFlag")]
        public bool CredentialsEmailSentFlag { get; set; }

        [JsonProperty("GUID")]
        public string Guid { get; set; }

        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        [JsonProperty("CreationDate")]
        public string CreationDate { get; set; }

        [JsonProperty("LastUpdatedBy")]
        public string LastUpdatedBy { get; set; }

        [JsonProperty("LastUpdateDate")]
        public string LastUpdateDate { get; set; }

        [JsonProperty("links")]
        public Link[] Links { get; set; }
    }

    public class Link
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
        public Properties Properties { get; set; }
    }

    public class Properties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }
}
