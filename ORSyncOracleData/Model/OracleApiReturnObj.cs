

using System.Globalization;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public partial class OracleApiReturnObj<T>
    {
        [JsonProperty("items")]
        public List<T> Items { get; set; }

        [JsonProperty("count")]
        public long Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public long Limit { get; set; }

        [JsonProperty("offset")]
        public long Offset { get; set; }

        [JsonProperty("links")]
        public OracleApiReturnObjLink[] Links { get; set; }
    }

    public partial class OracleApiReturnObjItem
    {
        [JsonProperty("PersonNameId")]
        public long PersonNameId { get; set; }

        [JsonProperty("EffectiveStartDate")]
        public DateTimeOffset EffectiveStartDate { get; set; }

        [JsonProperty("EffectiveEndDate")]
        public DateTimeOffset EffectiveEndDate { get; set; }

        [JsonProperty("LegislationCode")]
        public string LegislationCode { get; set; }

        /// <summary>
        /// 員工編號 EmpNo
        /// </summary>
        [JsonProperty("LastName")]
        public string LastName { get; set; }

        /// <summary>
        /// 員工中文姓名
        /// </summary>
        [JsonProperty("FirstName")]
        public object FirstName { get; set; }

        [JsonProperty("Title")]
        public object Title { get; set; }

        [JsonProperty("PreNameAdjunct")]
        public object PreNameAdjunct { get; set; }

        /// <summary>
        /// 員工英文姓名
        /// </summary>
        [JsonProperty("MiddleNames")]
        public object MiddleNames { get; set; }

        [JsonProperty("KnownAs")]
        public object KnownAs { get; set; }

        [JsonProperty("PreviousLastName")]
        public object PreviousLastName { get; set; }

        [JsonProperty("DisplayName")]
        public string DisplayName { get; set; }

        [JsonProperty("OrderName")]
        public string OrderName { get; set; }

        [JsonProperty("ListName")]
        public string ListName { get; set; }

        [JsonProperty("FullName")]
        public string FullName { get; set; }

        [JsonProperty("NameLanguage")]
        public string NameLanguage { get; set; }

        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        [JsonProperty("CreationDate")]
        public DateTimeOffset CreationDate { get; set; }

        [JsonProperty("LastUpdatedBy")]
        public string LastUpdatedBy { get; set; }

        [JsonProperty("LastUpdateDate")]
        public DateTimeOffset LastUpdateDate { get; set; }

        [JsonProperty("LocalPersonNameId")]
        public long LocalPersonNameId { get; set; }

        [JsonProperty("LocalEffectiveStartDate")]
        public DateTimeOffset LocalEffectiveStartDate { get; set; }

        [JsonProperty("LocalEffectiveEndDate")]
        public DateTimeOffset LocalEffectiveEndDate { get; set; }

        [JsonProperty("LocalLegislationCode")]
        public string LocalLegislationCode { get; set; }

        [JsonProperty("LocalLastName")]
        public string LocalLastName { get; set; }

        [JsonProperty("LocalFirstName")]
        public object LocalFirstName { get; set; }

        [JsonProperty("LocalTitle")]
        public object LocalTitle { get; set; }

        [JsonProperty("LocalPreNameAdjunct")]
        public object LocalPreNameAdjunct { get; set; }

        [JsonProperty("LocalMiddleNames")]
        public object LocalMiddleNames { get; set; }

        [JsonProperty("LocalKnownAs")]
        public object LocalKnownAs { get; set; }

        [JsonProperty("LocalPreviousLastName")]
        public object LocalPreviousLastName { get; set; }

        [JsonProperty("LocalDisplayName")]
        public string LocalDisplayName { get; set; }

        [JsonProperty("LocalOrderName")]
        public string LocalOrderName { get; set; }

        [JsonProperty("LocalListName")]
        public string LocalListName { get; set; }

        [JsonProperty("LocalFullName")]
        public string LocalFullName { get; set; }

        [JsonProperty("LocalNameLanguage")]
        public string LocalNameLanguage { get; set; }

        [JsonProperty("LocalCreatedBy")]
        public string LocalCreatedBy { get; set; }

        [JsonProperty("LocalCreationDate")]
        public DateTimeOffset LocalCreationDate { get; set; }

        [JsonProperty("LocalLastUpdatedBy")]
        public string LocalLastUpdatedBy { get; set; }

        [JsonProperty("LocalLastUpdateDate")]
        public DateTimeOffset LocalLastUpdateDate { get; set; }

        [JsonProperty("links")]
        public OracleApiReturnObjLink[] Links { get; set; }
    }

    public partial class OracleApiReturnObjLink
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
        public OracleApiReturnObjProperties Properties { get; set; }
    }

    public partial class OracleApiReturnObjProperties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }

}
