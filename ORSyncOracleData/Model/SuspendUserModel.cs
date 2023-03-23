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

    public class SuspendUserModel
    {
        [JsonProperty("schemas")]
        public string[] Schemas { get; set; }

        [JsonProperty("active")]
        public string Active { get; set; }
    }

    //public partial class SuspendUserModel
    //{
    //    public static SuspendUserModel FromJson(string json) => JsonConvert.DeserializeObject<SuspendUserModel>(json, QuickType.Converter.Settings);
    //}

    //public static class Serialize
    //{
    //    public static string ToJson(this SuspendUserModel self) => JsonConvert.SerializeObject(self, QuickType.Converter.Settings);
    //}

    //internal static class Converter
    //{
    //    public static readonly JsonSerializerSettings Settings = new JsonSerializerSettings
    //    {
    //        MetadataPropertyHandling = MetadataPropertyHandling.Ignore,
    //        DateParseHandling = DateParseHandling.None,
    //        Converters =
    //    {
    //        new IsoDateTimeConverter { DateTimeStyles = DateTimeStyles.AssumeUniversal }
    //    },
    //    };
    //}



}