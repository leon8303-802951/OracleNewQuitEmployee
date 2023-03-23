using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    using System.Globalization;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public class ApiResultModel
    {
        [JsonProperty("Result")]
        public bool Result { get; set; }

        [JsonProperty("ErrorCode")]
        public object ErrorCode { get; set; }

        [JsonProperty("ErrorMessages")]
        public string[] ErrorMessages { get; set; }
    }
}
