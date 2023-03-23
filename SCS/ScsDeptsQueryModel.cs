using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.SCS
{
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public partial class ScsDeptsQueryModel
    {
        public ScsDeptsQueryModel()
        {
            this.Value = new DeptsValue();
        }

        [JsonProperty("Action")]
        public string Action { get; set; }

        [JsonProperty("SessionGuid")]
        public string SessionGuid { get; set; }

        [JsonProperty("ProgID")]
        public string ProgId { get; set; }

        [JsonProperty("Value")]
        public DeptsValue Value { get; set; }
    }

    public partial class DeptsValue
    {
        [JsonProperty("$type")]
        public string Type { get; set; }

        [JsonProperty("SelectFields")]
        public string SelectFields { get; set; }

        [JsonProperty("FilterItems")]
        public object[] FilterItems { get; set; }

        [JsonProperty("SystemFilterOptions")]
        public string SystemFilterOptions { get; set; }

        [JsonProperty("IsBuildSelectedField")]
        public bool IsBuildSelectedField { get; set; }

        [JsonProperty("IsBuildFlowLightSignalField")]
        public bool IsBuildFlowLightSignalField { get; set; }
    }



}
