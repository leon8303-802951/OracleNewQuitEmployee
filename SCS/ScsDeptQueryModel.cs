using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.SCS
{
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public partial class ScsDeptQueryModel
    {

        public ScsDeptQueryModel()
        {
            this.Value = new DeptValue();
        }
        [JsonProperty("SessionGuid")]
        public string SessionGuid { get; set; }

        [JsonProperty("ProgID")]
        public string ProgId { get; set; }

        [JsonProperty("Value")]
        public DeptValue Value { get; set; }
    }

    public partial class DeptValue
    {
        [JsonProperty("$type")]
        public string Type { get; set; }

        [JsonProperty("FormID")]
        public string FormId { get; set; }

        [JsonProperty("SystemFilterOptions")]
        public string SystemFilterOptions { get; set; }
    }




}
