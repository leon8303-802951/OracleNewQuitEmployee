using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.SCS
{

    using System.Globalization;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public class ScsEmployeesQueryModel
    {
        public ScsEmployeesQueryModel()
        {
            this.Value = new SCSEmployeeQueryValue();
        }

        [JsonProperty("Action")]
        public string Action { get; set; }

        [JsonProperty("SessionGuid")]
        public string SessionGuid { get; set; }

        [JsonProperty("ProgID")]
        public string ProgId { get; set; }

        [JsonProperty("Value")]
        public SCSEmployeeQueryValue Value { get; set; }
    }

    public class SCSEmployeeQueryValue
    {
        [JsonProperty("$type")]
        public string Type { get; set; }

        [JsonProperty("UIType")]
        public string UiType { get; set; }

        [JsonProperty("ReportID")]
        public string ReportId { get; set; }

        [JsonProperty("ReportTailID")]
        public string ReportTailId { get; set; }

        [JsonProperty("FilterItems")]
        public FilterItem[] FilterItems { get; set; }

        [JsonProperty("UserFilter")]
        public string UserFilter { get; set; }
    }

    public class FilterItem
    {
        [JsonProperty("$type")]
        public string Type { get; set; }

        [JsonProperty("FieldName")]
        public string FieldName { get; set; }

        [JsonProperty("FilterValue")]
        public string FilterValue { get; set; }
    }

}
