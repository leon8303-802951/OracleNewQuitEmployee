using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.SCS
{
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public class ScsLoginModel
    {
        public ScsLoginModel()
        {
            this.Value = new SCSValue();
        }


        [JsonProperty("Action")]
        public string Action { get; set; }

        [JsonProperty("Value")]
        public SCSValue Value { get; set; }
    }

    public class SCSValue
    {
        [JsonProperty("$type")]
        public string Type { get; set; }

        [JsonProperty("CompanyID")]
        public string CompanyId { get; set; }

        [JsonProperty("UserID")]
        public string UserId { get; set; }

        [JsonProperty("Password")]
        public string Password { get; set; }

        [JsonProperty("LanguageID")]
        public string LanguageId { get; set; }
    }






}
