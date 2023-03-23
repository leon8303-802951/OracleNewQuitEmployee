using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.Models
{
    using System;
    using System.Collections.Generic;

    using System.Globalization;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public partial class SupplierModel
    {
        [JsonProperty("r_code")]
        public string RCode { get; set; }

        [JsonProperty("r_cname")]
        public string RCname { get; set; }

        [JsonProperty("r_dept")]
        public string RDept { get; set; }

        [JsonProperty("r_degress")]
        public string RDegress { get; set; }

        [JsonProperty("r_cell_phone")]
        public string RCellPhone { get; set; }

        [JsonProperty("r_birthday")]
        public string RBirthday { get; set; }

        [JsonProperty("r_sex")]
        public string RSex { get; set; }

        [JsonProperty("r_email")]
        public string REmail { get; set; }

        [JsonProperty("r_phone_ext")]
        public string RPhoneExt { get; set; }

        [JsonProperty("r_skype_id")]
        public string RSkypeId { get; set; }

        [JsonProperty("r_online_date")]
        public string ROnlineDate { get; set; }

        [JsonProperty("r_online")]
        public string ROnline { get; set; }

        [JsonProperty("r_offline_date")]
        public string ROfflineDate { get; set; }

        [JsonProperty("batchno")]
        public string Batchno { get; set; }

        [JsonProperty("r_ename")]
        public string REname { get; set; }

        [JsonProperty("departname")]
        public string Departname { get; set; }

        [JsonProperty("dutyid")]
        public string Dutyid { get; set; }

        [JsonProperty("workplaceid")]
        public string Workplaceid { get; set; }

        [JsonProperty("workplacename")]
        public string Workplacename { get; set; }

        [JsonProperty("idno")]
        public string Idno { get; set; }

        [JsonProperty("salaccountname")]
        public string Salaccountname { get; set; }

        [JsonProperty("bankcode")]
        public string Bankcode { get; set; }

        [JsonProperty("bankbranch")]
        public string Bankbranch { get; set; }

        [JsonProperty("salaccountid")]
        public string Salaccountid { get; set; }

        [JsonProperty("jobstatus")]
        public string Jobstatus { get; set; }
    }
}