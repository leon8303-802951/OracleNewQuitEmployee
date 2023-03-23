using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.SCS
{
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public class ScsDeptsReturnModel
    {
        [JsonProperty("DataTable")]
        public DeptDataTable[] DataTable { get; set; }
    }


    public class DeptDataTable
    {
        [JsonProperty("SYS_SELECTED")]
        public bool SysSelected { get; set; }

        [JsonProperty("SYS_VIEWID")]
        public string SysViewid { get; set; }

        [JsonProperty("SYS_NAME")]
        public string SysName { get; set; }

        [JsonProperty("SYS_ENGNAME")]
        public string SysEngname { get; set; }

        [JsonProperty("SYS_ID")]
        public string SysId { get; set; }
    }

}
