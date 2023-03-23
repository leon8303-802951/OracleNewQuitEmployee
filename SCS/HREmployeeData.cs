using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace OracleNewQuitEmployee.SCS
{
    public class HREmployeeData
    {
        /// <summary>
        /// 員編
        /// </summary>
        public string r_code { get; set; } //  員編
        /// <summary>
        /// 中文姓名
        /// </summary>
        public string r_cname { get; set; } //  中文姓名
        /// <summary>
        /// 英文姓名
        /// </summary>
        public string r_ename { get; set; } //  英文姓名
        /// <summary>
        /// 部門
        /// </summary>
        public string r_dept { get; set; } //  部門
        /// <summary>
        /// 職級
        /// </summary>
        public string r_degress { get; set; } //  職級
        /// <summary>
        /// 手機號碼
        /// </summary>
        public string r_cell_phone { get; set; } //  手機號碼
        /// <summary>
        /// 生日
        /// </summary>
        public string r_birthday { get; set; } //  生日
        /// <summary>
        /// 性別 1是男 2是女
        /// </summary>
        public string r_sex { get; set; } //   1是男 2是女
        /// <summary>
        /// EMAIL
        /// </summary>
        public string r_email { get; set; } //  EMAIL
        /// <summary>
        /// 分機
        /// </summary>
        public string r_phone_ext { get; set; } //  分機
        /// <summary>
        /// SKYPE ID
        /// </summary>
        public string r_skype_id { get; set; } //  SKYPE ID
        /// <summary>
        /// 到職日
        /// </summary>
        public string r_online_date { get; set; } //  到職日   "2018-10-01",
        /// <summary>
        ///  Y是 在職: N是非在職 無法分辨是否留職停薪,  也是在職
        ///  離職後才會是N
        /// </summary>
        public string r_online { get; set; } //  Y是 在職: N是非在職 無法分辨是否留職停薪
        /// <summary>
        /// 離職日
        /// </summary>
        public string r_offline_date { get; set; } //  離職日




        //部門名稱
        [JsonProperty("DEPARTNAME")]
        public string Departname { get; set; }

        [JsonProperty("DUTYID", NullValueHandling = NullValueHandling.Ignore)]
        public string Dutyid { get; set; }
        [JsonProperty("WORKPLACEID", NullValueHandling = NullValueHandling.Ignore)]
        public string Workplaceid { get; set; }
        //工作地點名稱
        [JsonProperty("WORKPLACENAME", NullValueHandling = NullValueHandling.Ignore)]
        public string Workplacename { get; set; }

        //帳戶身分證號
        [JsonProperty("IDNO", NullValueHandling = NullValueHandling.Ignore)]
        public string Idno { get; set; }

        //銀行帳戶戶名
        [JsonProperty("SALACCOUNTNAME", NullValueHandling = NullValueHandling.Ignore)]
        public string Salaccountname { get; set; }

        //銀行代碼
        [JsonProperty("BANKCODE", NullValueHandling = NullValueHandling.Ignore)]
        public string Bankcode { get; set; }

        //銀行分行代碼
        [JsonProperty("BANKBRANCH", NullValueHandling = NullValueHandling.Ignore)]
        public string Bankbranch { get; set; }

        //銀行收款帳號
        [JsonProperty("SALACCOUNTID", NullValueHandling = NullValueHandling.Ignore)]
        public string Salaccountid { get; set; }

        // 在職狀態
        public string Jobstatus { get; set; }




    }


}
