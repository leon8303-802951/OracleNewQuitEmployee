using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Newtonsoft.Json;

namespace OracleNewQuitEmployee.SCS
{
    public class FtEmployees
    {
        [JsonProperty("ReportID")]
        public string ReportId { get; set; }

        [JsonProperty("DataSet")]
        public DataSet DataSet { get; set; }
    }

    public class DataSet
    {
        [JsonProperty("ReportHeader")]
        public ReportHeader[] ReportHeader { get; set; }

        [JsonProperty("ReportBody")]
        public ReportBody[] ReportBody { get; set; }
    }

    public class ReportBody
    {
        //部門編號  "r_dept": 
        [JsonProperty("DEPARTID")]
        public string Departid { get; set; }

        //部門名稱
        [JsonProperty("DEPARTNAME")]
        public string Departname { get; set; }

        //員工工號  r_code
        [JsonProperty("EMPLOYEEID")]
        public string Employeeid { get; set; }

        // 中文姓名  r_cname
        [JsonProperty("EMPLOYEENAME")]
        public string Employeename { get; set; }

        //英文姓名 r_ename
        [JsonProperty("EMPLOYEEENGNAME")]
        public string Employeeengname { get; set; }

        /// <summary>
        /// 在職狀態 r_online
        /// 0=未就職(在職)
        /// 1=試用(在職)
        /// 2=正式(在職)
        /// 3=約聘(在職)
        /// 4=留職停薪(在職)
        /// 5=離職(離職)
        /// </summary>
        [JsonProperty("JOBSTATUS")]
        public string Jobstatus { get; set; }

        //生日 r_birthday
        [JsonProperty("BIRTHDATE")]
        public string Birthdate { get; set; }

        //性別 r_sex
        [JsonProperty("SEX")]
        public string Sex { get; set; }

        //公司EMAIL r_email
        [JsonProperty("PSNEMAIL")]
        public string Psnemail { get; set; }

        //行動電話 r_cell_phone
        [JsonProperty("MOIBLE")]
        public string Moible { get; set; }

        // 分機 r_phone_ext
        [JsonProperty("OFFICETEL1")]
        public string Officetel1 { get; set; }

        // Skype ID  r_skype_id
        [JsonProperty("OFFICETEL2")]
        public string Officetel2 { get; set; }

        //到職日期  r_online_date
        [JsonProperty("STARTDATE")]
        public string Startdate { get; set; }

        //職務編號
        [JsonProperty("DUTYID", NullValueHandling = NullValueHandling.Ignore)]
        public string Dutyid { get; set; }

        //職務名稱  r_degress
        [JsonProperty("DUTYNAME", NullValueHandling = NullValueHandling.Ignore)]
        public string Dutyname { get; set; }

        //工作地點編號
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


        //離職日期  r_offline_date
        [JsonProperty("SEPARATIONDATE", NullValueHandling = NullValueHandling.Ignore)]
        public string Separationdate { get; set; }
    }

    public class ReportHeader
    {
        [JsonProperty("COMPANYNAME")]
        public string Companyname { get; set; }

        [JsonProperty("COMPANYCURRENCYID")]
        public string Companycurrencyid { get; set; }

        [JsonProperty("COMPANYCURRENCYNAME")]
        public string Companycurrencyname { get; set; }

        [JsonProperty("COMPOINTPRICE")]
        public string Compointprice { get; set; }

        [JsonProperty("COMPOINTMONEY")]
        public string Compointmoney { get; set; }

        [JsonProperty("COMPOINTCURRRATE")]
        public string Compointcurrrate { get; set; }

        [JsonProperty("REPORTTITLE")]
        public string Reporttitle { get; set; }

        [JsonProperty("REPORTTAIL")]
        public string Reporttail { get; set; }

        [JsonProperty("REPORTUSERID")]
        public string Reportuserid { get; set; }

        [JsonProperty("REPORTUSERNAME")]
        public string Reportusername { get; set; }

        [JsonProperty("REPORTFILTER")]
        public string Reportfilter { get; set; }
    }




}
