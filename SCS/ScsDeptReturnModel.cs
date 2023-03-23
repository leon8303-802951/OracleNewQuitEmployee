using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.SCS
{

    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public class ScsDeptReturnModel
    {
        [JsonProperty("DataSet")]
        public DeptDataSet DataSet { get; set; }

        [JsonProperty("UpdateItems")]
        public object[] UpdateItems { get; set; }
    }

    public class DeptDataSet
    {
        [JsonProperty("HUM0010300")]
        public Hum0010300[] Hum0010300 { get; set; }
    }


    /// <summary>
    /// 部門詳細資料
    /// </summary>
    public class Hum0010300
    {
        /// <summary>
        /// SCS001
        /// </summary>
        [JsonProperty("SYS_COMPANYID")]
        public string SysCompanyid { get; set; }

        /// <summary>
        /// "SCS00120210628100005",
        /// </summary>
        [JsonProperty("SYS_ID")]
        public string SysId { get; set; }


        /// <summary>
        /// 部門編號
        /// </summary>
        [JsonProperty("SYS_VIEWID")]
        public string SysViewid { get; set; }

        /// <summary>
        /// 上階部門
        /// </summary>
        [JsonProperty("TMP_PDEPARTID")]
        public string TMP_PDEPARTID { get; set; }

        /// <summary>
        /// 虛擬部門(核決使用)
        /// </summary>
        [JsonProperty("VDEPART")]
        public bool Vdepart { get; set; }

        /// <summary>
        /// 部門名稱
        /// </summary>
        [JsonProperty("SYS_NAME")]
        public string SysName { get; set; }

        /// <summary>
        /// 英文名稱
        /// </summary>
        [JsonProperty("SYS_ENGNAME")]
        public string SysEngname { get; set; }

        /// <summary>
        /// 新增日期時間
        /// </summary>
        [JsonProperty("INS_DAT")]
        public string InsDat { get; set; }

        /// <summary>
        /// 新增的人
        /// </summary>
        [JsonProperty("INS_USR")]
        public string InsUsr { get; set; }

        /// <summary>
        /// 異動日期時間
        /// </summary>
        [JsonProperty("UPD_DAT")]
        public string UpdDat { get; set; }

        /// <summary>
        /// 異動的人
        /// </summary>
        [JsonProperty("UPD_USR")]
        public string UpdUsr { get; set; }

        /// <summary>
        /// 不明
        /// </summary>
        [JsonProperty("SYS_ISUSING")]
        public string SysIsusing { get; set; }

        /// <summary>
        /// 不明 不是所屬部門
        /// </summary>
        [JsonProperty("PDEPARTID")]
        public string Pdepartid { get; set; }

        /// <summary>
        /// 不明
        /// </summary>
        [JsonProperty("COSTCATEID")]
        public string Costcateid { get; set; }

        /// <summary>
        /// 不明
        /// </summary>
        [JsonProperty("PROFITID")]
        public string Profitid { get; set; }

        /// <summary>
        /// 不明
        /// </summary>
        [JsonProperty("MANAGERID")]
        public string Managerid { get; set; }

        /// <summary>
        /// 部門主管
        /// </summary>
        [JsonProperty("TMP_MANAGERID")]
        public string TmpManagerid { get; set; }

        /// <summary>
        /// 部門主管姓名
        /// </summary>
        [JsonProperty("TMP_MANAGERNAME")]
        public string TmpManagername { get; set; }


        /// <summary>
        /// 部門主管英文名字
        /// </summary>
        [JsonProperty("TMP_MANAGERENGNAME")]
        public string TmpManagerengname { get; set; }


        /// <summary>
        /// 不明
        /// </summary>
        [JsonProperty("DEPUTYMANAGERID")]
        public string Deputymanagerid { get; set; }

        /// <summary>
        /// 不明
        /// </summary>
        [JsonProperty("MDUTYID")]
        public string Mdutyid { get; set; }

        /// <summary>
        /// 編制人員
        /// </summary>
        [JsonProperty("BUDGETMEN")]
        public string Budgetmen { get; set; }

        /// <summary>
        /// 儲備人員
        /// </summary>
        [JsonProperty("BUDGETRESERVE")]
        public string Budgetreserve { get; set; }

        /// <summary>
        /// 編制計時
        /// </summary>
        [JsonProperty("BUDGETPT")]
        public string Budgetpt { get; set; }

        /// <summary>
        /// 實際計時
        /// </summary>
        [JsonProperty("REALPT")]
        public string Realpt { get; set; }

        /// <summary>
        /// 不明
        /// </summary>
        [JsonProperty("REALMEN")]
        public string Realmen { get; set; }

        /// <summary>
        /// 秘書1/姓名
        /// </summary>
        [JsonProperty("SECRETARY1")]
        public string Secretary1 { get; set; }

        /// <summary>
        /// 秘書2/姓名
        /// </summary>
        [JsonProperty("SECRETARY2")]
        public string Secretary2 { get; set; }

        /// <summary>
        /// 秘書3/姓名
        /// </summary>
        [JsonProperty("SECRETARY3")]
        public string Secretary3 { get; set; }

        [JsonProperty("CREATEDATE")]
        public string Createdate { get; set; }

        /// <summary>
        /// ** 生效日期
        /// </summary>
        [JsonProperty("STARTDATE")]
        public string Startdate { get; set; }

        /// <summary>
        /// ** 停用日期
        /// </summary>
        [JsonProperty("STOPDATE")]
        public string STOPDATE { get; set; }

        /// <summary>
        /// 等同部門
        /// </summary>
        [JsonProperty("TMP_SDEPARTID")]
        public string TMP_SDEPARTID { get; set; }




        /// <summary>
        /// 不明
        /// </summary>
        [JsonProperty("SDEPARTID")]
        public string Sdepartid { get; set; }

        /// <summary>
        /// 舊系統號
        /// </summary>
        [JsonProperty("SYS_VIEWIDOLD")]
        public string SysViewidold { get; set; }

        /// <summary>
        /// 申報公司 公司別
        /// 01  網家速配
        /// 02  連科通訊
        /// 03  國際連
        /// 04  基石創新
        /// 05  拍付
        /// 06  露天
        /// 06-1    露天職工福利委員會
        /// 07  網路家庭旅行社
        /// 08  易安網
        /// 09  雲坦數位
        /// 10  樂屋網
        /// 11  21世紀數位科技
        /// 12  網家數據科技
        /// 80  網家
        /// 90  集團
        /// </summary>
        [JsonProperty("TMP_DECCOMPANYID")]
        public string TMP_DECCOMPANYID { get; set; }

        /// <summary>
        /// 不明
        /// </summary>
        [JsonProperty("DECCOMPANYID")]
        public string Deccompanyid { get; set; }


        /// <summary>
        /// 部門階層
        /// </summary>
        [JsonProperty("ORGLEVEL")]
        public string Orglevel { get; set; }

        /// <summary>
        /// 識別證顯示名稱
        /// </summary>
        [JsonProperty("CARDNAME")]
        public string CARDNAME { get; set; }

        /// <summary>
        /// 直接/間接員工類型
        /// </summary>
        [JsonProperty("TMP_DIRINTYPEID")]
        public string TMP_DIRINTYPEID { get; set; }

        /// <summary>
        /// Report Org
        /// </summary>
        [JsonProperty("TMP_REPORTORGID")]
        public string TMP_REPORTORGID { get; set; }



        [JsonProperty("DIRINTYPEID")]
        public string Dirintypeid { get; set; }

        [JsonProperty("PROCATEGORYID")]
        public string Procategoryid { get; set; }

        [JsonProperty("MPROCATEGORYID")]
        public string Mprocategoryid { get; set; }

        [JsonProperty("SITEID")]
        public string Siteid { get; set; }

        [JsonProperty("WORKPLACEID")]
        public string Workplaceid { get; set; }

        [JsonProperty("SWIPEID")]
        public string Swipeid { get; set; }

        [JsonProperty("DEPHRBPID")]
        public string Dephrbpid { get; set; }

        [JsonProperty("ISHEAD")]
        public bool Ishead { get; set; }

        [JsonProperty("SIMPLENAME")]
        public string Simplename { get; set; }

        [JsonProperty("SIMPLEENGNAME")]
        public string Simpleengname { get; set; }

        [JsonProperty("REPORTORG")]
        public string Reportorg { get; set; }

        [JsonProperty("SAMELEVELORDERNO")]
        public string Samelevelorderno { get; set; }

        [JsonProperty("DEPARTCLASSID")]
        public string Departclassid { get; set; }

        [JsonProperty("SERIAL")]
        public string Serial { get; set; }

        [JsonProperty("ORGLONGNAME")]
        public string Orglongname { get; set; }

        [JsonProperty("ORGLONGENGNAME")]
        public string Orglongengname { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        [JsonProperty("NOTE")]
        public string Note { get; set; }

        [JsonProperty("APPROVELEAVE")]
        public string Approveleave { get; set; }

        /// <summary>
        /// 自訂欄位1
        /// </summary>
        [JsonProperty("SELFDEF1")]
        public string Selfdef1 { get; set; }

        /// <summary>
        /// 自訂欄位2
        /// 就是 Job Level
        /// 只要有填 Job Level 就有簽核權限
        /// (1, 2, 3, 4)
        /// </summary>
        [JsonProperty("SELFDEF2")]
        public string Selfdef2 { get; set; }

        /// <summary>
        /// 自訂欄位3
        /// 簽核主管若不是該部門主管，會另外寫在這裡
        /// 就是部門主管不具簽核權限，會用 自訂欄位3 來設定
        /// 自訂欄位3的優先權大於自訂欄位2
        /// </summary>
        [JsonProperty("SELFDEF3")]
        public string Selfdef3 { get; set; }

        /// <summary>
        /// 自訂欄位4
        /// </summary>
        [JsonProperty("SELFDEF4")]
        public string Selfdef4 { get; set; }

        /// <summary>
        /// 自訂欄位5
        /// </summary>
        [JsonProperty("SELFDEF5")]
        public string Selfdef5 { get; set; }


        /// <summary>
        /// 自訂欄位6
        /// </summary>
        [JsonProperty("SELFDEF6")]
        public string Selfdef6 { get; set; }

        /// <summary>
        /// 自訂欄位7
        /// </summary>
        [JsonProperty("SELFDEF7")]
        public string Selfdef7 { get; set; }

        /// <summary>
        /// 自訂欄位8
        /// </summary>
        [JsonProperty("SELFDEF8")]
        public string Selfdef8 { get; set; }

        /// <summary>
        /// 自訂欄位9
        /// </summary>
        [JsonProperty("SELFDEF9")]
        public string Selfdef9 { get; set; }


        /// <summary>
        /// 自訂欄位10
        /// </summary>
        [JsonProperty("SELFDEF10")]
        public string Selfdef10 { get; set; }
        /// <summary>
        /// 自訂欄位10
        /// </summary>
        [JsonProperty("TMP_SYS_DEPARTMENTID")]
        public string TMP_SYS_DEPARTMENTID { get; set; }


        [JsonProperty("BIZDEPARTID")]
        public string Bizdepartid { get; set; }

        [JsonProperty("SYS_DEPARTMENTID")]
        public string SysDepartmentid { get; set; }

        /// <summary>
        /// 部門批次異動單號
        /// </summary>
        [JsonProperty("SOURCEID")]
        public string Sourceid { get; set; }

        [JsonProperty("SYS_EDITPERMISSION")]
        public bool SysEditpermission { get; set; }

        [JsonProperty("SYS_DELETEPERMISSION")]
        public bool SysDeletepermission { get; set; }
    }





}
