using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace OracleNewQuitEmployee.SCS
{
    public class SCSUtils
    {
        SCSLog log;
        public SCSUtils()
        {
            log = new SCSLog();
            string scs_acc_file = $"{AppDomain.CurrentDomain.BaseDirectory}Ini/scs.config";
            try
            {
                // 直接讀取帳密
                // 第一行 帳號
                // 第二行 密碼
                // 第三行 company id
                // 第四行 ap01 ip
                // 第五行 ap02 ip
                if (System.IO.File.Exists(scs_acc_file))
                {
                    string[] content = System.IO.File.ReadAllLines(scs_acc_file);
                    UserName = content[0].Trim();
                    Password = content[1].Trim();
                    CompanyID = content[2].Trim();
                    AP01_Host = content[3].Trim();
                    AP02_Host = content[4].Trim();

                }
                else
                {
                    throw new Exception($"找不到 {scs_acc_file}");
                }
            }
            catch (Exception ex)
            {
                log.WriteLogAsync($"建構式錯誤:{ex.Message}{ex.InnerException}");
                throw;
            }
        }

        private string UserName;
        private string Password;
        private string CompanyID;
        private string AP01_Host;
        private string AP02_Host;



        #region oracle 使用

        /// <summary>
        /// 給我 oracle job 用的
        /// </summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        public List<staff_infoMd> GetEmployeesForOracle(string guid)
        {
            List<staff_infoMd> rst = new List<staff_infoMd>();
            FtEmployees rrtn = null;
            try
            {               
                string host = AP02_Host;
                string api = "/SCSRwd/api/businessobject/";
                string url = $"https://{host}{api}";


                HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
                req.Method = "POST";
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls |
                     (SecurityProtocolType)768//.Tls11
                     | (SecurityProtocolType)3072; //.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                req.ContentType = "application/json";
                ScsEmployeesQueryModel model = new ScsEmployeesQueryModel();
                model.Action = "ExecReport";
                model.SessionGuid = guid;
                model.ProgId = "RHUM002";
                model.Value.Type = "AIS.Define.TExecReportInputArgs, AIS.Define";
                model.Value.UiType = "Report";
                model.Value.ReportId = "";
                model.Value.ReportTailId = "";
                model.Value.FilterItems = new FilterItem[] {
                    new FilterItem() {
                        Type= "AIS.Define.TFilterItem, AIS.Define",
                        FieldName="ZA.SYS_CompanyID",
                        FilterValue=CompanyID //"SCS999"
                    } };
                model.Value.UserFilter = "";

                /*
{
  "Action": "ExecReport",
  "SessionGuid": "48ca52a2-b1f4-487d-ab93-24a4327ad166",
  "ProgID": "RHUM002",
  "Value": {
    "$type": "AIS.Define.TExecReportInputArgs, AIS.Define",
    "UIType": "Report",
    "ReportID": "",
    "ReportTailID": "",
    "FilterItems": [
      {
        "$type": "AIS.Define.TFilterItem, AIS.Define",
        "FieldName": "ZA.SYS_CompanyID",
        "FilterValue": "SCS999"
      }
    ],
    "UserFilter": ""
  }
}
                 */

                string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                byte[] contentByte = Encoding.UTF8.GetBytes(jsonStr);
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(contentByte, 0, contentByte.Length);
                }
                using (WebResponse res = req.GetResponse())
                {
                    string statusCode = $"{((HttpWebResponse)res).StatusCode}";
                    string statusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                    using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                    {
                        string rtnjson = reader.ReadToEnd();
                        //if (_Logger != null)
                        //{
                        //    _Logger.LogInformation("登入飛騰取得:{responseStr}", rst);
                        //}
                        rrtn = Newtonsoft.Json.JsonConvert.DeserializeObject<FtEmployees>(rtnjson);
                        //rst = $"{rtnobj.SessionGuid}";
                        if (rrtn.DataSet.ReportBody != null)
                        {
                            foreach (ReportBody row in rrtn.DataSet.ReportBody)
                            {
                                /*
                              就是想確認一下
下列這些狀態 那些是在職 那些是離職
JobStatus
0 = 未就職
1 = 試用
2 = 正式
3 = 約聘
4 = 留職停薪
5 = 離職
*/
                                var job = $"{row.Jobstatus}";
                                staff_infoMd staff = Getstaff_infoMdByReportBody(row);
                                rst.Add(staff);

                                //rst.Add(new staff_infoMd()
                                //{
                                //    r_code = $"{row.Employeeid}",
                                //    r_cname = $"{row.Employeename}",
                                //    r_ename = $"{row.Employeeengname}",
                                //    r_dept = $"{row.Departid}",
                                //    r_degress = $"{row.Jobcodename}",
                                //    r_cell_phone = $"{row.Moible}",
                                //    r_birthday = ($"{row.Birthdate}".Length >= 10) ? $"{row.Birthdate}".Substring(0, 10) : "",
                                //    r_sex = ($"{row.Sex}" == "0") ? "男" : ($"{row.Sex}" == "1") ? "女" : "未定",
                                //    r_email = $"{row.Psnemail}",
                                //    r_phone_ext = $"{row.Officetel1}",
                                //    r_skype_id = $"{row.Officetel2}",
                                //    r_online_date = ($"{row.Startdate}".Length > 0) ? $"{row.Startdate}".Substring(0, 10) : "",
                                //    r_online = ($"{row.Jobstatus}" == "5") ? "離職" : "在職"
                                //    , r_offline_date = ($"{row.SeparationDate}".Length > 0) ? $"{row.SeparationDate}".Substring(0, 10) : ""
                                //}); 

                            }

                        }
                        else
                        {
                            throw new Exception("取得資料失敗");
                        }

                        //                        if (rrtn.DataSet != null && rrtn.DataSet.Tables.Count > 0)
                        //                        {

                        //                            foreach (DataRow row in rrtn.DataSet.Tables["ReportBody"].Rows)
                        //                            {
                        //                                /*
                        //                                就是想確認一下
                        //下列這些狀態 那些是在職 那些是離職
                        // JobStatus
                        //0 = 未就職
                        //1 = 試用
                        //2 = 正式
                        //3 = 約聘
                        //4 = 留職停薪
                        //5 = 離職
                        //*/
                        //                                var job = $"{row["JobStatus"]}";

                        //                                rst.Add(new staff_infoMd()
                        //                                {
                        //                                    r_code = $"{row["EmployeeID"]}",
                        //                                    r_cname = $"{row["EmployeeNAME"]}",
                        //                                    r_ename = $"{row["EmployeeENGNAME"]}",
                        //                                    r_dept = $"{row["DEPARTID"]}",
                        //                                    r_degress = $"{row["JobCodeName"]}",
                        //                                    r_cell_phone = $"{row["Moible"]}",
                        //                                    r_birthday = ($"{row["Birthdate"]}".Length >= 10) ? $"{row["Birthdate"]}".Substring(0, 10) : "",
                        //                                    r_sex = ($"{row["Sex"]}" == "0") ? "男" : ($"{row["Sex"]}" == "1") ? "女" : "未定",
                        //                                    r_email = $"{row["PsnEMail"]}",
                        //                                    r_phone_ext = $"{row["OFFICETEL1"]}",
                        //                                    r_skype_id = $"{row["OFFICETEL2"]}",
                        //                                    r_online_date = ($"{row["StartDate"]}".Length > 0) ? $"{row["StartDate"]}".Substring(0, 10) : "",
                        //                                    r_online = ($"{row["JobStatus"]}" == "5") ? "離職" : "在職",
                        //                                    r_offline_date = ($"{row["SeparationDate"]}".Length > 0) ? $"{row["SeparationDate"]}".Substring(0, 10) : ""
                        //                                }); ;
                        //                            }

                        //                            /*
                        //                            Celine
                        //請問一下
                        //請問在職 是不是包含以下狀態的 試用 正式 約聘 留職停薪 ?

                        //0 = 未就職
                        //1 = 試用
                        //2 = 正式
                        //3 = 約聘
                        //4 = 留職停薪
                        //5 = 離職
                        //你這個是撈那個報表? 因為如果是"在職"，應該就不會有離職?
                        //         public string r_code { get; set; } // 員編
                        //         public string r_cname { get; set; } // 中文姓名
                        //         public string r_ename { get; set; } // 英文姓名
                        //         public string r_dept { get; set; } // 部門
                        //         public string r_degress { get; set; } // 職級
                        //         public string r_cell_phone { get; set; } // 手機號碼
                        //         public string r_birthday { get; set; } // 生日
                        //         public string r_sex { get; set; } // 1是男  2是女
                        //         public string r_email { get; set; } // EMAIL
                        //         public string r_phone_ext { get; set; } // 分機
                        //         public string r_skype_id { get; set; } // SKYPE ID
                        //         public string r_online_date { get; set; } // 到職日
                        //         public string r_online { get; set; } // Y是 在職: N是非在職 無法分辨是否留職停薪
                        //         public string r_offline_date { get; set; } // 離職日
                        //                               */

                        //                        }
                        //                        else
                        //                        {
                        //                            throw new Exception("取得資料失敗");
                        //                        }


                    }
                }
            }
            catch (WebException wx)
            {
                StringBuilder errmsg = new StringBuilder($"{wx.Message}{wx.InnerException}");
                if (wx.Response != null)
                {
                    using (WebResponse res = wx.Response)
                    {
                        HttpWebResponse hres = res as HttpWebResponse;
                        string statusCode = $"{hres.StatusCode}";
                        string statusDescription = $"{hres.StatusDescription}";
                        try
                        {
                            using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                            {
                                errmsg.AppendLine(reader.ReadToEnd());
                            }
                        }
                        catch (Exception ex2)
                        {
                            string msg = ex2.Message;
                            //_Logger.LogError(ex2, "SCS_utils 嘗試取得錯誤訊息失敗:{msg}", msg);
                        }
                    }
                }

                throw new Exception($"{errmsg}");
            }
            catch (Exception ex)
            {
                throw;
            }

            return rst;
        }

        /// <summary>
        /// 給我 oracle job 用的
        /// </summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        public List<staff_infoMd> GetEmployeesForOracleAsync(string guid)
        {
            List<staff_infoMd> rst = new List<staff_infoMd>();
            FtEmployees rrtn = null;
            try
            {                
                string host = AP02_Host; 
                string api = "/SCSRwd/api/businessobject/";
                string url = $"https://{host}{api}";


                HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
                req.Method = "POST";
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls |
                     (SecurityProtocolType)768 //.Tls11 
                     | (SecurityProtocolType)3072; //.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                req.ContentType = "application/json";
                ScsEmployeesQueryModel model = new ScsEmployeesQueryModel();
                model.Action = "ExecReport";
                model.SessionGuid = guid;
                model.ProgId = "RHUM002";
                model.Value.Type = "AIS.Define.TExecReportInputArgs, AIS.Define";
                model.Value.UiType = "Report";
                model.Value.ReportId = "";
                model.Value.ReportTailId = "";
                model.Value.FilterItems = new FilterItem[] {
                    new FilterItem() {
                        Type= "AIS.Define.TFilterItem, AIS.Define",
                        FieldName="ZA.SYS_CompanyID",
                        FilterValue=CompanyID //"SCS999"
                    } };
                model.Value.UserFilter = "";

                /*
{
  "Action": "ExecReport",
  "SessionGuid": "48ca52a2-b1f4-487d-ab93-24a4327ad166",
  "ProgID": "RHUM002",
  "Value": {
    "$type": "AIS.Define.TExecReportInputArgs, AIS.Define",
    "UIType": "Report",
    "ReportID": "",
    "ReportTailID": "",
    "FilterItems": [
      {
        "$type": "AIS.Define.TFilterItem, AIS.Define",
        "FieldName": "ZA.SYS_CompanyID",
        "FilterValue": "SCS999"
      }
    ],
    "UserFilter": ""
  }
}
                 */

                string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                byte[] contentByte = Encoding.UTF8.GetBytes(jsonStr);
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(contentByte, 0, contentByte.Length);
                }
                using (WebResponse res = req.GetResponse())
                {
                    string statusCode = $"{((HttpWebResponse)res).StatusCode}";
                    string statusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                    using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                    {
                        string rtnjson = reader.ReadToEnd();
                        //string rtnjson = reader.ReadToEnd();
                        //if (_Logger != null)
                        //{
                        //    _Logger.LogInformation("登入飛騰取得:{responseStr}", rst);
                        //}
                        rrtn = Newtonsoft.Json.JsonConvert.DeserializeObject<FtEmployees>(rtnjson);
                        //rst = $"{rtnobj.SessionGuid}";
                        if (rrtn.DataSet.ReportBody != null)
                        {
                            foreach (ReportBody row in rrtn.DataSet.ReportBody)
                            {
                                /*
                              就是想確認一下
下列這些狀態 那些是在職 那些是離職
JobStatus
0 = 未就職
1 = 試用
2 = 正式
3 = 約聘
4 = 留職停薪
5 = 離職
*/
                                var job = $"{row.Jobstatus}";
                                staff_infoMd staff = Getstaff_infoMdByReportBody(row);
                                rst.Add(staff);

                                //rst.Add(new staff_infoMd()
                                //{
                                //    r_code = $"{row.Employeeid}",
                                //    r_cname = $"{row.Employeename}",
                                //    r_ename = $"{row.Employeeengname}",
                                //    r_dept = $"{row.Departid}",
                                //    r_degress = $"{row.Jobcodename}",
                                //    r_cell_phone = $"{row.Moible}",
                                //    r_birthday = ($"{row.Birthdate}".Length >= 10) ? $"{row.Birthdate}".Substring(0, 10) : "",
                                //    r_sex = ($"{row.Sex}" == "0") ? "男" : ($"{row.Sex}" == "1") ? "女" : "未定",
                                //    r_email = $"{row.Psnemail}",
                                //    r_phone_ext = $"{row.Officetel1}",
                                //    r_skype_id = $"{row.Officetel2}",
                                //    r_online_date = ($"{row.Startdate}".Length > 0) ? $"{row.Startdate}".Substring(0, 10) : "",
                                //    r_online = ($"{row.Jobstatus}" == "5") ? "離職" : "在職"
                                //    , r_offline_date = ($"{row.SeparationDate}".Length > 0) ? $"{row.SeparationDate}".Substring(0, 10) : ""
                                //}); 

                            }

                        }
                        else
                        {
                            throw new Exception("取得資料失敗");
                        }

                        //                        if (rrtn.DataSet != null && rrtn.DataSet.Tables.Count > 0)
                        //                        {

                        //                            foreach (DataRow row in rrtn.DataSet.Tables["ReportBody"].Rows)
                        //                            {
                        //                                /*
                        //                                就是想確認一下
                        //下列這些狀態 那些是在職 那些是離職
                        // JobStatus
                        //0 = 未就職
                        //1 = 試用
                        //2 = 正式
                        //3 = 約聘
                        //4 = 留職停薪
                        //5 = 離職
                        //*/
                        //                                var job = $"{row["JobStatus"]}";

                        //                                rst.Add(new staff_infoMd()
                        //                                {
                        //                                    r_code = $"{row["EmployeeID"]}",
                        //                                    r_cname = $"{row["EmployeeNAME"]}",
                        //                                    r_ename = $"{row["EmployeeENGNAME"]}",
                        //                                    r_dept = $"{row["DEPARTID"]}",
                        //                                    r_degress = $"{row["JobCodeName"]}",
                        //                                    r_cell_phone = $"{row["Moible"]}",
                        //                                    r_birthday = ($"{row["Birthdate"]}".Length >= 10) ? $"{row["Birthdate"]}".Substring(0, 10) : "",
                        //                                    r_sex = ($"{row["Sex"]}" == "0") ? "男" : ($"{row["Sex"]}" == "1") ? "女" : "未定",
                        //                                    r_email = $"{row["PsnEMail"]}",
                        //                                    r_phone_ext = $"{row["OFFICETEL1"]}",
                        //                                    r_skype_id = $"{row["OFFICETEL2"]}",
                        //                                    r_online_date = ($"{row["StartDate"]}".Length > 0) ? $"{row["StartDate"]}".Substring(0, 10) : "",
                        //                                    r_online = ($"{row["JobStatus"]}" == "5") ? "離職" : "在職",
                        //                                    r_offline_date = ($"{row["SeparationDate"]}".Length > 0) ? $"{row["SeparationDate"]}".Substring(0, 10) : ""
                        //                                }); ;
                        //                            }

                        //                            /*
                        //                            Celine
                        //請問一下
                        //請問在職 是不是包含以下狀態的 試用 正式 約聘 留職停薪 ?

                        //0 = 未就職
                        //1 = 試用
                        //2 = 正式
                        //3 = 約聘
                        //4 = 留職停薪
                        //5 = 離職
                        //你這個是撈那個報表? 因為如果是"在職"，應該就不會有離職?
                        //         public string r_code { get; set; } // 員編
                        //         public string r_cname { get; set; } // 中文姓名
                        //         public string r_ename { get; set; } // 英文姓名
                        //         public string r_dept { get; set; } // 部門
                        //         public string r_degress { get; set; } // 職級
                        //         public string r_cell_phone { get; set; } // 手機號碼
                        //         public string r_birthday { get; set; } // 生日
                        //         public string r_sex { get; set; } // 1是男  2是女
                        //         public string r_email { get; set; } // EMAIL
                        //         public string r_phone_ext { get; set; } // 分機
                        //         public string r_skype_id { get; set; } // SKYPE ID
                        //         public string r_online_date { get; set; } // 到職日
                        //         public string r_online { get; set; } // Y是 在職: N是非在職 無法分辨是否留職停薪
                        //         public string r_offline_date { get; set; } // 離職日
                        //                               */

                        //                        }
                        //                        else
                        //                        {
                        //                            throw new Exception("取得資料失敗");
                        //                        }


                    }
                }
            }
            catch (WebException wx)
            {
                StringBuilder errmsg = new StringBuilder($"{wx.Message}{wx.InnerException}");
                if (wx.Response != null)
                {
                    using (WebResponse res = wx.Response)
                    {
                        HttpWebResponse hres = res as HttpWebResponse;
                        string statusCode = $"{hres.StatusCode}";
                        string statusDescription = $"{hres.StatusDescription}";
                        try
                        {
                            using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                            {
                                errmsg.AppendLine(reader.ReadToEnd());
                            }
                        }
                        catch (Exception ex2)
                        {
                            string msg = ex2.Message;
                            //_Logger.LogError(ex2, "SCS_utils 嘗試取得錯誤訊息失敗:{msg}", msg);
                        }
                    }
                }

                throw new Exception($"{errmsg}");
            }
            catch (Exception ex)
            {
                throw;
            }

            return rst;
        }


        // 用ReportBody產生staff_infoMd物件
        public staff_infoMd Getstaff_infoMdByReportBody(ReportBody row)
        {
            return new staff_infoMd()
            {
                r_code = $"{row.Employeeid}",
                r_cname = $"{row.Employeename}",
                r_ename = $"{row.Employeeengname}",
                r_dept = $"{row.Departid}",
                r_degress = $"{row.Dutyname}",
                r_cell_phone = $"{row.Moible}",
                r_birthday = ($"{row.Birthdate}".Length >= 10) ? $"{row.Birthdate}".Substring(0, 10) : "",
                r_sex = ($"{row.Sex}" == "0") ? "男" : ($"{row.Sex}" == "1") ? "女" : "未定",
                r_email = $"{row.Psnemail}",
                r_phone_ext = $"{row.Officetel1}",
                r_skype_id = $"{row.Officetel2}",
                r_online_date = ($"{row.Startdate}".Length > 0) ? $"{row.Startdate}".Substring(0, 10) : "",
                r_online = ($"{row.Jobstatus}" == "5") ? "離職" : "在職",
                r_offline_date = ($"{row.Separationdate}".Length > 0) ? $"{row.Separationdate}".Substring(0, 10) : ""
            };
        }

        /// <summary>
        /// 取一個員工給我 oracle job 用的
        /// </summary>
        /// <param name="guid"></param>
        /// <param name="EmpNo"></param>
        /// <returns></returns>
        public List<staff_infoMd> GetEmployeesByEmpNoForOracle(string guid, string EmpNo)
        {
            List<staff_infoMd> rst = new List<staff_infoMd>();
            FtEmployees rrtn = null;
            try
            {
                
                string host = AP02_Host; 
                string api = "/SCSRwd/api/businessobject/";
                string url = $"https://{host}{api}";


                HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
                req.Method = "POST";
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls |
                     (SecurityProtocolType)768 //.Tls11 
                     | (SecurityProtocolType)3072; //.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                req.ContentType = "application/json";
                ScsEmployeesQueryModel model = new ScsEmployeesQueryModel();
                model.Action = "ExecReport";
                model.SessionGuid = guid;
                model.ProgId = "RHUM002";
                model.Value.Type = "AIS.Define.TExecReportInputArgs, AIS.Define";
                model.Value.UiType = "Report";
                model.Value.ReportId = "";
                model.Value.ReportTailId = "";
                model.Value.FilterItems = new FilterItem[] {
                    new FilterItem() {
                        Type= "AIS.Define.TFilterItem, AIS.Define",
                        FieldName="ZA.SYS_CompanyID",
                        FilterValue=CompanyID // "SCS999"
                    },
                    new FilterItem() {
                        Type= "AIS.Define.TFilterItem, AIS.Define",
                        FieldName="ZA.EmployeeID",
                        FilterValue=EmpNo
                    }
                };
                model.Value.UserFilter = "";

                /*
{
  "Action": "ExecReport",
  "SessionGuid": "48ca52a2-b1f4-487d-ab93-24a4327ad166",
  "ProgID": "RHUM002",
  "Value": {
    "$type": "AIS.Define.TExecReportInputArgs, AIS.Define",
    "UIType": "Report",
    "ReportID": "",
    "ReportTailID": "",
    "FilterItems": [
      {
        "$type": "AIS.Define.TFilterItem, AIS.Define",
        "FieldName": "ZA.SYS_CompanyID",
        "FilterValue": "SCS999"
      }
    ],
    "UserFilter": ""
  }
}
                 */

                string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                byte[] contentByte = Encoding.UTF8.GetBytes(jsonStr);
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(contentByte, 0, contentByte.Length);
                }
                using (WebResponse res = req.GetResponse())
                {
                    string statusCode = $"{((HttpWebResponse)res).StatusCode}";
                    string statusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                    using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                    {
                        string rtnjson = reader.ReadToEnd();
                        //if (_Logger != null)
                        //{
                        //    _Logger.LogInformation("登入飛騰取得:{responseStr}", rst);
                        //}
                        rrtn = Newtonsoft.Json.JsonConvert.DeserializeObject<FtEmployees>(rtnjson);
                        //rst = $"{rtnobj.SessionGuid}";

                        if (rrtn.DataSet.ReportBody != null)
                        {
                            foreach (ReportBody row in rrtn.DataSet.ReportBody)
                            {
                                /*
                              就是想確認一下
下列這些狀態 那些是在職 那些是離職
JobStatus
0 = 未就職
1 = 試用
2 = 正式
3 = 約聘
4 = 留職停薪
5 = 離職
*/
                                var job = $"{row.Jobstatus}";
                                staff_infoMd staff = Getstaff_infoMdByReportBody(row);
                                rst.Add(staff);

                            }

                        }
                        else
                        {
                            throw new Exception("取得資料失敗");
                        }

                        //                        if (rrtn.DataSet != null && rrtn.DataSet.Tables.Count > 0)
                        //                        {

                        //                            foreach (DataRow row in rrtn.DataSet.Tables["ReportBody"].Rows)
                        //                            {
                        //                                /*
                        //                                就是想確認一下
                        //下列這些狀態 那些是在職 那些是離職
                        // JobStatus
                        //0 = 未就職
                        //1 = 試用
                        //2 = 正式
                        //3 = 約聘
                        //4 = 留職停薪
                        //5 = 離職
                        //*/
                        //                                var job = $"{row["JobStatus"]}";

                        //                                rst.Add(new staff_infoMd()
                        //                                {
                        //                                    r_code = $"{row["EmployeeID"]}",
                        //                                    r_cname = $"{row["EmployeeNAME"]}",
                        //                                    r_ename = $"{row["EmployeeENGNAME"]}",
                        //                                    r_dept = $"{row["DEPARTID"]}",
                        //                                    r_degress = $"{row["JobCodeName"]}",
                        //                                    r_cell_phone = $"{row["Moible"]}",
                        //                                    r_birthday = ($"{row["Birthdate"]}".Length >= 10) ? $"{row["Birthdate"]}".Substring(0, 10) : "",
                        //                                    r_sex = ($"{row["Sex"]}" == "0") ? "男" : ($"{row["Sex"]}" == "1") ? "女" : "未定",
                        //                                    r_email = $"{row["PsnEMail"]}",
                        //                                    r_phone_ext = $"{row["OFFICETEL1"]}",
                        //                                    r_skype_id = $"{row["OFFICETEL2"]}",
                        //                                    r_online_date = ($"{row["StartDate"]}".Length > 0) ? $"{row["StartDate"]}".Substring(0, 10) : "",
                        //                                    r_online = ($"{row["JobStatus"]}" == "5") ? "離職" : "在職",
                        //                                    r_offline_date = ($"{row["SeparationDate"]}".Length > 0) ? $"{row["SeparationDate"]}".Substring(0, 10) : ""
                        //                                }); ;
                        //                            }

                        //                            /*
                        //                            Celine
                        //請問一下
                        //請問在職 是不是包含以下狀態的 試用 正式 約聘 留職停薪 ?

                        //0 = 未就職
                        //1 = 試用
                        //2 = 正式
                        //3 = 約聘
                        //4 = 留職停薪
                        //5 = 離職
                        //你這個是撈那個報表? 因為如果是"在職"，應該就不會有離職?
                        //         public string r_code { get; set; } // 員編
                        //         public string r_cname { get; set; } // 中文姓名
                        //         public string r_ename { get; set; } // 英文姓名
                        //         public string r_dept { get; set; } // 部門
                        //         public string r_degress { get; set; } // 職級
                        //         public string r_cell_phone { get; set; } // 手機號碼
                        //         public string r_birthday { get; set; } // 生日
                        //         public string r_sex { get; set; } // 1是男  2是女
                        //         public string r_email { get; set; } // EMAIL
                        //         public string r_phone_ext { get; set; } // 分機
                        //         public string r_skype_id { get; set; } // SKYPE ID
                        //         public string r_online_date { get; set; } // 到職日
                        //         public string r_online { get; set; } // Y是 在職: N是非在職 無法分辨是否留職停薪
                        //         public string r_offline_date { get; set; } // 離職日
                        //                               */

                        //                        }
                        //                        else
                        //                        {
                        //                            throw new Exception("取得資料失敗");
                        //                        }

                    }
                }
            }
            catch (WebException wx)
            {
                StringBuilder errmsg = new StringBuilder($"{wx.Message}{wx.InnerException}");
                if (wx.Response != null)
                {
                    using (WebResponse res = wx.Response)
                    {
                        HttpWebResponse hres = res as HttpWebResponse;
                        string statusCode = $"{hres.StatusCode}";
                        string statusDescription = $"{hres.StatusDescription}";
                        try
                        {
                            using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                            {
                                errmsg.AppendLine(reader.ReadToEnd());
                            }
                        }
                        catch (Exception ex2)
                        {
                            string msg = ex2.Message;
                            //_Logger.LogError(ex2, "SCS_utils 嘗試取得錯誤訊息失敗:{msg}", msg);
                        }
                    }
                }

                throw new Exception($"{errmsg}");
            }
            catch (Exception ex)
            {
                throw;
            }

            return rst;
        }

        /// <summary>
        /// 取得單筆部門給 oracle 用
        /// </summary>
        /// <param name="guid"></param>
        /// <param name="DeptNo"></param>
        /// <returns></returns>
        public deptMd GetDeptByDeptNoForOracle(string guid, string DeptNo)
        {
            deptMd deptObj = null;
            ScsDeptsReturnModel Depts = GetDeptByDeptNoButNoDetail(guid, DeptNo);
            foreach (var dept in Depts.DataTable)
            {
                var deptdetails = GetDeptDetail(guid, dept.SysId);
                if (deptdetails.DataSet.Hum0010300.Length > 0)
                {
                    //deptlist.Add(deptdetails.DataSet.Hum0010300[0]);
                    deptObj = new deptMd()
                    {
                        r_dept_code = deptdetails.DataSet.Hum0010300[0].SysViewid,
                        r_cname = deptdetails.DataSet.Hum0010300[0].SysName,
                        r_belong = deptdetails.DataSet.Hum0010300[0].TMP_PDEPARTID
                    };
                }
            }
            return deptObj;
        }

        /// <summary>
        /// 取部門資料給 oracle 用
        /// </summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        public List<deptMd> GetDeptsForOracle(string guid)
        {
            string host = AP02_Host;
            List<deptMd> rst = new List<deptMd>();
            List<Hum0010300> deptlist = new List<Hum0010300>();
            ScsDeptsReturnModel Depts = GetDepts(guid);
            foreach (var dept in Depts.DataTable)
            {
                var deptdetails = GetDeptDetail(guid, dept.SysId);
                if (deptdetails.DataSet.Hum0010300.Length > 0)
                {
                    //deptlist.Add(deptdetails.DataSet.Hum0010300[0]);
                    rst.Add(
                        new deptMd()
                        {
                            r_dept_code = deptdetails.DataSet.Hum0010300[0].SysViewid,
                            r_cname = deptdetails.DataSet.Hum0010300[0].SysName,
                            r_belong = deptdetails.DataSet.Hum0010300[0].TMP_PDEPARTID
                        }
                        ); ;
                }
            }
            return rst;
        }
        #endregion

        #region 欄位對應

        /// <summary>
        /// 部門欄位對應檔 一定要是 utf8 檔案
        /// </summary>
        /// <returns></returns>
        public List<NameValueMapping> DeptColumnNameMapping()
        {
            string file_name = "dept_mapping.csv";
            return ColumnNameMapping(file_name);
        }

        /// <summary>
        ///  員工欄位對應檔 一定要是 utf8 檔案
        /// </summary>
        /// <returns></returns>
        public List<NameValueMapping> EmployeeColumnNameMapping()
        {
            string file_name = "employee_mapping.csv";
            return ColumnNameMapping(file_name);
        }

        public List<NameValueMapping> ColumnNameMapping(string file_name)
        {
            List<NameValueMapping> mapping = new List<NameValueMapping>();
            try
            {
                //string file_name = "employee_mapping.csv";
                string[] content = System.IO.File.ReadAllLines(file_name);
                string[] eng = content[0].Split(',');
                string[] ch = content[1].Split(',');
                for (int idx = 0; idx < eng.Length; idx++)
                {
                    mapping.Add(
                        new NameValueMapping()
                        {
                            EngFieldName = eng[idx],
                            ChFieldName = ch[idx]
                        });
                }
            }
            catch (Exception ex)
            {
            }
            return mapping;
        }

        #endregion

        /// <summary>
        /// 內部使用, 取得單筆部門 詳細資料 
        /// </summary>
        /// <param name="guid"></param>
        /// <param name="FormID"></param>
        /// <returns></returns>
        public ScsDeptReturnModel GetDeptDetail(string guid, string FormID)
        {
            string host = AP02_Host;
            ScsDeptReturnModel rst = null;
            ScsDeptQueryModel model = new ScsDeptQueryModel();
            model.SessionGuid = guid;
            model.ProgId = SCS_Const.ProgID; // "HUM0010300";
            model.Value.Type = SCS_Const.DeptDetailType; //  "AIS.Define.TMoveInputArgs, AIS.Define"
            model.Value.FormId = FormID; //sysid
            model.Value.SystemFilterOptions = SCS_Const.Dept_SystemFilterOptions; // "Session, DataPermission, EmployeeLevel"

            string json_str = Newtonsoft.Json.JsonConvert.SerializeObject(model);
            byte[] content_bytes = Encoding.UTF8.GetBytes(json_str);

          

            string api = "/SCSRwd/api/businessobject/";
            string url = $"https://{host}{api}";


            HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
            req.Timeout = 1000 * 60 * 30;//30分鐘
            req.Method = "POST";
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls |
                     (SecurityProtocolType)768 //.Tls11 
                     | (SecurityProtocolType)3072; //.Tls12;
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            req.ContentType = "application/json";

            string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(model);
            byte[] contentByte = Encoding.UTF8.GetBytes(jsonStr);
            using (Stream reqStream = req.GetRequestStream())
            {
                reqStream.Write(contentByte, 0, contentByte.Length);
            }
            using (WebResponse res = req.GetResponse())
            {
                string statusCode = $"{((HttpWebResponse)res).StatusCode}";
                string statusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                {
                    string rtnjson = reader.ReadToEnd();
                    log.WriteLogAsync($"取飛騰部門資料:{Environment.NewLine}{rtnjson}");
                    rst = Newtonsoft.Json.JsonConvert.DeserializeObject<ScsDeptReturnModel>(rtnjson);

                    // 看看哪些部門被停用了
                    //var stop_depts = from dept in rst.DataSet.Hum0010300
                    //                 where !string.IsNullOrWhiteSpace(dept.STOPDATE)
                    //                 select dept;

                    //if (_Logger != null)
                    //{
                    //    _Logger.LogInformation("登入飛騰取得:{responseStr}", rst);
                    //}
                    //rst = Newtonsoft.Json.JsonConvert.DeserializeObject<FtEmployees>(rtnjson);
                    ////rst = $"{rtnobj.SessionGuid}";
                    //if (rst.DataSet != null && rst.DataSet.Tables.Count > 0)
                    //{
                    //    return rst.DataSet.Tables["ReportBody"];
                    //}
                    //else
                    //{
                    //    throw new Exception("取得資料失敗");
                    //}
                }
            }


            return rst;


        }



        #region 外部使用


        /// <summary>
        /// 外部使用, 用 List 顯示部門清單詳細資料
        /// </summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        public List<Hum0010300> GetDeptsDetails(string guid)
        {
            string host = AP02_Host;
            List<Hum0010300> deptlist = new List<Hum0010300>();
            ScsDeptsReturnModel Depts = GetDepts(guid);
            foreach (var dept in Depts.DataTable)
            {
                var deptdetails = GetDeptDetail(guid, dept.SysId);
                if (deptdetails == null)
                {
                    throw new Exception("取得飛騰部門資料失敗!");
                }
                if (deptdetails.DataSet == null)
                {
                    throw new Exception("取得飛騰部門資料失敗! DataSet");
                }

                if (deptdetails.DataSet.Hum0010300.Length > 0)
                {
                    deptlist.Add(deptdetails.DataSet.Hum0010300[0]);
                }
            }

            return deptlist;
        }

        #endregion

        /// <summary>
        /// 內部使用, 取得[部門]SysID清單, 沒有詳細資料 
        /// </summary>
        /// <param name="host"></param>
        /// <param name="guid"></param>
        /// <returns></returns>
        public ScsDeptsReturnModel GetDepts(string guid)
        {
            string host = AP02_Host;
            /*
             {
  "Action": "Find",
  "SessionGuid": "a6c730bc-184d-4ef2-a4d7-64ef911384f9",
  "ProgID": "HUM0010300",
  "Value": {
    "$type": "AIS.Define.TFindInputArgs, AIS.Define",
    "SelectFields": "SYS_VIEWID,SYS_NAME,SYS_ENGNAME",
    "FilterItems": [],
    "SystemFilterOptions": "Session, DataPermission, EmployeeLevel",
    "IsBuildSelectedField": true,
    "IsBuildFlowLightSignalField": true
  }
}
             */

            ScsDeptsReturnModel rst = null;
            ScsDeptsQueryModel model = new ScsDeptsQueryModel();
            model.Action = "Find";
            model.SessionGuid = guid;
            model.ProgId = SCS_Const.ProgID;//   "HUM0010300",
            model.Value.Type = SCS_Const.DeptDetailType;// "AIS.Define.TFindInputArgs, AIS.Define";
            model.Value.SelectFields = SCS_Const.Depts_SelectFields; // "SYS_VIEWID,SYS_NAME,SYS_ENGNAME",
            model.Value.FilterItems = new object[] { };
            model.Value.SystemFilterOptions = SCS_Const.Depts_SystemFilterOptions;// "Session, DataPermission, EmployeeLevel";
            model.Value.IsBuildSelectedField = true;
            model.Value.IsBuildFlowLightSignalField = true;
            string json_str = Newtonsoft.Json.JsonConvert.SerializeObject(model);
            byte[] content_bytes = Encoding.UTF8.GetBytes(json_str);

            

            string api = "/SCSRwd/api/businessobject/";
            string url = $"https://{host}{api}";


            HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
            req.Method = "POST";
            req.Timeout = 1000 * 60 * 30; //30分鐘
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls |
                     (SecurityProtocolType)768 //.Tls11 
                     | (SecurityProtocolType)3072; //.Tls12;
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            req.ContentType = "application/json";

            string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(model);
            byte[] contentByte = Encoding.UTF8.GetBytes(jsonStr);
            using (Stream reqStream = req.GetRequestStream())
            {
                reqStream.Write(contentByte, 0, contentByte.Length);
            }
            using (WebResponse res = req.GetResponse())
            {
                string statusCode = $"{((HttpWebResponse)res).StatusCode}";
                string statusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                {
                    string rtnjson = reader.ReadToEnd();

                    rst = Newtonsoft.Json.JsonConvert.DeserializeObject<ScsDeptsReturnModel>(rtnjson);

                    //if (_Logger != null)
                    //{
                    //    _Logger.LogInformation("登入飛騰取得:{responseStr}", rst);
                    //}
                    //rst = Newtonsoft.Json.JsonConvert.DeserializeObject<FtEmployees>(rtnjson);
                    ////rst = $"{rtnobj.SessionGuid}";
                    //if (rst.DataSet != null && rst.DataSet.Tables.Count > 0)
                    //{
                    //    return rst.DataSet.Tables["ReportBody"];
                    //}
                    //else
                    //{
                    //    throw new Exception("取得資料失敗");
                    //}
                }
            }


            return rst;
        }


        /// <summary>
        /// 內部使用, 取得[單筆部門]SysID, 沒有詳細資料 
        /// </summary>
        /// <param name="guid"></param>
        /// <param name="DeptNo"></param>
        /// <returns></returns>
        public ScsDeptsReturnModel GetDeptByDeptNoButNoDetail(string guid, string DeptNo)
        {
            string host = AP02_Host;
            /*
             {
  "Action": "Find",
  "SessionGuid": "a6c730bc-184d-4ef2-a4d7-64ef911384f9",
  "ProgID": "HUM0010300",
  "Value": {
    "$type": "AIS.Define.TFindInputArgs, AIS.Define",
    "SelectFields": "SYS_VIEWID,SYS_NAME,SYS_ENGNAME",
    "FilterItems": [],
    "SystemFilterOptions": "Session, DataPermission, EmployeeLevel",
    "IsBuildSelectedField": true,
    "IsBuildFlowLightSignalField": true
  }
}
             */

            ScsDeptsReturnModel rst = null;
            ScsDeptsQueryModel model = new ScsDeptsQueryModel();
            model.Action = "Find";
            model.SessionGuid = guid;
            model.ProgId = SCS_Const.ProgID;//   "HUM0010300",
            model.Value.Type = SCS_Const.DeptDetailType;// "AIS.Define.TFindInputArgs, AIS.Define";
            model.Value.SelectFields = SCS_Const.Depts_SelectFields; // "SYS_VIEWID,SYS_NAME,SYS_ENGNAME",
            model.Value.FilterItems = new FilterItem[] {
                    new FilterItem() {
                        Type= "AIS.Define.TFilterItem, AIS.Define",
                        FieldName="SYS_VIEWID",
                        FilterValue=DeptNo
                    }
                };
            model.Value.SystemFilterOptions = SCS_Const.Depts_SystemFilterOptions;// "Session, DataPermission, EmployeeLevel";
            model.Value.IsBuildSelectedField = true;
            model.Value.IsBuildFlowLightSignalField = true;
            string json_str = Newtonsoft.Json.JsonConvert.SerializeObject(model);
            byte[] content_bytes = Encoding.UTF8.GetBytes(json_str);

            

            string api = "/SCSRwd/api/businessobject/";
            string url = $"https://{host}{api}";


            HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
            req.Method = "POST";
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                 | (SecurityProtocolType)768 //.Tls11 
                 | (SecurityProtocolType)3072; //.Tls12;
            ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
            req.ContentType = "application/json";

            string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(model);
            byte[] contentByte = Encoding.UTF8.GetBytes(jsonStr);
            using (Stream reqStream = req.GetRequestStream())
            {
                reqStream.Write(contentByte, 0, contentByte.Length);
            }
            using (WebResponse res = req.GetResponse())
            {
                string statusCode = $"{((HttpWebResponse)res).StatusCode}";
                string statusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                {
                    string rtnjson = reader.ReadToEnd();

                    rst = Newtonsoft.Json.JsonConvert.DeserializeObject<ScsDeptsReturnModel>(rtnjson);

                    //if (_Logger != null)
                    //{
                    //    _Logger.LogInformation("登入飛騰取得:{responseStr}", rst);
                    //}
                    //rst = Newtonsoft.Json.JsonConvert.DeserializeObject<FtEmployees>(rtnjson);
                    ////rst = $"{rtnobj.SessionGuid}";
                    //if (rst.DataSet != null && rst.DataSet.Tables.Count > 0)
                    //{
                    //    return rst.DataSet.Tables["ReportBody"];
                    //}
                    //else
                    //{
                    //    throw new Exception("取得資料失敗");
                    //}
                }
            }


            return rst;
        }

        /// <summary>
        /// 外部用 取得員工資料
        /// </summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        public ReportBody[] GetEmployees(string guid)
        {
            ReportBody[] tbl = null;
            FtEmployees rst = null;
            try
            {
                string host = this.AP02_Host;
                string api = "/SCSRwd/api/businessobject/";
                string url = $"https://{host}{api}";


                HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
                req.Method = "POST";
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls |
                     (SecurityProtocolType)768 //.Tls11 
                     | (SecurityProtocolType)3072; //.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                req.ContentType = "application/json";
                ScsEmployeesQueryModel model = new ScsEmployeesQueryModel();
                model.Action = "ExecReport";
                model.SessionGuid = guid;
                model.ProgId = "RHUM064";
                model.Value.Type = "AIS.Define.TExecReportInputArgs, AIS.Define";
                model.Value.UiType = "Report";
                model.Value.ReportId = "";
                model.Value.ReportTailId = "";
                model.Value.FilterItems = new FilterItem[] {
                    new FilterItem() {
                        Type= "AIS.Define.TFilterItem, AIS.Define",
                        FieldName="A.SYS_CompanyID",
                        FilterValue= CompanyID //"SCS999"
                    } };
                model.Value.UserFilter = "";

                /*
{
  "Action": "ExecReport",
  "SessionGuid": "48ca52a2-b1f4-487d-ab93-24a4327ad166",
  "ProgID": "RHUM002",
  "Value": {
    "$type": "AIS.Define.TExecReportInputArgs, AIS.Define",
    "UIType": "Report",
    "ReportID": "",
    "ReportTailID": "",
    "FilterItems": [
      {
        "$type": "AIS.Define.TFilterItem, AIS.Define",
        "FieldName": "ZA.SYS_CompanyID",
        "FilterValue": "SCS999"
      }
    ],
    "UserFilter": ""
  }
}
                 */

                string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                byte[] contentByte = Encoding.UTF8.GetBytes(jsonStr);
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(contentByte, 0, contentByte.Length);
                }
                using (WebResponse res = req.GetResponse())
                {
                    string statusCode = $"{((HttpWebResponse)res).StatusCode}";
                    string statusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                    using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                    {
                        string rtnjson = reader.ReadToEnd();
                        try
                        {
                            System.IO.File.WriteAllText("i:/alvin/FTEmp.json", rtnjson);
                        }
                        catch
                        {
                        }
                        //if (_Logger != null)
                        //{
                        //    _Logger.LogInformation("登入飛騰取得:{responseStr}", rst);
                        //}
                        //test
                        rst = Newtonsoft.Json.JsonConvert.DeserializeObject<FtEmployees>(rtnjson);
                        tbl = rst.DataSet.ReportBody;
                        //rst = $"{rtnobj.SessionGuid}";
                        //if (rst.DataSet != null && rst.DataSet.Tables.Count > 0)
                        //{
                        //    return rst.DataSet.Tables["ReportBody"];
                        //}
                        //else
                        //{
                        //    throw new Exception("取得資料失敗");
                        //}
                    }
                }
            }
            catch (WebException wx)
            {
                StringBuilder errmsg = new StringBuilder($"{wx.Message}{wx.InnerException}");
                if (wx.Response != null)
                {
                    using (WebResponse res = wx.Response)
                    {
                        HttpWebResponse hres = res as HttpWebResponse;
                        string statusCode = $"{hres.StatusCode}";
                        string statusDescription = $"{hres.StatusDescription}";
                        try
                        {
                            using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                            {
                                errmsg.AppendLine(reader.ReadToEnd());
                            }
                        }
                        catch (Exception ex2)
                        {
                            string msg = ex2.Message;
                            //_Logger.LogError(ex2, "SCS_utils 嘗試取得錯誤訊息失敗:{msg}", msg);
                        }
                    }
                }

                throw new Exception($"{errmsg}");
            }
            catch (Exception ex)
            {
                throw;
            }

            return tbl;
        }

        /// <summary>
        /// 外部用 取得員工資料
        /// </summary>
        /// <param name="guid"></param>
        /// <returns></returns>
        public ReportBody[] GetEmployeesAsync(string guid)
        {
            ReportBody[] tbl = null;
            FtEmployees rst = null;
            try
            {
               
                string host = "ap02.pchome.tw";
                string api = "/SCSRwd/api/businessobject/";
                string url = $"https://{host}{api}";


                HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
                req.Method = "POST";
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls |
                     (SecurityProtocolType)768 //.Tls11 
                     | (SecurityProtocolType)3072; //.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                req.ContentType = "application/json";
                ScsEmployeesQueryModel model = new ScsEmployeesQueryModel();
                model.Action = "ExecReport";
                model.SessionGuid = guid;
                model.ProgId = "RHUM002";
                model.Value.Type = "AIS.Define.TExecReportInputArgs, AIS.Define";
                model.Value.UiType = "Report";
                model.Value.ReportId = "";
                model.Value.ReportTailId = "";
                model.Value.FilterItems = new FilterItem[] {
                    new FilterItem() {
                        Type= "AIS.Define.TFilterItem, AIS.Define",
                        FieldName="ZA.SYS_CompanyID",
                        FilterValue= CompanyID //"SCS999"
                    } };
                model.Value.UserFilter = "";

                /*
{
  "Action": "ExecReport",
  "SessionGuid": "48ca52a2-b1f4-487d-ab93-24a4327ad166",
  "ProgID": "RHUM002",
  "Value": {
    "$type": "AIS.Define.TExecReportInputArgs, AIS.Define",
    "UIType": "Report",
    "ReportID": "",
    "ReportTailID": "",
    "FilterItems": [
      {
        "$type": "AIS.Define.TFilterItem, AIS.Define",
        "FieldName": "ZA.SYS_CompanyID",
        "FilterValue": "SCS999"
      }
    ],
    "UserFilter": ""
  }
}
                 */

                string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                byte[] contentByte = Encoding.UTF8.GetBytes(jsonStr);
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(contentByte, 0, contentByte.Length);
                }
                using (WebResponse res = req.GetResponse())
                {
                    string statusCode = $"{((HttpWebResponse)res).StatusCode}";
                    string statusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                    using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                    {
                        string rtnjson = reader.ReadToEnd();
                        //string rtnjson = reader.ReadToEnd();
                        //if (_Logger != null)
                        //{
                        //    _Logger.LogInformation("登入飛騰取得:{responseStr}", rst);
                        //}
                        rst = Newtonsoft.Json.JsonConvert.DeserializeObject<FtEmployees>(rtnjson);
                        tbl = rst.DataSet.ReportBody;
                        //rst = $"{rtnobj.SessionGuid}";
                        //if (rst.DataSet != null && rst.DataSet.Tables.Count > 0)
                        //{
                        //    return rst.DataSet.Tables["ReportBody"];
                        //}
                        //else
                        //{
                        //    throw new Exception("取得資料失敗");
                        //}
                    }
                }
            }
            catch (WebException wx)
            {
                StringBuilder errmsg = new StringBuilder($"{wx.Message}{wx.InnerException}");
                if (wx.Response != null)
                {
                    using (WebResponse res = wx.Response)
                    {
                        HttpWebResponse hres = res as HttpWebResponse;
                        string statusCode = $"{hres.StatusCode}";
                        string statusDescription = $"{hres.StatusDescription}";
                        try
                        {
                            using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                            {
                                errmsg.AppendLine(reader.ReadToEnd());
                            }
                        }
                        catch (Exception ex2)
                        {
                            string msg = ex2.Message;
                            //_Logger.LogError(ex2, "SCS_utils 嘗試取得錯誤訊息失敗:{msg}", msg);
                        }
                    }
                }

                throw new Exception($"{errmsg}");
            }
            catch (Exception ex)
            {
                throw;
            }

            return tbl;
        }


        public string GetLocalIP()
        {
            IPHostEntry host = Dns.GetHostEntry(Dns.GetHostName());
            var ippaddress = host
                .AddressList
                .FirstOrDefault(ip => ip.AddressFamily == AddressFamily.InterNetwork);
            return $"{ippaddress}";
        }

        /// <summary>
        /// 登入用
        /// </summary>
        /// <param name="loginModel"></param>
        /// <returns></returns>
        public string GetLoginToken()
        {          
            string rst = "";
            try
            {

                string host = AP01_Host; 
                
                string api = "/SCSRwd/api/systemobject/";
                string url = $"https://{host}{api}";


                HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
                req.Method = "POST";

                string local_ip = GetLocalIP();
                if (local_ip == "10.10.21.51" || local_ip == "10.10.21.52" || local_ip == "10.10.21.53")
                {
                    string MyProxyHost = "proxy";
                    int MyProxyPort = 3128;
                    req.Proxy = new WebProxy(MyProxyHost, MyProxyPort);
                }


                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls |
                     (SecurityProtocolType)768 //.Tls11 
                     | (SecurityProtocolType)3072; //.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                req.ContentType = "application/json";
                ScsLoginModel model = new ScsLoginModel();
                model.Action = "Login";
                model.Value.Type = "AIS.Define.TLoginInputArgs, AIS.Define";
                model.Value.UserId = UserName;
                model.Value.Password = Password;
                model.Value.CompanyId = CompanyID;

                model.Value.LanguageId = "zh-TW";

                string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                byte[] contentByte = Encoding.UTF8.GetBytes(jsonStr);
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(contentByte, 0, contentByte.Length);
                }
                using (WebResponse res = req.GetResponse())
                {
                    string statusCode = $"{((HttpWebResponse)res).StatusCode}";
                    string statusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                    using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                    {
                        string rtn_jsonStr = reader.ReadToEnd();
                        //if (_Logger != null)
                        //{
                        //    _Logger.LogInformation("登入飛騰取得:{responseStr}", rst);
                        //}
                        ScsLoginReturnModel rtnobj =
                            Newtonsoft.Json.JsonConvert.DeserializeObject<ScsLoginReturnModel>(rtn_jsonStr);
                        if (rtnobj.SessionGuid == "{00000000-0000-0000-0000-000000000000}")
                        {
                            throw new Exception(rtnobj.Message);
                        }
                        rst = $"{rtnobj.SessionGuid}";
                    }
                }
            }
            catch (WebException wx)
            {
                StringBuilder errmsg = new StringBuilder($"{wx.Message}{wx.InnerException}");
                if (wx.Response != null)
                {
                    using (WebResponse res = wx.Response)
                    {
                        HttpWebResponse hres = res as HttpWebResponse;
                        string statusCode = $"{hres.StatusCode}";
                        string statusDescription = $"{hres.StatusDescription}";
                        try
                        {
                            using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                            {
                                errmsg.AppendLine(reader.ReadToEnd());
                            }
                        }
                        catch (Exception ex2)
                        {
                            string msg = ex2.Message;
                            //_Logger.LogError(ex2, "SCS_utils 嘗試取得錯誤訊息失敗:{msg}", msg);
                            log.WriteLogAsync($"GET login token 取得 wx 失敗:{msg}", "ERROR");
                        }
                    }
                }
                log.WriteLogAsync($"GET login token 失敗:{errmsg}", "ERROR");
                throw new Exception($"{errmsg}");
            }
            catch (Exception ex)
            {
                log.WriteLogAsync($"GET login token 失敗:{ex.Message}{ex.InnerException}", "ERROR");
                throw;
            }

            return rst;
        }


        /// <summary>
        /// 登入用
        /// </summary>
        /// <param name="loginModel"></param>
        /// <returns></returns>
        public string GetLoginTokenAsync()
        {
           
            string rst = "";
            try
            {
               
                string host = AP01_Host; 
                string api = "/SCSRwd/api/systemobject/";
                string url = $"https://{host}{api}";


                HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
                req.Method = "POST";
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls |
                     (SecurityProtocolType)768 //.Tls11 
                     | (SecurityProtocolType)3072; //.Tls12;
                ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                req.ContentType = "application/json";
                ScsLoginModel model = new ScsLoginModel();
                model.Action = "Login";
                model.Value.Type = "AIS.Define.TLoginInputArgs, AIS.Define";
                model.Value.UserId = UserName;
                model.Value.Password = Password;
                model.Value.CompanyId = CompanyID;

                model.Value.LanguageId = "zh-TW";

                string jsonStr = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                byte[] contentByte = Encoding.UTF8.GetBytes(jsonStr);
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(contentByte, 0, contentByte.Length);
                }
                using (WebResponse res = req.GetResponse())
                {
                    string statusCode = $"{((HttpWebResponse)res).StatusCode}";
                    string statusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                    using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                    {
                        string rtn_jsonStr = reader.ReadToEnd();
                        //string rtn_jsonStr = reader.ReadToEnd();
                        //if (_Logger != null)
                        //{
                        //    _Logger.LogInformation("登入飛騰取得:{responseStr}", rst);
                        //}
                        ScsLoginReturnModel rtnobj =
                            Newtonsoft.Json.JsonConvert.DeserializeObject<ScsLoginReturnModel>(rtn_jsonStr);
                        rst = $"{rtnobj.SessionGuid}";
                    }
                }
            }
            catch (WebException wx)
            {
                StringBuilder errmsg = new StringBuilder($"{wx.Message}{wx.InnerException}");
                if (wx.Response != null)
                {
                    using (WebResponse res = wx.Response)
                    {
                        HttpWebResponse hres = res as HttpWebResponse;
                        string statusCode = $"{hres.StatusCode}";
                        string statusDescription = $"{hres.StatusDescription}";
                        try
                        {
                            using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                            {
                                errmsg.AppendLine(reader.ReadToEnd());
                            }
                        }
                        catch (Exception ex2)
                        {
                            string msg = ex2.Message;
                            //_Logger.LogError(ex2, "SCS_utils 嘗試取得錯誤訊息失敗:{msg}", msg);
                        }
                    }
                }

                throw new Exception($"{errmsg}");
            }
            catch (Exception ex)
            {
                throw;
            }

            return rst;
        }


    }

}
