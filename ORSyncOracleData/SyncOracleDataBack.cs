using AQuartzJobUTL;
using BoxmanBase;
using LogMgr;
using OracleNewQuitEmployee.ORSyncOracleData.Model;
using OracleNewQuitEmployee.ORSyncOracleData.Model.EmpAsm;
using OracleNewQuitEmployee.ORSyncOracleData.Model.Model1;
using OracleNewQuitEmployee.ORSyncOracleData.Model.SCIMUser;
using OracleNewQuitEmployee.ORSyncOracleData.Model.WorkAssignment;
using OracleNewQuitEmployee.ORSyncOracleData.Model.WorkerManagerInfo;
using OracleNewQuitEmployee.ToolModels;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;

namespace OracleNewQuitEmployee.ORSyncOracleData
{

    public class SyncOracleDataBack
    {

        keroroConnectBase.keroroConn conn;

        private LogOput log;


        public string Oracle_AP { get; set; }
        public string Oracle_Domain { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }


        string GetOracle_AP()
        {
            var rst = "";
            var fileName = "Oracle_AP.txt";
            try
            {
                var path = $"{AppDomain.CurrentDomain.BaseDirectory}ini/{fileName}";
                if (File.Exists(path))
                {
                    //rst = File.ReadAllText(path);
                    rst = File.ReadAllLines(path)[0];
                }
                else
                {
                    throw new Exception($"找不到 {fileName}");
                }
            }
            catch (Exception ex)
            {
                var errmsg = $"取得 {fileName} 失敗:{ex.Message}";
                log.WriteErrorLog(errmsg);
            }

            return rst;
        }

        string GetOracle_Domain()
        {
            var rst = "";
            var fileName = "Oracle_Domain.txt";
            try
            {
                var path = $"{AppDomain.CurrentDomain.BaseDirectory}ini/{fileName}";
                //var path = $"{AppDomain.CurrentDomain.BaseDirectory}{fileName}";
                if (File.Exists(path))
                {
                    var rst2 = File.ReadAllLines(path);
                    rst = rst2[0];
                }
                else
                {
                    throw new Exception($"找不到 {fileName}");
                }
            }
            catch (Exception ex)
            {
                var errmsg = $"取得 {fileName} 失敗:{ex.Message}";
                log.WriteErrorLog(errmsg);
            }

            return rst;
        }

        /// <summary>
        /// 取得 Oracle 帳密
        /// </summary>
        void FetchUserNamePwd()
        {
            try
            {
                //第一行 帳號
                //第二行 密碼
                var fileName = "OracleIDPW.txt";
                var path = $"{AppDomain.CurrentDomain.BaseDirectory}ini/{fileName}";
                var secrets = File.ReadAllLines(path);
                UserName = secrets[0];
                Password = secrets[1];
            }
            catch (Exception ex)
            {
                var errmsg = $"取得 Oracle 帳密 失敗:{ex.Message}";
                log.WriteErrorLog(errmsg);
            }
        }



        /// <summary>
        /// 建構式
        /// </summary>
        public SyncOracleDataBack(keroroConnectBase.keroroConn Conn, LogOput Log)
        {
            conn = Conn;
            log = Log;


            Oracle_AP = GetOracle_AP();
            Oracle_Domain = GetOracle_Domain();

            //取得 Oracle 帳/密
            FetchUserNamePwd();


            db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
            if (!Directory.Exists(db_path))
            {
                Directory.CreateDirectory(db_path);
            }
            db_file = $"{db_path}OracleData.sqlite";
            sqlite = new SQLiteUtl(db_file);

            try
            {
                var _53 = ConfigurationManager.AppSettings["ap53"];
                if (!string.IsNullOrWhiteSpace(_53))
                {
                    ap83位置 = _53;
                }
            }
            catch (Exception ex53)
            {
            }

        }

        /// <summary>
        /// 會讀取 ROOT 下的 53url.txt
        /// </summary>
        string ap83位置 = "";


        SQLiteUtl sqlite;
        DataTable tb;

        string db_path;
        string db_file;

        //        public void SuspendAccsIfNeeds(string BatchNo)
        //        {
        //            var empno = "";
        //            var success = "";
        //            var errmsg = "";
        //            //            var sql = $@"
        //            //select r_code
        //            //from supplier
        //            //where batchno = :0
        //            //and jobstatus in ('0','5')
        //            //";

        //            var sql = $@"

        //SELECT  
        //s.r_code
        //, w.lastname
        //, s.r_cname
        //, jobstatus 
        //		,[SuspendedFlag] 
        //		,s.*
        //	FROM supplier s
        //	left JOIN  [WorkerNames] W on s.r_code = w.lastname and s.batchno = w.batchno
        //	left join [OracleUsers] U ON U.PERSONID = W.PERSONID AND U.BATCHNO = W.BATCHNO	
        //	WHERE s.BATCHNO = :0 
        //	and jobstatus in ('0','5')
        //";


        //            DataTable tb;
        //            foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo))
        //            {
        //                empno = $"{row["r_code"]}";
        //                SuspendAccount(BatchNo, empno, "false", out success, out errmsg);
        //            }
        //        }



        public void SuspendAccountByGUID(string UserGUID, out string success)
        {
            var id = UserGUID;
            success = "false";
            try
            {
                log.WriteLog("5", "進入 SuspendAccountByGUID");

                var userName = "";
                if (All_Users != null)
                {
                    userName = (from user in All_Users
                                where user.GUID == UserGUID
                                select $"{user.Username}").FirstOrDefault();
                }


                log.WriteLog("5", $"準備停用 GUID:{UserGUID}, UserName:{userName}");

                var Enable = "false";
                SuspendUserModel model = new SuspendUserModel();
                model.Active = Enable;
                model.Schemas = new string[] { "urn:scim:schemas:core:2.0:User" };


                string url = $"{Oracle_Domain}/hcmRestApi/scim/Users/{id}";
                var mr = HttpPatchFromOracleAP(url, model);
                if (mr.StatusCode == "200")
                {
                    byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                    string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                    SuspendUserReturnModel usrrtn = Newtonsoft.Json.JsonConvert.DeserializeObject<SuspendUserReturnModel>(desc_str);
                    //log.WriteLog("5", $"SuspendAccount Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName}){Environment.NewLine}{desc_str}");
                    var rtn_enable = usrrtn.Active.ToUpper();
                    if (rtn_enable == Enable.ToUpper())
                    {
                        success = "true";
                        var word = "啟用";
                        if (rtn_enable == "TRUE")
                        {
                            word = "啟用";
                        }
                        if (rtn_enable == "FALSE")
                        {
                            word = "停用";
                        }

                        log.WriteLog("5", $"userName:{userName}, GUID:{UserGUID} 已 {word}");
                    }
                    else
                    {
                        log.WriteErrorLog($"SuspendAccountByGUID 操作失敗: Enable 狀態沒有改變!");

                    }
                }
                else
                {
                    throw new Exception($"{mr.StatusCode} {mr.StatusDescription} {mr.ReturnData} {mr.ErrorMessage}");
                }

            }
            catch (Exception ex)
            {
                var errmsg = $"SuspendAccountByGUID 失敗:GUID={UserGUID}, {ex.Message}{ex.InnerException}";
                log.WriteErrorLog(errmsg);
            }
        }

        /// <summary>
        /// 停用 Oracle user
        /// </summary>
        /// <param name="BATCHNO">不需填寫</param>
        /// <param name="EmpNo"></param>
        /// <param name="EmpName"></param>
        /// <param name="Enable"></param>
        /// <param name="success"></param>
        /// <param name="errmsg"></param>
        public void SuspendAccount(string BATCHNO, string EmpNo, string EmpName, string Enable, out string success, out string errmsg)
        {
            success = "false";
            errmsg = "";
            try
            {
                log.WriteLog("5", "準備停用/啟用 oracle user, 先取得Oracle User");
                var usrRtnObj = GetOracleUserByEmpNo(EmpNo);
                var tmpstr = Newtonsoft.Json.JsonConvert.SerializeObject(usrRtnObj);
                //log.WriteLog("5", $"取得 oracle user={tmpstr}");
                if (usrRtnObj.Resources.Count == 0)
                {
                    log.WriteLog("5", $"{EmpNo} 無 oracle 帳號, 略過停用!");
                    success = "true";
                    return;
                    //throw new Exception($"取得 oracle user 失敗! count=0");
                }
                var usrData = usrRtnObj.Resources[0];
                var id = usrData.Id;

                var oracle_user_active = "false";
                if (usrData.Active)
                {
                    oracle_user_active = "true";
                }
                var want_to_do_Enable = Enable.ToLower();
                if (oracle_user_active == want_to_do_Enable)
                {
                    var word = "停用";
                    if (usrData.Active)
                    {
                        word = "啟用";
                    }
                    log.WriteLog("5", $"{EmpNo} {EmpName} 在 oracle 的帳號已 {word}! 不需呼叫 api 操作.");
                }
                else
                {

                    if (!string.IsNullOrWhiteSpace(id))
                    {
                        //var tmp = "";
                        if (Enable.ToUpper() == "TRUE")
                        {
                            log.WriteLog("5", $"開始嘗試 啟用 {EmpNo} {EmpName}");
                        }
                        else if (Enable.ToUpper() == "FALSE")
                        {
                            log.WriteLog("5", $"開始嘗試 停用 {EmpNo} {EmpName}");
                        }
                        else
                        {
                            log.WriteLog("5", $"開始嘗試 Enable = {Enable} {EmpNo}");
                        }

                        SuspendUserModel model = new SuspendUserModel();
                        model.Active = Enable;
                        model.Schemas = new string[] { "urn:scim:schemas:core:2.0:User" };


                        string url = $"{Oracle_Domain}/hcmRestApi/scim/Users/{id}";
                        var mr = HttpPatchFromOracleAP(url, model);
                        tmpstr = Newtonsoft.Json.JsonConvert.SerializeObject(mr);
                        //log.WriteLog("5", $"呼叫 oracle 停用 api 得到={tmpstr}");
                        if (mr.StatusCode == "200")
                        {
                            byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                            string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                            //log.WriteLog("5", $"解 base64得到={desc_str}");
                            SuspendUserReturnModel usrrtn = Newtonsoft.Json.JsonConvert.DeserializeObject<SuspendUserReturnModel>(desc_str);
                            //log.WriteLog("5", $"SuspendAccount Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName}){Environment.NewLine}{desc_str}");
                            var rtn_enable = usrrtn.Active.ToUpper();
                            if (rtn_enable == Enable.ToUpper())
                            {
                                success = "true";
                                var word = "啟用";
                                if (rtn_enable == "TRUE")
                                {
                                    word = "啟用";
                                }
                                if (rtn_enable == "FALSE")
                                {
                                    word = "停用";
                                }

                                log.WriteLog("5", $"EmpNO:{EmpNo} {EmpName} 已 {word}");
                            }
                            else
                            {
                                throw new Exception($"SuspendAccount 操作失敗: Enable 狀態沒有改變!");
                            }
                        }
                        else
                        {
                            throw new Exception($"{mr.StatusCode} {mr.StatusDescription} {mr.ReturnData} {mr.ErrorMessage}");
                        }
                    }
                    else
                    {
                        log.WriteErrorLog($"user GUID is null!");
                    }
                }
            }
            catch (Exception ex)
            {
                errmsg = $"SuspendAccount 失敗:EmpNo={EmpNo}, {ex.Message}{ex.InnerException}";
                log.WriteErrorLog(errmsg);
            }
            finally
            {
                log.WriteLog("5", "停用/啟用 結束!");
            }
        }

        /// <summary>
        /// 取得 boxman ap 中的  _boxmanUtilitiesService.OracleApiUrl(); 值
        /// </summary>
        /// <returns></returns>
        public string TestGet()
        {
            string rst = "";
            ORGetData pars = new ORGetData();
            var ret = ApiOperation.CallApi(
                    new ApiRequestSetting()
                    {
                        MethodRoute = "api/OREMP/TestGet",
                        MethodType = "GET",
                        TimeOut = int30Mins
                    }
                    );
            if (ret.Success)
            {
                rst = ret.ResponseContent;
            }


            return rst;
        }
        /// <summary>
        /// 取得 boxman ap 中的  _boxmanUtilitiesService.OracleApiUrl(); 值
        /// </summary>
        /// <returns></returns>
        public string TestPost(string fn)
        {
            string rst = "";
            ORGetData pars = new ORGetData();
            pars.FunctionName = fn;
            var ret = ApiOperation.CallApi(
                    new ApiRequestSetting()
                    {
                        Data = pars,
                        MethodRoute = "api/OREMP/TestPost",
                        MethodType = "POST",
                        TimeOut = int30Mins
                    }
                    );
            if (ret.Success)
            {
                rst = ret.ResponseContent;
            }


            return rst;
        }

        /// <summary>
        /// 取得 boxman ap 中的  _boxmanUtilitiesService.OracleApiUrl(); 值
        /// </summary>
        /// <returns></returns>
        public string GetOracleApDomainName()
        {
            string rst = "";
            ORGetData pars = new ORGetData();
            pars.FunctionName = "GetOracleApDomainName";
            var ret = ApiOperation.CallApi(
                    new ApiRequestSetting()
                    {
                        Data = pars,
                        MethodRoute = "api/OREMP/ORGetData",
                        MethodType = "POST",
                        TimeOut = int30Mins
                    }
                    );
            if (ret.Success)
            {
                rst = ret.ResponseContent;
                ResultObject<string> _rst2 = Newtonsoft.Json.JsonConvert.DeserializeObject<ResultObject<string>>(rst);
                if (_rst2.Result)
                {
                    var url = _rst2.Data;
                    log.WriteLog("5", $"取得阿中的 oracle ap 位置:{url}");

                    var urls = url.Split(new string[] { "uri=" }, StringSplitOptions.None);
                    foreach (var _u in urls)
                    {
                        int pos1 = _u.ToLower().IndexOf("/api/");
                        if (pos1 >= 0)
                        {
                            rst = _u.Substring(0, pos1);
                            Oracle_AP = rst;
                        }
                    }
                }
            }


            return rst;
        }

        /// <summary>
        /// Create oracle 帳號，同時 update username/部門/主管員編/job level
        /// </summary>
        /// <param name="PersonNumber"></param>
        /// <param name="EmpNo"></param>
        /// <param name="EmpName"></param>
        /// <param name="EngName"></param>
        /// <param name="Email"></param>
        /// <param name="DeptNo">部門代碼</param>
        /// <param name="ManagerEmpNo">主管員編</param>
        /// <param name="JobLevel">0,1,2,3,4  空白=0</param>
        /// <param name="success"></param>
        /// <param name="errmsg"></param>
        public void CreateWorkerByMiddle(string PersonNumber, string EmpNo, string EmpName,
            string EngName, string Email, string DeptNo, string ManagerEmpNo, string JobLevel,
            out string success, out string errmsg)
        {
            success = "false";
            errmsg = "";
            //log.WriteLog("5", $"CreateWorkerByMiddle, EmpNo={EmpNo}, EmpName={EmpName} 準備新增 oracle worker");
            try
            {
                // ORACLE STAGE
                // for  oracle ap
                string email = Email;
                if (string.IsNullOrWhiteSpace(email))
                {
                    throw new Exception($"EmpNo={EmpNo}, EmpName={EmpName}, 沒有 email, 無法建立 Oracle Worker!");
                }
                else
                {
                    if (email.IndexOf('@') < 0)
                    {
                        email = $"{email}@staff.pchome.com.tw";
                    }
                }
                //CreateWorkerModel4 model = new CreateWorkerModel4();
                //model.Names[0].FirstName = EmpName;
                //model.Names[0].MiddleNames = EngName;
                //model.Names[0].LastName = EmpNo;
                //model.Emails[0].EmailType = "W1";
                //model.Emails[0].EmailAddress = email;

                CreateWorkerModel3 model = new CreateWorkerModel3();
                model.PersonNumber = PersonNumber;
                model.Names[0].FirstName = EmpName;
                model.Names[0].MiddleNames = EngName;
                model.Names[0].LastName = EmpNo;
                model.Emails[0].EmailType = "W1";
                model.Emails[0].EmailAddress = email;


                MiddleModel2 send_model2 = new MiddleModel2();
                string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/workers";
                //log.WriteLog("5", $"CreateWorkerByMiddle  url={url}");
                send_model2.URL = url;
                string payload = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                //log.WriteLog("5", $"CreateWorkerByMiddle payload={payload}");
                send_model2.SendingData = payload;
                send_model2.Method = "POST";
                //string username = this.UserName;
                //string password = this.Password;
                send_model2.UserName = this.UserName;
                send_model2.Password = this.Password;
                string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                //CredentialCache myCred = new CredentialCache();
                //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                //send_model2.Cred = myCred;
                send_model2.Timeout = int30Mins;


                // for BOXMAN API
                MiddleModel send_model = new MiddleModel();
                var _url = $"{Oracle_AP}/api/Middle/Call/";
                send_model.URL = _url;
                send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                send_model.Method = "POST";
                send_model.Timeout = int30Mins;
                var ret = ApiOperation.CallApi(
                        new ApiRequestSetting()
                        {
                            Data = send_model,
                            MethodRoute = "api/Middle/Call",
                            MethodType = "POST",
                            TimeOut = int30Mins
                        }
                        );
                if (ret.Success)
                {
                    string receive_str = ret.ResponseContent;
                    MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                    if (!string.IsNullOrWhiteSpace(mr2.ErrorMessage))
                    {
                        throw new Exception($"建立Worker失敗:{mr2.ErrorMessage}");
                    }
                    if (string.IsNullOrWhiteSpace(mr2.ReturnData))
                    {
                        throw new Exception("mr2.ReturnData is null, 伺服器回傳空白!");
                    }
                    MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2.ReturnData);
                    if (mr.StatusCode == "201")
                    {
                        byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                        string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                        Worker worker = Newtonsoft.Json.JsonConvert.DeserializeObject<Worker>(desc_str);
                        log.WriteLog("5", $"CreateWorkerByMiddle Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName})");
                        //log.WriteLog("5", $"CreateWorkerByMiddle Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName}){Environment.NewLine}{desc_str}");

                        //update username
                        ModifyOracleUserNameWhenDifferent(EmpNo, EmpNo);
                        //部門/主管員編/job level
                        Update_DefaultExpenseAccount_JobLevel_LineManager(EmpNo, DeptNo, ManagerEmpNo, JobLevel);
                        success = "true";
                    }
                    else
                    {
                        throw new Exception($"{mr.StatusCode} {mr.StatusDescription} {mr.ReturnData} {mr.ErrorMessage}");
                    }
                }
                else
                {
                    var _msg = $"CreateWorkerByMiddle Call Boxman Api {_url} 失敗:{ret.ErrorMessage}. {ret.ErrorException}";
                    log.WriteErrorLog(_msg);
                    throw new Exception(_msg);
                }
            }
            catch (Exception ex)
            {
                errmsg = $"CreateWorkerByMiddle 失敗:{ex.Message}{ex.InnerException}";
                log.WriteErrorLog(errmsg);
                throw;
            }

        }

        /// <summary>
        ///  新增 oracle 帳號, 但是沒有update部門/主管員編/job level
        /// </summary>
        /// <param name="PersonNumber"></param>
        /// <param name="EmpNo"></param>
        /// <param name="EmpName"></param>
        /// <param name="EngName"></param>
        /// <param name="Email"></param>
        /// <param name="success"></param>
        /// <param name="errmsg"></param>
        public void AddWorkerByMiddle(string PersonNumber, string EmpNo, string EmpName,
            string EngName, string Email, out string success, out string errmsg)
        {
            success = "false";
            errmsg = "";
            log.WriteLog("5", $"EmpNo={EmpNo}, EmpName={EmpName} 準備建立 oracle worker. (AddWorkerByMiddle)");
            try
            {
                // ORACLE STAGE
                // for  oracle ap
                string email = Email;

                //2021/10/29
                //沒有 email 就先用 xxx@staff.pchome.com.tw
                //到時再做發通知功能
                if (string.IsNullOrWhiteSpace(email))
                {
                    email = "xxx@staff.pchome.com.tw";
                    log.WriteLog("5", $"{EmpNo} 沒有email, 改用xxx@staff.pchome.com.tw");
                }

                if (string.IsNullOrWhiteSpace(email))
                {
                    throw new Exception($"EmpNo={EmpNo}, EmpName={EmpName}, 沒有 email, 無法建立 Oracle Worker!");
                }
                else
                {
                    if (email.IndexOf('@') < 0)
                    {
                        email = $"{email}@staff.pchome.com.tw";
                    }
                }
                //CreateWorkerModel4 model = new CreateWorkerModel4();
                //model.Names[0].FirstName = EmpName;
                //model.Names[0].MiddleNames = EngName;
                //model.Names[0].LastName = EmpNo;
                //model.Emails[0].EmailType = "W1";
                //model.Emails[0].EmailAddress = email;

                CreateWorkerModel3 model = new CreateWorkerModel3();
                model.PersonNumber = PersonNumber;
                model.Names[0].FirstName = EmpName;
                model.Names[0].MiddleNames = EngName;
                model.Names[0].LastName = EmpNo;
                model.Emails[0].EmailType = "W1";
                model.Emails[0].EmailAddress = email;


                MiddleModel2 send_model2 = new MiddleModel2();
                string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/workers";
                log.WriteLog("5", $"url={url}  (AddWorkerByMiddle)");
                send_model2.URL = url;
                string payload = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                log.WriteLog("5", $"payload={payload}  (AddWorkerByMiddle)");
                send_model2.SendingData = payload;
                send_model2.Method = "POST";
                //string username = this.UserName;
                //string password = this.Password;
                send_model2.UserName = this.UserName;
                send_model2.Password = this.Password;
                string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                //CredentialCache myCred = new CredentialCache();
                //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                //send_model2.Cred = myCred;
                send_model2.Timeout = int30Mins;


                // for BOXMAN API
                MiddleModel send_model = new MiddleModel();
                var _url = $"{Oracle_AP}/api/Middle/Call/";
                send_model.URL = _url;
                send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                send_model.Method = "POST";
                send_model.Timeout = int30Mins;
                var ret = ApiOperation.CallApi(
                        new ApiRequestSetting()
                        {
                            Data = send_model,
                            MethodRoute = "api/Middle/Call",
                            MethodType = "POST",
                            TimeOut = int30Mins
                        }
                        );
                if (ret.Success)
                {
                    string receive_str = ret.ResponseContent;
                    try
                    {
                        MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                        if (!string.IsNullOrWhiteSpace(mr2.ErrorMessage))
                        {
                            throw new Exception(mr2.ErrorMessage);
                        }
                        if (string.IsNullOrWhiteSpace(mr2.ReturnData))
                        {
                            throw new Exception("mr2.ReturnData is null, 伺服器回傳空白!");
                        }
                        MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2.ReturnData);
                        if (mr.StatusCode == "201")
                        {
                            byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                            string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                            Worker worker = Newtonsoft.Json.JsonConvert.DeserializeObject<Worker>(desc_str);
                            log.WriteLog("5", $"Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName})  (AddWorkerByMiddle )");
                            //log.WriteLog("5", $"AddWorkerByMiddle Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName}){Environment.NewLine}{desc_str}");
                            success = "true";

                        }
                        else
                        {
                            throw new Exception($"{mr.StatusCode} {mr.StatusDescription} {mr.ReturnData} {mr.ErrorMessage}");
                        }
                    }
                    catch (Exception exbs64)
                    {
                        log.WriteErrorLog($"AddWorkerByMiddle Create Worker 失敗:{exbs64.Message}{exbs64.InnerException}");
                    }
                }
                else
                {
                    log.WriteErrorLog($"AddWorkerByMiddle Call Boxman Api {_url} 失敗:{ret.ErrorMessage}. {ret.ErrorException}");
                }
            }
            catch (Exception ex)
            {
                errmsg = $"AddWorkerByMiddle 失敗:{ex.Message}{ex.InnerException}";
                log.WriteErrorLog(errmsg);
            }

        }

        /// <summary>
        ///  新增 oracle 帳號
        /// </summary>
        /// <param name="PersonNumber"></param>
        /// <param name="EmpNo"></param>
        /// <param name="EmpName"></param>
        /// <param name="EngName"></param>
        /// <param name="Email"></param>
        /// <param name="success"></param>
        /// <param name="errmsg"></param>
        public void AddWorkerByMiddle2(string PersonNumber, string EmpNo, string EmpName,
            string EngName, string Email, out string success, out string errmsg)
        {
            var fnName = "AddWorkerByMiddle2";
            success = "false";
            errmsg = "";
            log.WriteLog("5", $"{fnName}, EmpNo={EmpNo}, EmpName={EmpName} 準備新增 oracle worker");
            try
            {
                // ORACLE STAGE
                // for  oracle ap
                string email = Email;
                if (string.IsNullOrWhiteSpace(email))
                {
                    email = $"xxx@staff.pchome.com.tw";
                    log.WriteLog("5", $"{EmpNo} {EmpName} 沒有email, 改用xxx@staff.pchome.com.tw");
                    //throw new Exception($"EmpNo={EmpNo}, EmpName={EmpName}, 沒有 email, 無法建立 Oracle Worker!");
                }
                else
                {
                    if (email.IndexOf('@') < 0)
                    {
                        email = $"{email}@staff.pchome.com.tw";
                    }
                }
                //CreateWorkerModel4 model = new CreateWorkerModel4();
                //model.Names[0].FirstName = EmpName;
                //model.Names[0].MiddleNames = EngName;
                //model.Names[0].LastName = EmpNo;
                //model.Emails[0].EmailType = "W1";
                //model.Emails[0].EmailAddress = email;

                CreateWorkerModel3 model = new CreateWorkerModel3();
                model.PersonNumber = PersonNumber;
                model.Names[0].FirstName = EmpName;
                model.Names[0].MiddleNames = EngName;
                model.Names[0].LastName = EmpNo;
                model.Emails[0].EmailType = "W1";
                model.Emails[0].EmailAddress = email;

                string payload = Newtonsoft.Json.JsonConvert.SerializeObject(model);
                log.WriteLog("5", $"{fnName} payload={payload}");

                string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/workers";
                log.WriteLog("5", $"{fnName}  url={url}");

                //MiddleModel2 send_model2 = new MiddleModel2();
                //send_model2.URL = url;
                //send_model2.SendingData = payload;
                //send_model2.Method = "POST";
                //send_model2.UserName = this.UserName;
                //send_model2.Password = this.Password;
                //string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                //send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                //send_model2.Timeout = int30Mins;

                //var mr = HttpPostFromOracleAP(url, model);
                Dictionary<string, string> headers = new Dictionary<string, string>();
                headers.Add("effective-Of", "RangeStartDate=2000-01-01");
                var mr = HttpPostFromOracleAP_AndHeaders3(url, model, headers);
                var tmpstr = Newtonsoft.Json.JsonConvert.SerializeObject(mr);
                log.WriteLog("5", $"AddWorkerByMiddle2 Middle Return 取得:{tmpstr}");

                if ((mr.StatusCode == "201") || (mr.StatusCode == "200"))
                {
                    byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                    string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                    log.WriteLog("5", $"解base64得到:{desc_str}");
                    Worker worker = Newtonsoft.Json.JsonConvert.DeserializeObject<Worker>(desc_str);
                    log.WriteLog("5", $"{fnName} Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName})");
                    //log.WriteLog("5", $"AddWorkerByMiddle Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName}){Environment.NewLine}{desc_str}");
                    success = "true";

                }
                else
                {
                    throw new Exception($"{mr.StatusCode} {mr.StatusDescription} {mr.ReturnData} {mr.ErrorMessage}");
                }


                //// for BOXMAN API
                //MiddleModel send_model = new MiddleModel();
                //var _url = $"{Oracle_AP}/api/Middle/Call/";
                //send_model.URL = _url;
                //send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                //send_model.Method = "POST";
                //send_model.Timeout = int30Mins;
                //var ret = ApiOperation.CallApi(
                //        new ApiRequestSetting()
                //        {
                //            Data = send_model,
                //            MethodRoute = "api/Middle/Call",
                //            MethodType = "POST",
                //            TimeOut = int30Mins
                //        }
                //        );
                //if (ret.Success)
                //{
                //    string receive_str = ret.ResponseContent;
                //    try
                //    {
                //        MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                //        if (!string.IsNullOrWhiteSpace(mr2.ErrorMessage))
                //        {
                //            throw new Exception(mr2.ErrorMessage);
                //        }
                //        if (string.IsNullOrWhiteSpace(mr2.ReturnData))
                //        {
                //            throw new Exception("mr2.ReturnData is null, 伺服器回傳空白!");
                //        }
                //        MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2.ReturnData);
                //        if (mr.StatusCode == "201")
                //        {
                //            byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                //            string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                //            Worker worker = Newtonsoft.Json.JsonConvert.DeserializeObject<Worker>(desc_str);
                //            log.WriteLog("5", $"AddWorkerByMiddle Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName})");
                //            //log.WriteLog("5", $"AddWorkerByMiddle Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName}){Environment.NewLine}{desc_str}");
                //            success = "true";

                //        }
                //        else
                //        {
                //            throw new Exception($"{mr.StatusCode} {mr.StatusDescription} {mr.ReturnData} {mr.ErrorMessage}");
                //        }
                //    }
                //    catch (Exception exbs64)
                //    {
                //        log.WriteErrorLog($"AddWorkerByMiddle Create Worker 失敗:{exbs64.Message}{exbs64.InnerException}");
                //    }
                //}
                //else
                //{
                //    log.WriteErrorLog($"AddWorkerByMiddle Call Boxman Api {_url} 失敗:{ret.ErrorMessage}. {ret.ErrorException}");
                //}


            }
            catch (Exception ex)
            {
                errmsg = $"{fnName} 失敗:{ex.Message}{ex.InnerException}";
                log.WriteErrorLog(errmsg);
            }
        }

        private string GetGUIDByEmpNo(string EmpNo)
        {
            string rst = "";
            try
            {

                string sql = $@"

select u.guid
from workernames n
join oracleusers u on n.personnumber = u.personnumber
and n.batchno =  u.batchno 
where lastname = '{EmpNo}'
and u.batchno = (
select max(batchno)
from batch
)
";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                rst = sqlite.ExecuteScalarA(sql);
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"取得User GUID 失敗:{ex}");
            }
            return rst;
        }

        public List<OracleEmployee2> OracleEmployeesIncludeAssignmentsCollection;


        /// <summary>
        /// 取得部份Employee資料
        /// </summary>
        /// <param name="Success"></param>
        /// <param name="SomeEmpNo">用逗號分隔的員工編號</param>
        /// <returns></returns>
        public List<OracleEmployee2> GetSomeOracleEmployeesIncludeAssignments(out bool Success, string SomeEmpNo)
        {
            Success = false;
            OracleEmployeesIncludeAssignmentsCollection = new List<OracleEmployee2>();
            int offset = 0;
            var url = "";
            //var hasData = true;
            string[] SomeEmpNos = SomeEmpNo.Split(',');
            //while (hasData)
            foreach (var empno in SomeEmpNos)
            {

                url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/emps?q=LastName='{empno}'&expand=assignments";
                try
                {


                    //  //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                    MiddleReturn mr = HttpGetFromOracleAP(url);
                    if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                    {
                        throw new Exception($"GetSomeOracleEmployeesIncludeAssignments 失敗:{mr.ErrorMessage}");
                    }
                    var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                    var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                    OracleEmployeeIncludeAssignmentsCollection2 emps = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployeeIncludeAssignmentsCollection2>(desc_str);
                    if (emps != null && emps.Count > 0)
                    {
                        OracleEmployeesIncludeAssignmentsCollection.AddRange(emps.Items.ToList());
                        if (emps.HasMore)
                        {
                            offset += 500;
                        }
                    }
                }
                catch (Exception ex)
                {
                    log.WriteErrorLog($"GetSomeOracleEmployees 呼叫 api 失敗:api:{url}\r\n{ex.Message}{ex.InnerException}");
                    throw;
                }

            }
            Success = true;
            return OracleEmployeesIncludeAssignmentsCollection;

        }



        public List<OracleEmployee2> GetAllOracleEmployeesIncludeAssignments(out bool Success)
        {
            Success = false;
            OracleEmployeesIncludeAssignmentsCollection = new List<OracleEmployee2>();
            int offset = 0;
            var url = "";
            var hasData = true;
            while (hasData)
            {

                url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/emps?expand=assignments&limit=500&offset={offset}";
                try
                {


                    //  //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                    MiddleReturn mr = HttpGetFromOracleAP(url);
                    if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                    {
                        throw new Exception($"GetAllOracleEmployeesIncludeAssignments 失敗:{mr.ErrorMessage}");
                    }
                    var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                    var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                    OracleEmployeeIncludeAssignmentsCollection2 emps = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployeeIncludeAssignmentsCollection2>(desc_str);
                    if (emps != null && emps.Count > 0)
                    {
                        OracleEmployeesIncludeAssignmentsCollection.AddRange(emps.Items.ToList());
                        if (emps.HasMore)
                        {
                            hasData = true;
                            offset += 500;
                        }
                        else
                        {
                            hasData = false;
                        }
                    }
                    else
                    {
                        hasData = false;
                    }
                }
                catch (Exception ex)
                {
                    log.WriteErrorLog($"GetAllOracleEmployees 呼叫 api 失敗:api:{url}\r\n{ex.Message}{ex.InnerException}");
                    throw;
                }

            }
            Success = true;
            return OracleEmployeesIncludeAssignmentsCollection;
        }

        public List<OracleEmployee> GetAllOracleEmployees()
        {
            List<OracleEmployee> rst = new List<OracleEmployee>();
            int offset = 0;
            var url = "";
            var hasData = true;
            while (hasData)
            {

                url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/emps?limit=500&offset={offset}";
                try
                {


                    //  //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                    MiddleReturn mr = HttpGetFromOracleAP(url);
                    if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                    {
                        throw new Exception($"GetAllOracleEmployees 失敗:{mr.ErrorMessage}");
                    }
                    var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                    var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                    OracleEmployees emps = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployees>(desc_str);
                    if (emps != null && emps.Count > 0)
                    {
                        rst.AddRange(emps.Items.ToList());
                        if (emps.HasMore)
                        {
                            hasData = true;
                            offset += 500;
                        }
                        else
                        {
                            hasData = false;
                        }
                    }
                    else
                    {
                        hasData = false;
                    }
                }
                catch (Exception ex)
                {
                    log.WriteErrorLog($"GetAllOracleEmployees 呼叫 api 失敗:api:{url}\r\n{ex.Message}{ex.InnerException}");
                    throw;
                }

            }

            return rst;
        }


        public void UpdateLineManager(string EmployeeAssignmentsSelfUrl,
            string PersonId, string AssignmentId, out bool success, out string errmsg)
        {
            success = false;
            errmsg = "";
            try
            {
                var content1 = new
                {
                    ManagerAssignmentId = AssignmentId, //如果在Create Worker中已獲取並保存，
                                                        //则此步可不用执行；
                                                        //若未保存，则需要执行查询獲取。 
                    ActionCode = "MANAGER_CHANGE",
                    ManagerId = PersonId, //ManagerId为Manager這個用戶所對應的PersonId
                    ManagerType = "LINE_MANAGER"
                };
                var mr3 = HttpPatchFromOracleAP(EmployeeAssignmentsSelfUrl, content1);
                if (mr3.StatusCode == "200")
                {
                    //log.WriteLog("5", $"更新 Line Manager 為 {ManagerEmpNo} 成功!");
                    //byte[] bs64_bytes1 = Convert.FromBase64String(mr2.ReturnData);
                    //string desc_str1 = Encoding.UTF8.GetString(bs64_bytes1);
                    success = true;
                }
                else
                {
                    errmsg = $"{mr3.ErrorMessage}{mr3.ReturnData}";
                }
            }
            catch (Exception ex)
            {
                errmsg = $"{ex.Message}";
            }

        }

        public void GetEmployeePersonIDAssignmentIDByEmpNo(string EmpNo,
            out string PersonID, out string AssignmentID)
        {
            PersonID = "";
            AssignmentID = "";

            //這個 &expand=assignments 是取得 employee 後
            //json結構裡有個 assignments 的子結構
            //子結構就可以用 expand=xxx 來直接取得
            //不然子結構是一串url，就還要再打這個 url 才能取得
            var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/emps?q=LastName=\"{EmpNo}\"&expand=assignments";
            MiddleReturn mr = HttpGetFromOracleAP(url);
            var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
            var desc_str = Encoding.UTF8.GetString(bs64_bytes);
            OracleEmployeeIncludeAssignmentsCollection asms = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployeeIncludeAssignmentsCollection>(desc_str);
            if (asms.Count > 0)
            {
                foreach (var item in asms.Items)
                {
                    PersonID = item.PersonId;
                    foreach (var assignment in item.Assignments)
                    {
                        AssignmentID = assignment.AssignmentId;
                    }
                }
            }
        }

        /// <summary>
        /// 在用 hcmRestApi/resources//emps?LastName="800924"&expand=assignments
        /// 取得的Employee Json 結構裡會比用hcmRestApi/resources/emps?q=LastName="800924"
        /// 多出一個 assignments 節點，裡面的 links 在 rel=self, name=assignments 的 href 就代表這個 assignment
        /// 所以要更新 assignment 就要取這個 href 然後呼叫 patch http 方法, 參數就是跟 assignment 一樣，但只需
        /// 列出需修改的屬性
        /// 所以這個 url 是用來更新用的
        /// </summary>
        /// <param name="EmpNo"></param>
        /// <returns></returns>
        public string GetEmployeeAssignmentsSelfURLByEmpNo(string EmpNo)
        {

            var rst = "";
            //這個 &expand=assignments 是取得 employee 後
            //json結構裡有個 assignments 的子結構
            //子結構就可以用 expand=xxx 來直接取得
            //不然子結構是一串url，就還要再打這個 url 才能取得
            var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/emps?q=LastName=\"{EmpNo}\"&expand=assignments";
            MiddleReturn mr = HttpGetFromOracleAP(url);
            var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
            var desc_str = Encoding.UTF8.GetString(bs64_bytes);
            OracleEmployeeIncludeAssignmentsCollection asms = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployeeIncludeAssignmentsCollection>(desc_str);
            if (asms.Count > 0)
            {
                foreach (var item in asms.Items)
                {
                    foreach (var assignment in item.Assignments)
                    {
                        foreach (var link in assignment.Links)
                        {
                            if (link.Rel == "self" && link.Name == "assignments")
                            {
                                rst = link.Href;
                                break;
                            }
                        }
                    }
                }
            }

            return rst;
        }


        public string GetEmployeeAssignmentsSelfURLByEmpNo_2(string EmpNo)
        {

            var rst = "";
            //這個 &expand=assignments 是取得 employee 後
            //json結構裡有個 assignments 的子結構
            //子結構就可以用 expand=xxx 來直接取得
            //不然子結構是一串url，就還要再打這個 url 才能取得
            //var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/emps?q=LastName=\"{EmpNo}\"&expand=assignments";
            var url = $"{Oracle_Domain}/hcmRestApi/resources/latest/workers?q=PersonNumber={EmpNo}&expand=workRelationships.assignments";
            MiddleReturn mr = HttpGetFromOracleAP(url);
            var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
            var desc_str = Encoding.UTF8.GetString(bs64_bytes);
            TmpOracleReturnObj20230119_3 asms = Newtonsoft.Json.JsonConvert.DeserializeObject<TmpOracleReturnObj20230119_3>(desc_str);
            //if (asms.Count > 0)
            //{
            //    foreach (var item in asms.Items)
            //    {
            //        foreach (var ship in item.WorkRelationships)
            //        {
            //            foreach (var assignment in ship.Assignments)
            //            {
            //                foreach(var link in assignment.Links)
            //                if (link.Rel == "self" && link.Name == "assignments")
            //                {
            //                    rst = link.Href;
            //                    break;
            //                }
            //            }
            //        }
            //    }
            //}
            rst = (from item in asms.Items
                   from ship in item.WorkRelationships
                   from asgn in ship.Assignments
                   from link in asgn.Links
                   where link.Rel == "self" && link.Name == "assignments"
                   select link.Href).FirstOrDefault();

            return rst;
        }




        Dictionary<string, string> JobLevelCollection;
        public Dictionary<string, string> GetOracleJobLevelCollection()
        {
            if (JobLevelCollection == null)
            {
                JobLevelCollection = new Dictionary<string, string>();
                try
                {
                    var JobLevelName = "職員";
                    var jobid = GetJobIDByChinese(JobLevelName);
                    JobLevelCollection.Add(JobLevelName, jobid);
                    JobLevelName = "室/處長";
                    jobid = GetJobIDByChinese(JobLevelName);
                    JobLevelCollection.Add(JobLevelName, jobid);
                    JobLevelName = "部長";
                    jobid = GetJobIDByChinese(JobLevelName);
                    JobLevelCollection.Add(JobLevelName, jobid);
                    JobLevelName = "營運長";
                    jobid = GetJobIDByChinese(JobLevelName);
                    JobLevelCollection.Add(JobLevelName, jobid);
                    JobLevelName = "執行長/總經理";
                    jobid = GetJobIDByChinese(JobLevelName);
                    JobLevelCollection.Add(JobLevelName, jobid);
                }
                catch (Exception ex)
                {
                }
            }
            return JobLevelCollection;
        }





        private OracleEmployeeAssignmentSelf2 GetOracleEmployeeAssignmentSelf2ByURL(string EmpNo, string url)
        {
            OracleEmployeeAssignmentSelf2 rst = null;
            try
            {
                var mrass = HttpGetFromOracleAP(url);
                if (string.IsNullOrWhiteSpace(mrass.ReturnData))
                {
                    throw new Exception($"取得 Assignment Self Object 失敗! ReturnData 空白!");
                }
                var bs64_bytes2 = Convert.FromBase64String(mrass.ReturnData);
                var desc_str2 = Encoding.UTF8.GetString(bs64_bytes2);
                //OracleEmployeeAssignmentsModel
                rst = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployeeAssignmentSelf2>(desc_str2);
            }
            catch (Exception ex)
            {
                throw new Exception($"取得OracleEmployeeAssignment失敗:{ex.Message}");
            }
            return rst;
        }

        List<Employee_workRelationships_Assignment> EmployeeAssignmentCollection;
        private Employee_workRelationships_Assignment GetEmployee_workRelationships_assignments_ByPersonNumber(string PersonNumber)
        {
            Employee_workRelationships_Assignment rst = null;
            try
            {
                if (EmployeeAssignmentCollection == null)
                {
                    EmployeeAssignmentCollection = new List<Employee_workRelationships_Assignment>();
                }
                var asms = from mn in EmployeeAssignmentCollection
                           where mn.PersonNumber == PersonNumber
                           select mn;
                foreach (var asm in asms)
                    rst = asm;

                if (rst == null)
                {
                    var manager_url = $"{Oracle_Domain}/hcmRestApi/resources/latest/workers?q=PersonNumber={PersonNumber}&expand=workRelationships.assignments";
                    var mr2 = HttpGetFromOracleAP(manager_url);
                    if (string.IsNullOrWhiteSpace(mr2.ReturnData))
                    {
                        throw new Exception($"取得 workRelationships.assignments 失敗:{mr2.StatusCode} {mr2.ErrorMessage}");
                    }
                    var bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                    var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                    OracleEmployeeAssignmentSelfModel asmn = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployeeAssignmentSelfModel>(desc_str);

                    if (asmn.Count > 0)
                    {
                        rst = asmn.Items[0];
                        EmployeeAssignmentCollection.Add(rst);
                    }
                }

            }
            catch (Exception ex)
            {
                throw new Exception($"取得Employee_workRelationships_Assignment失敗:{ex.Message}");
            }

            return rst;
        }


        public bool UpdateUserNameEMailByOracleEmployee2(OracleEmployee2 oraEmp,
            string UserName, string EMail)
        {

            bool rst = false;
            foreach (var link in oraEmp.Links)
            {
                if (link.Rel == "self" && link.Name == "emps")
                {
                    var url = link.Href;
                    var par = new
                    {
                        WorkEmail = EMail,
                        UserName = UserName
                    };
                    MiddleReturn mr2 = HttpPatchFromOracleAP(url, par);
                    if (mr2.StatusCode == "200")
                    {
                        rst = true;
                        log.WriteLog("5", $"{oraEmp.LastName} 更新 UserName / email 成功:{UserName}");
                        //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                        //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                    }
                    else
                    {
                        log.WriteErrorLog($"{oraEmp.LastName} 更新 User Name / email 失敗:{mr2.ErrorMessage}{mr2.ReturnData}");
                    }
                    break;
                }
            }
            return rst;
        }






        /// <summary>
        /// 更新 Default Expense Account, 主管員編, Job Level
        /// 這是給新建帳號使用的
        /// </summary>
        /// <param name="EmpNo"></param>
        /// <param name="DeptNo"></param>
        /// <param name="ManagerEmpNo"></param>
        /// <param name="JobLevel"></param>
        /// <returns></returns>
        public bool Update_DefaultExpenseAccount_JobLevel_LineManager(string EmpNo, string DeptNo, string ManagerEmpNo,
            string JobLevel)
        {
            bool rst = false;
            var JobID = "";
            foreach (var item in GetJobLevelList())
            {
                if (JobLevel == item.ApprovalAuthority)
                {
                    JobID = item.JobId;
                }
            }


            log.WriteLog("5", $"準備更新 Job ID, Default Expense Account (UpdateDeptNoByEmployeeApi)");
            try
            {
                #region 先取得 Oracle 上的資料來供比對
                //正式環境User 的 Default Expense Account 的規則一樣比照Stage(DEV1), default account =6288099
                //完整  0001.< 員工所屬profit center >.< 員工所屬Department > .6288099.0000.000000000.0000000.0000
                //範例: 如 0001.POS.POS000000.6288099.0000.000000000.0000000.0000
                //取得EmployeeAssignment自已的url
                string EmployeeAssignmentsSelfUrl = GetEmployeeAssignmentsSelfURLByEmpNo(EmpNo);
                if (string.IsNullOrWhiteSpace(EmployeeAssignmentsSelfUrl))
                {
                    throw new Exception("Get Employee Assignments URL失敗:(空白)");
                }

                //先 GET 資料來看看
                OracleEmployeeAssignmentSelf2 empAsm = null;
                var jobid_in_oracle = "";
                var defExpenseAccount1_in_oracle = "";
                var ManagerAssignmentId1 = "";
                var ManagerID1 = "";
                try
                {
                    //取得 Employee Assignment
                    //var mrass = HttpGetFromOracleAP(EmployeeAssignmentsUrl);
                    //var bs64_bytes2 = Convert.FromBase64String(mrass.ReturnData);
                    //var desc_str2 = Encoding.UTF8.GetString(bs64_bytes2);
                    //asm2 = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployeeAssignmentSelf2>(desc_str2);


                    //取得 Employee Assignment
                    empAsm = GetOracleEmployeeAssignmentSelf2ByURL(EmpNo, EmployeeAssignmentsSelfUrl);

                    jobid_in_oracle = empAsm.JobId;
                    defExpenseAccount1_in_oracle = empAsm.DefaultExpenseAccount;
                    ManagerAssignmentId1 = empAsm.ManagerAssignmentId;
                    ManagerID1 = empAsm.ManagerId;
                    //if (asm2.Count > 0)
                    //{
                    //    var _asm = asm2.Items[0];
                    //    foreach (var asmm in _asm.Assignments)
                    //    {
                    //        jobid1 = asmm.JobId;
                    //        defExpenseAccount1 = asmm.DefaultExpenseAccount;
                    //        ManagerAssignmentId1 = asmm.ManagerAssignmentId;
                    //    }

                    //}
                }
                catch (Exception exass)
                {
                    log.WriteErrorLog($"先查資料時，取得 JobID, DefaultExpenseAccount, ManagerAssignmentId 失敗:{exass.Message}");
                }
                #endregion


                #region 檢查/更新 Job Level
                //var jobLevelName = JobLevelName;
                //if (string.IsNullOrWhiteSpace(jobLevelName))
                //{
                //    jobLevelName = "職員";
                //}

                //var Jobid = "";
                //try
                //{
                //    Jobid = JobLevelCollection[jobLevelName];
                //}
                //catch (Exception exlv)
                //{
                //}

                //if (string.IsNullOrWhiteSpace(Jobid))
                //{
                //    //重新準備 Job ID
                //    log.WriteErrorLog($"重新準備 Job ID");
                //    JobLevelCollection = null;
                //    GetOracleJobLevelCollection();
                //}
                //if (string.IsNullOrWhiteSpace(Jobid))
                //{
                //    log.WriteErrorLog($"還是取不到 Job Level 的 ID! 改用 api 傳中文取");
                //    Jobid = GetJobIDByChinese(jobLevelName);
                //}

                if (jobid_in_oracle != JobID)
                {
                    log.WriteLog("5", $"準備更新 jobid");

                    var content = new
                    {
                        JobId = JobID
                    };
                    var mr = HttpPatchFromOracleAP(EmployeeAssignmentsSelfUrl, content);
                    if (mr.StatusCode == "200")
                    {
                        rst = true;
                        log.WriteLog("5", $"更新 JobId 成功!");
                        //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                        //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                    }
                }
                else
                {
                    log.WriteLog("5", $"jobid 相同，不需更新");
                }
                #endregion

                #region 檢查/更新 Default Expense Account
                var deptNo = DeptNo.Substring(0, 9);
                var profit_center = deptNo.Substring(0, 3);
                var Default_Expense_Account = $"0001.{profit_center}.{deptNo}.6288099.0000.000000000.0000000.0000";
                if (defExpenseAccount1_in_oracle != Default_Expense_Account)
                {
                    log.WriteLog("5", $"準備更新 Default Expense Account!");

                    var content = new
                    {
                        DefaultExpenseAccount = Default_Expense_Account
                    };
                    var mr = HttpPatchFromOracleAP(EmployeeAssignmentsSelfUrl, content);
                    if (mr.StatusCode == "200")
                    {
                        rst = true;
                        log.WriteLog("5", $"更新 DefaultExpenseAccount 成功!");
                        //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                        //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                    }
                }
                else
                {
                    log.WriteLog("5", $"Default_Expense_Account 相同，不需更新");
                }
                #endregion

                #region 檢查/更新 Line Manager 主管
                if (string.IsNullOrWhiteSpace(ManagerEmpNo))
                {
                    log.WriteLog("5", $"傳入的主管員編為空，不更新 Line Manager");
                }
                else
                {
                    log.WriteLog("5", $"準備 Get ManagerEmpNo={ManagerEmpNo} 的 Employee 物件!");
                    var manager = GetOracleEmployeeByEmpNo(ManagerEmpNo);
                    if (manager == null)
                    {
                        log.WriteLog("5", $"取不到 Manager({ManagerEmpNo}) 的 Employee 物件");
                    }
                    else
                    {
                        ////取得workRelationships.assignments
                        //var manager_url = $"{Oracle_Domain}/hcmRestApi/resources/latest/workers?q=PersonNumber={manager.PersonNumber}&expand=workRelationships.assignments";
                        //var mr2 = HttpGetFromOracleAP(manager_url);
                        //if (string.IsNullOrWhiteSpace(mr2.ReturnData))
                        //{
                        //    throw new Exception($"在取得 Line Manager的 workRelationships.assignments 時失敗:{mr2.StatusCode} {mr2.ErrorMessage}");
                        //}
                        //var bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                        //var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                        //OracleEmployeeAssignmentSelfModel asmn = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployeeAssignmentSelfModel>(desc_str);


                        //取得workRelationships.assignments
                        Employee_workRelationships_Assignment asm = GetEmployee_workRelationships_assignments_ByPersonNumber(manager.PersonNumber);

                        var AssignmentID = "";
                        foreach (var relas in asm.WorkRelationships)
                        {
                            foreach (var ass in relas.Assignments)
                            {
                                AssignmentID = $"{ass.AssignmentId}";
                                break;
                            }
                        }

                        if (string.IsNullOrWhiteSpace(AssignmentID))
                        {
                            throw new Exception($"在取得 Line Manager 的 AssignmentID 時沒取到資料!");
                        }

                        var PersonID = $"{manager.PersonId}";

                        var content1 = new
                        {
                            ManagerAssignmentId = AssignmentID, //如果在Create Worker中已獲取並保存，
                                                                //则此步可不用执行；
                                                                //若未保存，则需要执行查询獲取。 
                            ActionCode = "MANAGER_CHANGE",
                            ManagerId = PersonID, //ManagerId为Manager這個用戶所對應的PersonId
                            ManagerType = "LINE_MANAGER"
                        };
                        var mr3 = HttpPatchFromOracleAP(EmployeeAssignmentsSelfUrl, content1);
                        if (mr3.StatusCode == "200")
                        {
                            rst = true;
                            log.WriteLog("5", $"EmpNO={EmpNo} 更新 Line Manager 為 {ManagerEmpNo} 成功!");
                            //byte[] bs64_bytes1 = Convert.FromBase64String(mr2.ReturnData);
                            //string desc_str1 = Encoding.UTF8.GetString(bs64_bytes1);

                        }
                        else
                        {
                            log.WriteErrorLog($"更新 Line Manager 失敗:{mr3.ErrorMessage}{mr3.ReturnData} (主管的AssignmentID={AssignmentID}, PersonID={PersonID})");
                        }


                    }


                }
                #endregion
            }
            catch (Exception exall)
            {
                log.WriteErrorLog($"Update_DefaultExpenseAccount_JobLevel_LineManager ERROR:{exall.Message}");
                //throw;
            }



            return rst;
        }

        //{{oracle_test_domain}}/hcmRestApi/resources/latest/workers?q=PersonNumber=801524&expand=workRelationships.assignments
        //Get Assignment URL
        public void Get_workrelationship_assignments_Info(string EmpNo,
            out string workrelationship_assignments_self_href,
            out string EffectiveStartDate,
            out string EffectiveEndDate,
            out string jobid_in_oracle,
            out string jobCode_in_oracle,
            out string defExpenseAccount1_in_oracle,
            out string AssignmentId
            )
        {
            workrelationship_assignments_self_href = "";
            EffectiveStartDate = "";
            EffectiveEndDate = "";
            jobid_in_oracle = "";
            jobCode_in_oracle = "";
            defExpenseAccount1_in_oracle = "";
            AssignmentId = "";
            try
            {
                var url = $"{Oracle_Domain}/hcmRestApi/resources/latest/workers?q=PersonNumber={EmpNo}&expand=workRelationships.assignments";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                TmpOracleReturnObj20230119_3 rtnobj = Newtonsoft.Json.JsonConvert.DeserializeObject<TmpOracleReturnObj20230119_3>(desc_str);
                if (rtnobj.Count > 0)
                {
                    OracleWorkerAssignment wa = rtnobj.Items[0];
                    workrelationship_assignments_self_href = (from link in wa.WorkRelationships[0].Assignments[0].Links
                                                              where link.Rel == "self"
                                                                && link.Name == "assignments"
                                                              select link.Href).FirstOrDefault();
                    EffectiveStartDate = wa.WorkRelationships[0].Assignments[0].EffectiveStartDate;
                    EffectiveEndDate = wa.WorkRelationships[0].Assignments[0].EffectiveEndDate;
                    jobCode_in_oracle = $"{wa.WorkRelationships[0].Assignments[0].JobCode}";
                    jobid_in_oracle = $"{wa.WorkRelationships[0].Assignments[0].JobId}";
                    defExpenseAccount1_in_oracle = $"{wa.WorkRelationships[0].Assignments[0].DefaultExpenseAccount}";
                    AssignmentId = $"{wa.WorkRelationships[0].Assignments[0].AssignmentId}";
                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"Get_workrelationship_assignments_Info error:取得Work資料失敗:{ex.Message}");
            }
        }


        /// <summary>
        /// 取得員工的主管資訊       
        /// /** 获取URL邏輯說明
        /// 需要先檢查當前用戶是否指派經理，並據此作為判斷結果，而會使用到不同的Manager Assignment URL，及後續的執行方法。檢查過程如下：
        /// 1）獲取返回結果中的items.workRelationships.assignments.Managers節點
        /// 2）如果該節點為空，則代表目前這個用戶還沒有分派經理，需獲取items.workRelationships.assignments.links下的 rel='child' and name='managers'節點所對應的href，見下截圖，並且後續需要使用步驟2.2.4.4；
        /// 3）如果該節點不為空，則代表目前這個用戶已經分派經理，需獲取items.workRelationships.assignments.managers.links下的 rel='self' and name='managers'節點所對應的href，見下截圖並且後續需要使用步驟2.2.4.3；
        /// **/
        /// </summary>
        /// <param name="EmpNo"></param>
        /// <param name="ManagerUpdateURL"></param>
        public void AddOrUpdateManager(string EmpNo, string ManagerEmpNo,
            string startDt = "", string endDt = "")
        {
            var managerEmpNo = ManagerEmpNo.Replace("A", "");
            try
            {

                //若沒有就取得 EmpNo 的資料
                var url_EmpNo = "";
                var startDt_EmpNo = "";
                var endDt_EmpNo = "";
                var jobid_in_oracle_EmpNo = "";
                var jobCode_in_oracle_EmpNo = "";
                var defExpenseAccount1_in_oracle_EmpNo = "";
                var ManagerAssignmentId_EmpNo = "";
                if (string.IsNullOrWhiteSpace(startDt))
                {
                    Get_workrelationship_assignments_Info(EmpNo, out url_EmpNo,
                        out startDt_EmpNo, out endDt_EmpNo,
                       out jobid_in_oracle_EmpNo, out jobCode_in_oracle_EmpNo,
                       out defExpenseAccount1_in_oracle_EmpNo,
                       out ManagerAssignmentId_EmpNo);
                }
                else
                {
                    startDt_EmpNo = startDt;
                    endDt_EmpNo = endDt;
                }


                //取得主管資訊
                var url_Manager = "";
                var startDt_Manager = "";
                var endDt_Manager = "";
                var jobid_in_oracle_Manager = "";
                var jobCode_in_oracle_Manager = "";
                var defExpenseAccount1_in_oracle_Manager = "";
                var ManagerAssignmentId = "";
                Get_workrelationship_assignments_Info(managerEmpNo, out url_Manager,
                    out startDt_Manager, out endDt_Manager,
                   out jobid_in_oracle_Manager, out jobCode_in_oracle_Manager,
                   out defExpenseAccount1_in_oracle_Manager,
                   out ManagerAssignmentId);



                //取得 EmpNo 該員的主管資料
                var url = $"{Oracle_Domain}/hcmRestApi/resources/latest/workers?q=PersonNumber={EmpNo}&expand=workRelationships.assignments.managers";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                TmpReturnObj20230130 rtnobj = Newtonsoft.Json.JsonConvert.DeserializeObject<TmpReturnObj20230130>(desc_str);
                if (rtnobj.Count > 0)
                {
                    OracleWorkerManagerInfo ma = rtnobj.Items[0];
                    if (ma.WorkRelationships[0].Assignments[0].Managers.Length == 0)
                    {
                        //2）如果該節點為空，則代表目前這個用戶還沒有分派經理，
                        //需獲取items.workRelationships.assignments.links下的
                        //rel='child' and name='managers'節點所對應的href，
                        //見下截圖，並且後續需要使用步驟2.2.4.4；
                        var AddManagerURL = (from urls in ma.WorkRelationships[0].Assignments[0].Links
                                             where urls.Rel == "child"
                                                && urls.Name == "managers"
                                             select urls.Href).FirstOrDefault();
                        Dictionary<string, string> header = new Dictionary<string, string>();
                        header.Add("Effective-Of", $"RangeStartDate={startDt_EmpNo};RangeEndDate={endDt_EmpNo}");
                        var pars = new
                        {
                            ManagerAssignmentId = ManagerAssignmentId,
                            ActionCode = "HIRE",
                            ManagerType = "LINE_MANAGER"
                        };
                        var mr2 = HttpPostFromOracleAP_2(AddManagerURL, pars, header);
                        ///** 創建成功時，返回的Http Status Code=201 **/
                        if ((mr2.StatusCode == "201") || (mr2.StatusCode == "204"))
                        {
                            log.WriteLog("5", $"新增 LINE_MANAGER 成功!");
                        }
                        else
                        {
                            throw new Exception($"新增 LINE_MANAGER 失敗! {mr2.StatusCode} {mr2.ErrorMessage} {mr2.ReturnData}");
                        }
                    }
                    else
                    {
                        //3）如果該節點不為空，則代表目前這個用戶已經分派經理，
                        //需獲取items.workRelationships.assignments.managers.links下的
                        //rel='self' and name='managers'節點所對應的href，
                        //見下截圖並且後續需要使用步驟2.2.4.3；
                        var UpdateManagerURL = (from urls in ma.WorkRelationships[0].Assignments[0].Managers[0].Links
                                                where urls.Rel == "self"
                                                    && urls.Name == "managers"
                                                select urls.Href).FirstOrDefault();
                        Dictionary<string, string> header = new Dictionary<string, string>();
                        header.Add("Effective-Of", $"RangeMode=CORRECTION;RangeStartDate={startDt_EmpNo};RangeEndDate={endDt_EmpNo}");
                        var pars = new
                        {
                            ManagerAssignmentId = ManagerAssignmentId,
                            ManagerType = "LINE_MANAGER"
                        };
                        var mr2 = HttpPatchFromOracleAP(UpdateManagerURL, pars, header);
                        if ((mr2.StatusCode == "200") || (mr2.StatusCode == "204"))
                        {
                            log.WriteLog("5", $"更新 LINE_MANAGER 成功!");
                        }
                        else
                        {
                            throw new Exception($"新增 LINE_MANAGER 失敗! {mr2.StatusCode} {mr2.ErrorMessage} {mr2.ReturnData}");
                        }
                    }
                }
                else
                {
                    throw new Exception($"找不到主管 {managerEmpNo} 的資料! ");
                }

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"設定 {EmpNo} 的主管 {managerEmpNo} 資料失敗:{ex.Message}");
            }

        }

        /// <summary>
        /// emp api 廢除了，改用這個來更新
        /// </summary>
        /// <param name="EmpNo"></param>
        /// <param name="DeptNo"></param>
        /// <param name="ManagerEmpNo"></param>
        /// <param name="JobLevel"></param>
        /// <returns></returns>
        public bool Update_DefaultExpenseAccount_JobLevel_LineManager_2(string EmpNo, string DeptNo, string ManagerEmpNo,
            string JobLevel)
        {
            bool rst = false;
            var FTJobID = "";
            foreach (var item in GetJobLevelList())
            {
                if (JobLevel == item.ApprovalAuthority)
                {
                    FTJobID = item.JobId;
                    break;
                }
            }


            log.WriteLog("5", $"準備更新 Job ID, Default Expense Account");
            //Get Assignment URL

            try
            {
                #region 先取得 Oracle 上的資料來供比對
                ////正式環境User 的 Default Expense Account 的規則一樣比照Stage(DEV1), default account =6288099
                ////完整  0001.< 員工所屬profit center >.< 員工所屬Department > .6288099.0000.000000000.0000000.0000
                ////範例: 如 0001.POS.POS000000.6288099.0000.000000000.0000000.0000
                ////取得EmployeeAssignment自已的url
                //string EmployeeAssignmentsSelfUrl = GetEmployeeAssignmentsSelfURLByEmpNo_2(EmpNo);
                //if (string.IsNullOrWhiteSpace(EmployeeAssignmentsSelfUrl))
                //{
                //    throw new Exception("Get Employee Assignments URL失敗:(空白)");
                //}

                //先 GET 資料來看看
                OracleEmployeeAssignmentSelf2 empAsm = null;
                var jobid_in_oracle = "";
                var jobCode_in_oracle = "";
                var defExpenseAccount1_in_oracle = "";
                var AssignmentId = "";
                var ManagerAssignmentId = "";
                var url = "";
                var startDt = "";
                var endDt = "";
                Dictionary<string, string> header = new Dictionary<string, string>();

                try
                {
                    Get_workrelationship_assignments_Info(EmpNo, out url, out startDt, out endDt,
                        out jobid_in_oracle, out jobCode_in_oracle,
                        out defExpenseAccount1_in_oracle, out AssignmentId);
                    header.Add("Effective-Of", $"RangeMode=CORRECTION;RangeStartDate={startDt};RangeEndDate={endDt}");


                    ////取得 Employee Assignment
                    //empAsm = GetOracleEmployeeAssignmentSelf2ByURL(EmpNo, EmployeeAssignmentsSelfUrl);

                    //jobid_in_oracle = empAsm.JobId;
                    //defExpenseAccount1_in_oracle = empAsm.DefaultExpenseAccount;
                    //ManagerAssignmentId1 = empAsm.ManagerAssignmentId;
                    //ManagerID1 = empAsm.ManagerId;
                }
                catch (Exception exass)
                {
                    log.WriteErrorLog($"先查資料時，取得 JobID, DefaultExpenseAccount, ManagerAssignmentId 失敗:{exass.Message}");
                }
                #endregion


                #region 檢查/更新 Job Level
                //var jobLevelName = JobLevelName;
                //if (string.IsNullOrWhiteSpace(jobLevelName))
                //{
                //    jobLevelName = "職員";
                //}

                //var Jobid = "";
                //try
                //{
                //    Jobid = JobLevelCollection[jobLevelName];
                //}
                //catch (Exception exlv)
                //{
                //}

                //if (string.IsNullOrWhiteSpace(Jobid))
                //{
                //    //重新準備 Job ID
                //    log.WriteErrorLog($"重新準備 Job ID");
                //    JobLevelCollection = null;
                //    GetOracleJobLevelCollection();
                //}
                //if (string.IsNullOrWhiteSpace(Jobid))
                //{
                //    log.WriteErrorLog($"還是取不到 Job Level 的 ID! 改用 api 傳中文取");
                //    Jobid = GetJobIDByChinese(jobLevelName);
                //}


                log.WriteLog("5", $"準備更新 jobid");
                if (FTJobID == jobid_in_oracle)
                {
                    log.WriteLog("5", $"JobId 相同，不需更新!");
                }
                else
                {
                    try
                    {

                        if (string.IsNullOrWhiteSpace(FTJobID))
                        {
                            throw new Exception("飛騰 JobID 是空的!");
                        }
                        var content = new
                        {
                            JobId = FTJobID
                        };
                        var mr = HttpPatchFromOracleAP(url, content, header);
                        if (mr.StatusCode == "200")
                        {
                            rst = true;
                            log.WriteLog("5", $"更新 JobId 成功!");
                            //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                            //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                        }
                        else
                        {
                            var errmsg = $"更新 JobId 失敗: {mr.StatusCode} {mr.ErrorMessage} {mr.ReturnData}";
                            log.WriteErrorLog(errmsg);
                        }

                    }
                    catch (Exception exJobID)
                    {
                        log.WriteErrorLog($"Update JobID 失敗:{exJobID.Message}");
                    }
                }

                #endregion

                #region 檢查/更新 Default Expense Account
                var deptNo = DeptNo.Substring(0, 9);
                var profit_center = deptNo.Substring(0, 3);
                var Default_Expense_Account = $"0001.{profit_center}.{deptNo}.6288099.0000.000000000.0000000.0000";
                if (defExpenseAccount1_in_oracle != Default_Expense_Account)
                {
                    log.WriteLog("5", $"準備更新 Default Expense Account!");

                    var content_2 = new
                    {
                        DefaultExpenseAccount = Default_Expense_Account
                    };
                    var mr2 = HttpPatchFromOracleAP(url, content_2, header);
                    if (mr2.StatusCode == "200")
                    {
                        rst = true;
                        log.WriteLog("5", $"更新 DefaultExpenseAccount 成功!");
                        //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                        //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                    }
                    else
                    {
                        log.WriteLog("5", $"更新 DefaultExpenseAccount 失敗! {mr2.StatusCode} {mr2.ErrorMessage} {mr2.ReturnData}");
                    }
                }
                else
                {
                    log.WriteLog("5", $"Default_Expense_Account 相同，不需更新");
                }
                #endregion

                #region 檢查/更新 Line Manager 主管
                if (string.IsNullOrWhiteSpace(ManagerEmpNo))
                {
                    log.WriteLog("5", $"傳入的主管員編為空，不更新 Line Manager");
                }
                else
                {
                    log.WriteLog("5", $"準備 Get ManagerEmpNo={ManagerEmpNo} 的 Employee 物件!");
                    //var manager = GetOracleEmployeeByEmpNo(ManagerEmpNo);

                    //2023/01/29 做到這裡 startDt, out endDt,
                    AddOrUpdateManager(EmpNo, ManagerEmpNo, startDt, endDt);


                    //if (string.IsNullOrWhiteSpace(url_Manager))
                    //{
                    //    log.WriteLog("5", $"查無 {ManagerEmpNo}) 的 資料!");
                    //}
                    //else
                    //{

                    //    ////取得workRelationships.assignments
                    //    //Employee_workRelationships_Assignment asm = GetEmployee_workRelationships_assignments_ByPersonNumber(manager.PersonNumber);

                    //    //var AssignmentID = "";
                    //    //foreach (var relas in asm.WorkRelationships)
                    //    //{
                    //    //    foreach (var ass in relas.Assignments)
                    //    //    {
                    //    //        AssignmentID = $"{ass.AssignmentId}";
                    //    //        break;
                    //    //    }
                    //    //}

                    //    //if (string.IsNullOrWhiteSpace(AssignmentID))
                    //    //{
                    //    //    throw new Exception($"在取得 Line Manager 的 AssignmentID 時沒取到資料!");
                    //    //}

                    //    //var PersonID = $"{manager.PersonId}";

                    //    //var content1 = new
                    //    //{
                    //    //    ManagerAssignmentId = AssignmentID, //如果在Create Worker中已獲取並保存，
                    //    //                                        //则此步可不用执行；
                    //    //                                        //若未保存，则需要执行查询獲取。 
                    //    //    ActionCode = "MANAGER_CHANGE",
                    //    //    ManagerId = PersonID, //ManagerId为Manager這個用戶所對應的PersonId
                    //    //    ManagerType = "LINE_MANAGER"
                    //    //};
                    //    //var mr3 = HttpPatchFromOracleAP(url, content1);
                    //    //if (mr3.StatusCode == "200")
                    //    //{
                    //    //    rst = true;
                    //    //    log.WriteLog("5", $"EmpNO={EmpNo} 更新 Line Manager 為 {ManagerEmpNo} 成功!");
                    //    //    //byte[] bs64_bytes1 = Convert.FromBase64String(mr2.ReturnData);
                    //    //    //string desc_str1 = Encoding.UTF8.GetString(bs64_bytes1);

                    //    //}
                    //    //else
                    //    //{
                    //    //    log.WriteErrorLog($"更新 Line Manager 失敗:{mr3.ErrorMessage}{mr3.ReturnData} (主管的AssignmentID={AssignmentID}, PersonID={PersonID})");
                    //    //}


                    //}


                }
                #endregion
            }
            catch (Exception exall)
            {
                log.WriteErrorLog($"Update_DefaultExpenseAccount_JobLevel_LineManager ERROR:{exall.Message}");
                //throw;
            }



            return rst;
        }



        /// <summary>
        /// 用 Employee api 更新 username
        /// 還是要先 get, 因為需要它的 self url
        /// </summary>
        /// <param name="EmpNo"></param>
        /// <param name="UserName">要更新成的UserName</param>
        /// <returns></returns>
        public bool ModifyOracleUserNameWhenDifferent(string EmpNo, string UserName)
        {
            bool rst = false;
            log.WriteLog("5", $"準備更新 {EmpNo} 的 UserName為  {UserName}. (Into UpdateUserNameByEmployeeApi)");
            try
            {
                var emp = GetOracleEmployeeByEmpNo(EmpNo);
                if (emp == null)
                {
                    log.WriteErrorLog($"取得 {EmpNo} oracle employee 物件失敗!");
                }
                else
                {
                    var username = $"{emp.UserName}";
                    var personNumber = $"{emp.PersonNumber}";
                    var UserName和personNumber都相同 = (username.CompareTo(UserName) == 0) && (personNumber.CompareTo(EmpNo) == 0);
                    if (UserName和personNumber都相同)
                    {
                        //不需更新
                        rst = true;
                    }
                    else
                    {
                        //需更新
                        foreach (var link in emp.Links)
                        {
                            if (link.Rel == "self" && link.Name == "emps")
                            {
                                var url = link.Href;
                                var par = new
                                {
                                    UserName = UserName,
                                    PersonNumber = personNumber
                                };
                                MiddleReturn mr2 = HttpPatchFromOracleAP(url, par);
                                if (mr2.StatusCode == "200")
                                {
                                    rst = true;
                                    log.WriteLog("5", $"{EmpNo} 更新 UserName 和 PersonNumber 成功:{UserName}");
                                    //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                                    //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                                }
                                break;
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"UpdateUserNameByEmployeeApi 更新 UserName 失敗:{ex.Message}");
            }

            return rst;
        }


        /// <summary>
        /// 更新 User 的登入帳號/員工姓名/員工編號/員工英文名稱
        /// </summary>
        /// <param name="EmpNo">員編</param>
        /// <param name="UserName">oracle的登入帳號(員編)</param>
        /// <param name="EmpName">員工姓名</param>
        /// <param name="EmpEngName">員工英文名稱</param>
        /// <returns></returns>
        public bool ModifyOracleUserNameAndEmpNameWhenDifferent(string EmpNo,
            string UserName, string EmpName, string EmpEngName, string email)
        {
            var _email = email;
            if (string.IsNullOrWhiteSpace(_email))
            {
                _email = "xxx@staff.pchome.com.tw";
            }
            bool rst = false;
            log.WriteLog("5", $"準備更新 {EmpNo} {EmpName} 的 User 資料. (Into ModifyOracleUserNameAndEmpNameWhenDifferent)");
            try
            {
                var emp = GetOracleEmployeeByEmpNo(EmpNo);
                if (emp == null)
                {
                    log.WriteErrorLog($"取得 {EmpNo} {EmpName} oracle employee 物件失敗!");
                }
                else
                {

                    var oracleData = $"{emp.UserName},{emp.PersonNumber},{emp.FirstName},{emp.MiddleName},{emp.LastName},{emp.WorkEmail}";
                    var ftData = $"{UserName},{EmpNo},{EmpName},{EmpEngName},{EmpNo},{_email}";
                    log.WriteLog("5", "UserName,EmpNo,員工姓名,英文名稱,員工編號,email");
                    log.WriteLog("5", $"oracleData={oracleData}");
                    log.WriteLog("5", $"飛騰資料={ftData}");

                    if (oracleData == ftData)
                    {
                        //不需更新
                        rst = true;
                        log.WriteLog("5", $"資料一樣，不需更新");
                    }
                    else
                    {
                        log.WriteLog("5", $"資料不同，需要更新");
                        //需更新
                        foreach (var link in emp.Links)
                        {
                            if (link.Rel == "self" && link.Name == "emps")
                            {
                                var url = link.Href;
                                log.WriteLog("5", $"更新 emp url={url}");
                                var par = new
                                {
                                    UserName = UserName,
                                    PersonNumber = EmpNo,
                                    FirstName = EmpName,
                                    MiddleName = EmpEngName,
                                    LastName = EmpNo,
                                    WorkEmail = _email
                                };
                                var _par = Newtonsoft.Json.JsonConvert.SerializeObject(par);
                                log.WriteLog("5", $"url={url}{Environment.NewLine}{_par}");
                                MiddleReturn mr2 = HttpPatchFromOracleAP(url, par);
                                if (mr2.StatusCode == "200")
                                {
                                    rst = true;
                                    log.WriteLog("5", $"{EmpNo} {EmpName} 更新 User 資料為 飛騰資料 成功!");
                                    //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                                    //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                                }
                                break;
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"ModifyOracleUserNameAndEmpNameWhenDifferent 更新 User 資料失敗:{ex.Message}");
            }

            return rst;
        }


        public bool UpdateUserName_2(string EmpNo, string EmpName, string UpdateURL)
        {
            bool rst = false;
            try
            {
                var par = new
                {
                    Username = EmpNo
                };
                MiddleReturn mr2 = HttpPatchFromOracleAP(UpdateURL, par);
                if (mr2.StatusCode == "200")
                {
                    rst = true;
                    //log.WriteLog("5", $"{EmpNo} {EmpName} 更新 Worker 中文姓名/英文姓名 資料為 飛騰資料 成功!");
                    //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                    //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                }

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{EmpNo} {EmpName} 更新 username 失敗:{ex.Message}");
            }
            return rst;
        }

        /// <summary>
        /// 更新 UserName / email
        /// 因為 employee api 要廢除了
        /// 所以要改用其他 api 執行
        /// </summary>
        /// <param name="EmpNo">員工編號</param>
        /// <param name="UserName">員工編號</param>
        /// <param name="EmpName">中文姓名</param>
        /// <param name="EmpEngName">英文姓名</param>
        /// <param name="email"></param>
        /// <returns></returns>
        public bool ModifyOracleUserNameAndEmpNameWhenDifferent_2(string EmpNo,
            string UserName, string EmpName, string EmpEngName, string email)
        {
            var _email = email;
            if (string.IsNullOrWhiteSpace(_email))
            {
                _email = "xxx@staff.pchome.com.tw";
            }
            bool rst = false;
            log.WriteLog("5", $"準備更新 {EmpNo} {EmpName} 的 User 資料. (Into ModifyOracleUserNameAndEmpNameWhenDifferent_2)");
            try
            {
                var oracleUserName = "";
                var oraclePersonNumber = "";
                var UserNameUpdateURL = "";
                var oracleUser = GetOracleUserByEmpNo_2(EmpNo);
                if (oracleUser != null)
                {
                    oracleUserName = oracleUser.Username;
                    oraclePersonNumber = oracleUser.PersonNumber;
                    if (oracleUserName != EmpNo)
                    {

                        var tmp = (from link in oracleUser.Links
                                   where link.Rel == "self"
                                   select link.Href).FirstOrDefault();
                        if (!string.IsNullOrWhiteSpace(tmp))
                        {
                            UserNameUpdateURL = tmp;
                        }
                        UpdateUserName_2(EmpNo, EmpName, UserNameUpdateURL);
                    }
                }




                var oracleEmpName = "";
                var oracleEmpEngName = "";
                var oracleEmail = "";
                var oracleNamesUpdateURL = "";
                var EffectiveStartDate = "";
                var EffectiveEndDate = "";
                var oracleEmailUpdateURL = "";
                GetOracleWorkerFirstNameMiddleNameLastNameEmail(EmpNo, out oracleEmpName, out oracleEmpEngName,
                    out oracleEmail, out oracleNamesUpdateURL, out oracleEmailUpdateURL,
                    out EffectiveStartDate, out EffectiveEndDate);

                var oracleData = $"{oracleUserName},{oraclePersonNumber},{oracleEmpName},{oracleEmpEngName},{oracleEmail}";
                var ftData = $"{UserName},{EmpNo},{EmpName},{EmpEngName},{_email}";
                log.WriteLog("5", "UserName,EmpNo(PersonNumber),員工姓名,英文名稱,email");
                log.WriteLog("5", $"oracleData={oracleData}");
                log.WriteLog("5", $"飛騰資料={ftData}");

                if (oracleData == ftData)
                {
                    //不需更新
                    rst = true;
                    log.WriteLog("5", $"資料一樣，不需更新");
                }
                else
                {
                    log.WriteLog("5", $"資料不同，需要更新");
                    //需更新


                    if (!string.IsNullOrWhiteSpace(oracleNamesUpdateURL))
                    {
                        //更新 名稱 
                        Dictionary<string, string> addHeaders = new Dictionary<string, string>();
                        addHeaders.Add("Effective-Of", $"RangeMode=CORRECTION;RangeStartDate={EffectiveStartDate};RangeEndDate={EffectiveEndDate}");
                        var par = new
                        {
                            LastName = EmpNo,
                            FirstName = EmpName,
                            MiddleNames = EmpEngName
                        };
                        MiddleReturn mr2 = HttpPatchFromOracleAP(oracleNamesUpdateURL, par, addHeaders);
                        if (mr2.StatusCode == "200")
                        {
                            rst = true;
                            log.WriteLog("5", $"{EmpNo} {EmpName} 更新 Worker 中文姓名/英文姓名 資料為 飛騰資料 成功!");
                            //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                            //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                        }
                    }

                    if (!string.IsNullOrWhiteSpace(oracleEmailUpdateURL))
                    {
                        //更新 email
                        var par = new
                        {
                            EmailAddress = _email
                        };
                        MiddleReturn mr2 = HttpPatchFromOracleAP(oracleEmailUpdateURL, par);
                        if (mr2.StatusCode == "200")
                        {
                            rst = true;
                            log.WriteLog("5", $"{EmpNo} {EmpName} 更新 Worker Email 資料為 飛騰資料 成功!");
                            //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                            //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                        }
                    }

                    //foreach (var link in oracleUser.Links)
                    //{
                    //    if (link.Rel == "self" && link.Name == "emps")
                    //    {
                    //        var url = link.Href;
                    //        log.WriteLog("5", $"更新 emp url={url}");
                    //        var par = new
                    //        {
                    //            UserName = UserName,
                    //            PersonNumber = EmpNo,
                    //            FirstName = EmpName,
                    //            MiddleName = EmpEngName,
                    //            LastName = EmpNo,
                    //            WorkEmail = _email
                    //        };
                    //        var _par = Newtonsoft.Json.JsonConvert.SerializeObject(par);
                    //        log.WriteLog("5", $"url={url}{Environment.NewLine}{_par}");
                    //        MiddleReturn mr2 = HttpPatchFromOracleAP(url, par);
                    //        if (mr2.StatusCode == "200")
                    //        {
                    //            rst = true;
                    //            log.WriteLog("5", $"{EmpNo} {EmpName} 更新 User 資料為 飛騰資料 成功!");
                    //            //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                    //            //string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                    //        }
                    //        break;
                    //    }
                    //}
                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"ModifyOracleUserNameAndEmpNameWhenDifferent_2 更新 User 資料失敗:{ex.Message}");
            }

            return rst;
        }



        /// <summary>
        /// 用 user api 更新員工編號
        /// </summary>
        /// <param name="OldEmpNo"></param>
        /// <param name="NewUserName"></param>
        public bool UpdateUserName(string OldEmpNo, string NewUserName)
        {
            bool rst = false;
            if (!string.IsNullOrWhiteSpace(OldEmpNo))
            {
                if (!string.IsNullOrWhiteSpace(NewUserName))
                {
                    try
                    {
                        // ORACLE STAGE
                        // for  oracle ap
                        string guid = GetGUIDByEmpNo(OldEmpNo);
                        if (string.IsNullOrWhiteSpace(guid))
                        {
                            throw new Exception("guid 空白!");
                        }
                        var par = new
                        {
                            Username = NewUserName
                        };
                        MiddleModel2 send_model2 = new MiddleModel2();
                        string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/userAccounts/{guid}";
                        send_model2.URL = url;
                        send_model2.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(par);
                        send_model2.Method = "PATCH";
                        //string username = this.UserName;
                        //string password = this.Password;
                        send_model2.UserName = this.UserName;
                        send_model2.Password = this.Password;
                        string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                        send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                        //CredentialCache myCred = new CredentialCache();
                        //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                        //send_model2.Cred = myCred;
                        send_model2.Timeout = int30Mins;


                        // for BOXMAN API
                        MiddleModel send_model = new MiddleModel();
                        var _url = $"{Oracle_AP}/api/Middle/Call/";
                        send_model.URL = _url;
                        send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                        send_model.Method = "POST";
                        send_model.Timeout = int30Mins;
                        var ret = ApiOperation.CallApi(
                                new ApiRequestSetting()
                                {
                                    Data = send_model,
                                    MethodRoute = "api/Middle/Call",
                                    MethodType = "POST",
                                    TimeOut = int30Mins
                                }
                                );
                        if (ret.Success)
                        {
                            string receive_str = ret.ResponseContent;
                            try
                            {
                                MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                                MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2.ReturnData);
                                byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                                string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                                OracleUser usr = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleUser>(desc_str);
                                if (usr != null)
                                {
                                    if (usr.Username.CompareTo(NewUserName) == 0)
                                    {
                                        rst = true;
                                    }
                                }

                            }
                            catch (Exception exbs64)
                            {
                                log.WriteErrorLog($"Get Workers By Page 失敗:{exbs64.Message}{exbs64.InnerException}");
                            }
                        }
                        else
                        {
                            log.WriteErrorLog($"Call Boxman api {_url} 失敗:{ret.ErrorMessage}. {ret.ErrorException}");
                        }
                    }
                    catch (Exception ex)
                    {
                        log.WriteErrorLog($"Call UpdateEmpNo 失敗:{ex.Message}. {ex.InnerException}");

                    }
                }
            }
            return rst;
        }


        public List<OracleUser> GetAllOracleUsers()
        {
            bool hasData = true;
            int iCurrent_page = 1;
            int iRecCntPerPage = 100;
            List<OracleUser> all_users = new List<OracleUser>();
            while (hasData)
            {
                string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/userAccounts?limit={iRecCntPerPage}";
                if (iCurrent_page > 1)
                {
                    int offect = iRecCntPerPage * (iCurrent_page - 1);// iRecCntPerPage * (iCurrent_page - 1);

                    url = $"{url}&offset={offect}";
                }
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                {
                    throw new Exception($"Get Oracle users 失敗:{mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                if (string.IsNullOrWhiteSpace(desc_str))
                {
                    throw new Exception($"取得 oracle users 失敗, 回傳空白!");
                }
                OracleApiReturnObj<OracleUser> rtnobj = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleApiReturnObj<OracleUser>>(desc_str);
                if (rtnobj == null)
                {
                    throw new Exception("解析失敗!");
                }
                hasData = rtnobj.HasMore;
                all_users.AddRange(rtnobj.Items);
                iCurrent_page++;
            }
            return all_users;
        }

        public OracleAllUsersBySCIM AllUsersBySCIM { get; set; }


        /// <summary>
        /// 2022/06/29
        /// 因為新建帳號就不會包含在這裡面，所以還不如動態去抓，
        /// 所以這個廢除不用了
        /// 2022/05/18
        /// 這個準備取代 : 每次都要把所有 oracle user 寫入 sqlite 
        /// </summary>
        public void GetAllUserBySCIM()
        {
            try
            {
                log.WriteLog("5", "準備透過 SCIM api 取得 oracle SCIM user 回來");

                string url = $"{Oracle_Domain}/hcmRestApi/scim/Users";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                {
                    throw new Exception($"Get Oracle SCIM users 失敗:{mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                //log.WriteLog("5", $"Get Oracle SCIM users 得到:{desc_str}");
                if (string.IsNullOrWhiteSpace(desc_str))
                {
                    throw new Exception($"取得 oracle SCIM users 失敗, 回傳空白!");
                }
                AllUsersBySCIM = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleAllUsersBySCIM>(desc_str);
                if (AllUsersBySCIM == null)
                {
                    throw new Exception("解析 oracle SCIM users 失敗!");
                }
                else
                {
                    log.WriteLog("5", $"總共取得 {AllUsersBySCIM.ItemsPerPage} 位 SCIM user");
                }

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"SyncUsers 失敗:{ex.Message}{ex.InnerException}");
            }

        }


        public OracleEmployeeResultModel<OracleEmployeeWithAssignments> GetOracleEmployeeWithAssignmentByEmpNo(string EmpNo)
        {
            var desc_str = "";
            OracleEmployeeResultModel<OracleEmployeeWithAssignments> rtnobj = null;
            try
            {
                string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/emps?q=LastName=\"{EmpNo}\"&expand=assignments";
                log.WriteLog("5", $"GetOracleEmployeeWithAssignmentByEmpNo url={url}");
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                {
                    throw new Exception($"Get Oracle users 失敗:{mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                desc_str = Encoding.UTF8.GetString(bs64_bytes);
                //log.WriteLog("5", $"取得:{desc_str}");
                if (string.IsNullOrWhiteSpace(desc_str))
                {
                    throw new Exception($"取得 oracle users 失敗, 回傳空白!");
                }
                rtnobj =
                    Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployeeResultModel<OracleEmployeeWithAssignments>>(desc_str);
                if (rtnobj == null)
                {
                    throw new Exception("解析失敗!");
                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"GetOracleEmployeeWithAssignmentByEmpNo Error:{ex.Message} {desc_str}");
            }
            return rtnobj;
        }

        /// <summary>
        /// 取得 user 的中文姓名/英文姓名
        /// </summary>
        /// <param name="EmpNo"></param>
        /// <param name="EmpName"></param>
        /// <param name="EmpEngName"></param>
        public void GetOracleWorkerFirstNameMiddleNameLastNameEmail(string EmpNo, out string EmpName, out string EmpEngName,
            out string Email, out string NamesUpdateURL, out string oracleEmailUpdateURL,
            out string EffectiveStartDate, out string EffectiveEndDate)
        {
            var desc_str = "";
            EmpName = "";
            EmpEngName = "";
            Email = "";
            NamesUpdateURL = "";
            oracleEmailUpdateURL = "";
            EffectiveStartDate = "";
            EffectiveEndDate = "";
            try
            {
                //string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/userAccounts?q=PersonNumber={EmpNo}";
                string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/workers?q=PersonNumber={EmpNo}&expand=names,emails";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                {
                    throw new Exception($"Get Oracle users 失敗:{mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                desc_str = Encoding.UTF8.GetString(bs64_bytes);
                //log.WriteLog("5", $"取得:{desc_str}");
                if (string.IsNullOrWhiteSpace(desc_str))
                {
                    throw new Exception($"取得 oracle users 失敗, 回傳空白!");
                }
                var rtnobj =
                    Newtonsoft.Json.JsonConvert.DeserializeObject<OracleNewQuitEmployee.ORSyncOracleData.Model.OracleWorker.TmpReturnObj20230119_1>(desc_str);
                if (rtnobj == null)
                {
                    throw new Exception("解析失敗!");
                }
                if (rtnobj.Count > 0)
                {
                    OracleNewQuitEmployee.ORSyncOracleData.Model.OracleWorker.OracleWorker worker = rtnobj.Items[0];
                    if (worker.Names.Length > 0)
                    {
                        EmpName = worker.Names[0].FirstName;
                        EmpEngName = worker.Names[0].MiddleNames;
                        EffectiveStartDate = worker.Names[0].EffectiveStartDate;
                        EffectiveEndDate = worker.Names[0].EffectiveEndDate;
                        var tmp = (from link in worker.Names[0].Links
                                   where link.Rel == "self" && link.Name == "names"
                                   select link.Href).FirstOrDefault();
                        if (!string.IsNullOrWhiteSpace(tmp))
                        {
                            NamesUpdateURL = tmp;
                        }
                    }
                    if (worker.Emails.Length > 0)
                    {
                        Email = worker.Emails[0].EmailAddress;
                        var tmp = (from link in worker.Emails[0].Links
                                   where link.Rel == "self" && link.Name == "emails"
                                   select link.Href).FirstOrDefault();
                        if (!string.IsNullOrWhiteSpace(tmp))
                        {
                            oracleEmailUpdateURL = tmp;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"GetOracleUserNames Error:{ex.Message} {desc_str}");
            }
        }

        /// <summary>
        /// 這是另一個取得 oracle user 的方式
        /// </summary>
        /// <param name="EmpNo"></param>
        public OracleUser2 GetOracleUserByEmpNo_2(string EmpNo)
        {
            var desc_str = "";
            OracleUser2 usr = null;
            try
            {
                string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/userAccounts?q=PersonNumber={EmpNo}";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                {
                    throw new Exception($"Get Oracle users 失敗:{mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                desc_str = Encoding.UTF8.GetString(bs64_bytes);
                //log.WriteLog("5", $"取得:{desc_str}");
                if (string.IsNullOrWhiteSpace(desc_str))
                {
                    throw new Exception($"取得 oracle users 失敗, 回傳空白!");
                }
                var rtnobj =
                    Newtonsoft.Json.JsonConvert.DeserializeObject<TmpResult20230119>(desc_str);
                if (rtnobj == null)
                {
                    throw new Exception("解析失敗!");
                }
                if (rtnobj.Count > 0)
                {
                    usr = rtnobj.Items[0];
                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"GetOracleUserByEmpNo Error:{ex.Message} {desc_str}");
            }
            return usr;
        }

        public OracleResultModel2<OracleUserResultModel> GetOracleUserByEmpNo(string EmpNo)
        {
            var desc_str = "";
            OracleResultModel2<OracleUserResultModel> rtnobj = null;
            try
            {
                string url = $"{Oracle_Domain}/hcmRestApi/scim/Users?filter=name.familyName eq \"{EmpNo}\"";
                log.WriteLog("5", $"GetOracleUserByEmpNo url={url}");
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                {
                    throw new Exception($"Get Oracle users 失敗:{mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                desc_str = Encoding.UTF8.GetString(bs64_bytes);
                //log.WriteLog("5", $"取得:{desc_str}");
                if (string.IsNullOrWhiteSpace(desc_str))
                {
                    throw new Exception($"取得 oracle users 失敗, 回傳空白!");
                }
                rtnobj =
                    Newtonsoft.Json.JsonConvert.DeserializeObject<OracleResultModel2<OracleUserResultModel>>(desc_str);
                if (rtnobj == null)
                {
                    throw new Exception("解析失敗!");
                }

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"GetOracleUserByEmpNo Error:{ex.Message} {desc_str}");
            }
            return rtnobj;
        }


        public List<OracleUser> All_Users { get; set; }


        /// <summary>
        /// 同步 oracle users 回來
        /// </summary>
        /// <param name="BatchNo"></param>
        public void SyncUsers(string BatchNo)
        {
            try
            {
                log.WriteLog("4", "準備同步 oracle user 回來");
                bool hasData = true;
                int iCurrent_page = 1;
                int iRecCntPerPage = 100;
                if (All_Users == null)
                {
                    All_Users = new List<OracleUser>();
                }
                All_Users.Clear();
                while (hasData)
                {
                    string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/userAccounts?limit={iRecCntPerPage}";
                    if (iCurrent_page > 1)
                    {
                        int offect = iRecCntPerPage * (iCurrent_page - 1);// iRecCntPerPage * (iCurrent_page - 1);

                        url = $"{url}&offset={offect}";
                    }
                    log.WriteLog("5", $"url={url}");
                    MiddleReturn mr = HttpGetFromOracleAP(url);
                    if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                    {
                        throw new Exception($"Get Oracle users 失敗:{mr.ErrorMessage}");
                    }
                    var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                    var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                    if (string.IsNullOrWhiteSpace(desc_str))
                    {
                        throw new Exception($"取得 oracle users 失敗, 回傳空白!");
                    }
                    OracleApiReturnObj<OracleUser> rtnobj = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleApiReturnObj<OracleUser>>(desc_str);
                    if (rtnobj == null)
                    {
                        throw new Exception("解析失敗!");
                    }
                    hasData = rtnobj.HasMore;
                    All_Users.AddRange(rtnobj.Items);
                    iCurrent_page++;





                    //var par = new
                    //{
                    //    iCurrent_page = iCurrent_page,
                    //    iRecCntPerPage = iRecCntPerPage
                    //};
                    ////var ret = ApiOperation.CallApi<string>("api/BoxmanOracleEmployee/BoxmanGetUsersByPagesAsync", WebRequestMethods.Http.Post, par);
                    //DateTime dt_start = DateTime.Now;
                    //var ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                    //{
                    //    MethodRoute = "api/BoxmanOracleEmployee/BoxmanGetUsersByPagesAsync",
                    //    Data = par,
                    //    MethodType = "POST",
                    //    TimeOut = 1000 * 60 * 20
                    //}
                    //);
                    //var diff = DateTime.Now.Subtract(dt_start).TotalSeconds;
                    //Console.WriteLine($"api call use {diff} sec(s).");



                    //if (string.IsNullOrWhiteSpace(ret.ErrorMessage))
                    //{
                    //    OracleApiReturnObj<OracleUser> rtnobj = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleApiReturnObj<OracleUser>>(ret.Data);
                    //    if (rtnobj == null)
                    //    {
                    //        throw new Exception("解析失敗!");
                    //    }
                    //    hasData = rtnobj.HasMore;
                    //    all_users.AddRange(rtnobj.Items);
                    //    iCurrent_page++;

                    //}
                    //else
                    //{
                    //    string errmsg = $"呼叫 api/BoxmanOracleEmployee/BoxmanGetUsersByPagesAsync 失敗:{ret.StatusCode} {ret.StatusDescription} {ret.ErrorMessage}";
                    //    //Console.WriteLine(errmsg);
                    //    throw new Exception(errmsg);
                    //}
                }

                log.WriteLog("4", $"準備寫入 {All_Users.Count} 筆 user 到 SQLite");
                SaveOracleUsersIntoSQLite(All_Users, BatchNo);
                //string content = Newtonsoft.Json.JsonConvert.SerializeObject(all_users);
                //string all_oracle_users_json_obj_file_path = $@"{AppDomain.CurrentDomain.BaseDirectory}all_oracle_users_json_obj_file_path.json";
                //System.IO.File.WriteAllText(all_oracle_users_json_obj_file_path, content);
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"SyncUsers 失敗:{ex.Message}{ex.InnerException}");
            }
        }


        public void SyncWorkers(string BatchNo)
        {
            try
            {
                // Get all workers

                //var par = new
                //{
                //    companyCustNo = "H1660610201",
                //    shipNoList = new object[] { }
                //};

                log.WriteLog("4", $"準備同步 oracle Worker 回來");

                bool hasData = true;
                int iCurrent_page = 1;
                int iRecCntPerPage = 100;
                List<Worker> Workers = new List<Worker>();
                int current_cnt = 0;
                bool hasMore = true;
                while ((hasMore) && (current_cnt < 100))
                {
                    current_cnt++;


                    // ORACLE STAGE
                    // for  oracle ap
                    MiddleModel2 send_model2 = new MiddleModel2();
                    string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/workers/?limit={iRecCntPerPage}";
                    if (iCurrent_page > 1)
                    {
                        int offect = iRecCntPerPage * (iCurrent_page - 1);// iRecCntPerPage * (iCurrent_page - 1);

                        url = $"{url}&offset={offect}";

                    }
                    log.WriteLog("4", $"準備呼叫 {url}");
                    send_model2.URL = url;
                    //send_model2.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(par);
                    send_model2.Method = "GET";
                    //string username = this.UserName;
                    //string password = this.Password;
                    send_model2.UserName = this.UserName;
                    send_model2.Password = this.Password;
                    string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                    send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                    //CredentialCache myCred = new CredentialCache();
                    //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                    //send_model2.Cred = myCred;
                    send_model2.Timeout = int30Mins;


                    // for BOXMAN API
                    MiddleModel send_model = new MiddleModel();
                    var _url = $"{Oracle_AP}/api/Middle/Call/";
                    send_model.URL = _url;
                    send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                    send_model.Method = "POST";
                    send_model.Timeout = int30Mins;
                    var ret = ApiOperation.CallApi(
                            new ApiRequestSetting()
                            {
                                Data = send_model,
                                MethodRoute = "api/Middle/Call",
                                MethodType = "POST",
                                TimeOut = int30Mins
                            }
                            );
                    if (ret.Success)
                    {
                        string receive_str = ret.ResponseContent;
                        try
                        {
                            MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                            if (!string.IsNullOrWhiteSpace(mr2.ErrorMessage))
                            {
                                var _errmsg = $"SyncWorkers 失敗:{mr2.StatusCode} {mr2.StatusDescription} {mr2.ReturnData} {mr2.ErrorMessage}";
                                throw new Exception(_errmsg);
                            }

                            MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2.ReturnData);
                            if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                            {
                                var _errmsg = $"SyncWorkers 失敗:{mr.StatusCode} {mr.StatusDescription} {mr.ReturnData} {mr.ErrorMessage}";
                                throw new Exception(_errmsg);
                            }

                            byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                            string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                            OracleResponseWorkersObj obj = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleResponseWorkersObj>(desc_str);
                            Workers.AddRange(obj.Items);
                            iCurrent_page++;

                            hasMore = obj.HasMore;
                        }
                        catch (Exception exbs64)
                        {
                            log.WriteErrorLog($"Call {url} 失敗:{exbs64.Message}{exbs64.InnerException}");
                        }
                    }
                    else
                    {
                        log.WriteErrorLog($"Call Boxman api {_url} 失敗:{ret.ErrorMessage}. {ret.ErrorException}");
                    }

                }

                log.WriteLog("4", $"準備寫入 {Workers.Count} 筆 worker 資料到 SQLite");
                SaveOracleWorkersIntoSQLite(Workers, BatchNo);

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"Get Workers 失敗:{ex.Message}{ex.InnerException}");
            }

        }

        // 30分鐘 的 毫秒
        const int int30Mins = 1000 * 60 * 30;

        /// <summary>
        /// 經由 TEST BOXMAN AP => ORACLE AP => ORACLE CLOUD
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public MiddleReturn HttpGetFromOracleAP(string url)
        {
            try
            {
                // ORACLE STAGE
                // for  oracle ap
                //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                //string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/Suppliers/?limit={iRecCntPerPage}";
                MiddleModel2 send_model2 = new MiddleModel2();
                send_model2.URL = url;
                //send_model2.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(par);
                send_model2.Method = "GET";
                //string username = this.UserName;
                //string password = this.Password;
                send_model2.UserName = this.UserName;
                send_model2.Password = this.Password;
                string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                //CredentialCache myCred = new CredentialCache();
                //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                //send_model2.Cred = myCred;
                send_model2.Timeout = int30Mins;
               
                MiddleModel send_model = new MiddleModel();
                send_model.URL = $"{Oracle_AP}/api/Middle/Call/";
                send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                send_model.Method = "POST";
                send_model.Timeout = int30Mins;

                // for BOXMAN API
                MiddleModel3 send_model_83 = new MiddleModel3();
                send_model_83.URL = $"{ap83位置}/api/Middle/Call/";
                //send_model53.URL = $"{ap53位置}/api/Middle/TestMiddle/";
                send_model_83.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model);
                send_model_83.Method = "POST";
                //send_model53.Method = "GET";
                send_model_83.Timeout = int30Mins;
                MiddleReturn ret = HttpCall(send_model_83);




                //var _MethodRoute = "api/Middle/Call";
                //var _MethodType = "POST";
                ////var _MethodRoute = "api/Middle/TestMiddle";
                ////var _MethodType = "GET";


                //var ret = ApiOperation.CallApi(
                //        new ApiRequestSetting()
                //        {
                //            Data = send_model53,
                //            MethodRoute = _MethodRoute,
                //            MethodType = _MethodType,
                //            TimeOut = int30Mins
                //        }
                //        );
                if (ret.StatusCode == "200")
                {
                    string receive_str = ret.ReturnData;
                    if (string.IsNullOrWhiteSpace(receive_str))
                    {
                        throw new Exception("HttpGetFromOracleAP 取得空白的回應!");
                    }
                    string mr2rec = "";
                    try
                    {
                        MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                        mr2rec = mr2.ReturnData;
                    }
                    catch (Exception exmr2)
                    {
                        throw new Exception($"HttpGetFromOracleAP 轉成 MiddleReturn mr2 失敗:{exmr2.Message}{exmr2.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }
                    string mr3rec = "";
                    try
                    {
                        MiddleReturn mr3 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2rec);
                        return mr3;
                        //var bs64 = mr3.ReturnData;
                        //var bytes = Convert.FromBase64String(bs64);
                        //mr3rec = Encoding.UTF8.GetString(bytes);
                    }
                    catch (Exception exmr3)
                    {
                        throw new Exception($"HttpGetFromOracleAP 轉成 MiddleReturn mr3 失敗:{exmr3.Message}{exmr3.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }


                    //try
                    //{
                    //    MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr3rec);
                    //    return mr;
                    //}
                    //catch (Exception exmr)
                    //{
                    //    throw new Exception($"HttpGetFromOracleAP 轉成 MiddleReturn mr 失敗:{exmr.Message}{exmr.InnerException}{Environment.NewLine}收到的資料:{mr2rec}");
                    //}
                }
                else
                {
                    throw new Exception($"HttpGetFromOracleAP 呼叫api失敗:{ret.StatusCode} {ret.StatusDescription} {ret.ReturnData} {ret.ErrorMessage}");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        public MiddleReturn HttpCall(MiddleModel3 model)
        {
            MiddleReturn rtn = new MiddleReturn();
            string url = model.URL;
            try
            {

                DateTime time_start = DateTime.Now;

                HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
                if (url.Substring(0, 5).ToLower() == "https")
                {
                    // 為https做準備

                    System.Net.ServicePointManager.SecurityProtocol =
                        (SecurityProtocolType)12288 |
                        (SecurityProtocolType)3072 |
                        (SecurityProtocolType)768 |
                        SecurityProtocolType.Tls;
                    ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                    if (model.Cred != null)
                    {
                        req.Credentials = model.Cred;
                    }
                }

                if (model.AddHeaders != null)
                {
                    foreach (string key in model.AddHeaders.AllKeys)
                    {
                        string val = model.AddHeaders[key];
                        req.Headers.Add(key, val);
                    }
                }

                if (model.Timeout.HasValue)
                {
                    req.Timeout = model.Timeout.Value;
                }
                req.Method = model.Method;
                req.ContentType = "application/json";
                if (!string.IsNullOrWhiteSpace(model.ContnetType))
                {
                    req.ContentType = model.ContnetType;
                }

                string SendingData = model.SendingData;
                //_Logger.LogInformation("SendingData={SendingData}", SendingData);

                // 寫入BODY
                if (!string.IsNullOrWhiteSpace(SendingData))
                {
                    //byte[] contentBytes = Convert.FromBase64String(model.SendingData);

                    byte[] contentBytes = Encoding.UTF8.GetBytes(SendingData);
                    using (Stream reqStream = req.GetRequestStream())
                    {
                        reqStream.Write(contentBytes, 0, contentBytes.Length);
                    }
                }

                // 發送 request
                using (WebResponse res = req.GetResponse())
                {
                    rtn.StatusCode = $"{(int)((HttpWebResponse)res).StatusCode}";
                    rtn.StatusDescription = $"{((HttpWebResponse)res).StatusDescription}";
                    using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                    {
                        rtn.ReturnData = reader.ReadToEnd();
                        // 2021/08/16
                        // 飛騰 會用到這個 因為資料過大 寫 elk 會失敗，佔用太多資源 所以世益叫我拿掉
                        //if (_Logger != null)
                        //{
                        //    _Logger.LogInformation("HttpCall Return:{@rtn}", rtn);
                        //}
                    }
                }
                rtn.UsingSecs = DateTime.Now.Subtract(time_start).TotalSeconds;

            }
            catch (WebException wx)
            {
                string errmsg = $"呼叫{url}失敗:{wx.Message} {wx.InnerException}";
                if (wx.Response != null)
                {
                    using (WebResponse wr = wx.Response)
                    {
                        try
                        {
                            HttpWebResponse httpWebResponse = (HttpWebResponse)wx.Response;
                            rtn.StatusCode = $"{(int)httpWebResponse.StatusCode}";
                            rtn.StatusDescription = httpWebResponse.StatusDescription;
                            try
                            {
                                using (StreamReader sr = new StreamReader(httpWebResponse.GetResponseStream()))
                                {
                                    errmsg = $"{errmsg}{Environment.NewLine}{sr.ReadToEnd()}";
                                }
                            }
                            catch (Exception ex2)
                            {
                                errmsg = $"{errmsg} {ex2.Message} {ex2.InnerException}";
                            }
                        }
                        catch (Exception ex3)
                        {
                            errmsg = $"{errmsg} {ex3.Message} {ex3.InnerException}";
                        }
                    }
                }
                rtn.ErrorMessage = $"wx errmsg:{errmsg}";


            }
            catch (Exception ex)
            {
                string errmsg = $"呼叫{url}失敗:{ex.Message} {ex.InnerException}";
                rtn.ErrorMessage = $"ex errmsg:{errmsg}";
            }
            return rtn;
        }


        public MyHttpResult myPost(string url, string content)
        {
            MyHttpResult rtn = new MyHttpResult();


            try
            {
                DateTime t1 = DateTime.Now;


                ////少了這個等號,參數就傳不過去!!!!
                //json_val = $"={json_val}";
                ////byte[] contentbytes = Encoding.UTF8.GetBytes(json_val);



                byte[] contentbytes = Encoding.UTF8.GetBytes(content);


                var _url = url.ToLower();

                var prefix = _url.Substring(0, 5);
                if (prefix == "https")
                {
                   
                    System.Net.ServicePointManager.SecurityProtocol =
                        (SecurityProtocolType)12288 |
                        (SecurityProtocolType)3072 |
                        (SecurityProtocolType)768 |
                        SecurityProtocolType.Tls;

                    //走 https
                    // 強制認為憑證都是通過的，特殊情況再使用
                    ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                }



                HttpWebRequest req = HttpWebRequest.Create(url) as HttpWebRequest;
                req.Method = "POST";
                req.Timeout = 1000 * 60 * 30;

                //req.Method = "GET";
                req.ContentType = "application/json";
                //req.ContentType = "application/x-www-form-urlencoded";

                req.ContentLength = contentbytes.Length;
                using (Stream reqStream = req.GetRequestStream())
                {
                    reqStream.Write(contentbytes, 0, contentbytes.Length);
                    //reqStream.Flush();
                }
                using (WebResponse res = req.GetResponse())
                {
                    rtn.StatusCode = string.Format("{0:d}", ((HttpWebResponse)res).StatusCode);
                    rtn.StatusDescription = ((HttpWebResponse)res).StatusDescription;
                    using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                    {
                        //txtResponse.AppendText(string.Format("Response:{0}", reader.ReadToEnd()));
                        rtn.Result = reader.ReadToEnd();
                        //rtn= string.Format("StatusCode:{0}\r\nStatusDescription:{1}\r\nResponse:{2}", sstatusCode, statusDescription, reader.ReadToEnd());
                    }
                }

                DateTime t2 = DateTime.Now;
            }
            catch (WebException wx)
            {
                if (wx.Response != null)
                {
                    using (WebResponse res = wx.Response)
                    {
                        using (StreamReader reader = new StreamReader(res.GetResponseStream()))
                        {
                            //rtn.StatusCode = string.Format("{0:d}", ((HttpWebResponse)res).StatusCode);
                            //rtn.StatusDescription = ((HttpWebResponse)res).StatusDescription;
                            var errmsg = reader.ReadToEnd();
                            throw new Exception(errmsg);
                        }
                    }
                }
                else
                {
                    throw;
                }
            }
            catch (Exception ex)
            {
                throw;
            }

            return rtn;
        }



        /// <summary>
        /// 經由 TEST BOXMAN AP => ORACLE AP => ORACLE CLOUD
        /// </summary>
        /// <param name="url"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        public MiddleReturn HttpPatchFromOracleAP(string url, object content,
            Dictionary<string, string> AddHeaders = null)
        {
            try
            {
                // ORACLE STAGE
                // for  oracle ap
                MiddleModel2 send_model2 = new MiddleModel2();
                //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                //string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/Suppliers/?limit={iRecCntPerPage}";


                send_model2.URL = url;
                send_model2.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(content);
                send_model2.Method = "PATCH";
                //string username = this.UserName;
                //string password = this.Password;
                send_model2.UserName = this.UserName;
                send_model2.Password = this.Password;
                string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                //CredentialCache myCred = new CredentialCache();
                //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                //send_model2.Cred = myCred;
                if (AddHeaders != null)
                {
                    foreach (var header in AddHeaders)
                    {
                        send_model2.AddHeaders.Add(header.Key, header.Value);
                    }
                }
                send_model2.Timeout = int30Mins;

                
                MiddleModel3 send_model = new MiddleModel3();
                send_model.URL = $"{Oracle_AP}/api/Middle/Call/";
                send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                send_model.Method = "POST";
                send_model.Timeout = int30Mins;


                //這裡應該直接 call job 機
                // 呼叫 83
                MiddleModel3 send_model_83 = new MiddleModel3();
                send_model_83.URL = $"{ap83位置}/api/Middle/Call/";
                send_model_83.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model);
                send_model_83.Method = "POST";
                send_model_83.Timeout = int30Mins;
                MiddleReturn ret = HttpCall(send_model_83);


                //// for BOXMAN API
                //// boxman api 會把難字自動解成 &#12345; 例如 "勳" => &#21234;                
                //MiddleModel send_model53 = new MiddleModel();
                //send_model53.URL = $"{ap53位置}/api/Middle/Call/";
                //send_model53.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model);
                //send_model53.Method = "POST";
                //send_model53.Timeout = int30Mins;
                //var ret = ApiOperation.CallApi(
                //        new ApiRequestSetting()
                //        {
                //            Data = send_model53,
                //            MethodRoute = "api/Middle/Call",
                //            MethodType = "POST",
                //            TimeOut = int30Mins
                //        }
                //        );
                if (ret.StatusCode == "200")
                {
                    if (!string.IsNullOrWhiteSpace(ret.ErrorMessage))
                    {
                        throw new Exception(ret.ErrorMessage);
                    }

                    string receive_str = ret.ReturnData;
                    if (string.IsNullOrWhiteSpace(receive_str))
                    {
                        throw new Exception("HttpPatchFromOracleAP 取得空白的回應!");
                    }
                    string mr2rec = "";
                    try
                    {
                        MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                        mr2rec = mr2.ReturnData;
                        // Oracle Ap 回傳的錯誤
                        if (mr2.StatusCode != "200")
                        {
                            throw new Exception(mr2.ErrorMessage);
                        }
                    }
                    catch (Exception exmr2)
                    {
                        throw new Exception($"HttpPatchFromOracleAP 轉成 MiddleReturn mr2 失敗:{exmr2.Message}{exmr2.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }

                    string mr3rec = "";
                    try
                    {
                        MiddleReturn mr3 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2rec);
                        //mr3rec = mr3.ReturnData;
                        //var tmp = Encoding.UTF8.GetString(Convert.FromBase64String(mr3rec));

                        //if ((mr3.StatusCode != "200") && (mr3.StatusCode != "201") && (mr3.StatusCode != "204"))
                        //{
                        //    var errmsg = $"{mr3.StatusCode} {mr3.ReturnData} {mr3.ErrorMessage}";
                        //}
                        return mr3;
                    }
                    catch (Exception exmr3)
                    {
                        throw new Exception($"HttpGetFromOracleAP 轉成 MiddleReturn mr3 失敗:{exmr3.Message}{exmr3.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }


                    try
                    {
                        MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr3rec);
                        //return mr;
                    }
                    catch (Exception exmr)
                    {
                        throw new Exception($"HttpPatchFromOracleAP 轉成 MiddleReturn mr 失敗:{exmr.Message}{exmr.InnerException}{Environment.NewLine}收到的資料:{mr2rec}");
                    }
                }
                else
                {
                    throw new Exception($"{ret.StatusCode} {ret.StatusDescription} {ret.ReturnData} {ret.ErrorMessage}");
                }

            }
            catch (Exception ex)
            {
                throw;
            }
        }


        /// <summary>
        /// 這個 API 會直接呼叫 JOB 機上的 WEB SITE
        /// </summary>
        /// <param name="url"></param>
        /// <param name="content"></param>
        /// <param name="AddHeaders"></param>
        /// <returns></returns>
        public MiddleReturn HttpPostFromOracleAP_2(string url, object content,
            Dictionary<string, string> AddHeaders = null)
        {
            try
            {
                // ORACLE STAGE
                // for  oracle ap
                MiddleModel2 send_model2 = new MiddleModel2();
                //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                //string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/Suppliers/?limit={iRecCntPerPage}";


                send_model2.URL = url;
                send_model2.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(content);
                send_model2.Method = "POST";
                //string username = this.UserName;
                //string password = this.Password;
                send_model2.UserName = this.UserName;
                send_model2.Password = this.Password;
                string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                //CredentialCache myCred = new CredentialCache();
                //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                //send_model2.Cred = myCred;
                if (AddHeaders != null)
                {
                    foreach (var header in AddHeaders)
                    {
                        send_model2.AddHeaders.Add(header.Key, header.Value);
                    }
                }
                send_model2.Timeout = int30Mins;

               
                MiddleModel3 send_model = new MiddleModel3();
                send_model.URL = $"{Oracle_AP}/api/Middle/Call/";
                send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                send_model.Method = "POST";
                send_model.Timeout = int30Mins;


                //這裡應該直接 call job 機
                // 呼叫 83
                MiddleModel3 send_model_83 = new MiddleModel3();
                send_model_83.URL = $"{ap83位置}/api/Middle/Call/";
                send_model_83.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model);
                send_model_83.Method = "POST";
                send_model_83.Timeout = int30Mins;
                MiddleReturn ret = HttpCall(send_model_83);


                //// for BOXMAN API
                //// boxman api 會把難字自動解成 &#12345; 例如 "勳" => &#21234;                
                //MiddleModel send_model53 = new MiddleModel();
                //send_model53.URL = $"{ap53位置}/api/Middle/Call/";
                //send_model53.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model);
                //send_model53.Method = "POST";
                //send_model53.Timeout = int30Mins;
                //var ret = ApiOperation.CallApi(
                //        new ApiRequestSetting()
                //        {
                //            Data = send_model53,
                //            MethodRoute = "api/Middle/Call",
                //            MethodType = "POST",
                //            TimeOut = int30Mins
                //        }
                //        );
                if (ret.StatusCode == "200")
                {
                    if (!string.IsNullOrWhiteSpace(ret.ErrorMessage))
                    {
                        throw new Exception(ret.ErrorMessage);
                    }

                    string receive_str = ret.ReturnData;
                    if (string.IsNullOrWhiteSpace(receive_str))
                    {
                        throw new Exception("HttpPatchFromOracleAP 取得空白的回應!");
                    }
                    string mr2rec = "";
                    try
                    {
                        MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                        mr2rec = mr2.ReturnData;
                        // Oracle Ap 回傳的錯誤
                        if (mr2.StatusCode != "200")
                        {
                            throw new Exception(mr2.ErrorMessage);
                        }
                    }
                    catch (Exception exmr2)
                    {
                        throw new Exception($"HttpPatchFromOracleAP 轉成 MiddleReturn mr2 失敗:{exmr2.Message}{exmr2.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }

                    string mr3rec = "";
                    try
                    {
                        MiddleReturn mr3 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2rec);
                        mr3rec = mr3.ReturnData;
                        var tmp = Encoding.UTF8.GetString(Convert.FromBase64String(mr3rec));

                        return mr3;
                    }
                    catch (Exception exmr3)
                    {
                        throw new Exception($"HttpGetFromOracleAP 轉成 MiddleReturn mr3 失敗:{exmr3.Message}{exmr3.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }


                    try
                    {
                        MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr3rec);
                        //return mr;
                    }
                    catch (Exception exmr)
                    {
                        throw new Exception($"HttpPatchFromOracleAP 轉成 MiddleReturn mr 失敗:{exmr.Message}{exmr.InnerException}{Environment.NewLine}收到的資料:{mr2rec}");
                    }
                }
                else
                {
                    throw new Exception($"{ret.StatusCode} {ret.StatusDescription} {ret.ReturnData} {ret.ErrorMessage}");
                }

            }
            catch (Exception ex)
            {
                throw;
            }
        }


        /// <summary>
        /// 這個 API 會走 boxman
        /// </summary>
        /// <param name="url"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        public MiddleReturn HttpPostFromOracleAP(string url, object content)
        {
            try
            {
                // ORACLE STAGE
                // for  oracle ap
                MiddleModel2 send_model2 = new MiddleModel2();
                //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                //string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/Suppliers/?limit={iRecCntPerPage}";


                send_model2.URL = url;
                send_model2.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(content);
                send_model2.Method = "POST";
                //string username = this.UserName;
                //string password = this.Password;
                send_model2.UserName = this.UserName;
                send_model2.Password = this.Password;
                string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                //CredentialCache myCred = new CredentialCache();
                //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                //send_model2.Cred = myCred;
                send_model2.Timeout = int30Mins;



                
                MiddleModel send_model = new MiddleModel();
                send_model.URL = $"{Oracle_AP}/api/Middle/Call/";
                send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                send_model.Method = "POST";
                send_model.Timeout = int30Mins;


                // for BOXMAN API
                MiddleModel send_model53 = new MiddleModel();
                send_model53.URL = $"{ap83位置}/api/Middle/Call/";
                send_model53.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model);
                send_model53.Method = "POST";
                send_model53.Timeout = int30Mins;
                var ret = ApiOperation.CallApi(
                        new ApiRequestSetting()
                        {
                            Data = send_model53,
                            MethodRoute = "api/Middle/Call",
                            MethodType = "POST",
                            TimeOut = int30Mins
                        }
                        );
                if (ret.Success)
                {
                    string receive_str = ret.ResponseContent;
                    if (string.IsNullOrWhiteSpace(receive_str))
                    {
                        throw new Exception("HttpPatchFromOracleAP 取得空白的回應!");
                    }
                    string mr2rec = "";
                    try
                    {
                        MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                        mr2rec = mr2.ReturnData;
                    }
                    catch (Exception exmr2)
                    {
                        throw new Exception($"HttpPatchFromOracleAP 轉成 MiddleReturn mr2 失敗:{exmr2.Message}{exmr2.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }

                    string mr3rec = "";
                    try
                    {
                        MiddleReturn mr3 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2rec);
                        return mr3;
                        //mr3rec = mr3.ReturnData;
                    }
                    catch (Exception exmr3)
                    {
                        throw new Exception($"HttpGetFromOracleAP 轉成 MiddleReturn mr3 失敗:{exmr3.Message}{exmr3.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }


                    //try
                    //{
                    //    MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr3rec);
                    //    return mr;
                    //}
                    //catch (Exception exmr)
                    //{
                    //    throw new Exception($"HttpPatchFromOracleAP 轉成 MiddleReturn mr 失敗:{exmr.Message}{exmr.InnerException}{Environment.NewLine}收到的資料:{mr2rec}");
                    //}
                }
                else
                {
                    throw new Exception($"HttpPatchFromOracleAP 呼叫api失敗:{ret.ErrorMessage}. {ret.ErrorException}");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// 這個 fn 會加上 header 參數
        /// </summary>
        /// <param name="url"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        public MiddleReturn HttpPostFromOracleAP_AndHeaders(string url, object content,
            Dictionary<string, string> Headers)
        {
            try
            {
                // ORACLE STAGE
                // for  oracle ap
                MiddleModel2 send_model2 = new MiddleModel2();
                //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                //string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/Suppliers/?limit={iRecCntPerPage}";

                send_model2.URL = url;
                send_model2.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(content);
                send_model2.Method = "POST";
                //string username = this.UserName;
                //string password = this.Password;
                send_model2.UserName = this.UserName;
                send_model2.Password = this.Password;
                string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                foreach (var item in Headers)
                {
                    send_model2.AddHeaders.Add(item.Key, Headers[item.Key]);
                }

                //CredentialCache myCred = new CredentialCache();
                //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                //send_model2.Cred = myCred;
                send_model2.Timeout = int30Mins;



                
                MiddleModel send_model = new MiddleModel();
                send_model.URL = $"{Oracle_AP}/api/Middle/Call/";
                send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                send_model.Method = "POST";
                send_model.Timeout = int30Mins;


                // for BOXMAN API
                MiddleModel send_model53 = new MiddleModel();
                send_model53.URL = $"{ap83位置}/api/Middle/Call/";
                send_model53.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model);
                send_model53.Method = "POST";
                send_model53.Timeout = int30Mins;
                var ret = ApiOperation.CallApi(
                        new ApiRequestSetting()
                        {
                            Data = send_model53,
                            MethodRoute = "api/Middle/Call",
                            MethodType = "POST",
                            TimeOut = int30Mins
                        }
                        );
                if (ret.Success)
                {
                    string receive_str = ret.ResponseContent;
                    if (string.IsNullOrWhiteSpace(receive_str))
                    {
                        throw new Exception("HttpPatchFromOracleAP 取得空白的回應!");
                    }
                    string mr2rec = "";
                    try
                    {
                        MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                        mr2rec = mr2.ReturnData;
                        if (string.IsNullOrWhiteSpace(mr2rec))
                        {
                            throw new Exception($"{mr2.ErrorMessage}");
                        }
                    }
                    catch (Exception exmr2)
                    {
                        throw new Exception($"HttpPatchFromOracleAP 轉成 MiddleReturn mr2 失敗:{exmr2.Message}{exmr2.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }

                    string mr3rec = "";
                    try
                    {
                        MiddleReturn mr3 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2rec);
                        return mr3;
                        //mr3rec = mr3.ReturnData;
                        //if (string.IsNullOrWhiteSpace(mr3rec))
                        //{
                        //    throw new Exception($"{mr3.ErrorMessage}");
                        //}
                    }
                    catch (Exception exmr3)
                    {
                        throw new Exception($"HttpGetFromOracleAP 轉成 MiddleReturn mr3 失敗:{exmr3.Message}{exmr3.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }


                    //try
                    //{
                    //    MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr3rec);
                    //    return mr;
                    //}
                    //catch (Exception exmr)
                    //{
                    //    throw new Exception($"HttpPatchFromOracleAP 轉成 MiddleReturn mr 失敗:{exmr.Message}{exmr.InnerException}{Environment.NewLine}收到的資料:{mr2rec}");
                    //}
                }
                else
                {
                    throw new Exception($"HttpPatchFromOracleAP 呼叫api失敗:{ret.ErrorMessage}. {ret.ErrorException}");
                }
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        /// <summary>
        /// 這個 fn 會把 sending data 用 base64 先編碼
        /// </summary>
        /// <param name="url"></param>
        /// <param name="content"></param>
        /// <param name="Headers"></param>
        /// <returns></returns>
        public MiddleReturn HttpPostFromOracleAP_AndHeaders3(string url, object content,
            Dictionary<string, string> Headers)
        {
            try
            {
                // ORACLE STAGE
                // for  oracle ap
                MiddleModel2 send_model2 = new MiddleModel2();
                //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                //string url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/Suppliers/?limit={iRecCntPerPage}";

                send_model2.URL = url;
                send_model2.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(content);
                send_model2.Method = "POST";
                //string username = this.UserName;
                //string password = this.Password;
                send_model2.UserName = this.UserName;
                send_model2.Password = this.Password;
                string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                foreach (var item in Headers)
                {
                    send_model2.AddHeaders.Add(item.Key, Headers[item.Key]);
                }

                //CredentialCache myCred = new CredentialCache();
                //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                //send_model2.Cred = myCred;
                send_model2.Timeout = int30Mins;



              
                MiddleModel send_model = new MiddleModel();
                send_model.URL = $"{Oracle_AP}/api/Middle/Call/";
                send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                send_model.Method = "POST";
                send_model.Timeout = int30Mins;



                // for job website
                MiddleModel3 send_model_83 = new MiddleModel3();
                send_model_83.URL = $"{ap83位置}/api/Middle/Call/";
                //send_model53.URL = $"{ap53位置}/api/Middle/TestMiddle/";
                send_model_83.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model);
                send_model_83.Method = "POST";
                send_model_83.Timeout = int30Mins;

                MiddleReturn ret = HttpCall(send_model_83);

                if (ret.StatusCode == "200")
                {
                    string receive_str = ret.ReturnData;
                    if (string.IsNullOrWhiteSpace(receive_str))
                    {
                        throw new Exception("HttpPostFromOracleAP 取得空白的回應!");
                    }
                    string mr2rec = "";
                    try
                    {
                        MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                        mr2rec = mr2.ReturnData;
                    }
                    catch (Exception exmr2)
                    {
                        throw new Exception($"HttpPostFromOracleAP 轉成 MiddleReturn mr2 失敗:{exmr2.Message}{exmr2.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }
                    string mr3rec = "";
                    try
                    {
                        MiddleReturn mr3 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2rec);
                        //var bs64 = mr3.ReturnData;
                        //var bytes = Convert.FromBase64String(bs64);
                        //mr3rec = Encoding.UTF8.GetString(bytes);
                        return mr3;
                    }
                    catch (Exception exmr3)
                    {
                        throw new Exception($"HttpGetFromOracleAP 轉成 MiddleReturn mr3 失敗:{exmr3.Message}{exmr3.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                    }
                }
                else
                {
                    throw new Exception($"HttpGetFromOracleAP 呼叫api失敗:{ret.StatusCode} {ret.StatusDescription} {ret.ReturnData} {ret.ErrorMessage}");
                }


                //// for BOXMAN API
                //MiddleModel send_model53 = new MiddleModel();
                //send_model53.URL = $"{ap83位置}/api/Middle/Call/";
                //var tmp = Newtonsoft.Json.JsonConvert.SerializeObject(send_model);
                //var bytes = Encoding.UTF8.GetBytes(tmp);
                //send_model53.SendingData = Convert.ToBase64String(bytes);
                ////send_model53.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model);
                //send_model53.Method = "POST";
                //send_model53.Timeout = int30Mins;
                //var ret = ApiOperation.CallApi(
                //        new ApiRequestSetting()
                //        {
                //            Data = send_model53,
                //            MethodRoute = "api/Middle/Call3",
                //            MethodType = "POST",
                //            TimeOut = int30Mins
                //        }
                //        );
                //if (ret.Success)
                //{
                //    string receive_str = ret.ResponseContent;
                //    if (string.IsNullOrWhiteSpace(receive_str))
                //    {
                //        throw new Exception("HttpPatchFromOracleAP 取得空白的回應!");
                //    }
                //    string mr2rec = "";
                //    try
                //    {
                //        MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                //        mr2rec = mr2.ReturnData;
                //        if (string.IsNullOrWhiteSpace(mr2rec))
                //        {
                //            throw new Exception($"{mr2.ErrorMessage}");
                //        }
                //    }
                //    catch (Exception exmr2)
                //    {
                //        throw new Exception($"HttpPatchFromOracleAP 轉成 MiddleReturn mr2 失敗:{exmr2.Message}{exmr2.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                //    }

                //    string mr3rec = "";
                //    try
                //    {
                //        MiddleReturn mr3 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2rec);
                //        return mr3;
                //        //mr3rec = mr3.ReturnData;
                //        //if (string.IsNullOrWhiteSpace(mr3rec))
                //        //{
                //        //    throw new Exception($"{mr3.ErrorMessage}");
                //        //}
                //    }
                //    catch (Exception exmr3)
                //    {
                //        throw new Exception($"HttpGetFromOracleAP 轉成 MiddleReturn mr3 失敗:{exmr3.Message}{exmr3.InnerException}{Environment.NewLine}收到的資料={receive_str}");
                //    }


                //    //try
                //    //{
                //    //    MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr3rec);
                //    //    return mr;
                //    //}
                //    //catch (Exception exmr)
                //    //{
                //    //    throw new Exception($"HttpPatchFromOracleAP 轉成 MiddleReturn mr 失敗:{exmr.Message}{exmr.InnerException}{Environment.NewLine}收到的資料:{mr2rec}");
                //    //}
                //}
                //else
                //{
                //    throw new Exception($"HttpPatchFromOracleAP 呼叫api失敗:{ret.ErrorMessage}. {ret.ErrorException}");
                //}
            }
            catch (Exception ex)
            {
                throw;
            }
        }


        void CreateSQLiteFileIfNotExists_Supplier_1()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_Supplier_1";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "OracleSupplier";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users

                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"
CREATE TABLE [{table_name}] (
    [BatchNo] nvarchar(30) ,
	[SupplierId] nvarchar(20) NOT NULL, 
	[SupplierPartyId] nvarchar(20), 
	[Supplier] nvarchar(100),
	[SupplierNumber] nvarchar(100),
	[AlternateName] nvarchar(100),
	[SupplierTypeCode] nvarchar(100),
	[SupplierType] nvarchar(100),
	[Status] nvarchar(100)
)
";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (SupplierNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        void CreateSQLiteFileIfNotExists_Supplier_2_address()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_Supplier_2_address";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "OracleSupplierAddr";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users

                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"
CREATE TABLE [{table_name}] (
    [BatchNo] nvarchar(30) ,
	[SupplierNumber] nvarchar(100),
	[SupplierAddressId] nvarchar(100),
	[AddressName] nvarchar(100),
	[CountryCode] nvarchar(100),
	[Country] nvarchar(100),
	[AddressLine1] nvarchar(100)
)
";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (SupplierNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        void CreateSQLiteFileIfNotExists_Supplier_3_site()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_Supplier_3_site";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "OracleSupplierSite";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users

                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"
CREATE TABLE [{table_name}] (
    [BatchNo] nvarchar(30) ,
	[SupplierNumber] nvarchar(100),

	[SupplierSiteId] nvarchar(100),
	[SupplierSite] nvarchar(100),
	[SupplierAddressId] nvarchar(100),
	[SupplierAddressName] nvarchar(100)

)
";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (SupplierNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        void CreateSQLiteFileIfNotExists_Supplier_4_Assignment()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_Supplier_4_Assignment";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "OracleSupplierAssignment";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users

                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"
CREATE TABLE [{table_name}] (
    [BatchNo] nvarchar(30) ,
	[SupplierNumber] nvarchar(100),

	[AssignmentId] nvarchar(100),
	[Status] nvarchar(20) 

)
";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (SupplierNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        void CreateSQLiteFileIfNotExists_Supplier_5_Payee()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_Supplier_5_Payee";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "OracleSupplierPayee";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users

                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"
CREATE TABLE [{table_name}] (
    [BatchNo] nvarchar(30) ,
	[SupplierNumber] nvarchar(100),

	[PayeeId] nvarchar(100),
	[PayeePartyIdentifier] nvarchar(100),
	[PartyName] nvarchar(100),
	[PayeePartyNumber] nvarchar(100),
	[PayeePartySiteIdentifier] nvarchar(100),
	[SupplierSiteCode] nvarchar(100),
	[SupplierSiteIdentifier] nvarchar(100),
	[PayeePartySiteNumber] nvarchar(100)

)
";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (SupplierNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        void CreateSQLiteFileIfNotExists_Supplier_5_BankAccount()
        {

            //5.2   Get Supplier Bank Account(
            //  第一個用5.1步驟獲得的 PayeePartyIdentifier，
            //  第二個用5.1步驟獲得的 PayeePartySiteIdentifier)


            string fn_name = "CreateSQLiteFileIfNotExists_Supplier_5_BankAccount";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "OracleSupplierBankAccount";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users

                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"
CREATE TABLE [{table_name}] (
    [BatchNo] nvarchar(30) ,
	[SupplierNumber] nvarchar(100),

	[BankAccountId] nvarchar(100),
	[AccountNumber] nvarchar(100),
	[AccountName] nvarchar(100),
	[BankName] nvarchar(100),
	[BranchName] nvarchar(100)

)
";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (SupplierNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }


        }


        void CreateSQLiteFileIfNotExists_HR_Oracle_Mapping()
        {

            // 新舊員編對應表
            string fn_name = "CreateSQLiteFileIfNotExists_HR_Oracle_Mapping";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "HR_Oracle_Mapping";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users

                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"
CREATE TABLE [{table_name}] (
	[OEmpNo] nvarchar(20), 
	[NEmpNo] nvarchar(20)
)
";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    string file_path = $"{AppDomain.CurrentDomain.BaseDirectory}/Jobs/ORSyncOracleData/rcode_mapping.csv";
                    string[] mapping_table = System.IO.File.ReadAllLines(file_path);
                    for (int idx = 1; idx < mapping_table.Length; idx++)
                    {
                        var vals = mapping_table[idx];
                        string[] onempno = vals.Split(',');
                        var oldEmpno = onempno[0];
                        var newEmpNo = onempno[1];
                        sql = $@"
INSERT INTO [HR_Oracle_Mapping]
		([OEmpNo]
		,[NEmpNo])
	VALUES
(:0, :1);
";
                        try
                        {
                            int cnt1 = sqlite.ExecuteByCmd(sql, oldEmpno, newEmpNo);
                        }
                        catch (Exception exins)
                        {
                        }
                    }

                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }


        }




        public MiddleReturn GetOracleSupplierByEmpNo(string EmpNo)
        {
            string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
            log.WriteLog("5", $"{EmpNo} supplier url:{url}");
            return HttpGetFromOracleAP(url);
        }

        // 2-Get Supplier Address(第一步得到的 Supplier ID)
        public MiddleReturn GetOracleSupplierAddrByEmpNo(string EmpNo, string BatchNo, string SupplierId = "")
        {
            string _SupplierId = SupplierId;
            if (string.IsNullOrWhiteSpace(_SupplierId))
            {

                string sql = $@"
select SupplierId
from  [OracleSupplier]
where batchno = :0
and SupplierNumber = :1
";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                DataTable tb;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo, EmpNo))
                {
                    _SupplierId = $"{row["SupplierId"]}";
                    break;
                }
            }

            if (!string.IsNullOrWhiteSpace(_SupplierId))
            {
                string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers/{_SupplierId}/child/addresses";
                //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                log.WriteLog("5", $"{EmpNo} addresses url:{url}");
                return HttpGetFromOracleAP(url);
            }
            else
            {
                throw new Exception($"EmpNo:{EmpNo} 找不到 SupplierId");
            }

        }

        // 3-Get Supplier Site(第一步得到的 Supplier ID)
        public MiddleReturn GetOracleSupplierSiteByEmpNo(string EmpNo, string BatchNo, string SupplierId = "")
        {
            string _SupplierId = SupplierId;
            if (string.IsNullOrWhiteSpace(_SupplierId))
            {

                string sql = $@"
select SupplierId
from  [OracleSupplier]
where batchno = :0
and SupplierNumber = :1
";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                DataTable tb;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo, EmpNo))
                {
                    _SupplierId = $"{row["SupplierId"]}";
                    break;
                }
            }

            if (!string.IsNullOrWhiteSpace(_SupplierId))
            {
                string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers/{_SupplierId}/child/sites";
                //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                log.WriteLog("5", $"{EmpNo} sites url:{url}");
                return HttpGetFromOracleAP(url);
            }
            else
            {
                throw new Exception($"EmpNo:{EmpNo} 找不到 SupplierId");
            }

        }

        //4-Get Supplier Assignment(第一個用第一步獲得的 Supplier ID，第二個用第三步獲得的 SupplierSiteId)
        public MiddleReturn GetOracleSupplierAssignmentByEmpNo(string EmpNo, string BatchNo
            , string SupplierId = "", string SupplierSiteId = "")
        {
            string _SupplierId = SupplierId;
            string _SupplierSiteId = SupplierSiteId;

            string sql = "";
            //SQLiteUtl sqlite = new SQLiteUtl(db_file);
            DataTable tb;
            if (string.IsNullOrWhiteSpace(_SupplierId))
            {

                sql = $@"
select SupplierId
from  [OracleSupplier]
where batchno = :0
and SupplierNumber = :1
";
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo, EmpNo))
                {
                    _SupplierId = $"{row["SupplierId"]}";
                    break;
                }
            }

            if (string.IsNullOrWhiteSpace(_SupplierSiteId))
            {

                sql = $@"
select SupplierSiteId
from  [OracleSupplierSite]
where batchno =:0
and SupplierNumber = :1
";

                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo, EmpNo))
                {
                    _SupplierSiteId = $"{row["SupplierSiteId"]}";
                    break;
                }
            }

            if (!string.IsNullOrWhiteSpace(_SupplierId))
            {
                string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers/{_SupplierId}/child/sites/{_SupplierSiteId}/child/assignments";
                //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                log.WriteLog("5", $"{EmpNo} assignments url:{url}");
                return HttpGetFromOracleAP(url);
            }
            else
            {
                throw new Exception($"EmpNo:{EmpNo} 找不到 SupplierId");
            }

        }

        // 5-1.Get Supplier Bank Account-5.1   Get Payee Party ID(第一步獲得的SupplierPartyID)
        public MiddleReturn GetOracleSupplierPayeeByEmpNo(string EmpNo, string BatchNo, string SupplierPartyId = "")
        {
            //5-1.Get Supplier Bank Account-5.1   Get Payee Party ID(第一步獲得的 SupplierPartyID)
            string _SupplierPartyId = SupplierPartyId;
            if (string.IsNullOrWhiteSpace(_SupplierPartyId))
            {

                string sql = $@"
select SupplierPartyId
from  [OracleSupplier]
where batchno = :0
and SupplierNumber = :1
";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                DataTable tb;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo, EmpNo))
                {
                    _SupplierPartyId = $"{row["SupplierPartyId"]}";
                    break;
                }
            }



            if (!string.IsNullOrWhiteSpace(_SupplierPartyId))
            {
                string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/paymentsExternalPayees?finder=ExternalPayeeSearch;PayeePartyIdentifier={_SupplierPartyId},Intent=Supplier";
                //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                return HttpGetFromOracleAP(url);
            }
            else
            {
                throw new Exception($"EmpNo:{EmpNo} 找不到 SupplierId");
            }

        }

        // 5.2   Get Supplier Bank Account
        //  (第一個用5.1步驟獲得的 PayeePartyIdentifier，
        //    第二個用5.1步驟獲得的 PayeePartySiteIdentifier)
        public MiddleReturn GetOracleSupplierBankAccountByEmpNo(string EmpNo, string BatchNo,
            string PayeePartyIdentifier = "", string PayeePartySiteIdentifier = "")
        {

            string _PayeePartyIdentifier = PayeePartyIdentifier;
            string _PayeePartySiteIdentifier = PayeePartySiteIdentifier;

            if (string.IsNullOrWhiteSpace(_PayeePartyIdentifier) || string.IsNullOrWhiteSpace(_PayeePartySiteIdentifier))
            {

                string sql = $@"
select PayeePartyIdentifier, PayeePartySiteIdentifier
from  [OracleSupplierPayee]
where batchno = :0
and SupplierNumber = :1
";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                DataTable tb;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo, EmpNo))
                {
                    _PayeePartyIdentifier = $"{row["PayeePartyIdentifier"]}";
                    _PayeePartySiteIdentifier = $"{row["PayeePartySiteIdentifier"]}";
                    break;
                }

            }



            if (!string.IsNullOrWhiteSpace(_PayeePartyIdentifier))
            {
                if (!string.IsNullOrWhiteSpace(_PayeePartySiteIdentifier))
                {

                    string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/payeeBankAccountsLOV?finder=AvailablePayeeBankAccountsFinder;PaymentFunction=\"PAYABLES_DISB\",PayeePartyId={_PayeePartyIdentifier},SupplierSiteId={_PayeePartySiteIdentifier}";
                    //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                    log.WriteLog("5", $"{EmpNo} 5.2   Get Supplier Bank Account url:{url}");
                    return HttpGetFromOracleAP(url);
                }
                else
                {
                    throw new Exception($"EmpNo:{EmpNo} PayeePartySiteIdentifier is null");
                }
            }
            else
            {
                throw new Exception($"EmpNo:{EmpNo} PayeePartyIdentifier is null");
            }

        }


        string GetOldEmpNo(string EmpNo)
        {
            string empNo = EmpNo.ToUpper();
            string rst = EmpNo;
            try
            {
                if (empNo.Substring(0, 2) != "PC")
                {

                    db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                    if (!Directory.Exists(db_path))
                    {
                        Directory.CreateDirectory(db_path);
                    }
                    db_file = $"{db_path}OracleData.sqlite";
                    //SQLiteUtl sqlite = new SQLiteUtl(db_file);


                    DateTime time_start = DateTime.Now;
                    //string BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                    try
                    {
                        string sql = "";
                        sql = $@"
SELECT [OEmpNo]
		,[NEmpNo]
	FROM [HR_Oracle_Mapping]
where [NEmpNo] = :0
";

                        DataTable tb = null;
                        foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, empNo))
                        {
                            rst = $"{row["OEmpNo"]}";
                        }

                    }
                    catch (Exception exhr)
                    {
                        log.WriteErrorLog($"GetOldEmpNo 失敗:{exhr.Message}{exhr.InnerException}");
                    }


                }
            }
            catch (Exception ex)
            {
                rst = empNo;
            }

            return rst;
        }

        /// <summary>
        /// 取得 oracle 的銀行資料
        /// </summary>
        /// <returns></returns>
        public IEnumerable<OracleApiChkBankBranchesMd> GetOracleBankData()
        {
            bool rst = false;
            try
            {
                // for BOXMAN API
                var par = new
                {
                    dept = "ec"
                };
                var ret = ApiOperation.CallApi(
                        new ApiRequestSetting()
                        {
                            Data = par,
                            MethodRoute = "api/OREMP/GetOracleBankData",
                            MethodType = "POST",
                            TimeOut = int30Mins
                        }
                        );
                if (ret.Success)
                {
                    string receive_str = ret.ResponseContent;
                    if (string.IsNullOrWhiteSpace(receive_str))
                    {
                        throw new Exception("HttpGetFromOracleAP 取得空白的回應!");
                    }

                    ResultObject<IEnumerable<OracleApiChkBankBranchesMd>> mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<ResultObject<IEnumerable<OracleApiChkBankBranchesMd>>>(receive_str);
                    return mr2.Data;



                }
                else
                {
                    throw new Exception($"HttpGetFromOracleAP 呼叫api失敗:{ret.ErrorMessage}. {ret.ErrorException}");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"HttpGetFromOracleAP 呼叫api失敗:{ex.Message}. {ex.InnerException}");
            }
            return null;
        }



        /// <summary>
        /// 同步oracle的銀行名稱/分行名稱資料回來，供修改銀行帳號資料時使用
        /// </summary>
        /// <returns></returns>
        public bool SyncOracleBankData()
        {
            bool rst = false;
            try
            {
                // for BOXMAN API
                var par = new
                {
                    dept = "ec"
                };
                var ret = ApiOperation.CallApi(
                        new ApiRequestSetting()
                        {
                            Data = par,
                            MethodRoute = "api/OREMP/GetOracleBankData",
                            MethodType = "POST",
                            TimeOut = int30Mins
                        }
                        );
                if (ret.Success)
                {
                    string receive_str = ret.ResponseContent;
                    if (string.IsNullOrWhiteSpace(receive_str))
                    {
                        throw new Exception("HttpGetFromOracleAP 取得空白的回應!");
                    }
                    string mr2rec = "";
                    try
                    {
                        ResultObject<IEnumerable<OracleApiChkBankBranchesMd>> mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<ResultObject<IEnumerable<OracleApiChkBankBranchesMd>>>(receive_str);
                        IEnumerable<OracleApiChkBankBranchesMd> bankdata = mr2.Data;
                        CreateSQLiteFileIfNotExists_OracleBankDataBranchData();
                        if (bankdata == null)
                            throw new Exception("取得 Oracle Bank / Branch Data 失敗!");

                        foreach (var bank in bankdata)
                        {
                            var sql = "";
                            try
                            {

                                sql = $@"DELETE FROM OracleBankData
WHERE BANK_NAME = :0
AND BANK_NUMBER = :1
AND BANK_BRANCH_NAME  = :2
AND BRANCH_NUMBER  = :3
";
                                int cnt = sqlite.ExecuteByCmd(sql,
                                    bank.BANK_NAME,
                                    bank.BANK_NUMBER,
                                    bank.BANK_BRANCH_NAME,
                                    bank.BRANCH_NUMBER
                                          );

                                // insert
                                sql = $@"
INSERT INTO [OracleBankData]
		(
   BANK_NAME ,
BANK_NUMBER ,
BANK_BRANCH_NAME  ,
BRANCH_NUMBER 
)
	VALUES(
:0
,:1
,:2
,:3
)";
                                cnt = sqlite.ExecuteByCmd(sql,
                                  bank.BANK_NAME,
                                  bank.BANK_NUMBER,
                                  bank.BANK_BRANCH_NAME,
                                  bank.BRANCH_NUMBER
                                        );

                            }
                            catch (Exception exx)
                            {
                                var _msg = $@"BANK_NAME={bank.BANK_NAME}, BANK_NUMBER={bank.BANK_NUMBER}, BANK_BRANCH_NAME={bank.BANK_BRANCH_NAME}, BRANCH_NUMBER={bank.BRANCH_NUMBER}";
                                throw new Exception($"寫入 oracle bank data,{_msg} 失敗:{exx.Message}{exx.InnerException}");
                            }
                        }
                    }
                    catch (Exception exmr2)
                    {
                        throw new Exception($"同步 oracle bank data 失敗:{exmr2.Message}{exmr2.InnerException}");
                    }


                }
                else
                {
                    throw new Exception($"HttpGetFromOracleAP 呼叫api失敗:{ret.ErrorMessage}. {ret.ErrorException}");
                }
            }
            catch (Exception ex)
            {
            }
            return rst;
        }



        //public IEnumerable<OracleApiChkBankBranchesMd> GetBankDataFromOracle()
        //{
        //    var par = new
        //    {
        //        dept = "ec"
        //    };

        //    var ret = ApiOperation.CallApi(
        //            new ApiRequestSetting()
        //            {
        //                Data = par,
        //                MethodRoute = "api/OREMP/GetOracleBankData",
        //                MethodType = "POST",
        //                TimeOut = int30Mins
        //            }
        //            );
        //    if (ret.Success)
        //    {
        //        string receive_str = ret.ResponseContent;
        //        if (string.IsNullOrWhiteSpace(receive_str))
        //        {
        //            throw new Exception("HttpGetFromOracleAP 取得空白的回應!");
        //        }
        //        string mr2rec = "";
        //        try
        //        {
        //            ResultObject<IEnumerable<OracleApiChkBankBranchesMd>> mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<ResultObject<IEnumerable<OracleApiChkBankBranchesMd>>>(receive_str);
        //            return mr2.Data;
        //        }
        //        catch (Exception exmr2)
        //        {
        //            throw new Exception($"HttpGetFromOracleAP 轉成 MiddleReturn mr2 失敗:{exmr2.Message}{exmr2.InnerException}{Environment.NewLine}收到的資料={receive_str}");
        //        }


        //    }
        //    else
        //    {
        //        throw new Exception($"HttpGetFromOracleAP 呼叫api失敗:{ret.ErrorMessage}. {ret.ErrorException}");
        //    }
        //}

        /// <summary>
        /// 更新員工供應商銀行資料
        /// </summary>
        /// <param name="EmpNo"></param>
        /// <param name="EmpName"></param>
        /// <param name="TW_ID"></param>
        /// <param name="BankNumber"></param>
        /// <param name="BankName"></param>
        /// <param name="ACCOUNT_NUMBER"></param>
        /// <param name="ACCOUNT_NAME"></param>
        /// <param name="BranchNumber"></param>
        /// <param name="BranchName"></param>
        /// <param name="OldBankName"></param>
        /// <param name="OldBranchName"></param>
        /// <param name="OldAccountNumber"></param>
        public void UpdateSupplierBankDataByEmpNo(
                 string EmpNo, // 員編
                 string EmpName, // 中文姓名
                 string TW_ID, // 帳戶身分證號
                 string BankNumber, // 銀行代碼
                 string BankName, // 銀行名稱
                 string ACCOUNT_NUMBER, // 銀行收款帳號
                 string ACCOUNT_NAME, // 銀行帳戶戶名
                 string BranchNumber, // 銀行分行代碼
                 string BranchName, // 銀行分行名稱
                 string OldBankName,  //舊的銀行名稱
                 string OldBranchName,  //舊的分行名稱
                 string OldAccountNumber //舊的銀行帳號
              )
        {
            //準備update  
            OREmpBankData par1 = new OREmpBankData()
            {
                EmpNo = EmpNo, // 員編
                EmpName = EmpName, // 中文姓名
                TW_ID = TW_ID, // 帳戶身分證號
                BankNumber = BankNumber, // 銀行代碼
                ACCOUNT_NUMBER = ACCOUNT_NUMBER, // 銀行收款帳號
                ACCOUNT_NAME = ACCOUNT_NAME, // 銀行帳戶戶名
                BranchNumber = BranchNumber, // 銀行分行代碼
                ApplyMan = "Job",
                ReDoStep = "",
                BankName = BankName,
                BranchName = BranchName,
                OldBankName = OldBankName,
                OldBranchName = OldBranchName,
                OldAccountNumber = OldAccountNumber
            };

            //var send_model = "";
            var ret1 = ApiOperation.CallApi(
                    new ApiRequestSetting()
                    {
                        Data = par1,
                        MethodRoute = "api/OREMP/UpdateBanksData",
                        MethodType = "POST",
                        TimeOut = 1000 * 60 * 120
                    }
                    );
            if (ret1.Success)
            {
                string receive_str = ret1.ResponseContent;
                try
                {
                    //GenFileResult mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<GenFileResult>(receive_str);

                    //log.WriteErrorLog(string.Join(Environment.NewLine, mr2.ErrorMessages.ToArray()));
                    //foreach (var file in mr2.Data)
                    //{
                    //    byte[] content = Convert.FromBase64String(file.FileBase64Str);
                    //    var path = $"I:/Alvin/Oracle供應商Supplier/產生supplier匯入檔給顧問/{folder}";
                    //    var zipFolder = $"{path}/zip/";
                    //    if (!Directory.Exists(zipFolder))
                    //    {
                    //        Directory.CreateDirectory(zipFolder);
                    //    }
                    //    //var sort_idx = index.ToString("D4");
                    //    var zipFilename = $"{zipFolder}{file.MailFileName}";
                    //    var csvFolder = $"{path}/csv/";
                    //    if (!Directory.Exists(csvFolder))
                    //    {
                    //        Directory.CreateDirectory(csvFolder);
                    //    }
                    //    var tmp = $"{Path.GetFileNameWithoutExtension(file.MailFileName)}.csv";
                    //    var csvFilename = $"{csvFolder}{index}-{tmp}";
                    //    System.IO.File.WriteAllBytes(zipFilename, content);

                    
                    //    //FileInfo fileToDecompress = new FileInfo(zipFilename);
                    //    //using (FileStream originalFileStream = fileToDecompress.OpenRead())
                    //    //{
                    //    //    string currentFileName = fileToDecompress.FullName;
                    //    //    string newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length);
                    //    //    newFileName = csvFilename;

                    //    //    using (FileStream decompressedFileStream = File.Create(newFileName))
                    //    //    {
                    //    //        using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                    //    //        {
                    //    //            decompressionStream.CopyTo(decompressedFileStream);
                    //    //            //Console.WriteLine($"Decompressed: {fileToDecompress.Name}");
                    //    //        }
                    //    //    }
                    //    //}
                    //}

                }
                catch (Exception exb)
                {
                    log.WriteErrorLog($"AddWorkerByMiddle Create Worker 失敗:{exb.Message}{exb.InnerException}");
                }
            }
            else
            {
                log.WriteErrorLog($"呼叫 api/OREMP/EMPACCFILES 失敗:{ret1.ResponseContent}{ret1.ErrorMessage}. {ret1.ErrorException}");
            }

        }



        /// <summary>
        /// 建立員工供應商
        /// </summary>
        /// <param name="EmpNo"></param>
        /// <param name="EmpName"></param>
        /// <param name="TW_ID"></param>
        /// <param name="BankNumber"></param>
        /// <param name="ACCOUNT_NUMBER"></param>
        /// <param name="ACCOUNT_NAME"></param>
        /// <param name="BranchNumber"></param>
        /// <returns></returns>
        public ApiResultModel CreateSupplier(
          string EmpNo, // 員編
          string EmpName, // 中文姓名
          string TW_ID, // 帳戶身分證號
          string BankNumber, // 銀行代碼
          string ACCOUNT_NUMBER, // 銀行收款帳號
          string ACCOUNT_NAME, // 銀行帳戶戶名
          string BranchNumber, // 銀行分行代碼
            string DoStep // 要打哪些
            )
        {
            ApiResultModel rst = new ApiResultModel();
            List<OREmpNoData> suppliers = new List<OREmpNoData>();
            suppliers.Add(new OREmpNoData()
            {
                EmpNo = EmpNo, // 員編
                EmpName = $"{EmpName}", // 中文姓名
                TW_ID = TW_ID, // 帳戶身分證號
                BankNumber = BankNumber, // 銀行代碼
                ACCOUNT_NUMBER = ACCOUNT_NUMBER, // 銀行收款帳號
                ACCOUNT_NAME = ACCOUNT_NAME, // 銀行帳戶戶名
                BranchNumber = BranchNumber, // 銀行分行代碼
                ApplyMan = "Job",
                ReDoStep = DoStep
            });
            var par = new
            {
                Data = suppliers
            };
            var ppar = Newtonsoft.Json.JsonConvert.SerializeObject(par);
            //log.WriteLog("4", $"準備推送員工供應商 {EmpNo} {EmpName}, 先暫停 3 秒鐘{Environment.NewLine}Payload={ppar}");
            Thread.Sleep(1000 * 3);
            var ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
            {
                MethodRoute = "api/OREMP/EMPACC/",
                Data = par,
                MethodType = "POST",
                TimeOut = int30Mins  // 5分鐘竟然不夠
            });
            if (ret.Success)
            {
                ApiResultModel model = Newtonsoft.Json.JsonConvert.DeserializeObject<ApiResultModel>(ret.Data);
                if (model.Result)
                {
                    //log.WriteLog("5", $"{EmpNo},{EmpName} Supplier push 成功! {model.ErrorCode} {model.ErrorMessages}");
                    rst.Result = model.Result;
                }
                else
                {
                    rst.ErrorCode = model.ErrorCode;
                    rst.ErrorMessages = model.ErrorMessages;
                    //log.WriteErrorLog($"{EmpNo},{EmpName}  Supplier push 失敗! {ret.Data} {ret.ErrorMessage} {ret.ErrorException}");
                }
                /*
{"Result":false,"ErrorCode":null,"ErrorMessages":["Supplier:ERROR","SupplierAddress:ERROR","SupplierSite:ERROR","SupplierSiteAssignment:ERROR","SupplierBank:ERROR"]}

                 * 
                 * 
                 * */

            }
            else
            {
                rst.ErrorMessages = new string[] { ret.ErrorMessage };
                //log.WriteLog("5", $"{EmpNo},{EmpName}  Supplier push 失敗! {ret.Data} {ret.ErrorMessage} {ret.ErrorException}");
            }

            return rst;
        }

        public bool IsOracleAccExists(string EmpNo)
        {
            bool rst = false;
            var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/emps?q=LastName=\"{EmpNo}\"";
            try
            {


                //  //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                OracleEmployees emps = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployees>(desc_str);
                if (emps.Count > 0)
                {
                    //employee = emps.Items[0];
                    rst = true;
                }

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"呼叫 api 失敗:api:{url}\r\n{ex.Message}{ex.InnerException}");
            }
            return rst;
        }


        //public bool UpdateOracleUserNameByUrl(string url, string UserName)
        //{
        //    bool rst = false;

        //    return rst;
        //}

        /// <summary>
        /// 用 oracle employee api 傳入員編取得 employee
        /// LastName=員編
        /// </summary>
        /// <param name="EmpNo"></param>
        /// <returns></returns>
        public OracleEmployee GetOracleEmployeeByEmpNo(string EmpNo)
        {
            OracleEmployee employee = null;

            log.WriteLog("5", $"準備取得 EmpNo={EmpNo} 的 oracle employee 物件. (Into GetOracleEmployeeByEmpNo)");

            var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/emps?q=LastName=\"{EmpNo}\"";
            try
            {


                //  //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (string.IsNullOrWhiteSpace(mr.ReturnData))
                {
                    throw new Exception($"{mr.StatusCode} {mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                OracleEmployees emps = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEmployees>(desc_str);

                if (emps != null && emps.Count > 0)
                {
                    employee = emps.Items[0];
                }

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"GetOracleEmployeeByEmpNo 呼叫 api 失敗:api:{url}\r\n{ex.Message}{ex.InnerException}");
                throw;
            }

            return employee;
        }

        /// <summary>
        /// 用員編取得 Assignment Number 給 create worker 用
        /// </summary>
        /// <param name="EmpNo"></param>
        /// <returns></returns>
        public string GetAssignmentNumberByEmpNo(string EmpNo)
        {
            OracleEmployee emp = GetOracleEmployeeByEmpNo(EmpNo);
            var assignment_number = GetAssignmentNumberByPersonNumber(emp.PersonNumber);
            return assignment_number;
        }

        /// <summary>
        /// 用 person number取得 assignment number
        /// </summary>
        /// <param name="PersonNumber"></param>
        /// <returns></returns>
        public string GetAssignmentNumberByPersonNumber(string PersonNumber)
        {
            string rst = "";
            var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/workers?q=PersonNumber={PersonNumber}&expand=workRelationships.assignments";
            try
            {


                //  //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                OracleReturnWorkers workers = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleReturnWorkers>(desc_str);
                foreach (var worker in workers.Items)
                {
                    foreach (var relationship in worker.WorkRelationships)
                    {
                        foreach (var assignment in relationship.Assignments)
                        {
                            rst = assignment.AssignmentNumber;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"呼叫 api 失敗:api:{url}\r\n{ex.Message}{ex.InnerException}");
                throw;
            }
            return rst;
        }

        /// <summary>
        /// 重新同步 Oracle 上的 Job Level List 回來
        /// </summary>
        /// <returns></returns>
        public List<OracleJobID> RefreshOracleJobLevelList()
        {
            if (JobLevelList == null)
            {
                JobLevelList = new List<OracleJobID>();
            }
            JobLevelList.Clear();

            var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/jobs";
            try
            {


                //  //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (string.IsNullOrWhiteSpace(mr.ReturnData))
                {
                    throw new Exception($"取 Job ID 失敗:{mr.StatusCode}{mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                OracleJobIDReturnModel ids = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleJobIDReturnModel>(desc_str);

                foreach (var item in ids.Items)
                {
                    JobLevelList.Add(item);
                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"GetJobLevelItemByNumber Get Job Level Item 失敗:api:{url}\r\n{ex.Message}{ex.InnerException}");
                throw new Exception($"GetJobLevelItemByNumber Get Job Level Item 失敗:{ex.Message}{ex.InnerException}");
            }


            return JobLevelList;
        }


        private List<OracleJobID> JobLevelList;
        public List<OracleJobID> GetJobLevelList()
        {
            if (JobLevelList == null)
            {
                RefreshOracleJobLevelList();
            }
            return JobLevelList;
        }

        /// <summary>
        /// Job Level 
        /// 用 1, 2, 3, 4, 5 等字串打 api 取得 oracle 對應的 job id
        /// 如果對應不到 就是職員
        /// 職員	0.
        /// 室/處長	1.
        /// 部長	2.
        /// 營運長	3.
        /// 執行長/總經理	4.
        /// 董事長  5.
        /// </summary>
        /// <param name="_number"></param>
        /// <returns></returns>
        public string GetJobIDByNumberAuto(string Number)
        {
            var JobID = "";
            var _number = Number;
            if (string.IsNullOrWhiteSpace(_number))
            {
                _number = "0";
            }

            if (JobLevelList == null)
            {
                JobLevelList = new List<OracleJobID>();
            }
            var tmp = (from item in JobLevelList
                       where item.ApprovalAuthority == _number
                       select item).FirstOrDefault();
            if (tmp != null)
            {
                //如果已經取過，直接拿來回
                return tmp.JobId;
            }
            else
            {
                //如果未取過，就打 api 去取，然後存下
                var jobLevelItem = GetJobLevelItemByNumber(_number);
                if (jobLevelItem == null)
                {
                    //如果對應不到就給他職員
                    log.WriteLog("5", $"Job Level:[{_number}]對應不到，給他0");
                    jobLevelItem = GetJobLevelItemByNumber("0");
                }
                else
                {
                    //如果對應到 如果不存在就存下來
                    var tmpitem = from item in JobLevelList
                                  where item.ApprovalAuthority == _number
                                  select item;
                    if (tmpitem == null)
                    {
                        JobLevelList.Add(jobLevelItem);
                        log.WriteLog("5", $"加入 Job Level Number={_number}, JobID={JobID}");
                        log.WriteLog("5", "目前總共的 Job Level List 有:");
                        foreach (var item in JobLevelList)
                        {
                            log.WriteLog("5", $"ApprovalAuthority={item.ApprovalAuthority}, Name={item.Name}, JobID={item.JobId}");
                        }
                    }
                }
            }

            return JobID;
        }

        public OracleJobID GetJobLevelItemByNumber(string Number)
        {

            OracleJobID rst = null;
            var _number = Number;
            if (string.IsNullOrWhiteSpace(_number))
            {
                _number = "0";
            }
            //var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/workers?q=PersonNumber={PersonNumber}&expand=workRelationships.assignments";
            var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/jobs?q=ApprovalAuthority=\"{_number}\";SetId=0";
            try
            {


                //  //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (string.IsNullOrWhiteSpace(mr.ReturnData))
                {
                    throw new Exception($"取 Job ID 失敗:{mr.StatusCode}{mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                log.WriteLog("5", $"取得 oracle Job Level/Job ID 對應表:{desc_str}");
                OracleJobIDReturnModel ids = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleJobIDReturnModel>(desc_str);
                if (ids.Count > 0)
                {
                    rst = ids.Items[0];
                }

                //foreach (var jobid in ids.Items)
                //{
                //    rst = jobid.JobId;
                //}
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"GetJobLevelItemByNumber Get Job Level Item 失敗:api:{url}\r\n{ex.Message}{ex.InnerException}");
                throw new Exception($"GetJobLevelItemByNumber Get Job Level Item 失敗:{ex.Message}{ex.InnerException}");
            }



            return rst;
        }



        /// <summary>
        /// Job Level 
        /// 用 1, 2, 3, 4, 5 等字串打 api 取得 oracle 對應的 job id
        /// 如果對應不到 就是空白
        /// 職員	0.
        /// 室/處長	1.
        /// 部長	2.
        /// 營運長	3.
        /// 執行長/總經理	4.
        /// 董事長  5.
        /// </summary>
        /// <param name="ChineseName"></param>
        /// <returns></returns>
        public string GetJobIDByNumber(string Number)
        {


            var _number = Number;
            if (string.IsNullOrWhiteSpace(_number))
            {
                _number = "0";
            }
            string rst = "";
            //var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/workers?q=PersonNumber={PersonNumber}&expand=workRelationships.assignments";
            var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/jobs?q=ApprovalAuthority=\"{_number}\";SetId=0";
            try
            {


                //  //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (string.IsNullOrWhiteSpace(mr.ReturnData))
                {
                    throw new Exception($"取 Job ID 失敗:{mr.StatusCode}{mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                OracleJobIDReturnModel ids = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleJobIDReturnModel>(desc_str);
                foreach (var jobid in ids.Items)
                {
                    rst = jobid.JobId;
                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"GetJobIDByNumber Get Job ID 失敗:api:{url}\r\n{ex.Message}{ex.InnerException}");
                throw new Exception($"GetJobIDByNumber Get Job ID 失敗:{ex.Message}{ex.InnerException}");
            }



            return rst;
        }



        /// <summary>
        /// Job Level 
        /// 用中文取得 oracle 對應的 id
        /// 職員	0.
        /// 室/處長	1.
        /// 部長	2.
        /// 營運長	3.
        /// 執行長/總經理	4.
        /// </summary>
        /// <param name="ChineseName"></param>
        /// <returns></returns>
        public string GetJobIDByChinese(string ChineseName)
        {
            var cname = ChineseName;
            if (string.IsNullOrWhiteSpace(cname))
            {
                cname = "職員";
            }
            string rst = "";
            //var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/workers?q=PersonNumber={PersonNumber}&expand=workRelationships.assignments";
            var url = $"{Oracle_Domain}/hcmRestApi/resources/11.13.18.05/jobs?q=Name=\"{cname}\";SetId=0";
            try
            {


                //  //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                MiddleReturn mr = HttpGetFromOracleAP(url);
                if (string.IsNullOrWhiteSpace(mr.ReturnData))
                {
                    throw new Exception($"取 Job ID 失敗:{mr.StatusCode}{mr.ErrorMessage}");
                }
                var bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                var desc_str = Encoding.UTF8.GetString(bs64_bytes);
                OracleJobIDReturnModel ids = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleJobIDReturnModel>(desc_str);
                foreach (var jobid in ids.Items)
                {
                    rst = jobid.JobId;
                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"Get Job ID 失敗:api:{url}\r\n{ex.Message}{ex.InnerException}");
                throw new Exception($"Get Job ID 失敗:{ex.Message}{ex.InnerException}");
            }

            return rst;
        }


        public SupplierBankData GetSupplierDataByEmpNo(string EmpNo)
        {
            SupplierBankData rst = null;
            string url = "";
            // step 1.       Get Supplier
            MiddleReturn mr = GetOracleSupplierByEmpNo(EmpNo);
            byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
            string desc_str = Encoding.UTF8.GetString(bs64_bytes);

            SupplierResponseModel obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierResponseModel>(desc_str);
            if (obj.Count == 0) return rst;
            var SupplierId = obj.Items[0].SupplierId;
            var SupplierNumber = obj.Items[0].SupplierNumber;
            var SupplierPartyId = obj.Items[0].SupplierPartyId;

            //Step 2.       Get Supplier Address(第一步得到的Supplier ID)
            //  url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers/{SupplierId}/child/addresses";
            //log.WriteLog("5", $"{EmpNo} addresses url:{url}");
            //mr = HttpGetFromOracleAP(url);
            //bs64_bytes = Convert.FromBase64String(mr.ReturnData);
            //desc_str = Encoding.UTF8.GetString(bs64_bytes);
            //SupplierAddrReturnModel objAddr = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierAddrReturnModel>(desc_str);
            //var AddressName = "";
            //if (objAddr.Count > 0)
            //{
            //    AddressName = objAddr.Items[0].AddressName;
            //}

            //Step 3.       Get Supplier Site(第一步得到的Supplier ID)
            //url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers/{SupplierId}/child/sites";
            ////string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
            //log.WriteLog("5", $"{EmpNo} sites url:{url}");
            //mr = HttpGetFromOracleAP(url);
            //bs64_bytes = Convert.FromBase64String(mr.ReturnData);
            //desc_str = Encoding.UTF8.GetString(bs64_bytes);
            //SupplierSiteReturnModel obj1 = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierSiteReturnModel>(desc_str);
            //var _SupplierSiteId = "";
            //if (obj.Count > 0)
            //{
            //    _SupplierSiteId = obj1.Items[0].SupplierSiteId;
            //}

            //Step 4.       Get Supplier Assignment(第一個用第一步獲得的Supplier ID，第二個用第三步獲得的SupplierSiteId)
            //if (!string.IsNullOrWhiteSpace(_SupplierSiteId))
            //{
            //    url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers/{SupplierId}/child/sites/{_SupplierSiteId}/child/assignments";
            //    //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
            //    log.WriteLog("5", $"{EmpNo} assignments url:{url}");
            //    mr = HttpGetFromOracleAP(url);
            //    bs64_bytes = Convert.FromBase64String(mr.ReturnData);
            //    desc_str = Encoding.UTF8.GetString(bs64_bytes);

            //    SupplierAssignmentReturnModel objAssignment = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierAssignmentReturnModel>(desc_str);
            //    var AssignmentId = "";
            //    if (objAssignment.Count > 0)
            //    {
            //        AssignmentId = objAssignment.Items[0].AssignmentId;
            //    }
            //}

            //5.       Get Supplier Bank Account
            //                5.1   Get Payee Party ID(第一步獲得的SupplierPartyID)
            url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/paymentsExternalPayees?finder=ExternalPayeeSearch;PayeePartyIdentifier={SupplierPartyId},Intent=Supplier";
            //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
            mr = HttpGetFromOracleAP(url);
            bs64_bytes = Convert.FromBase64String(mr.ReturnData);
            desc_str = Encoding.UTF8.GetString(bs64_bytes);
            SupplierPayeeReturnModel objPayee = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierPayeeReturnModel>(desc_str);
            var PayeePartyIdentifier = "";
            var PayeePartySiteIdentifier = "";
            //if (objPayee.Count > 0)
            foreach (var item in objPayee.Items)
            {
                PayeePartyIdentifier = item.PayeePartyIdentifier;
                PayeePartySiteIdentifier = $"{item.PayeePartySiteIdentifier}";

                //Step 5.2   Get Supplier Bank Account(第一個用5.1步驟獲得的PayeePartyIdentifier，
                //第二個用5.1步驟獲得的PayeePartySiteIdentifier)
                if ((!string.IsNullOrWhiteSpace(PayeePartyIdentifier)) && (!string.IsNullOrWhiteSpace(PayeePartySiteIdentifier)))
                {

                    url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/payeeBankAccountsLOV?finder=AvailablePayeeBankAccountsFinder;PaymentFunction=\"PAYABLES_DISB\",PayeePartyId={PayeePartyIdentifier},SupplierSiteId={PayeePartySiteIdentifier}";
                    //string url = $"{Oracle_Domain}/fscmRestApi/resources/11.13.18.05/suppliers?q=SupplierNumber '{EmpNo}'";
                    mr = HttpGetFromOracleAP(url);
                    bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                    desc_str = Encoding.UTF8.GetString(bs64_bytes);
                    var BankAccountId = "";
                    SupplierBankAccountReturnModel objBankAccount = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierBankAccountReturnModel>(desc_str);
                    if (objBankAccount.Count > 0)
                    {
                        rst = new SupplierBankData();
                        BankAccountId = objBankAccount.Items[0].BankAccountId;
                        rst.AccountNumber = objBankAccount.Items[0].AccountNumber;
                        rst.AccountName = objBankAccount.Items[0].AccountName;
                        rst.BankName = objBankAccount.Items[0].BankName;
                        rst.BranchName = objBankAccount.Items[0].BranchName;
                        break;
                    }
                }

            }



            return rst;
        }

        public void SyncOneSupplierByEmpNo(string EmpNo)
        {
            try
            {
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);



                // 因為同步一個供應商 約需5分鐘
                // 假設有1000個員工就要5000分鐘約3.5天
                // 太久了  所以要改成有同步過的就不同步
                // 因為也不會改(有要改再說)
                // 要同步之前 先查一下自己有沒有資料
                // 沒有再同步
                string SupplierNumber = "";
                string SupplierPartyId = "";
                string SupplierSiteId = "";
                string SupplierId = "";
                string PayeePartyIdentifier = "";
                string PayeePartySiteIdentifier = "";
                bool has_this_supplier = false;
                string sql = "";
                string _BatchNo = "Supplier"; //因為同步太慢了 就不分批 只要有資料就記下來備查
                bool need_sync = false;
                // 1. supplier
                try
                {

                    sql = $@"
select count(1) from  [OracleSupplier]
	WHERE [SupplierNumber] = :0
";
                    int cnts = int.Parse(sqlite.ExecuteScalarA(sql, EmpNo));
                    if (cnts == 0)
                    {
                        need_sync = true;
                    }
                    else
                    {
                        sql = $@"
select count(1) from  [OracleSupplierAddr]
	WHERE [SupplierNumber] = :0
";
                        cnts = int.Parse(sqlite.ExecuteScalarA(sql, EmpNo));
                        if (cnts == 0)
                        {
                            need_sync = true;
                        }
                        else
                        {
                            sql = $@"
select count(1) from  [OracleSupplierSite]
	WHERE [SupplierNumber] = :0
";
                            cnts = int.Parse(sqlite.ExecuteScalarA(sql, EmpNo));
                            if (cnts == 0)
                            {
                                need_sync = true;
                            }
                            else
                            {
                                sql = $@"
select count(1) from  [OracleSupplierAssignment]
	WHERE [SupplierNumber] = :0
";
                                cnts = int.Parse(sqlite.ExecuteScalarA(sql, EmpNo));
                                if (cnts == 0)
                                {
                                    need_sync = true;
                                }
                                else
                                {
                                    sql = $@"
select count(1) from  [OracleSupplierBankAccount]
	WHERE [SupplierNumber] = :0
";
                                    cnts = int.Parse(sqlite.ExecuteScalarA(sql, EmpNo));
                                    if (cnts == 0)
                                    {
                                        need_sync = true;
                                    }
                                }
                            }
                        }
                    }

                    if (need_sync)
                    {
                        log.WriteLog("5", $"準備從oracle同步員工供應商回來:{EmpNo}");
                        MiddleReturn mr = GetOracleSupplierByEmpNo(EmpNo);
                        byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                        string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                        SupplierResponseModel obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierResponseModel>(desc_str);

                        if (obj.Count > 0)
                        {
                            foreach (var item in obj.Items)
                            {
                                SupplierId = item.SupplierId;
                                SupplierNumber = item.SupplierNumber;
                                SupplierPartyId = item.SupplierPartyId;

                                sql = $@"
UPDATE [OracleSupplier]
	SET [BatchNo] = :0
		,[SupplierId] = :1
		,[SupplierPartyId] = :2
		,[Supplier] = :3
		,[AlternateName] = :4
		,[SupplierTypeCode] = :5
		,[SupplierType] = :6
		,[Status] = :7
	WHERE [SupplierNumber] = :8
";
                                int cntu = sqlite.ExecuteByCmd(sql,
                                         _BatchNo,
                                         item.SupplierId,
                                         item.SupplierPartyId,
                                         item.Supplier,
                                         item.AlternateName,
                                         item.SupplierTypeCode,
                                         item.SupplierType,
                                         item.Status,
                                         item.SupplierNumber
                                         );
                                if (cntu == 0)
                                {

                                    sql = $@"
INSERT INTO [OracleSupplier]
		(
    [BatchNo] ,
	[SupplierId] ,
	[SupplierPartyId] ,
	[Supplier] ,
	[SupplierNumber] ,
	[AlternateName] ,
	[SupplierTypeCode] ,
	[SupplierType] ,
	[Status] 
)
	VALUES(
:0
,:1
,:2
,:3
,:4
,:5
,:6
,:7
,:8
)";
                                    int cnt = sqlite.ExecuteByCmd(sql, _BatchNo,
                                              item.SupplierId,
                                              item.SupplierPartyId,
                                              item.Supplier,
                                              item.SupplierNumber,
                                              item.AlternateName,
                                              item.SupplierTypeCode,
                                              item.SupplierType,
                                              item.Status
                                              );
                                }

                                has_this_supplier = true;
                            }
                        }
                        log.WriteLog("5", $"同步了 {obj.Count} 筆");


                    }



                }
                catch (Exception ex)
                {
                    log.WriteErrorLog($"取得 1.Supplier:{EmpNo} 失敗:{ex.Message}{ex.InnerException}");
                }


                // 2. supplier address
                // 2-Get Supplier Address(第一步得到的Supplier ID)
                if (has_this_supplier)
                {

                    try
                    {
                        if (string.IsNullOrWhiteSpace(SupplierId))
                        {
                            throw new Exception($"SupplierId is null!");
                        }

                        MiddleReturn mr = GetOracleSupplierAddrByEmpNo(EmpNo, _BatchNo, SupplierId);
                        byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                        string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                        SupplierAddrReturnModel obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierAddrReturnModel>(desc_str);

                        if (obj.Count > 0)
                        {
                            foreach (var item in obj.Items)
                            {

                                sql = $@"
UPDATE [OracleSupplierAddr]
	SET [BatchNo] = :0
		,[SupplierAddressId] = :1
		,[AddressName] = :2
		,[CountryCode] = :3
		,[Country] = :4
		,[AddressLine1] = :5
	WHERE [SupplierNumber] = :6
";
                                int cntu = sqlite.ExecuteByCmd(sql,
                                       _BatchNo,
                                       item.SupplierAddressId,
                                       item.AddressName,
                                       item.CountryCode,
                                       item.Country,
                                       item.AddressLine1,
                                       SupplierNumber
                                       );


                                if (cntu == 0)
                                {

                                    sql = $@"
INSERT INTO [OracleSupplierAddr]
		(
    [BatchNo] ,
	[SupplierNumber] ,
	[SupplierAddressId] ,
	[AddressName] ,
	[CountryCode] ,
	[Country] ,
	[AddressLine1]  
)
	VALUES(
:0
,:1
,:2
,:3
,:4
,:5
,:6 )
";
                                    int cnt = sqlite.ExecuteByCmd(sql, _BatchNo,
                                              SupplierNumber,
                                              item.SupplierAddressId,
                                              item.AddressName,
                                              item.CountryCode,
                                              item.Country,
                                              item.AddressLine1
                                              );
                                }

                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log.WriteErrorLog($"取得 2.Supplier addr:{EmpNo} 失敗:{ex.Message}{ex.InnerException}");
                    }
                }


                // 3. supplier site
                // 3-Get Supplier Site(第一步得到的Supplier ID)
                if (has_this_supplier)
                {

                    try
                    {
                        if (string.IsNullOrWhiteSpace(SupplierId))
                        {
                            throw new Exception($"SupplierId is null!");
                        }


                        MiddleReturn mr = GetOracleSupplierSiteByEmpNo(EmpNo, _BatchNo, SupplierId);
                        byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                        string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                        SupplierSiteReturnModel obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierSiteReturnModel>(desc_str);

                        if (obj.Count > 0)
                        {
                            foreach (var item in obj.Items)
                            {
                                var _SupplierSiteId = $"{item.SupplierSiteId}";
                                if (!string.IsNullOrWhiteSpace(_SupplierSiteId))
                                {
                                    SupplierSiteId = _SupplierSiteId;
                                }

                                sql = $@"
UPDATE [OracleSupplierSite]
	SET [BatchNo] = :0
		,[SupplierSiteId] = :1
		,[SupplierSite] = :2
		,[SupplierAddressId] = :3
		,[SupplierAddressName] = :4
	WHERE 	[SupplierNumber] = :5
";
                                int cntu = sqlite.ExecuteByCmd(sql,
                                       _BatchNo,
                                       item.SupplierSiteId,
                                       item.SupplierSite,
                                       item.SupplierAddressId,
                                       item.SupplierAddressName,
                                       SupplierNumber
                                       );

                                if (cntu == 0)
                                {

                                    sql = $@"
INSERT INTO [OracleSupplierSite]
		(
    [BatchNo] ,
	[SupplierNumber] ,

	[SupplierSiteId] ,
	[SupplierSite],
	[SupplierAddressId] ,
	[SupplierAddressName]  
)
	VALUES(
:0
,:1
,:2
,:3
,:4
,:5
)
";
                                    int cnt = sqlite.ExecuteByCmd(sql,
                                              _BatchNo,
                                              SupplierNumber,
                                              item.SupplierSiteId,
                                              item.SupplierSite,
                                              item.SupplierAddressId,
                                              item.SupplierAddressName
                                              );
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log.WriteErrorLog($"取得 3.Supplier Site:{EmpNo} 失敗:{ex.Message}{ex.InnerException}");
                    }
                }


                // 4. supplier Assignment
                // 4-Get Supplier Assignment( 第一個用第一步獲得的 Supplier ID，
                //     第二個用第三步獲得的 SupplierSiteId)
                if (has_this_supplier)
                {

                    try
                    {
                        if (string.IsNullOrWhiteSpace(SupplierId))
                        {
                            throw new Exception($"SupplierId is null!");
                        }

                        if (string.IsNullOrWhiteSpace(SupplierSiteId))
                        {
                            throw new Exception($"SupplierSiteId is null!");
                        }


                        MiddleReturn mr = GetOracleSupplierAssignmentByEmpNo(EmpNo, _BatchNo, SupplierId, SupplierSiteId);
                        byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                        string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                        SupplierAssignmentReturnModel obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierAssignmentReturnModel>(desc_str);

                        if (obj.Count > 0)
                        {
                            foreach (var item in obj.Items)
                            {
                                sql = $@"
UPDATE [OracleSupplierAssignment]
	SET [BatchNo] = :0
		,[AssignmentId] = :1
		,[Status] = :2
	WHERE
		[SupplierNumber] =  :3
";
                                int cntu = sqlite.ExecuteByCmd(sql,
                                        _BatchNo,
                                        item.AssignmentId,
                                        item.Status,
                                        SupplierNumber
                                        );
                                if (cntu == 0)
                                {

                                    sql = $@"
INSERT INTO [OracleSupplierAssignment]
		(
    [BatchNo] ,
	[SupplierNumber] ,
	[AssignmentId] ,
	[Status] 
)
	VALUES(
:0
,:1
,:2
,:3

)
";
                                    int cnt = sqlite.ExecuteByCmd(sql,
                                              _BatchNo,
                                              SupplierNumber,
                                              item.AssignmentId,
                                              item.Status
                                              );
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log.WriteErrorLog($"取得 4.Supplier Assignment:{EmpNo} 失敗:{ex.Message}{ex.InnerException}");
                    }
                }

                // 5-1. supplier Payee
                //5-1.Get Supplier Bank Account-5.1   Get Payee Party ID(第一步獲得的 SupplierPartyID)
                if (has_this_supplier)
                {

                    try
                    {
                        if (string.IsNullOrWhiteSpace(SupplierPartyId))
                        {
                            throw new Exception($"SupplierPartyId is null!");
                        }

                        MiddleReturn mr = GetOracleSupplierPayeeByEmpNo(EmpNo, _BatchNo, SupplierPartyId);
                        byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                        string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                        SupplierPayeeReturnModel obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierPayeeReturnModel>(desc_str);

                        if (obj.Count > 0)
                        {
                            foreach (var item in obj.Items)
                            {

                                //  (第一個用5.1步驟獲得的 PayeePartyIdentifier，
                                //    第二個用5.1步驟獲得的 PayeePartySiteIdentifier)
                                var _PayeePartyIdentifier = $"{item.PayeePartyIdentifier}";
                                var _PayeePartySiteIdentifier = $"{item.PayeePartySiteIdentifier}";
                                if (!string.IsNullOrWhiteSpace(_PayeePartyIdentifier))
                                {
                                    PayeePartyIdentifier = _PayeePartyIdentifier;
                                    if (!string.IsNullOrWhiteSpace(_PayeePartySiteIdentifier))
                                    {
                                        PayeePartySiteIdentifier = _PayeePartySiteIdentifier;

                                        sql = $@"
UPDATE [OracleSupplierPayee]
	SET [BatchNo] = :0
		,[PayeeId] = :1
		,[PayeePartyIdentifier] = :2
		,[PartyName] = :3
		,[PayeePartyNumber] = :4
		,[PayeePartySiteIdentifier] = :5
		,[SupplierSiteCode] = :6
		,[SupplierSiteIdentifier] = :7
		,[PayeePartySiteNumber] = :8
	WHERE
		[SupplierNumber] = :9
";
                                        int cntu = sqlite.ExecuteByCmd(sql,
                                               _BatchNo,
                                               item.PayeeId,
                                               item.PayeePartyIdentifier,
                                               item.PartyName,
                                               item.PayeePartyNumber,
                                               item.PayeePartySiteIdentifier,
                                               item.SupplierSiteCode,
                                               item.SupplierSiteIdentifier,
                                               item.PayeePartySiteNumber,
                                               SupplierNumber
                                               );
                                        if (cntu == 0)
                                        {

                                            sql = $@"
INSERT INTO [OracleSupplierPayee]
		(
    [BatchNo] ,
	[SupplierNumber] ,

	[PayeeId] ,
	[PayeePartyIdentifier] ,
	[PartyName] ,
	[PayeePartyNumber] ,
	[PayeePartySiteIdentifier] ,
	[SupplierSiteCode] ,
	[SupplierSiteIdentifier] ,
	[PayeePartySiteNumber]
)
	VALUES(
:0
,:1
,:2
,:3
,:4
,:5
,:6
,:7
,:8
,:9

)
";
                                            int cnt = sqlite.ExecuteByCmd(sql, _BatchNo,
                                                      SupplierNumber,
                                                      item.PayeeId,
                                                      item.PayeePartyIdentifier,
                                                      item.PartyName,
                                                      item.PayeePartyNumber,
                                                      item.PayeePartySiteIdentifier,
                                                      item.SupplierSiteCode,
                                                      item.SupplierSiteIdentifier,
                                                      item.PayeePartySiteNumber
                                                      );
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log.WriteErrorLog($"取得 5-1.Supplier Payee:{EmpNo} 失敗:{ex.Message}{ex.InnerException}");
                    }
                }

                // 5.2   Get Supplier Bank Account
                //  (第一個用5.1步驟獲得的 PayeePartyIdentifier，
                //    第二個用5.1步驟獲得的 PayeePartySiteIdentifier)
                if (has_this_supplier)
                {

                    try
                    {
                        if (string.IsNullOrWhiteSpace(PayeePartyIdentifier))
                        {
                            throw new Exception($"PayeePartyIdentifier is null!");
                        }
                        if (string.IsNullOrWhiteSpace(PayeePartySiteIdentifier))
                        {
                            throw new Exception($"PayeePartySiteIdentifier is null!");
                        }



                        MiddleReturn mr = GetOracleSupplierBankAccountByEmpNo(EmpNo, _BatchNo,
                            PayeePartyIdentifier, PayeePartySiteIdentifier);
                        byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                        string desc_str = Encoding.UTF8.GetString(bs64_bytes);

                        SupplierBankAccountReturnModel obj = Newtonsoft.Json.JsonConvert.DeserializeObject<SupplierBankAccountReturnModel>(desc_str);

                        if (obj.Count > 0)
                        {
                            foreach (var item in obj.Items)
                            {
                                sql = $@"
UPDATE [OracleSupplierBankAccount]
	SET [BatchNo] = :0
		,[BankAccountId] = :1
		,[AccountNumber] = :2
		,[AccountName] = :3
		,[BankName] = :4
		,[BranchName] = :5
	WHERE 
		[SupplierNumber] = :6
";
                                int cntu = sqlite.ExecuteByCmd(sql,
                                        _BatchNo,
                                        item.BankAccountId,
                                        item.AccountNumber,
                                        item.AccountName,
                                        item.BankName,
                                        item.BranchName,
                                        SupplierNumber
                                        );

                                if (cntu == 0)
                                {

                                    sql = $@"
INSERT INTO [OracleSupplierBankAccount]
		(
    [BatchNo] ,
	[SupplierNumber] ,

	[BankAccountId] ,
	[AccountNumber] ,
	[AccountName] ,
	[BankName] ,
	[BranchName]  
)
	VALUES(
:0
,:1
,:2
,:3
,:4
,:5
,:6 

)
";
                                    int cnt = sqlite.ExecuteByCmd(sql, _BatchNo,
                                              SupplierNumber,

                                              item.BankAccountId,
                                              item.AccountNumber,
                                              item.AccountName,
                                              item.BankName,
                                              item.BranchName
                                              );
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        log.WriteErrorLog($"取得 5.2   Get Supplier Bank Account:{EmpNo} 失敗:{ex.Message}{ex.InnerException}");
                    }
                }

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"SyncOneSupplierByEmpNo 同步一個供應商失敗:{ex.Message}{ex.InnerException}");
            }
        }

        /// <summary>
        /// 飛騰要先同步 這樣才知道有哪些 supplier 要查
        /// </summary>
        /// <param name="_BatchNo"></param>
        public void SyncSuppliers(string BatchNo)
        {

            CreateSQLiteFileIfNotExists_Supplier_1();
            CreateSQLiteFileIfNotExists_Supplier_2_address();
            CreateSQLiteFileIfNotExists_Supplier_3_site();
            CreateSQLiteFileIfNotExists_Supplier_4_Assignment();
            CreateSQLiteFileIfNotExists_Supplier_5_Payee();
            CreateSQLiteFileIfNotExists_Supplier_5_BankAccount();
            CreateSQLiteFileIfNotExists_HR_Oracle_Mapping();

            try
            {
                string sql = $@"
SELECT [BatchNo]
		,[r_code]
		,[r_cname]
		,[r_ename]
		,[r_dept]
		,[r_degress]
		,[r_cell_phone]
		,[r_birthday]
		,[r_sex]
		,[r_email]
		,[r_phone_ext]
		,[r_skype_id]
		,[r_online_date]
		,[r_online]
		,[r_offline_date]
		,[departname]
		,[dutyid]
		,[workplaceid]
		,[Workplacename]
		,[idno]
		,[salaccountname]
		,[bankcode]
		,[bankbranch]
		,[salaccountid]
		,[jobstatus]
	FROM [Supplier] s
	where BatchNo = :0
	and s.r_code like '8%'	
";

                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                DataTable tb;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo))
                {
                    //供應商不要分批次
                    //只要有資料就記下來
                    //供下次查詢
                    //因為同步一次的時間要好幾天

                    string NewEmpNo = $"{row["r_code"]}";
                    string EmpNo = NewEmpNo;

                    // stage 上用的是舊員編
                    EmpNo = GetOldEmpNo(NewEmpNo);

                    SyncOneSupplierByEmpNo(EmpNo);
                }

                //SaveOracleSuppliersIntoSQLite(Suppliers, BatchNo);

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"Get Supplier 失敗:{ex.Message}{ex.InnerException}");
            }

        }



        /// <summary>
        /// 把 oracle worker 存進 sqlite
        /// </summary>
        /// <param name="Workers"></param>
        void SaveOracleWorkersIntoSQLite(List<Worker> Workers, string BatchNo)
        {
            try
            {
                // 檢查一下 如果 sqlite 檔案沒有就自動建立
                CreateSQLiteFileIfNotExists_Batch();
                CreateSQLiteFileIfNotExists_OracleWorkers();
                CreateSQLiteFileIfNotExists_WorkerLinks();
                CreateSQLiteFileIfNotExists_WorkerProperties();
                CreateSQLiteFileIfNotExists_WorkerNames();
                CreateSQLiteFileIfNotExists_WorkerNamesLink();
                InsertIntoWorkersTable(Workers, BatchNo);
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"SaveOracleWorkersIntoSQLite 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        void InsertIntoWorkersTable(List<Worker> Workers, string BatchNo)
        {
            try
            {
                //string BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string sql = "";

                //SQLiteUtl sqlite = new SQLiteUtl(db_file);

                foreach (Worker worker in Workers)
                {
                    sql = $@"
INSERT INTO [OracleWorkers]
		([BatchNo]
		,[PersonId]
		,[PersonNumber]
		,[CorrespondenceLanguage]
		,[BloodType]
		,[DateOfBirth]
		,[DateOfDeath]
		,[CountryOfBirth]
		,[RegionOfBirth]
		,[TownOfBirth]
		,[ApplicantNumber]
		,[CreatedBy]
		,[CreationDate]
		,[LastUpdatedBy]
		,[LastUpdateDate])
	VALUES
		(
:0,
:1,
:2,
:3,
:4,
:5,
:6,
:7,
:8,
:9,
:10,
:11,
:12,
:13,
:14
);
";
                    sqlite.ExecuteByCmd(sql,
                        BatchNo,
                        worker.PersonId,
                        worker.PersonNumber,
                        worker.CorrespondenceLanguage,
                        worker.BloodType,
                        worker.DateOfBirth,
                        worker.DateOfDeath,
                        worker.CountryOfBirth,
                        worker.RegionOfBirth,
                        worker.TownOfBirth,
                        worker.ApplicantNumber,
                        worker.CreatedBy,
                        worker.CreationDate,
                        worker.LastUpdatedBy,
                        worker.LastUpdateDate
                        );

                    foreach (WorkerLink link in worker.Links)
                    {
                        try
                        {
                            sql = $@"

INSERT INTO [WorkerLinks]
		([BatchNo]
        ,[PersonId]
		,[PersonNumber]
		,[Rel]
		,[Href]
		,[Name]
		,[Kind])
	VALUES
		(
:0
,:1
,:2
,:3
,:4
,:5
,:6
)
";
                            sqlite.ExecuteByCmd(sql,
                                    BatchNo,
                                    worker.PersonId,
                                    worker.PersonNumber,
                                    link.Rel,
                                    link.Href,
                                    link.Name,
                                    link.Kind
                                );

                            // 抓取 worker names 寫入 workerNames table
                            try
                            {
                                if (link.Name == "names")
                                {
                                    string url = link.Href;

                                    // ORACLE STAGE
                                    // for  oracle ap
                                    MiddleModel2 send_model2 = new MiddleModel2();
                                    send_model2.URL = url;
                                    //send_model2.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(par);
                                    send_model2.Method = "GET";
                                    //string username = this.UserName;
                                    //string password = this.Password;
                                    send_model2.UserName = this.UserName;
                                    send_model2.Password = this.Password;
                                    string usernamePassword = send_model2.UserName + ":" + send_model2.Password;
                                    send_model2.AddHeaders.Add("Authorization", "Basic " + Convert.ToBase64String(Encoding.UTF8.GetBytes(usernamePassword)));
                                    //CredentialCache myCred = new CredentialCache();
                                    //myCred.Add(new Uri(send_model2.URL), "Basic", new NetworkCredential(username, password));
                                    //send_model2.Cred = myCred;
                                    send_model2.Timeout = int30Mins;


                                    // for BOXMAN API
                                    MiddleModel send_model = new MiddleModel();
                                    var _url = $"{Oracle_AP}/api/Middle/Call/";
                                    send_model.URL = _url;
                                    send_model.SendingData = Newtonsoft.Json.JsonConvert.SerializeObject(send_model2);
                                    send_model.Method = "POST";
                                    send_model.Timeout = int30Mins;
                                    var ret = ApiOperation.CallApi(
                                            new ApiRequestSetting()
                                            {
                                                Data = send_model,
                                                MethodRoute = "api/Middle/Call",
                                                MethodType = "POST",
                                                TimeOut = int30Mins
                                            }
                                            );
                                    if (ret.Success)
                                    {
                                        string receive_str = ret.ResponseContent;
                                        try
                                        {
                                            MiddleReturn mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(receive_str);
                                            if (!string.IsNullOrWhiteSpace(mr2.ErrorMessage))
                                            {
                                                var _errmsg = $"呼叫 {url} 失敗:{mr2.StatusCode} {mr2.StatusDescription} {mr2.ReturnData} {mr2.ErrorMessage}";
                                                throw new Exception(_errmsg);
                                            }

                                            MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2.ReturnData);
                                            if (!string.IsNullOrWhiteSpace(mr.ErrorMessage))
                                            {
                                                var _errmsg = $"呼叫 {url} 失敗:{mr.StatusCode} {mr.StatusDescription} {mr.ReturnData} {mr.ErrorMessage}";
                                                throw new Exception(_errmsg);
                                            }

                                            byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                                            string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                                            WorkerNamesResponseObj obj = Newtonsoft.Json.JsonConvert.DeserializeObject<WorkerNamesResponseObj>(desc_str);
                                            if (obj.Count > 0)
                                            {
                                                WorkerName work_Name = obj.Items[0];
                                                sql = $@"

INSERT INTO [WorkerNames]
		([BatchNo]
        ,[PersonId]
		,[PersonNumber]
		,[PersonNameId]
		,[EffectiveStartDate]
		,[EffectiveEndDate]
		,[LegislationCode]
		,[LastName]
		,[FirstName]
		,[Title]
		,[MiddleNames]
		,[DisplayName]
		,[OrderName]
		,[ListName]
		,[FullName]
		,[NameLanguage]
		,[CreatedBy]
		,[CreationDate]
		,[LastUpdatedBy]
		,[LastUpdateDate]
)
	VALUES
		(
:0
,:1
,:2
,:3
,:4
,:5
,:6
,:7
,:8
,:9
,:10
,:11
,:12
,:13
,:14
,:15
,:16
,:17
,:18
,:19

)
";
                                                int cnt_names = sqlite.ExecuteByCmd(sql,
                                                    BatchNo,
                                                    worker.PersonId,
                                                    worker.PersonNumber,
                                                    work_Name.PersonNameId,
                                                    work_Name.EffectiveStartDate,
                                                    work_Name.EffectiveEndDate,
                                                    work_Name.LegislationCode,
                                                    work_Name.LastName,
                                                    work_Name.FirstName,
                                                    work_Name.Title,
                                                    work_Name.MiddleNames,
                                                    work_Name.DisplayName,
                                                    work_Name.OrderName,
                                                    work_Name.ListName,
                                                    work_Name.FullName,
                                                    work_Name.NameLanguage,
                                                    work_Name.CreatedBy,
                                                    work_Name.CreationDate,
                                                    work_Name.LastUpdatedBy,
                                                    work_Name.LastUpdateDate
                                                    );

                                                foreach (var work_name_link in work_Name.Links)
                                                {
                                                    if (work_name_link.Rel == "self" && work_name_link.Name == "names")
                                                    {
                                                        try
                                                        {
                                                            sql = $@"
INSERT INTO [WorkerNamesLink]
(
BatchNo ,
PersonId ,
PersonNumber ,
PersonNameId ,
Href 
)
VALUES
(
:0
,:1
,:2
,:3
,:4
)
";
                                                            int cnt_names_link = sqlite.ExecuteByCmd(sql,
                                                                 BatchNo,
                                                                 worker.PersonId,
                                                                 worker.PersonNumber,
                                                                 work_Name.PersonNameId,
                                                                 work_name_link.Href
                                                                 );
                                                        }
                                                        catch (Exception exWorkNameLink)
                                                        {
                                                            log.WriteErrorLog($"Insert WorkerNameLink 失敗:{exWorkNameLink.Message}{exWorkNameLink.InnerException}");

                                                        }
                                                    }
                                                }
                                            }

                                        }
                                        catch (Exception exWorkNames)
                                        {
                                            log.WriteErrorLog($"呼叫 {url} 失敗:{exWorkNames.Message}{exWorkNames.InnerException}");
                                        }
                                    }
                                    else
                                    {
                                        log.WriteErrorLog($"Call Boxman Api {_url} 失敗:{ret.ErrorMessage} {ret.ErrorException}");
                                    }
                                }
                            }
                            catch (Exception wnEx)
                            {
                                log.WriteErrorLog($"寫入 Worker names 失敗:{wnEx}");
                            }
                        }
                        catch (Exception exlink)
                        {
                            log.WriteErrorLog($"寫入 WorkerLinks 失敗:{exlink}");
                        }


                        if (link.Properties != null)
                        {
                            try
                            {
                                sql = $@"
INSERT INTO [WorkerProperties]
		([BatchNo]
        ,[PersonId]
		,[PersonNumber]
		,[ChangeIndicator])
	VALUES
		(
:0
,:1
,:2
,:3
)
";
                                sqlite.ExecuteByCmd(sql,
                                     BatchNo,
                                     worker.PersonId,
                                     worker.PersonNumber,
                                     link.Properties.ChangeIndicator
                         );
                            }
                            catch (Exception exProp)
                            {
                                log.WriteErrorLog($"寫入 WorkerProperties 失敗:{exProp}");
                            }

                        }
                    }
                }


                sql = $@"
                INSERT INTO [Batch]
                		([BatchNo])
                	VALUES
                		(:0)
                ";
                int cnt = sqlite.ExecuteByCmd(sql, BatchNo);

                //int cnt = 0;

                // 留下 3 天的資料
                string B4BatchNo = DateTime.Today.AddDays(-3).ToString("yyyy-MM-dd HH:mm:ss.fff");
                sql = $@"
DELETE FROM  [Batch] WHERE [BatchNo] <= :0
";
                cnt = sqlite.ExecuteByCmd(sql, B4BatchNo);

                sql = $@"
                DELETE FROM   [OracleWorkers]  WHERE [BatchNo] <= :0
                ";
                cnt = sqlite.ExecuteByCmd(sql, B4BatchNo);

                sql = $@"
                DELETE FROM   [WorkerLinks]  WHERE [BatchNo] <= :0
                ";
                cnt = sqlite.ExecuteByCmd(sql, B4BatchNo);

                sql = $@"
                DELETE FROM   [WorkerNames]  WHERE [BatchNo] <= :0
                ";
                cnt = sqlite.ExecuteByCmd(sql, B4BatchNo);

                sql = $@"
DELETE FROM  [WorkerNamesLink] WHERE [BatchNo] <= :0
";
                cnt = sqlite.ExecuteByCmd(sql, B4BatchNo);

                sql = $@"
DELETE FROM  [WorkerProperties] WHERE [BatchNo] <= :0
";
                cnt = sqlite.ExecuteByCmd(sql, B4BatchNo);

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"InsertIntoTable 失敗:{ex.Message}{ex.InnerException}");
            }

        }


        /// <summary>
        /// 建立放oracle bank資料的table
        /// </summary>
        void CreateSQLiteFileIfNotExists_OracleBankDataBranchData()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_OracleBankDataBranchData";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "OracleBankData";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users

                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"CREATE TABLE [{table_name}] (
BANK_NAME nvarchar(100),
BANK_NUMBER nvarchar(10),
BANK_BRANCH_NAME nvarchar(100),
BRANCH_NUMBER nvarchar(10)

)";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                }



            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }


        /// <summary>
        /// 自動建立 oracle workers table
        /// </summary>
        void CreateSQLiteFileIfNotExists_OracleWorkers()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_OracleWorkers";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "OracleWorkers";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users

                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"CREATE TABLE [{table_name}] (
BatchNo nvarchar(30) ,
PersonId nvarchar(100),
PersonNumber nvarchar(100),
CorrespondenceLanguage nvarchar(50),
BloodType nvarchar(50),
DateOfBirth nvarchar(50),
DateOfDeath nvarchar(50),
CountryOfBirth nvarchar(50),
RegionOfBirth nvarchar(50),
TownOfBirth nvarchar(100),
ApplicantNumber nvarchar(100),
CreatedBy nvarchar(100),
CreationDate nvarchar(100),
LastUpdatedBy nvarchar(100),
LastUpdateDate nvarchar(100)
)";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (PersonNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        /// <summary>
        /// create oracle worker links table
        /// </summary>
        void CreateSQLiteFileIfNotExists_WorkerLinks()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_WorkerLinks";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "WorkerLinks";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    int cnt = 0;
                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"CREATE TABLE [{table_name}] (
BatchNo nvarchar(30) ,
PersonId nvarchar(100),
PersonNumber nvarchar(50),
Rel nvarchar(100),
Href nvarchar(4000),
Name nvarchar(100),
Kind nvarchar(100)
)";

                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (PersonNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        void CreateSQLiteFileIfNotExists_WorkerProperties()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_WorkerProperties";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "WorkerProperties";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table links
                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"CREATE TABLE [{table_name}] (
BatchNo nvarchar(30) ,
PersonId nvarchar(100),
PersonNumber nvarchar(50),
ChangeIndicator nvarchar(4000)
)";
                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (PersonNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }

        }

        void CreateSQLiteFileIfNotExists_WorkerNames()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_WorkerNames";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "WorkerNames";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table links
                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"CREATE TABLE [{table_name}] (
BatchNo nvarchar(30) ,
PersonId nvarchar(100),
PersonNumber nvarchar(50),
PersonNameId nvarchar(100),
EffectiveStartDate nvarchar(50),
EffectiveEndDate nvarchar(50),
LegislationCode nvarchar(10),
LastName nvarchar(100),
FirstName nvarchar(100),
Title nvarchar(100),
MiddleNames nvarchar(100),
DisplayName nvarchar(100),
OrderName nvarchar(100),
ListName nvarchar(100),
FullName nvarchar(100),
NameLanguage nvarchar(100),
CreatedBy nvarchar(100),
CreationDate nvarchar(100),
LastUpdatedBy nvarchar(100),
LastUpdateDate nvarchar(100)
)";
                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (PersonNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }

        }
        void CreateSQLiteFileIfNotExists_WorkerNamesLink()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_WorkerNamesLink";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "WorkerNamesLink";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table links
                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"CREATE TABLE [{table_name}] (
BatchNo nvarchar(30) ,
PersonId nvarchar(100),
PersonNumber nvarchar(50),
PersonNameId nvarchar(100),
Href nvarchar(4000)
)";
                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (PersonNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }

        }


        void SaveOracleUsersIntoSQLite(List<OracleUser> all_users, string BatchNo)
        {
            try
            {
                // 檢查一下 如果 sqlite 檔案沒有就自動建立
                CreateSQLiteFileIfNotExists_Batch();
                CreateSQLiteFileIfNotExists_OracleUsers();
                CreateSQLiteFileIfNotExists_Links();
                CreateSQLiteFileIfNotExists_Properties();
                InsertIntoOracleUsersTable(all_users, BatchNo);
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"SaveOracleUsersIntoSQLite 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        void InsertIntoOracleUsersTable(List<OracleUser> all_users, string BatchNo)
        {
            try
            {
                //string BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                string sql = "";

                //SQLiteUtl sqlite = new SQLiteUtl(db_file);

                foreach (OracleUser user in all_users)
                {
                    sql = $@"
INSERT INTO [OracleUsers]
		([BatchNo]
		,[UserId]
		,[Username]
		,[SuspendedFlag]
		,[PersonId]
		,[PersonNumber]
		,[CredentialsEmailSentFlag]
		,[GUID]
		,[CreatedBy]
		,[CreationDate]
		,[LastUpdatedBy]
		,[LastUpdateDate])
	VALUES
		(
:0,
:1,
:2,
:3,
:4,
:5,
:6,
:7,
:8,
:9,
:10,
:11);
";
                    sqlite.ExecuteByCmd(sql,
                        BatchNo,
                        user.UserId,
                        user.Username,
                        user.SuspendedFlag,
                        user.PersonId,
                        user.PersonNumber,
                        user.CredentialsEmailSentFlag,
                        user.GUID,
                        user.CreatedBy,
                        user.CreationDate,
                        user.LastUpdatedBy,
                        user.LastUpdateDate
                        );

                    foreach (OracleApiReturnObjLink link in user.Links)
                    {
                        sql = $@"

INSERT INTO [Links]
		([BatchNo]
        ,[PersonId]
		,[PersonNumber]
		,[Rel]
		,[Href]
		,[Name]
		,[Kind])
	VALUES
		(
:0
,:1
,:2
,:3
,:4
,:5
,:6
)
";
                        sqlite.ExecuteByCmd(sql,
                                BatchNo,
                                user.PersonId,
                                user.PersonNumber,
                                link.Rel,
                                link.Href,
                                link.Name,
                                link.Kind
                            );

                        if (link.Properties != null)
                        {
                            sql = $@"
INSERT INTO [Properties]
		([BatchNo]
        ,[PersonId]
		,[PersonNumber]
		,[ChangeIndicator])
	VALUES
		(
:0
,:1
,:2
,:3
)
";
                            sqlite.ExecuteByCmd(sql,
                                 BatchNo,
                                 user.PersonId,
                                 user.PersonNumber,
                                 link.Properties.ChangeIndicator
                     );
                        }
                    }
                }


                sql = $@"
                INSERT INTO [Batch]
                		([BatchNo])
                	VALUES
                		(:0)
                ";
                int cnt = sqlite.ExecuteByCmd(sql, BatchNo);
                //int cnt = 0;

                // 留下 3 天的資料
                string B4BatchNo = DateTime.Today.AddDays(-3).ToString("yyyy-MM-dd HH:mm:ss.fff");
                sql = $@"
DELETE FROM  [Batch] WHERE [BatchNo] <= :0
";
                cnt = sqlite.ExecuteByCmd(sql, B4BatchNo);
                sql = $@"
DELETE FROM   [OracleUsers]  WHERE [BatchNo] <= :0
";
                cnt = sqlite.ExecuteByCmd(sql, B4BatchNo);
                sql = $@"
DELETE FROM  [Links] WHERE [BatchNo] <= :0
";
                cnt = sqlite.ExecuteByCmd(sql, B4BatchNo);
                sql = $@"
DELETE FROM  [Properties] WHERE [BatchNo] <= :0
";
                cnt = sqlite.ExecuteByCmd(sql, B4BatchNo);

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"InsertIntoTable 失敗:{ex.Message}{ex.InnerException}");
            }
        }



        /// <summary>
        /// 如果 sqlite 檔案沒有就自動建立
        /// </summary>
        void CreateSQLiteFileIfNotExists_Batch()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_Batch";
            try
            {

                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "Batch";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users
                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"CREATE TABLE [{table_name}] (
BatchNo nvarchar(30) 
)";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        /// <summary>
        /// 自動建立 oracle users table
        /// </summary>
        void CreateSQLiteFileIfNotExists_OracleUsers()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_OracleUsers";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "OracleUsers";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table users

                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"CREATE TABLE [{table_name}] (
BatchNo nvarchar(30) ,
UserId nvarchar(100),
Username nvarchar(100),
SuspendedFlag nvarchar(10),
PersonId nvarchar(50),
PersonNumber nvarchar(50),
CredentialsEmailSentFlag nvarchar(50),
GUID nvarchar(100),
CreatedBy nvarchar(100),
CreationDate nvarchar(50),
LastUpdatedBy nvarchar(100),
LastUpdateDate nvarchar(50)
)";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (PersonNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        /// <summary>
        /// 自動建立 oracle users links table
        /// </summary>
        void CreateSQLiteFileIfNotExists_Links()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_Links";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "Links";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    int cnt = 0;
                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"CREATE TABLE [{table_name}] (
BatchNo nvarchar(30) ,
PersonId nvarchar(50),
PersonNumber nvarchar(50),
Rel nvarchar(100),
Href nvarchar(4000),
Name nvarchar(100),
Kind nvarchar(100)
)";

                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (PersonNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }

        /// <summary>
        /// 自動建立 oracle users properties table
        /// </summary>
        void CreateSQLiteFileIfNotExists_Properties()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_Properties";
            try
            {
                db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                if (!Directory.Exists(db_path))
                {
                    Directory.CreateDirectory(db_path);
                }
                db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "Properties";
                string sql = $"SELECT * FROM sqlite_master WHERE type='table' AND name=:0 ";
                DataTable tb = null;
                bool hasNoTable = true;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, table_name))
                {
                    hasNoTable = false;
                }
                if (hasNoTable)
                {
                    // create table links
                    //string sqlite_sql = $"DROP TABLE IF EXISTS {table_name}";
                    //sqlite.ExecuteScalarA(sqlite_sql);
                    string sqlite_sql = $@"CREATE TABLE [{table_name}] (
BatchNo nvarchar(30) ,
PersonId nvarchar(50),
PersonNumber nvarchar(50),
ChangeIndicator nvarchar(4000)
)";
                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (PersonNumber);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM Batch ORDER BY BatchNo DESC";
                    string BatchNo = "";
                    int cnt = 0;
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, null))
                    {
                        if (cnt > 2)
                        {
                            BatchNo = $"{row[0]}";
                            break;
                        }
                        cnt++;
                    }
                    if (!string.IsNullOrWhiteSpace(BatchNo))
                    {
                        sql = $"DELETE FROM {table_name} WHERE BATCHNO <= :0";
                        sqlite.ExecuteByCmd(sql, BatchNo);
                    }
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }
        }



    }



}
