using AQuartzJobUTL;
using BoxmanBase;
using LogMgr;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OracleNewQuitEmployee.Models;
using OracleNewQuitEmployee.ORSyncOracleData;
using OracleNewQuitEmployee.SCS;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using WinnieJob;
using System.IO.Compression;
using OracleNewQuitEmployee.ORSyncOracleData.Model.EmpAsm;
using System.Configuration;

namespace OracleNewQuitEmployee
{
    /// <summary>
    /// 備忘總覽
    /// 部門的自訂欄位2是為主管設定個人 Job Level用的
    /// 部門的自訂欄位3是設定這個部門的 oracle 上級簽核主管用的
    /// </summary>
    public class AddNewEmployees : Autojob
    {

        keroroConnectBase.keroroConn conn;

        /// <summary>
        /// 同步所有到職日在這天以後的同仁
        /// </summary>
        //public string SyncOnDutyDay { get; set; }
        public string only { get; set; }
        public string JLOG_LOGMODE { get; set; }//LOG使用 須宣告的參數
        public string 不做停用 { get; set; }//有值代表不做oracle帳號停用功能
        public string FUNC { get; set; }//LOG使用 須宣告的參數
        private LogOput log;
        //public string OldEmpNo { get; set; }
        //public string NewEmpNo { get; set; }
        public string EmpNosBycomma { get; set; }
        //string OR_PD;
        //string OR_PDCALLAPI;
        /// <summary>
        /// 建構式
        /// </summary>
        public AddNewEmployees()
        {
            only = "one";
            JLOG_LOGMODE = "5";//5 = 所有log都寫

            log = new LogOput();
            conn = new keroroConnectBase.keroroConn();


            db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
            iniPath = $@"{AppDomain.CurrentDomain.BaseDirectory}ini/";

            if (!Directory.Exists(db_path))
            {
                Directory.CreateDirectory(db_path);
            }


            try
            {
                var _83 = ConfigurationManager.AppSettings["ap53"];
                if (!string.IsNullOrWhiteSpace(_83))
                {
                    ap83位置 = _83;
                }
            }
            catch (Exception ex53)
            {
                log.WriteErrorLog($"取得 AppSettings ap53 失敗:{ex53}");
                throw;
            }

            try
            {
                var adlogin_domain_name_fileName = "adlogin_domain_name.txt";
                adlogin_domain_name = File.ReadAllText(adlogin_domain_name_fileName);
            }
            catch (Exception exadlogin)
            {

            }

        }

        ORSyncOracleData.SyncOracleDataBack sync_oracle;
        SQLiteUtl sqlite;


        string ap83位置 = "";
        string iniPath = "";


        public void ExecJob(Dictionary<string, string> datamap)
        {



            /////////////////
            //流程:
            //-.取飛騰資料
            //-.檢查ORACLE是否有此帳號
            //--沒有就建立
            //-有則查詢部門/主管/JOB LEVEL
            //--不對就更新
            //-檢查有沒有員工供應商
            //--沒有就建立
            //--有就檢查帳號是否相同
            //---不相同就更新
            //-已停職的帳號做停用
            /////////////////

            try
            {

                log.setLog(this.GetType(), datamap);
                log.WriteLog("5", "開始");




                //記錄一下 datamap 裡面到底有些什麼
                //log.WriteLog("5", "記錄一下 datamap 裡面到底有些什麼?");
                //foreach (var key in datamap.Keys)
                //{
                //    try
                //    {
                //        var str = $"{key}={datamap[key]}";
                //        log.WriteLog("5", str);
                //    }
                //    catch (Exception ex)
                //    {

                //    }
                //}
                //log.WriteLog("5", "=== end of datamap ===");


                try
                {

                    db_file = $"{db_path}OracleData.sqlite";
                    sqlite = new SQLiteUtl(db_file);
                }
                catch (Exception exsqlite)
                {
                    log.WriteErrorLog($"建立 sqlite 失敗:{exsqlite}{Environment.NewLine}{db_file}");
                    throw;
                }

                try
                {
                    sync_oracle = new ORSyncOracleData.SyncOracleDataBack(conn, log);
                    log.WriteLog("5", $"sync_oracle.Oracle_AP={sync_oracle.Oracle_AP}");
                }
                catch (Exception exSyncOracleDataBack)
                {
                    log.WriteErrorLog($"Create ORSyncOracleData.SyncOracleDataBack 失敗:{exSyncOracleDataBack}");
                    throw;
                }

                //test
                //var success1 = "";
                //var errmsg1 = "";
                //sync_oracle.AddWorkerByMiddle2("F807156", "F807156",
                //    "F807156-Name", "F807156-EngName", "",
                //    out success1, out errmsg1);

                ////TEST
                //this.EmpNosBycomma = "806452";
                //sync_oracle.AddOrUpdateManager("806452", "801833");
                //return;

                //test
                //this.EmpNosBycomma = "801524";
                //sync_oracle.Update_DefaultExpenseAccount_JobLevel_LineManager_2("801524", "POD090303", "801634", "0");

                //test
                //var oracleEmpName = "陳松柏-11";
                //var oracleEmpEngName = "Alvin-11";
                //var oracleEmail = "alvin-11@staff.pchome.com.tw";
                //var oracleNamesUpdateURL = "";
                //sync_oracle.ModifyOracleUserNameAndEmpNameWhenDifferent_2("801524", "Alvin",
                //    oracleEmpName, oracleEmpEngName, oracleEmail);
                //return;


                var 總經理員編 = "800551"; //目前預設 蔡凱文
                var 總經理員編file = $"{iniPath}總經理員編.txt";
                if (File.Exists(總經理員編file))
                {
                    var lines = File.ReadAllLines(總經理員編file);
                    if (lines.Length > 0)
                    {
                        總經理員編 = lines[0];
                    }
                }



                //要先取得信義上俊中的 oracle ap 位置, 不然打不出去
                //sync_oracle.Oracle_AP = sync_oracle.GetOracleApDomainName();
                //log.WriteLog("5", $"sync_oracle.Oracle_AP={sync_oracle.Oracle_AP}");





                // 先把ORACLE上的 user 同步回來
                // 方便等一下可以停用 oracle 上的帳號
                string BatchNo2 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                //2022/06/29
                //不要 sync user 了，需要就直接抓 
                //sync_oracle.SyncUsers(BatchNo2);


                //2022/06/29
                //不要 sync user 了，需要就直接抓 
                //執行這行之後，sync_oracle.AllUsersBySCIM 才會有資料(not null)
                //sync_oracle.GetAllUserBySCIM();


                //var nm = sync_oracle.GetAssignmentNumberByEmpNo("801524");
                //return;

                //string vsuccess = "";
                //string verrmsg = "";
                //var tmp_empno = "vickywei02";
                //sync_oracle.AddWorkerByMiddle(tmp_empno, tmp_empno, tmp_empno, tmp_empno,
                //    "vickywei@staff.pchome.com.tw", out vsuccess, out verrmsg);

                //sync_oracle.先檢查Oracle的UserName不一樣才更新(tmp_empno, tmp_empno);
                //return;


                //2022/06/29
                //不用先抓了，因為新建的也不會在這裡面
                //需要時再動態抓取
                //
                //如果有指定 EmpNo 就只做 EmpNo
                //沒有就取全部 Oracle Employees 包含 assignments
                //List<OracleEmployee2> allOracleEmps = null;
                if (!string.IsNullOrWhiteSpace(EmpNosBycomma))
                {
                    #region 廢除先抓 oracle emp
                    //bool 取得所有oracleEmployee成功否 = false;
                    //int try三次 = 1;
                    //while (try三次 <= 3)
                    //{
                    //    try三次++;
                    //    try
                    //    {
                    //        allOracleEmps = sync_oracle.GetSomeOracleEmployeesIncludeAssignments(out 取得所有oracleEmployee成功否, EmpNosBycomma);
                    //        取得所有oracleEmployee成功否 = true;
                    //        break;
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        log.WriteErrorLog($"{ex.Message}");
                    //        Thread.Sleep(3 * 1000);
                    //    }
                    //}
                    //if (!取得所有oracleEmployee成功否)
                    //{
                    //    throw new Exception("取得 oracle 部份 employee 失敗!");
                    //}

                    //log.WriteLog("5", $"取得 oracle 部份 employee 數量 = {allOracleEmps.Count}");
                    #endregion
                }
                else
                {

                    //2022/06/29
                    //不用先抓了，要就動態去抓
                    //因為新建的不會在這裡面
                    #region 廢除先抓 oracle emp
                    //-取全部Oracle Employees回來
                    //var all_oracle_emps = sync_oracle.GetAllOracleEmployees();
                    //bool 取得所有oracleEmployee成功否 = false;
                    //int try三次 = 1;
                    //while (try三次 <= 3)
                    //{
                    //    try三次++;
                    //    try
                    //    {
                    //        allOracleEmps = sync_oracle.GetAllOracleEmployeesIncludeAssignments(out 取得所有oracleEmployee成功否);
                    //        取得所有oracleEmployee成功否 = true;
                    //        break;
                    //    }
                    //    catch (Exception ex)
                    //    {
                    //        log.WriteErrorLog($"{ex.Message}");
                    //        Thread.Sleep(3 * 1000);
                    //    }
                    //}
                    //if (!取得所有oracleEmployee成功否)
                    //{
                    //    throw new Exception("取得 oracle 所有 employee 失敗!");
                    //}

                    //log.WriteLog("5", $"all_oracle_emps count = {allOracleEmps.Count}");
                    #endregion
                }

                //先準備 Job ID
                log.WriteLog("5", $"準備取得 Oracle Job ID..");
                //sync_oracle.GetOracleJobLevelCollection();
                sync_oracle.RefreshOracleJobLevelList();

                //-.取飛騰資料
                log.WriteLog("5", $"準備取得 飛騰 token");
                SCSUtils scs = new SCSUtils();
                var guid = scs.GetLoginToken();
                log.WriteLog("5", $"飛騰 Token={guid}");
                log.WriteLog("5", $"準備取得 飛騰 部門資料..");
                List<Hum0010300> depts = scs.GetDeptsDetails(guid).Where(
                    d => string.IsNullOrWhiteSpace(d.STOPDATE)
                    && d.TMP_DECCOMPANYID == "80"  //目前只做網家 子公司還沒串
                    ).ToList();
                log.WriteLog("5", $"取得 飛騰 部門資料 {depts.Count} 筆");

                // 部門代號清單
                var deptNos = (from dp in depts
                               select dp.SysViewid).ToList();


                //var 有自訂欄位3的部門們 = depts.Where(
                //    d => !string.IsNullOrWhiteSpace(d.Selfdef3)).ToList();

                log.WriteLog("5", $"準備取得 飛騰 人員資料..");
                ReportBody[] emps = scs.GetEmployees(guid);
                log.WriteLog("5", $"取得 飛騰 人員資料 {emps.Length} 筆");

                if (string.IsNullOrWhiteSpace(不做停用))
                {

                    //先做停用
                    List<ReportBody> disable_emps = (from emp in emps
                                                         //where emp.Employeeid.StartsWith("8")
                                                     where (deptNos.Contains(emp.Departid)
                                                       && emp.Jobstatus == "5")
                                                     select emp).ToList();

                    if (!string.IsNullOrWhiteSpace(EmpNosBycomma))
                    {
                        string[] 指定的員編們 = EmpNosBycomma.Split(',');
                        disable_emps = disable_emps.Where(item => 指定的員編們.Contains(item.Employeeid)).ToList();
                    }


                    log.WriteLog("5", $"準備停用離職人員..");
                    foreach (var _emp in disable_emps)
                    {
                        var success = "";
                        var errmsg = "";
                        var EmpNo = $"{_emp.Employeeid}";
                        var EmpName = $"{_emp.Employeename}";
                        log.WriteLog("5", $"-- 準備停用:{EmpNo} {EmpName} --");
                        sync_oracle.SuspendAccount(BatchNo2, EmpNo, EmpName, "false", out success, out errmsg);

                    }
                }
                else
                {
                    log.WriteLog("5", $"不做停用.");
                }

                List<ReportBody> pchome_emps = null;
                if (string.IsNullOrWhiteSpace(EmpNosBycomma))
                {
                    //以前可以用8開頭來判斷是否為網家員工
                    //但之後若有其他子公司調來網家的，他的員編也不會改變
                    //會維持他原來的員編，就不會是8開頭的
                    //此時就要用他的部門的備註來判別此人屬於哪間公司
                    //
                    pchome_emps = (from emp in emps
                                       //where emp.Employeeid.StartsWith("8")
                                   where deptNos.Contains(emp.Departid)
                                     && emp.Jobstatus != "5"
                                   select emp).ToList();
                }
                else
                {
                    pchome_emps = (from emp in emps
                                   where emp.Jobstatus != "5"
                                   select emp).ToList();

                    string[] EmpNos_ThisTime = EmpNosBycomma.Split(',');
                    pchome_emps = pchome_emps.Where(item => EmpNos_ThisTime.Contains(item.Employeeid)).ToList();

                }

                log.WriteLog("5", $"整理網家飛騰資料筆數:{pchome_emps.Count()}");

                //排序一下 我想讓號碼大的先做(倒著排)
                pchome_emps.Sort(delegate (ReportBody x, ReportBody y)
                {
                    int ix = 0;
                    int iy = 0;
                    int rst = -1;
                    try
                    {
                        ix = int.Parse(x.Employeeid);
                        iy = int.Parse(y.Employeeid);
                        if (ix > iy) return -1;
                        if (ix == iy) return 0;
                        if (ix < iy) return 1;
                    }
                    catch (Exception)
                    { }
                    return rst;
                });
                //建立帳號/相關資料
                foreach (var _emp in pchome_emps)
                {

                    var success = "";
                    var errmsg = "";

                    //飛騰的資料
                    var DeptNo = $"{_emp.Departid}";
                    var EmpNo = $"{_emp.Employeeid}";
                    var PersonNumber = EmpNo;
                    var EmpName = $"{_emp.Employeename}";
                    var EmpEngName = $"{_emp.Employeeengname}";
                    var email = $"{_emp.Psnemail}";
                    var ManagerEmpNo = "";
                    //var JobLevelName = "職員";

                    //自從要處理停用的員工後，發生如下錯誤:
                    //  工作者處理序無法啟動時就會發生這種錯誤。 這可能是因為不正確的識別或設定所造成，或是已經達到同時並存要求的上限。 
                    //應該是達到飽和了，發生在 02:00:03 - 02:00:05
                    //雖然短短2秒鐘，但是已經瞬間 略過幾十/上百個使用者了
                    //所以讓它每個user都停個0.5秒，讓iis消化一下
                    log.WriteLog("5", "暫停0.5秒，讓iis消化一下...");
                    Thread.Sleep(500);
                    log.WriteLog("5", $"-- 開始處理:{EmpNo} {EmpName} --");
                    log.WriteLog("5", $"飛騰的資料:部門代碼:{DeptNo}");
                    log.WriteLog("5", $"飛騰的資料:email:{email}");



                    try
                    {
                        //2021/10/06
                        //取得主管的邏輯，改由飛騰的部門欄位: SelfDef2=1,2,3,4 來決定
                        //例如我的主館是家康，但是家康部門的 SelfDef2=0,
                        //那就再往上找逸彬，但是逸彬的部門的 SelfDef2=0,
                        //再往上找儒鋒，儒鋒的部門的 SelfDef2=1
                        //所以我的簽核主管就是儒鋒
                        //
                        //然後再是自已的 Job Level
                        #region 決定簽核主管的邏輯
                        log.WriteLog("5", "開始決定簽核主管的邏輯");
                        // 決定簽核主管的邏輯=如果部門的自訂欄位3有值(有指定簽核主管，例:805385A)
                        // 就以自訂欄位三為準，否則以自訂欄位2為準
                        int Loop_Insurance = 0; //避免無窮迴圈

                        var Tmp_EmpNo = DeptNo;
                        while (ManagerEmpNo == "")
                        {
                            Loop_Insurance++;
                            if (Loop_Insurance > 20)
                            {
                                break;
                            }
                            var find_dept = (from dept in depts
                                             where dept.SysViewid == Tmp_EmpNo
                                             select dept).FirstOrDefault();

                            //_dept = find_dept.TMP_PDEPARTID;
                            if (find_dept != null)
                            {
                                log.WriteLog("5", $"找到飛騰部門:{find_dept.SysViewid} {find_dept.SysName}");
                                //部門主管先看部門的自訂欄位3
                                //若有值就用自訂欄位3的值
                                //若沒有就看部門的自訂欄位2的值
                                var self3 = $"{find_dept.Selfdef3}";
                                if (!string.IsNullOrWhiteSpace(self3))
                                {
                                    log.WriteLog("5", $"{EmpNo} {EmpName} 的簽核主管是 (Selfdef3) {self3}, 自訂欄位3 優先權大於 自訂欄位2");
                                    ManagerEmpNo = self3;
                                    break;
                                }
                                else
                                {
                                    log.WriteLog("5", $"自訂欄位2=[{find_dept.Selfdef2}]");
                                    var the_Manager_of_this_empno_and_Manager_s_JobLevel_in_FT = $"{find_dept.Selfdef2}".Trim();
                                    if (string.IsNullOrWhiteSpace(the_Manager_of_this_empno_and_Manager_s_JobLevel_in_FT))
                                    {
                                        log.WriteLog("5", $"{find_dept.SysViewid} {find_dept.SysName} 的 Job Level(自訂欄位2) 空白, 無簽核權限，再找上階部門!");
                                    }
                                    else
                                    {
                                        int tmpLevel = 0;
                                        if (int.TryParse(the_Manager_of_this_empno_and_Manager_s_JobLevel_in_FT, out tmpLevel))
                                        {
                                            if (tmpLevel > 0)
                                            {
                                                //ALVIN 2022/07/07
                                                //不要限定 JOB LEVEL 為 1-4 因為隨時可能會增加                                                
                                                if (find_dept.TmpManagerid != EmpNo)
                                                {
                                                    ManagerEmpNo = find_dept.TmpManagerid; //部門主管
                                                    log.WriteLog("5", $"判斷 {EmpNo} {EmpName} 的簽核主管為={ManagerEmpNo}, 找到 部門編號={Tmp_EmpNo}, 部門名稱={find_dept.SysName} 自訂欄位2(Selfdef2) = {the_Manager_of_this_empno_and_Manager_s_JobLevel_in_FT}");
                                                    break;//找到主管了可以離開迴圈了
                                                }
                                            }
                                            else if (tmpLevel == 0)
                                            {
                                                log.WriteLog("5", $"部門編號={Tmp_EmpNo}, 部門名稱={find_dept.SysName} 自訂欄位2(Selfdef2) = {the_Manager_of_this_empno_and_Manager_s_JobLevel_in_FT}");
                                            }
                                            else
                                            {
                                                log.WriteErrorLog($"自訂欄位2設定錯誤:[{find_dept.Selfdef2}]");
                                            }
                                        }
                                        else
                                        {
                                            log.WriteErrorLog($"飛騰的 {find_dept.SysViewid} {find_dept.SysName} 自訂欄位2 設定錯誤! ={the_Manager_of_this_empno_and_Manager_s_JobLevel_in_FT}");
                                        }
                                    }
                                }

                                //上階部門代碼
                                Tmp_EmpNo = find_dept.TMP_PDEPARTID;
                            }
                            else
                            {
                                log.WriteErrorLog($"沒有找到飛騰的部門:{Tmp_EmpNo}");
                                break;
                            }
                        }


                        //如果主管是空的，看看是不是總經理
                        //若是就不管他                        
                        if (string.IsNullOrWhiteSpace(ManagerEmpNo))
                        {
                            if (EmpNo == 總經理員編)
                            {
                                // 如果是總經理就算了 總經理 沒有主管
                            }
                            else
                            {
                                log.WriteErrorLog($"怎麼回事?? {EmpNo} {EmpName} 的主管竟然是空的?? ");
                            }

                        }

                        log.WriteLog("5", "結束決定簽核主管的邏輯");
                        #endregion


                        #region 把此員工飛騰上的 Job Level 抓出來
                        log.WriteLog("5", "開始把此員工飛騰上的 Job Level 抓出來");

                        //這個員工的最大 Job Level
                        var theMaxJobLevelOfThisEmpNoInFT = 0;
                        try
                        {
                            //員工的 Job Level 是依自已當主管的那個部門的自訂欄位2決定的
                            //值如下: 1 2 3 4 5
                            //JobLevelName
                            //表示這個 EmpNo 就是這個部門的主管
                            //那就可以查看他的 Job Level
                            //1 室/處長
                            //2 部長
                            //3 營運長
                            //4 執行長/總經理
                            //5 董事長
                            //但是不要限定只有4 可能隨時會增加

                            //這個員工當部門主管的所有部門
                            var TheDeptsInChargeOfThisEmpNo = from dp in depts
                                                              where dp.TmpManagerid == EmpNo
                                                                     && string.IsNullOrWhiteSpace(dp.STOPDATE)
                                                              group dp by dp.Selfdef2 into g
                                                              select g.OrderByDescending(s => s.Selfdef2).FirstOrDefault();

                            log.WriteLog("5", "這個員工當部門主管的所有部門:");
                            foreach (var _部門 in TheDeptsInChargeOfThisEmpNo)
                            {
                                var str = $" {EmpNo}\t{EmpName}\t{_部門.SysViewid}\t{_部門.SysName}\t自訂欄位2=[{_部門.Selfdef2}]";

                                #region 記錄這個員工的 Job Level 用來判斷要不要在新建帳號時停用
                                var Selfdef2 = _部門.Selfdef2;
                                int tmpint = 0;
                                if (int.TryParse(Selfdef2, out tmpint))
                                {
                                    if (tmpint > theMaxJobLevelOfThisEmpNoInFT)
                                    {
                                        theMaxJobLevelOfThisEmpNoInFT = tmpint;
                                    }
                                }
                                #endregion

                                log.WriteLog("5", str);
                            }
                            log.WriteLog("5", $"{EmpNo} {EmpName} 的最大 Job Level = {theMaxJobLevelOfThisEmpNoInFT}");

                            //if (TheDeptsInChargeOfThisEmpNo != null)
                            //{
                            //    var 最大的部門的jobLevel = (from 部門 in TheDeptsInChargeOfThisEmpNo
                            //                          group 部門 by 部門.Selfdef2 into g
                            //                          orderby g.Key descending
                            //                          select g.Key
                            //                          ).FirstOrDefault();

                            //    FT_EmpNo_JobLevel = 最大的部門的jobLevel.Trim();
                            //    if (string.IsNullOrWhiteSpace(FT_EmpNo_JobLevel))
                            //    {
                            //        FT_EmpNo_JobLevel = "0";
                            //    }
                            //    //if (!string.IsNullOrWhiteSpace(最大的部門的jobLevel))
                            //    //{
                            //    //    JobLevelName = GetJobLevelName(最大的部門的jobLevel);
                            //    //}
                            //}

                        }
                        catch (Exception exGetJobLevel)
                        {
                            log.WriteErrorLog($"取得飛騰上的 Job Level 失敗:{exGetJobLevel.Message}");
                        }
                        log.WriteLog("5", "結束把此員工飛騰上的 Job Level 抓出來");
                        #endregion

                        //2022/06/29
                        //不用預先抓的，要用再動態抓
                        //var oraEmp = (from oraemp in allOracleEmps
                        //              where oraemp.LastName == EmpNo
                        //              select oraemp).FirstOrDefault();

                        //-檢查ORACLE是否有此帳號
                        log.WriteLog("5", "檢查ORACLE是否有此帳號");
                        var oraEmp2 = sync_oracle.GetOracleUserByEmpNo(EmpNo);


                        //oracle 的 employee
                        var ora_PersonNumber = "";
                        var ora_PersonID = "";
                        int _cnt = oraEmp2.Resources.Count;
                        if (_cnt == 0)
                        {
                            //如果 oracle 沒有此帳號 就建立
                            log.WriteLog("5", "oracle 沒有此帳號 準備建立");
                            success = "";
                            errmsg = "";
                            //log.WriteLog("5", $"Oracle沒有Employee {EmpNo}, 準備建立Worker");
                            sync_oracle.AddWorkerByMiddle2(PersonNumber, EmpNo, EmpName, EmpEngName, email, out success, out errmsg);
                            if (success != "true")
                            {
                                throw new Exception(errmsg);
                            }
                            //sync_oracle.AddWorkerByMiddle(PersonNumber, EmpNo, EmpName, EmpEngName, email, out success, out errmsg);

                            //更新 username(oracle 登入帳號(員編)) / 員工中文姓名/英文姓名/員編
                            sync_oracle.ModifyOracleUserNameAndEmpNameWhenDifferent_2(
                                EmpNo, EmpNo, EmpName, EmpEngName, email);

                            //更新使用新的 api 更新 Job Level(Job ID) / Default Expense Account / Line Manager
                            sync_oracle.Update_DefaultExpenseAccount_JobLevel_LineManager_2(EmpNo, DeptNo, ManagerEmpNo, $"{theMaxJobLevelOfThisEmpNoInFT}");

                            ////Default Expense Account / Line Manager / Job Level
                            ////log.WriteLog("5", $"準備更新 Default Expense Account / Line Manager / Job Level");
                            //sync_oracle.Update_DefaultExpenseAccount_JobLevel_LineManager(EmpNo, DeptNo, ManagerEmpNo, $"{theMaxJobLevelOfThisEmpNoInFT}");

                            //add by alvin on 2022/04/08
                            //新帳號預設要是 disable
                            //2022/08/05 Alvin
                            //但是要先判斷是否為 Oracle 簽核主管，若是就不要把他停用掉
                            if (theMaxJobLevelOfThisEmpNoInFT > 0)
                            {
                                log.WriteLog("5", $"此員工{EmpNo} {EmpName} 的 Job Level={theMaxJobLevelOfThisEmpNoInFT} 不做新建帳號停用。");
                            }
                            else
                            {
                                log.WriteLog("5", $"{EmpNo} {EmpName} Job Level = {theMaxJobLevelOfThisEmpNoInFT}, 可以停用，準備進行新建帳號停用程序");
                                var s_success = "";
                                var s_errmsg = "";
                                sync_oracle.SuspendAccount("", EmpNo, EmpName, "false", out s_success, out s_errmsg);
                                if (s_success == "true")
                                {
                                    log.WriteLog("5", $"{EmpNo} {EmpName} 使用帳號失敗! {s_errmsg}");
                                }
                                else
                                {
                                    log.WriteErrorLog($"{EmpNo} {EmpName} 使用帳號失敗! {s_errmsg}");
                                }
                            }
                        }
                        else
                        {
                            log.WriteLog("5", "oracle 有此帳號");
                            //如果 oracle 有此帳號
                            //而且沒有被停用的話
                            //就update:
                            // UserName,
                            // 員工姓名
                            // 英文姓名
                            // Default Expanse Account, (部門)
                            // Job Level, 職等
                            // Line Manager 主管

                            //var scimUsersActive = (from user in sync_oracle.AllUsersBySCIM.Resources
                            //                       where user.UserName == EmpNo
                            //                       select user.Active).FirstOrDefault();


                            var isActive = oraEmp2.Resources[0].Active;
                            log.WriteLog("5", $"{EmpNo} {EmpName} Active = {isActive}, 檢查是否需要更新資料");

                            #region 更新員工基本資料
                            //sync_oracle.ModifyOracleUserNameAndEmpNameWhenDifferent(
                            //    EmpNo, EmpNo, EmpName, EmpEngName, email);

                            //更新 username(oracle 登入帳號(員編)) / 員工中文姓名/英文姓名/員編
                            sync_oracle.ModifyOracleUserNameAndEmpNameWhenDifferent_2(
                                EmpNo, EmpNo, EmpName, EmpEngName, email);
                            #endregion

                            //更新使用新的 api 更新 Job Level(Job ID) / Default Expense Account / Line Manager
                            sync_oracle.Update_DefaultExpenseAccount_JobLevel_LineManager_2(EmpNo, DeptNo, ManagerEmpNo, $"{theMaxJobLevelOfThisEmpNoInFT}");


                            //#region 更新 Job Level 職等
                            //log.WriteLog("5", "進入更新 Job Level 職等程序");


                            //var FT_EmpNo_JobID = sync_oracle.GetJobIDByNumberAuto($"{theMaxJobLevelOfThisEmpNoInFT}");
                            //log.WriteLog("5", $"從 oracle 取得 Job Level/Job ID對應:{theMaxJobLevelOfThisEmpNoInFT}={FT_EmpNo_JobID}");

                            ////取得 oracle 的 JobID
                            //var OraEmpReturnObj = sync_oracle.GetOracleEmployeeWithAssignmentByEmpNo(EmpNo);


                            //if (OraEmpReturnObj.Count == 0)
                            //{
                            //    var errmsg1 = $"{EmpNo} {EmpName} 呼叫 api 取得 Employee 失敗! Count=0";
                            //    log.WriteErrorLog(errmsg1);
                            //    throw new Exception(errmsg1);
                            //}

                            //var oraEmp = OraEmpReturnObj.Items[0];
                            ////取得EmployeeAssignment自已的url 以供更新使用
                            //string EmployeeAssignmentsSelfUrl = (from asm in oraEmp.Assignments
                            //                                     from link in asm.Links
                            //                                     where link.Rel == "self" && link.Name == "assignments"
                            //                                     select link.Href).FirstOrDefault();


                            //var Oracle_EmpNo_JobID = OraEmpReturnObj.Items[0].Assignments[0].JobId;
                            //log.WriteLog("5", $"取得 Oracle 上的 JobID={Oracle_EmpNo_JobID}");
                            //var Oracle_EmpNo_JobLevelName = "";
                            //var Oracle_EmpNo_Job_Level = "";
                            //foreach (var item in sync_oracle.GetJobLevelList())
                            //{
                            //    if (item.JobId == Oracle_EmpNo_JobID)
                            //    {
                            //        Oracle_EmpNo_JobLevelName = item.Name;
                            //        Oracle_EmpNo_Job_Level = item.ApprovalAuthority;
                            //        break;
                            //    }
                            //}


                            //log.WriteLog("5", $"飛騰上 {EmpNo} 的 Job Level = {theMaxJobLevelOfThisEmpNoInFT}");
                            //log.WriteLog("5", $"Oracle 上 {EmpNo} 的 Job Level = {Oracle_EmpNo_Job_Level} {Oracle_EmpNo_JobLevelName}");
                            //if ($"{theMaxJobLevelOfThisEmpNoInFT}" != Oracle_EmpNo_Job_Level)
                            //{
                            //    log.WriteLog("5", $"準備更新 {EmpNo} 的 Job Level = {theMaxJobLevelOfThisEmpNoInFT}");
                            //    var content = new
                            //    {
                            //        JobId = FT_EmpNo_JobID
                            //    };


                            //    log.WriteLog("5", $"取得EmployeeAssignmentsSelfUrl={EmployeeAssignmentsSelfUrl}");

                            //    var mr = sync_oracle.HttpPatchFromOracleAP(EmployeeAssignmentsSelfUrl, content);
                            //    if (mr.StatusCode == "200")
                            //    {
                            //        log.WriteLog("5", $"更新 JobId 成功!");
                            //        //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                            //        //string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                            //    }
                            //    else
                            //    {
                            //        log.WriteErrorLog($"更新 JobId 失敗:{mr.ErrorMessage}{mr.ReturnData}");
                            //    }
                            //}
                            //else
                            //{
                            //    log.WriteLog("5", $"Job ID 相同，不更新。");
                            //}
                            //log.WriteLog("5", "結束更新 Job Level 職等");
                            //#endregion

                            //#region 更新 Default Expense Account 部門
                            //log.WriteLog("5", "進入更新 Default Expense Account 部門");
                            ////正式環境User 的 Default Expense Account 的規則一樣比照Stage(DEV1), default account =6288099
                            ////完整  0001.< 員工所屬profit center >.< 員工所屬Department > .6288099.0000.000000000.0000000.0000
                            ////範例: 如 0001.POS.POS000000.6288099.0000.000000000.0000000.0000

                            //var deptNo = DeptNo.Substring(0, 9);
                            //var profitCenter = deptNo.Substring(0, 3);
                            //var defaultExpenseAccount = $"0001.{profitCenter}.{deptNo}.6288099.0000.000000000.0000000.0000";
                            //// 寫到這裡 比較一下 defaultExpenseAccount
                            //if (defaultExpenseAccount == oraEmp.Assignments[0].DefaultExpenseAccount)
                            //{
                            //    log.WriteLog("5", $"Default Expense Account 相同，不更新!");
                            //}
                            //else
                            //{
                            //    log.WriteLog("5", $"準備更新 {EmpNo} {EmpName} 的 Default Expense Account = {defaultExpenseAccount}");
                            //    var content = new
                            //    {
                            //        DefaultExpenseAccount = defaultExpenseAccount
                            //    };
                            //    var mr = sync_oracle.HttpPatchFromOracleAP(EmployeeAssignmentsSelfUrl, content);
                            //    if (mr.StatusCode == "200")
                            //    {
                            //        log.WriteLog("5", $"更新 DefaultExpenseAccount 成功!");
                            //        //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                            //        //string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                            //    }
                            //    else
                            //    {
                            //        log.WriteErrorLog($"更新 DefaultExpenseAccount 失敗:{mr.ErrorMessage}{mr.ReturnData}");
                            //    }
                            //}

                            //log.WriteLog("5", "結束更新 Default Expense Account 部門");
                            //#endregion

                            //#region 更新 Line Manager 主管
                            //log.WriteLog("5", $"進入更新 Line Manager 主管程序");
                            //try
                            //{

                            //    if (string.IsNullOrWhiteSpace(ManagerEmpNo))
                            //    {
                            //        log.WriteErrorLog($"{EmpNo} {EmpName} 的 Line Manager 為空!");
                            //    }
                            //    else
                            //    {
                            //        //2022/06/29
                            //        //不要用舊的
                            //        //改用動態打 api 去抓

                            //        if (ManagerEmpNo == EmpNo)
                            //        {
                            //            log.WriteLog("5", $"{EmpNo} 的主管等於自已，不更新(oracle 也不給更新)。");
                            //        }
                            //        else
                            //        {
                            //            log.WriteLog("5", $"準備更新 Line Manager 為 {ManagerEmpNo}");
                            //            //var manager = sync_oracle.GetOracleEmployeeByEmpNo(ManagerEmpNo);
                            //            var OraManagerReturnObj = sync_oracle.GetOracleEmployeeWithAssignmentByEmpNo(ManagerEmpNo);
                            //            if (OraManagerReturnObj.Count == 0)
                            //            {
                            //                //自訂欄位3 員編最後面會加'A'代表是兼任主管
                            //                //要寫下來才不會忘記
                            //                //
                            //                if (ManagerEmpNo.ToUpper().IndexOf('A') > 0)
                            //                {
                            //                    log.WriteLog("5", $"ManagerEmpNo 是自訂欄位3指定的 {ManagerEmpNo}, 需向Oracle查詢此 Employee");
                            //                    var self3_manager = sync_oracle.GetOracleEmployeeByEmpNo(ManagerEmpNo);
                            //                    if (self3_manager == null)
                            //                    {
                            //                        log.WriteErrorLog($"{EmpNo} {EmpName} 的自訂欄位3指定的簽核主管 {ManagerEmpNo} 在 Oracle 找不到! 現在建立!");

                            //                        var mEmpNo = ManagerEmpNo.ToUpper().Replace("A", "");
                            //                        var mPersonNumber = mEmpNo;
                            //                        var mEmpName = "";
                            //                        var mEmpEngName = "";
                            //                        var mEmail = "";
                            //                        var tmpgrp = (from tmpEmp in pchome_emps
                            //                                      where tmpEmp.Employeeid == mEmpNo
                            //                                      select tmpEmp).FirstOrDefault();
                            //                        if (tmpgrp != null)
                            //                        {
                            //                            mEmpName = $"{tmpgrp.Employeename}1";
                            //                            mEmpEngName = $"{tmpgrp.Employeeengname} 1";
                            //                            mEmail = tmpgrp.Psnemail;
                            //                            var succ = "";
                            //                            var errmsg3 = "";
                            //                            sync_oracle.AddWorkerByMiddle2(mPersonNumber, mEmpNo, mEmpName, mEmpEngName, mEmail,
                            //                                out succ, out errmsg3);

                            //                            if (succ == "true")
                            //                            {
                            //                                log.WriteLog("5", $"兼任主管:{mEmpNo} {mEmpName} Create Worker 成功!");

                            //                                // 準備更新 Line Manager
                            //                                var suc = false;
                            //                                var _errmsg = "";
                            //                                var _PersonID = "";
                            //                                var _AssignmentID = "";
                            //                                sync_oracle.GetEmployeePersonIDAssignmentIDByEmpNo(
                            //                                     ManagerEmpNo, out _PersonID, out _AssignmentID);
                            //                                sync_oracle.UpdateLineManager(EmployeeAssignmentsSelfUrl,
                            //                                     _PersonID, _AssignmentID,
                            //                                     out suc, out _errmsg);
                            //                                if (suc)
                            //                                {
                            //                                    log.WriteLog("5", $"更新 Line Manager 為 {ManagerEmpNo} 成功!");
                            //                                }
                            //                                if (!string.IsNullOrWhiteSpace(_errmsg))
                            //                                {
                            //                                    log.WriteErrorLog($"更新 Line Manager 失敗:{_errmsg}");
                            //                                }
                            //                            }
                            //                            else
                            //                            {
                            //                                log.WriteErrorLog($"兼任主管:{mEmpNo} {mEmpName} Create Worker 失敗! {errmsg3}");
                            //                            }
                            //                        }
                            //                        else
                            //                        {
                            //                            log.WriteErrorLog($"飛騰資料找不到 {mEmpNo} !");
                            //                        }
                            //                    }



                            //                }
                            //                else
                            //                {

                            //                    var pchome_manager_emps = from emp in emps
                            //                                              where emp.Employeeid == ManagerEmpNo
                            //                                              select emp;
                            //                    foreach (var man_emp in pchome_manager_emps)
                            //                    {
                            //                        log.WriteLog("5", $"準備建立 {EmpNo} 的主管 {ManagerEmpNo} {man_emp.Employeename}");

                            //                        sync_oracle.AddWorkerByMiddle2(EmpNo, EmpNo, man_emp.Employeename,
                            //                            man_emp.Employeeengname, man_emp.Psnemail,
                            //                            out success, out errmsg);

                            //                        var suc = false;
                            //                        var _errmsg = "";
                            //                        var _PersonID = "";
                            //                        var _AssignmentID = "";
                            //                        sync_oracle.GetEmployeePersonIDAssignmentIDByEmpNo(
                            //                             ManagerEmpNo, out _PersonID, out _AssignmentID);
                            //                        sync_oracle.UpdateLineManager(EmployeeAssignmentsSelfUrl,
                            //                             _PersonID, _AssignmentID,
                            //                             out suc, out _errmsg);
                            //                        if (!string.IsNullOrWhiteSpace(_errmsg))
                            //                        {
                            //                            log.WriteErrorLog($"更新 Line Manager 失敗:{_errmsg}");
                            //                        }
                            //                    }

                            //                }
                            //            }
                            //            else
                            //            {
                            //                var manager = OraManagerReturnObj.Items[0];
                            //                log.WriteLog("5", $"oraEmp.Assignments[0].ManagerAssignmentId={oraEmp.Assignments[0].ManagerAssignmentId}");
                            //                log.WriteLog("5", $"ManagerReturnObj.Items[0].Assignments[0].AssignmentId={OraManagerReturnObj.Items[0].Assignments[0].AssignmentId}");
                            //                if (oraEmp.Assignments[0].ManagerAssignmentId == OraManagerReturnObj.Items[0].Assignments[0].AssignmentId)
                            //                {
                            //                    log.WriteLog("5", $"Line Manager 相同，不更新!");
                            //                }
                            //                else
                            //                {
                            //                    log.WriteLog("5", $"準備更新 {EmpNo} 的 Line Manager 為 {ManagerEmpNo} ");
                            //                    var suc = false;
                            //                    var _errmsg = "";
                            //                    sync_oracle.UpdateLineManager(EmployeeAssignmentsSelfUrl,
                            //                        manager.PersonId, manager.Assignments[0].AssignmentId,
                            //                        out suc, out _errmsg);
                            //                    if (suc)
                            //                    {
                            //                        log.WriteLog("5", $"更新 Line Manager 為 {ManagerEmpNo} 成功!");
                            //                    }
                            //                    else
                            //                    {
                            //                        log.WriteErrorLog($"更新 Line Manager 為 {ManagerEmpNo} 失敗:{_errmsg}");
                            //                    }
                            //                }
                            //            }
                            //        }
                            //    }

                            //}
                            //catch (Exception exLineManager)
                            //{
                            //    log.WriteErrorLog($"執行更新 Line Manager 失敗:{exLineManager.Message}");
                            //}
                            //log.WriteLog("5", "結束更新 Line Manager 主管");
                            //#endregion



                        }
                    }
                    catch (Exception exemp1)
                    {
                        log.WriteErrorLog($"EmpNo={EmpNo}, PersonNumber={PersonNumber} 帳號建立/更新時錯誤:{exemp1.Message}");
                    }

                }
                log.WriteLog("5", "結束");
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"執行失敗:{ex.Message}");
            }


        }




        public string AddWorker(
            string PersonNumber,
            string EmpNo,
            string EmpName,
            string EmpEngName,
            string email)
        {
            string rst = "";
            try
            {

                // string r_online_date = $"{row["r_online_date"]}";  //到職日
                if (!string.IsNullOrWhiteSpace(PersonNumber))
                {
                    string success = "false";
                    string errmsg = "";
                    //ORSyncOracleData.SyncOracleDataBack sync_oracle = new ORSyncOracleData.SyncOracleDataBack(conn, log);
                    sync_oracle.AddWorkerByMiddle(PersonNumber, EmpNo, EmpName, EmpEngName, email,
                        out success, out errmsg);

                    //// 在 ORACLE 中找不到 準備新增
                    //// add user to oracle
                    //AddWorkerCommand worder = new AddWorkerCommand();
                    //worder.EmpNo = EmpNo;
                    //worder.EmpName = EmpName;
                    //worder.EngName = EmpEngName;
                    //worder.Email = string.IsNullOrWhiteSpace(email) ? "" : $"{email}@staff.pchome.com.tw";
                    ////var ret = ApiOperation.CallApi<string>("api/BoxmanOracleEmployee/BoxmanGetUsersByPagesAsync", WebRequestMethods.Http.Post, par);
                    //DateTime dt_start = DateTime.Now;
                    //ResponseData<string> ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                    //{
                    //    MethodRoute = "api/BoxmanOracleEmployee/BoxmanAddOracleWorker",
                    //    Data = worder,
                    //    MethodType = "POST",
                    //    TimeOut = 1000 * 60 * 5
                    //}
                    //);

                    //// 檢查 回傳值

                    //if (ret.StatusCode == 200)
                    //{
                    //    string rtndata = ret.Data;
                    //    string _str = Encoding.UTF8.GetString(Convert.FromBase64String(ret.Data));
                    //    // 收到 ORACLE 的回傳值
                    //    WebRequestResult ret2 = Newtonsoft.Json.JsonConvert.DeserializeObject<WebRequestResult>(_str);
                    //    if (ret2.StatusCode == HttpStatusCode.Created)
                    //    {
                    //        log.WriteLog($"{EmpNo} {EmpName} Oracle帳號建立成功:{ret2.ReturnData}");


                    //    }
                    //    else
                    //    {
                    //        throw new Exception($"{ret2.ErrorMessage}");
                    //    }
                    //}
                    //else
                    //{
                    //    throw new Exception($"{ret.ErrorMessage}");
                    //}
                }
            }
            catch (Exception exCreate)
            {
                log.WriteErrorLog($"{EmpNo} {EmpName} Oracle帳號建立失敗:{exCreate}");
            }


            return rst;
        }

        /// <summary>
        /// 用 PCA 員編建立帳號
        /// </summary>
        /// <param name="EEmpNo"></param>
        /// <param name="BatchNo"></param>
        public void CreateWorker(string EEmpNo, string BatchNo)
        {
            try
            {
                log.WriteLog("4", $"準備 create worker {EEmpNo} 到 oracle");
                //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                //if (!Directory.Exists(db_path))
                //{
                //    Directory.CreateDirectory(db_path);
                //}
                //db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);

                try
                {
                    string sql = "";
                    sql = $@"
select *
from [Supplier]
		where [BatchNo] =  :0
and R_CODE = :1
";




                    DataTable tb = null;
                    List<OREmpNoData> suppliers = new List<OREmpNoData>();
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo, EEmpNo))
                    {
                        string new_empno = $"{row["r_code"]}";
                        string old_empno = GetOldEmpNo(new_empno);
                        // 測試區用舊員編
                        // for test
                        string empno = old_empno;

                        string PersonNumber = empno;
                        //string PersonNumber = $"{row["idno"]}";
                        string EmpNo = empno;
                        string EmpName = $"{row["r_cname"]}";
                        string EmpEngName = $"{row["r_ename"]}";
                        string email = $"{row["r_email"]}";

                        try
                        {
                            string success = "false";
                            string errmsg = "";
                            //ORSyncOracleData.SyncOracleDataBack sync_oracle = new ORSyncOracleData.SyncOracleDataBack(null, log);
                            sync_oracle.AddWorkerByMiddle(PersonNumber, EmpNo, EmpName, EmpEngName, email, out success, out errmsg);


                            // string r_online_date = $"{row["r_online_date"]}";  //到職日
                            //if (!string.IsNullOrWhiteSpace(PersonNumber))
                            //{
                            //    //// 在 ORACLE 中找不到 準備新增
                            //    //// add user to oracle
                            //    //AddWorkerCommand worder = new AddWorkerCommand();
                            //    //worder.EmpNo = EmpNo;
                            //    //worder.EmpName = EmpName;
                            //    //worder.EngName = EmpEngName;
                            //    //worder.Email = string.IsNullOrWhiteSpace(email) ? "" : $"{email}@staff.pchome.com.tw";
                            //    ////var ret = ApiOperation.CallApi<string>("api/BoxmanOracleEmployee/BoxmanGetUsersByPagesAsync", WebRequestMethods.Http.Post, par);
                            //    //DateTime dt_start = DateTime.Now;
                            //    //ResponseData<string> ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                            //    //{
                            //    //    MethodRoute = "api/BoxmanOracleEmployee/BoxmanAddOracleWorker",
                            //    //    Data = worder,
                            //    //    MethodType = "POST",
                            //    //    TimeOut = 1000 * 60 * 5
                            //    //}
                            //    //);

                            //    //// 檢查 回傳值

                            //    //if (ret.StatusCode == 200)
                            //    //{
                            //    //    string rtndata = ret.Data;
                            //    //    string _str = Encoding.UTF8.GetString(Convert.FromBase64String(ret.Data));
                            //    //    // 收到 ORACLE 的回傳值
                            //    //    WebRequestResult ret2 = Newtonsoft.Json.JsonConvert.DeserializeObject<WebRequestResult>(_str);
                            //    //    if (ret2.StatusCode == HttpStatusCode.Created)
                            //    //    {
                            //    //        log.WriteLog($"{EmpNo} {EmpName} Oracle帳號建立成功:{ret2.ReturnData}");


                            //    //    }
                            //    //    else
                            //    //    {
                            //    //        throw new Exception($"{ret2.ErrorMessage}");
                            //    //    }
                            //    //}
                            //    //else
                            //    //{
                            //    //    throw new Exception($"{ret.ErrorMessage}");
                            //    //}
                            //}


                        }
                        catch (Exception exCreate)
                        {
                            log.WriteErrorLog($"{EmpNo} {EmpName} Oracle帳號建立失敗:{exCreate}");
                        }



                    }

                    //List<OREmpNoData> sups = new List<OREmpNoData>();
                    //foreach (OREmpNoData emp in suppliers)
                    //{
                    //    sups.Clear();
                    //    sups.Add(emp);
                    //    try
                    //    {
                    //        var par = new
                    //        {
                    //            Data = sups
                    //        };
                    //        var ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                    //        {
                    //            MethodRoute = "api/OREMP/EMPACC/",
                    //            Data = par,
                    //            MethodType = "POST",
                    //            TimeOut = 1000 * 60 * 10  // 5分鐘竟然不夠
                    //        });
                    //        if (ret.Success)
                    //        {
                    //            foreach (var _sup in sups)
                    //            {
                    //                log.WriteLog("5", $"{_sup.EmpNo},{_sup.EmpName} Supplier push 成功!");
                    //            }
                    //        }
                    //        else
                    //        {
                    //            foreach (var _sup in sups)
                    //            {
                    //                log.WriteLog("5", $"{_sup.EmpNo},{_sup.EmpName} Supplier push 失敗! {ret.ErrorMessage} {ret.ErrorException}");
                    //            }
                    //        }
                    //    }
                    //    catch (Exception exPush)
                    //    {
                    //        foreach (var _sup in sups)
                    //        {
                    //            log.WriteLog("5", $"{_sup.EmpNo},{_sup.EmpName} Supplier push 失敗! {exPush.Message} {exPush.InnerException}");
                    //        }
                    //    }




                    //}





                }
                catch (Exception exhr)
                {
                    log.WriteErrorLog($"create worker 失敗:{exhr.Message}{exhr.InnerException}");
                }

                //using (SQLiteConnection conn = new SQLiteConnection(string.Format("Data Source={0}", db_file)))
                //{
                //    conn.Open();
                //    SQLiteCommand cmd = new SQLiteCommand(conn);

                //    DateTime time_start = DateTime.Now;
                //    //string BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                //    conn.Close();
                //}

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"create worker Error:{ex.Message}{ex.InnerException}");
            }

        }

        /// <summary>
        /// 上正式用，員編用801234建立帳號
        /// </summary>
        /// <param name="EEmpNo"></param>
        /// <param name="BatchNo"></param>
        public void CreateWorkerNewEmpNo(string EEmpNo, string BatchNo)
        {
            try
            {
                log.WriteLog("4", $"準備 Create Worker {EEmpNo} 到 oracle");

                //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                //if (!Directory.Exists(db_path))
                //{
                //    Directory.CreateDirectory(db_path);
                //}
                //db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);


                using (SQLiteConnection conn = new SQLiteConnection(string.Format("Data Source={0}", db_file)))
                {
                    conn.Open();
                    SQLiteCommand cmd = new SQLiteCommand(conn);

                    DateTime time_start = DateTime.Now;
                    //string BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                    try
                    {
                        string sql = "";
                        sql = $@"
select *
from [Supplier]
		where [BatchNo] =  :0
and R_CODE = :1
";




                        DataTable tb = null;
                        List<OREmpNoData> suppliers = new List<OREmpNoData>();
                        foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo, EEmpNo))
                        {
                            string new_empno = $"{row["r_code"]}";
                            //string old_empno = GetOldEmpNo(new_empno);
                            // 測試區用舊員編
                            // for test
                            string empno = new_empno;

                            string PersonNumber = empno;
                            //string PersonNumber = $"{row["idno"]}";
                            string EmpNo = empno;
                            string EmpName = $"{row["r_cname"]}";
                            string EmpEngName = $"{row["r_ename"]}";
                            string email = $"{row["r_email"]}";

                            try
                            {

                                // string r_online_date = $"{row["r_online_date"]}";  //到職日
                                if (!string.IsNullOrWhiteSpace(PersonNumber))
                                {
                                    string success = "false";
                                    string errmsg = "";
                                    //ORSyncOracleData.SyncOracleDataBack sync_oracle = new ORSyncOracleData.SyncOracleDataBack(null, log);
                                    sync_oracle.AddWorkerByMiddle(PersonNumber, EmpNo, EmpName, EmpEngName, email, out success, out errmsg);

                                    //// 在 ORACLE 中找不到 準備新增
                                    //// add user to oracle
                                    //AddWorkerCommand worder = new AddWorkerCommand();
                                    //worder.EmpNo = EmpNo;
                                    //worder.EmpName = EmpName;
                                    //worder.EngName = EmpEngName;
                                    //worder.Email = string.IsNullOrWhiteSpace(email) ? "" : $"{email}@staff.pchome.com.tw";
                                    ////var ret = ApiOperation.CallApi<string>("api/BoxmanOracleEmployee/BoxmanGetUsersByPagesAsync", WebRequestMethods.Http.Post, par);
                                    //DateTime dt_start = DateTime.Now;
                                    //ResponseData<string> ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                                    //{
                                    //    MethodRoute = "api/BoxmanOracleEmployee/BoxmanAddOracleWorker",
                                    //    Data = worder,
                                    //    MethodType = "POST",
                                    //    TimeOut = 1000 * 60 * 5
                                    //}
                                    //);

                                    //// 檢查 回傳值

                                    //if (ret.StatusCode == 200)
                                    //{
                                    //    string rtndata = ret.Data;
                                    //    string _str = Encoding.UTF8.GetString(Convert.FromBase64String(ret.Data));
                                    //    // 收到 ORACLE 的回傳值
                                    //    WebRequestResult ret2 = Newtonsoft.Json.JsonConvert.DeserializeObject<WebRequestResult>(_str);
                                    //    if (ret2.StatusCode == HttpStatusCode.Created)
                                    //    {
                                    //        log.WriteLog($"{EmpNo} {EmpName} Oracle帳號建立成功:{ret2.ReturnData}");


                                    //    }
                                    //    else
                                    //    {
                                    //        throw new Exception($"{ret2.ErrorMessage}");
                                    //    }
                                    //}
                                    //else
                                    //{
                                    //    throw new Exception($"{ret.ErrorMessage}");
                                    //}
                                }
                            }
                            catch (Exception exCreate)
                            {
                                log.WriteErrorLog($"{EmpNo} {EmpName} Oracle帳號建立失敗:{exCreate}");
                            }



                        }

                        //List<OREmpNoData> sups = new List<OREmpNoData>();
                        //foreach (OREmpNoData emp in suppliers)
                        //{
                        //    sups.Clear();
                        //    sups.Add(emp);
                        //    try
                        //    {
                        //        var par = new
                        //        {
                        //            Data = sups
                        //        };
                        //        var ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                        //        {
                        //            MethodRoute = "api/OREMP/EMPACC/",
                        //            Data = par,
                        //            MethodType = "POST",
                        //            TimeOut = 1000 * 60 * 10  // 5分鐘竟然不夠
                        //        });
                        //        if (ret.Success)
                        //        {
                        //            foreach (var _sup in sups)
                        //            {
                        //                log.WriteLog("5", $"{_sup.EmpNo},{_sup.EmpName} Supplier push 成功!");
                        //            }
                        //        }
                        //        else
                        //        {
                        //            foreach (var _sup in sups)
                        //            {
                        //                log.WriteLog("5", $"{_sup.EmpNo},{_sup.EmpName} Supplier push 失敗! {ret.ErrorMessage} {ret.ErrorException}");
                        //            }
                        //        }
                        //    }
                        //    catch (Exception exPush)
                        //    {
                        //        foreach (var _sup in sups)
                        //        {
                        //            log.WriteLog("5", $"{_sup.EmpNo},{_sup.EmpName} Supplier push 失敗! {exPush.Message} {exPush.InnerException}");
                        //        }
                        //    }




                        //}


                        conn.Close();



                    }
                    catch (Exception exhr)
                    {
                        log.WriteErrorLog($"create worker 失敗:{exhr.Message}{exhr.InnerException}");
                    }
                }

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"create worker Error:{ex.Message}{ex.InnerException}");
            }

        }

        public void CreateWorkerNewOracleAccForTest(string EEmpNo, string EEmpName, string EEmpEngName, string Eemail)
        {
            try
            {
                log.WriteLog("4", $"準備 Create Worker {EEmpNo} 到 oracle");

                //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                //if (!Directory.Exists(db_path))
                //{
                //    Directory.CreateDirectory(db_path);
                //}
                //db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);

                try
                {

                    string new_empno = EEmpNo;
                    //string old_empno = GetOldEmpNo(new_empno);
                    // 測試區用舊員編
                    // for test
                    string empno = new_empno;

                    string PersonNumber = "";
                    //string PersonNumber = $"{row["idno"]}";
                    string EmpNo = empno;
                    string EmpName = EEmpName;
                    string EmpEngName = EEmpEngName;
                    string email = Eemail;

                    try
                    {

                        string success = "false";
                        string errmsg = "";
                        //ORSyncOracleData.SyncOracleDataBack sync_oracle = new ORSyncOracleData.SyncOracleDataBack(null, log);
                        sync_oracle.AddWorkerByMiddle(PersonNumber, EmpNo, EmpName, EmpEngName, email, out success, out errmsg);




                    }
                    catch (Exception exCreate)
                    {
                        log.WriteErrorLog($"{EmpNo} {EmpName} Oracle帳號建立失敗:{exCreate}");
                    }






                }
                catch (Exception exhr)
                {
                    log.WriteErrorLog($"create worker 失敗:{exhr.Message}{exhr.InnerException}");
                }

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"create worker Error:{ex.Message}{ex.InnerException}");
            }

        }

        /// <summary>
        /// 把Oracle 上的 supplier 同步回來
        /// </summary>
        /// <param name="BatchNo"></param>
        void SyncSupplierFromOracle(string BatchNo)
        {
            try
            {
                sync_oracle.SyncSuppliers(BatchNo);

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"把Oracle 上的 supplier 同步回來失敗:{ex.Message}{ex.InnerException}");
            }
        }

        // create oracle worker for new empno
        public void CreateWorkerIfNotExists(string BatchNo)
        {
            try
            {
                //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                //db_file = $"{db_path}OracleData.sqlite";
                string sql = $@"
SELECT R_CODE, R_CNAME, N.LASTNAME
FROM SUPPLIER S
LEFT JOIN [WorkerNames] N ON S.R_CODE = N.LASTNAME
AND S.BATCHNO = N.BATCHNO
WHERE S.BATCHNO = :0
AND N.LASTNAME IS NULL
ORDER BY R_CODE
";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                DataTable tb;
                foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, BatchNo))
                {
                    var empno = $"{row["R_CODE"]}";
                    CreateWorkerNewEmpNo(empno, BatchNo);
                }
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"CreateWorkerIfNotExists 建立 Oracle Worker 執行失敗!{ex.Message}{ex.InnerException}");
            }
        }


        //public void SuspendAccount(string BatchNo)
        //{
        //    try
        //    {

        //        string success = "";
        //        string errmsg = "";
        //        SyncFromHRFT(BatchNo);
        //        sync_oracle.SyncUsers(BatchNo);
        //        sync_oracle.SuspendAccsIfNeeds(BatchNo);

        //    }
        //    catch (Exception ex)
        //    {
        //        log.WriteErrorLog($"SuspendAccount 停用 Oracle 帳號 執行失敗!{ex.Message}{ex.InnerException}");
        //    }
        //}

        public class OREMPUploadOracleItemDetailMd
        {
            public string Source { get; set; }
            public string MailFileName { get; set; }
            public string FileBase64Str { get; set; }

            public ResultObject ResponseResult { get; set; }
        }

        public class ResultObject
        {
            public ResultObject()
            {
                this.ErrorMessages = new List<string>();
            }
            /// <summary>
            /// 執行成功 / 失敗  (自訂 預設為失敗)
            /// </summary>
            public bool Result { get; set; }


            /// <summary>
            /// 錯誤代碼 (自訂)
            /// </summary>
            public string ErrorCode { get; set; }

            /// <summary>
            /// 錯誤原因 (自訂)
            /// </summary>
            public List<string> ErrorMessages { get; set; }

            public string PaseToString()
            {
                return $"{string.Join("\r\n", this.ErrorMessages)}";
            }

        }

        public class OracleApiChkBankBranchesMd
        {
            public OracleApiChkBankBranchesMd()
            {
                this.BANK_NAME_VALRESLUT = false;
                this.BANK_NUMBER_VALRESLUT = false;
                this.BRANCH_NUMBER_VALRESLUT = false;
                this.BANK_BRANCH_NAME_VALRESLUT = false;
            }
            public string BANK_NAME { get; set; }
            public string BANK_NUMBER { get; set; }
            public string BANK_BRANCH_NAME { get; set; }
            public string BRANCH_NUMBER { get; set; }


            public bool BANK_NAME_VALRESLUT { get; set; }

            public bool BANK_NUMBER_VALRESLUT { get; set; }

            public bool BRANCH_NUMBER_VALRESLUT { get; set; }

            public bool BANK_BRANCH_NAME_VALRESLUT { get; set; }


        }

        public class GenFileResult
        {
            public List<OREMPUploadOracleItemDetailMd> Data { get; set; }
            public IEnumerable<OracleApiChkBankBranchesMd> BankData { get; set; }
            public List<string> ErrorMessages { get; set; }
        }

        void GET_SUPPLIER_FILES()
        {


            try
            {
                //從飛騰抓資料
                SCSUtils scs = new SCSUtils();
                var guid = scs.GetLoginToken();

                List<OREmpNoData> suppliers = new List<OREmpNoData>();
                ReportBody[] emps = scs.GetEmployees(guid);
                var pchome_emps = from emp in emps
                                  where emp.Employeeid.StartsWith("8")
                                  && emp.Jobstatus != "5"
                                  select emp;
                int cnt = 0;
                int index = 0;
                var folder = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                foreach (var emp in pchome_emps)
                {
                    var empno = $"{emp.Employeeid}";
                    var name = $"{emp.Employeename}";
                    var Idno = $"{emp.Idno}";
                    var Bankcode = $"{emp.Bankcode}";
                    var Salaccountid = $"{emp.Salaccountid}";
                    var Salaccountname = $"{emp.Salaccountname}";
                    var Bankbranch = $"{emp.Bankbranch}";
                    if (string.IsNullOrWhiteSpace(empno)) continue;
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (string.IsNullOrWhiteSpace(Idno)) continue;
                    if (string.IsNullOrWhiteSpace(Bankcode)) continue;
                    if (string.IsNullOrWhiteSpace(Salaccountid)) continue;
                    if (string.IsNullOrWhiteSpace(Salaccountname)) continue;
                    if (string.IsNullOrWhiteSpace(Bankbranch)) continue;

                    cnt++;
                    suppliers.Add(
                        new OREmpNoData()
                        {
                            EmpNo = $"{emp.Employeeid}", // 員編
                            EmpName = $"{emp.Employeename}", // 中文姓名
                            TW_ID = $"{emp.Idno}", // 帳戶身分證號
                            BankNumber = $"{emp.Bankcode}", // 銀行代碼
                            ACCOUNT_NUMBER = $"{emp.Salaccountid}", // 銀行收款帳號
                            ACCOUNT_NAME = $"{emp.Salaccountname}", // 銀行帳戶戶名
                            BranchNumber = $"{emp.Bankbranch}", // 銀行分行代碼
                            ApplyMan = "Job",
                            ReDoStep = ""
                        }
                        );

                    if (cnt >= 5000)
                    {
                        index++;
                        //準備產檔            
                        var par = new
                        {
                            Data = suppliers
                        };
                        //var send_model = "";
                        var ret = ApiOperation.CallApi(
                                new ApiRequestSetting()
                                {
                                    Data = par,
                                    MethodRoute = "api/OREMP/EMPACCFILES",
                                    MethodType = "POST",
                                    TimeOut = 1000 * 60 * 30
                                }
                                );
                        if (ret.Success)
                        {
                            string receive_str = ret.ResponseContent;
                            try
                            {
                                //List<OREMPUploadOracleItemDetailMd> mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<List<OREMPUploadOracleItemDetailMd>>(receive_str);
                                GenFileResult mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<GenFileResult>(receive_str);
                                //var rrr = mr2.BankData.Where(p => p.BANK_NUMBER == "011" && p.BRANCH_NUMBER == "0554").FirstOrDefault();

                                log.WriteErrorLog(string.Join(Environment.NewLine, mr2.ErrorMessages.ToArray()));
                                foreach (var file in mr2.Data)
                                {
                                    byte[] content = Convert.FromBase64String(file.FileBase64Str);
                                    var path = $"I:/Alvin/Oracle供應商Supplier/產生supplier匯入檔給顧問/{folder}";
                                    var zipFolder = $"{path}/zip/";
                                    if (!Directory.Exists(zipFolder))
                                    {
                                        Directory.CreateDirectory(zipFolder);
                                    }
                                    var sort_idx = index.ToString("D4");
                                    var zipFilename = $"{zipFolder}{sort_idx}-{file.MailFileName}";
                                    var csvFolder = $"{path}/csv/";
                                    if (!Directory.Exists(csvFolder))
                                    {
                                        Directory.CreateDirectory(csvFolder);
                                    }
                                    var tmp = $"{Path.GetFileNameWithoutExtension(file.MailFileName)}.csv";
                                    var csvFilename = $"{csvFolder}{index}-{tmp}";
                                    System.IO.File.WriteAllBytes(zipFilename, content);

                                    
                                    //FileInfo fileToDecompress = new FileInfo(zipFilename);
                                    //using (FileStream originalFileStream = fileToDecompress.OpenRead())
                                    //{
                                    //    string currentFileName = fileToDecompress.FullName;
                                    //    string newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length);
                                    //    newFileName = csvFilename;

                                    //    using (FileStream decompressedFileStream = File.Create(newFileName))
                                    //    {
                                    //        using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                                    //        {
                                    //            decompressionStream.CopyTo(decompressedFileStream);
                                    //            //Console.WriteLine($"Decompressed: {fileToDecompress.Name}");
                                    //        }
                                    //    }
                                    //}
                                }

                                //if (!string.IsNullOrWhiteSpace(mr2.ErrorMessage))
                                //{
                                //    throw new Exception(mr2.ErrorMessage);
                                //}
                                //if (string.IsNullOrWhiteSpace(mr2.ReturnData))
                                //{
                                //    throw new Exception("mr2.ReturnData is null, 伺服器回傳空白!");
                                //}
                                //MiddleReturn mr = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(mr2.ReturnData);
                                //if (mr.StatusCode == "201")
                                //{
                                //    byte[] bs64_bytes = Convert.FromBase64String(mr.ReturnData);
                                //    string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                                //    Worker worker = Newtonsoft.Json.JsonConvert.DeserializeObject<Worker>(desc_str);
                                //    log.WriteLog("5", $"AddWorkerByMiddle Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName})");
                                //    //log.WriteLog("5", $"AddWorkerByMiddle Create Worker 成功(EmpNo={EmpNo}, EmpName={EmpName}){Environment.NewLine}{desc_str}");
                                //    success = "true";

                                //}
                                //else
                                //{
                                //    throw new Exception($"{mr.StatusCode} {mr.StatusDescription} {mr.ReturnData} {mr.ErrorMessage}");
                                //}
                            }
                            catch (Exception exbs64)
                            {
                                log.WriteErrorLog($"AddWorkerByMiddle Create Worker 失敗:{exbs64.Message}{exbs64.InnerException}");
                            }
                        }
                        else
                        {
                            log.WriteErrorLog($"呼叫 api/OREMP/EMPACCFILES 失敗:{ret.ResponseContent}{ret.ErrorMessage}. {ret.ErrorException}");
                        }

                        suppliers.Clear();
                        cnt = 0;
                    }
                }



                //準備產檔            
                var par1 = new
                {
                    Data = suppliers
                };
                //var send_model = "";
                var ret1 = ApiOperation.CallApi(
                        new ApiRequestSetting()
                        {
                            Data = par1,
                            MethodRoute = "api/OREMP/EMPACCFILES",
                            MethodType = "POST",
                            TimeOut = 1000 * 60 * 120
                        }
                        );
                if (ret1.Success)
                {
                    string receive_str = ret1.ResponseContent;
                    try
                    {
                        //List<OREMPUploadOracleItemDetailMd> mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<List<OREMPUploadOracleItemDetailMd>>(receive_str);
                        GenFileResult mr2 = Newtonsoft.Json.JsonConvert.DeserializeObject<GenFileResult>(receive_str);
                        //var rrr = mr2.BankData.Where(p => p.BANK_NUMBER == "011" && p.BRANCH_NUMBER == "0554").FirstOrDefault();

                        log.WriteErrorLog(string.Join(Environment.NewLine, mr2.ErrorMessages.ToArray()));
                        foreach (var file in mr2.Data)
                        {
                            byte[] content = Convert.FromBase64String(file.FileBase64Str);
                            var path = $"I:/Alvin/Oracle供應商Supplier/產生supplier匯入檔給顧問/{folder}";
                            var zipFolder = $"{path}/zip/";
                            if (!Directory.Exists(zipFolder))
                            {
                                Directory.CreateDirectory(zipFolder);
                            }
                            //var sort_idx = index.ToString("D4");
                            var zipFilename = $"{zipFolder}{file.MailFileName}";
                            var csvFolder = $"{path}/csv/";
                            if (!Directory.Exists(csvFolder))
                            {
                                Directory.CreateDirectory(csvFolder);
                            }
                            var tmp = $"{Path.GetFileNameWithoutExtension(file.MailFileName)}.csv";
                            var csvFilename = $"{csvFolder}{index}-{tmp}";
                            System.IO.File.WriteAllBytes(zipFilename, content);

                            
                            //FileInfo fileToDecompress = new FileInfo(zipFilename);
                            //using (FileStream originalFileStream = fileToDecompress.OpenRead())
                            //{
                            //    string currentFileName = fileToDecompress.FullName;
                            //    string newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length);
                            //    newFileName = csvFilename;

                            //    using (FileStream decompressedFileStream = File.Create(newFileName))
                            //    {
                            //        using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                            //        {
                            //            decompressionStream.CopyTo(decompressedFileStream);
                            //            //Console.WriteLine($"Decompressed: {fileToDecompress.Name}");
                            //        }
                            //    }
                            //}
                        }

                    }
                    catch (Exception exbs64)
                    {
                        log.WriteErrorLog($"AddWorkerByMiddle Create Worker 失敗:{exbs64.Message}{exbs64.InnerException}");
                    }
                }
                else
                {
                    log.WriteErrorLog($"呼叫 api/OREMP/EMPACCFILES 失敗:{ret1.ResponseContent}{ret1.ErrorMessage}. {ret1.ErrorException}");
                }


            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"GET_SUPPLIER_FILES error:{ex.Message}");
            }

        }

        public class SupplierUpdateBankAccountsParameterMd
        {
            public string SupplierNumber { get; set; }
            public string BankAccountNumber { get; set; }
            public string BankName { get; set; }
            public string BankBranchName { get; set; }
            public string UpdateBankAccountName { get; set; }
            public string UpdateBankAccountNumber { get; set; }
            public string UpdateBankName { get; set; }
            public string UpdateBankBranchName { get; set; }
            public string user { get; set; }
            public string dept { get; set; }
        }
        public class ORStep5UploadApiObject
        {
            public ORStep5UploadApiObject()
            {
                supbp = new List<SupplierUpdateBankAccountsParameterMd>();
            }
            public List<SupplierUpdateBankAccountsParameterMd> supbp { get; set; }

            //public List<SupplierUpdateBankChargerBearerCodeParameterMd> subcbcp { get; set; }
        }


        void Update_Supplier_Bank_Data()
        {
            try
            {


                //從飛騰抓資料
                SCSUtils scs = new SCSUtils();
                var guid = scs.GetLoginToken();

                //List<OREmpNoData> suppliers = new List<OREmpNoData>();
                ReportBody[] emps = scs.GetEmployees(guid);
                var pchome_emps = from emp in emps
                                  where emp.Employeeid.StartsWith("8")
                                  select emp;
                foreach (var emp in pchome_emps)
                {

                    var empno = $"{emp.Employeeid}";
                    var name = $"{emp.Employeename}";
                    var Idno = $"{emp.Idno}";
                    var Bankcode = $"{emp.Bankcode}";
                    var Salaccountid = $"{emp.Salaccountid}";
                    var Salaccountname = $"{emp.Salaccountname}";
                    var Bankbranch = $"{emp.Bankbranch}";
                    if (string.IsNullOrWhiteSpace(empno)) continue;
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (string.IsNullOrWhiteSpace(Idno)) continue;
                    if (string.IsNullOrWhiteSpace(Bankcode)) continue;
                    if (string.IsNullOrWhiteSpace(Salaccountid)) continue;
                    if (string.IsNullOrWhiteSpace(Salaccountname)) continue;
                    if (string.IsNullOrWhiteSpace(Bankbranch)) continue;

                    //準備update  
                    OREmpNoData par1 = new OREmpNoData()
                    {
                        EmpNo = $"{emp.Employeeid}", // 員編
                        EmpName = $"{emp.Employeename}", // 中文姓名
                        TW_ID = $"{emp.Idno}", // 帳戶身分證號
                        BankNumber = $"{emp.Bankcode}", // 銀行代碼
                        ACCOUNT_NUMBER = $"{emp.Salaccountid}", // 銀行收款帳號
                        ACCOUNT_NAME = $"{emp.Salaccountname}", // 銀行帳戶戶名
                        BranchNumber = $"{emp.Bankbranch}", // 銀行分行代碼
                        ApplyMan = "Job",
                        ReDoStep = ""
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




            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"GET_SUPPLIER_FILES error:{ex.Message}");
            }

        }





        private string GetJobLevelName(string num)
        {
            //JobLevelName
            string rst = "職員";
            if (num == "1")
                rst = "室/處長";
            if (num == "2")
                rst = "部長";
            if (num == "3")
                rst = "營運長";
            if (num == "4")
                rst = "執行長/總經理";
            return rst;
        }



        bool get_oracle_has_this_lastname(string old_empno, string BatchNo)
        {
            bool rst = false;
            try
            {

                //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                //if (!Directory.Exists(db_path))
                //{
                //    Directory.CreateDirectory(db_path);
                //}
                //db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);

                try
                {
                    string sql = "";
                    sql = $@"
SELECT *
FROM
WORKERNAMES
WHERE LASTNAME = :0
AND 	  [BatchNo] =  :1
";




                    DataTable tb = null;
                    List<OREmpNoData> suppliers = new List<OREmpNoData>();
                    foreach (DataRow row in sqlite.QueryOkWithDataRows(out tb, sql, old_empno, BatchNo))
                    {
                        rst = true;
                    }
                }
                catch (Exception exhr)
                {
                    log.WriteErrorLog($"get_oracle_has_this_lastname 失敗:{exhr.Message}{exhr.InnerException}");
                }

                //using (SQLiteConnection conn = new SQLiteConnection(string.Format("Data Source={0}", db_file)))
                //{
                //    conn.Open();
                //    SQLiteCommand cmd = new SQLiteCommand(conn);

                //    DateTime time_start = DateTime.Now;
                //    //string BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");

                //    conn.Close();
                //}

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"get_oracle_has_this_lastname Error:{ex.Message}{ex.InnerException}");
            }

            return rst;
        }

        void SyncSupplierFromOracleByEmpNo(string EmpNo)
        {
            sync_oracle.SyncOneSupplierByEmpNo(EmpNo);
        }




        string GetOldEmpNo(string EmpNo)
        {
            string empNo = EmpNo.ToUpper();
            string rst = EmpNo;
            try
            {
                if (empNo.Substring(0, 2) != "PC")
                {

                    //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                    //if (!Directory.Exists(db_path))
                    //{
                    //    Directory.CreateDirectory(db_path);
                    //}
                    //db_file = $"{db_path}OracleData.sqlite";
                    //SQLiteUtl sqlite = new SQLiteUtl(db_file);


                    using (SQLiteConnection conn = new SQLiteConnection(string.Format("Data Source={0}", db_file)))
                    {
                        conn.Open();
                        SQLiteCommand cmd = new SQLiteCommand(conn);

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
                            conn.Close();
                        }
                        catch (Exception exhr)
                        {
                            log.WriteErrorLog($"GetOldEmpNo 失敗:{exhr.Message}{exhr.InnerException}");
                        }

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
        /// 從 oracle 同步回來
        /// </summary>
        void SyncFromOracle1(string BatchNo)
        {
            //ORSyncOracleData.SyncOracleDataBack sync_oracle = new ORSyncOracleData.SyncOracleDataBack(conn, log);
            sync_oracle.SyncWorkers(BatchNo);
            sync_oracle.SyncUsers(BatchNo);
        }

        /// <summary>
        /// 從 oracle 同步回來 => 不要用了 但是還是可以用  把程式碼 COPY 回來直接執行了
        /// </summary>
        void SyncFromOracle()
        {
            try
            {
                // 先執行同步
                Dictionary<string, string> ORSyncOracleData_datamap = null;
                string sync_dll = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData.dll";
                if (File.Exists(sync_dll))
                {

                    string ORSyncOracleData_datamap_file = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData_datamap_file.json";
                    string datamap_json = @"{""JLOG_LOGMODE"":""5"",""RunNow"":"""",""only"":""one"",""dll_name"":""ORSyncOracleData"",""job_type"":""ORSyncOracleData.SyncOracleDataBack"",""job_name"":""同步Oracle的資料回來"",""job_once"":""True""}";
                    if (File.Exists(ORSyncOracleData_datamap_file))
                    {
                        datamap_json = System.IO.File.ReadAllText(ORSyncOracleData_datamap_file);
                    }
                    ORSyncOracleData_datamap = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, string>>(datamap_json);


                    Assembly asm = Assembly.LoadFrom(sync_dll);
                    foreach (Type tp in asm.GetTypes())
                    {
                        Type _tp = tp.GetInterface("Autojob");
                        if (_tp != null)
                        {
                            object obj = asm.CreateInstance(tp.FullName);
                            Autojob job = obj as Autojob;
                            if (job != null)
                            {
                                job.ExecJob(ORSyncOracleData_datamap);
                                break;
                            }
                        }
                    }





                }



            }
            catch (Exception exo)
            {
                log.WriteErrorLog($"從 oracle 同步回來 失敗:{exo.Message}{exo.InnerException}");
            }

        }

        void SyncFromHRFT(string BatchNo)
        {
            try
            {
                //MiddleModel model = new MiddleModel();
                //model.SendingData = "";
                //model.Method = "GET";
                //ResponseData<string> ret = ApiOperation.CallApi<string>("api/ORHR/GetHRData", WebRequestMethods.Http.Get, null);

                ParseFnModel model = new ParseFnModel();
                model.FnName = "GetEmployeeForOracle";
                model.SendBase64Data = "";
                var ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                {
                    MethodRoute = "api/ORHR/ADAPI/",
                    Data = model,
                    MethodType = "POST",
                    TimeOut = 1000 * 60 * 5
                });
                if (ret.Success)
                {

                    List<SupplierModel> rtnobj = Newtonsoft.Json.JsonConvert.DeserializeObject<List<SupplierModel>>(ret.Data);
                    //SaveHREmployeesToSqlite(rtnobj, BatchNo);
                    SaveSupplierToSqlite(rtnobj, BatchNo);
                }
                else
                {
                    throw new Exception($"{ret.StatusCode} {ret.StatusDescription} {ret.ErrorMessage} {ret.ErrorException}");
                }


            }
            catch (Exception exhr)
            {
                log.WriteErrorLog($"從 HR 同步回來 失敗:{exhr.Message}{exhr.InnerException}");
            }

        }


        void SyncFromHR1(string BatchNo)
        {
            try
            {
                ResponseData<string> ret = ApiOperation.CallApi<string>("api/ORHR/GetHRData", WebRequestMethods.Http.Get, null);
                List<HREmployeeModel> rtnobj = Newtonsoft.Json.JsonConvert.DeserializeObject<List<HREmployeeModel>>(ret.Data);
                SaveHREmployeesToSqlite(rtnobj, BatchNo);
            }
            catch (Exception exhr)
            {
                log.WriteErrorLog($"從 HR 同步回來 失敗:{exhr.Message}{exhr.InnerException}");
            }

        }


        /// <summary>
        /// 取自 adlogin_domain_name.txt
        /// </summary>
        public string adlogin_domain_name { get; set; }

        /// <summary>
        /// 從 HR 同步回來
        /// </summary>
        void SyncFromHR(string BatchNo)
        {
            try
            {

                MiddleModel model = new MiddleModel();
                model.URL = $"https://{adlogin_domain_name}/ADOrg/api/Employees/";
                model.SendingData = "";
                model.Method = "GET";
                ResponseData<string> ret = ApiOperation.CallApi<string>("api/Middle/Call", WebRequestMethods.Http.Post, model);

                //ResponseData<string> ret = ApiOperation.CallApi<string>("api/ORHR/GetHRData", WebRequestMethods.Http.Get, null);

                if (string.IsNullOrWhiteSpace(ret.ErrorMessage))
                {
                    MiddleReturn middle_return = Newtonsoft.Json.JsonConvert.DeserializeObject<MiddleReturn>(ret.Data);
                    if (string.IsNullOrWhiteSpace(middle_return.ErrorMessage))
                    {

                        List<HREmployeeModel> rtnobj = Newtonsoft.Json.JsonConvert.DeserializeObject<List<HREmployeeModel>>(middle_return.ReturnData);
                        SaveHREmployeesToSqlite(rtnobj, BatchNo);

                    }
                    else
                    {
                        throw new Exception($"{middle_return.StatusCode} {middle_return.StatusDescription} {middle_return.ErrorMessage}");

                    }
                }
                else
                {
                    throw new Exception($"{ret.StatusCode} {ret.StatusDescription} {ret.ErrorMessage}");
                }

            }
            catch (Exception exhr)
            {
                log.WriteErrorLog($"從 HR 同步回來 失敗:{exhr.Message}{exhr.InnerException}");
            }
        }

        void SaveSupplierToSqlite(List<SupplierModel> emps, string BatchNo)
        {
            try
            {

                CreateSQLiteFileIfNotExists_HR_Batch();
                CreateMappingTable();
                CreateSQLiteFileIfNotExists_Supplier();
                InsertIntoSupplierTable(emps, BatchNo);
            }
            catch (Exception exhr)
            {
                log.WriteErrorLog($"SaveHREmployeesToSqlite 失敗:{exhr.Message}{exhr.InnerException}");
            }
        }

        void SaveHREmployeesToSqlite(List<HREmployeeModel> emps, string BatchNo)
        {
            try
            {

                CreateSQLiteFileIfNotExists_HR_Batch();
                CreateSQLiteFileIfNotExists_HR_Employee();
                CreateMappingTable();
                InsertIntoHRTable2(emps, BatchNo);
            }
            catch (Exception exhr)
            {
                log.WriteErrorLog($"SaveHREmployeesToSqlite 失敗:{exhr.Message}{exhr.InnerException}");
            }
        }

        /// <summary>
        /// 這個Mapping Table 做舊員編與新員編的對應 
        /// table name : OracleHRMap
        /// sqlite FILE: OracleData.sqlite
        /// </summary>
        void CreateMappingTable()
        {

            string fn_name = "CreateMappingTable";
            try
            {
                //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                //if (!Directory.Exists(db_path))
                //{
                //    Directory.CreateDirectory(db_path);
                //}
                //db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);

                string table_name = "OracleHRMap";
                db_Batch_file = table_name;
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
	[PersonNumber] nvarchar(50), 
	[OEmpNo] nvarchar(20), 
	[NEmpNo] nvarchar(20)
)
";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }

            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"{fn_name} 失敗:{ex.Message}{ex.InnerException}");
            }

        }

        void InsertIntoHRTable(List<HREmployeeModel> emps)
        {
            //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
            //if (!Directory.Exists(db_path))
            //{
            //    Directory.CreateDirectory(db_path);
            //}
            //db_file = $"{db_path}OracleData.sqlite";
            //SQLiteUtl sqlite = new SQLiteUtl(db_file);

            DateTime time_start = DateTime.Now;
            string BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            try
            {
                string sql = "";
                foreach (HREmployeeModel emp in emps)
                {
                    sql = $@"
INSERT INTO [HR_Employee]
		([BatchNo]
		,[r_code]
		,[r_cname]
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
)
	VALUES (
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
	:13
)
";

                    int cnt = sqlite.ExecuteByCmd(sql, BatchNo,
                        emp.r_code,
                        emp.r_cname,
                        emp.r_dept,
                        emp.r_degress,
                        emp.r_cell_phone,
                        emp.r_birthday,
                        emp.r_sex,
                        emp.r_email,
                        emp.r_phone_ext,
                        emp.r_skype_id,
                        emp.r_online_date,
                        emp.r_online,
                        emp.r_offline_date
                        );
                }

                //                sql = $@"
                //INSERT INTO [HR_Batch]
                //		([BatchNo])
                //	VALUES
                //		(:0)
                //";
                //                sqlite.ExecuteByCmd(sql, BatchNo);

                //                // 這裡要再加上 刪除3天以前的資料
                //                string B4BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                //                sql = $@"DELETE FROM [HR_Batch]    WHERE [BatchNo] <= :0";
                //                sqlite.ExecuteByCmd(sql, B4BatchNo);
                //                sql = $@"DELETE FROM [HR_Employee]    WHERE [BatchNo] <= :0";
                //                sqlite.ExecuteByCmd(sql, B4BatchNo);

                //                Console.WriteLine($"insert need {DateTime.Now.Subtract(time_start).TotalSeconds} sec(s).");

            }
            catch (Exception exhr)
            {
                log.WriteErrorLog($"InsertIntoHRTable 失敗:{exhr.Message}{exhr.InnerException}");
            }
        }

        /// <summary>
        /// 比 InsertIntoHRTable 快一些些
        /// </summary>
        /// <param name="emps"></param>
        void InsertIntoHRTable2(List<HREmployeeModel> emps, string BatchNo)
        {
            db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
            if (!Directory.Exists(db_path))
            {
                Directory.CreateDirectory(db_path);
            }
            db_file = $"{db_path}OracleData.sqlite";
            using (SQLiteConnection conn = new SQLiteConnection(string.Format("Data Source={0}", db_file)))
            {
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand(conn);

                DateTime time_start = DateTime.Now;
                //string BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                try
                {
                    string sql = "";
                    sql = $@"
INSERT INTO [HR_Employee]
		([BatchNo]
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
)
	VALUES (
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
)
";
                    cmd.CommandText = sql;
                    foreach (HREmployeeModel emp in emps)
                    {
                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SQLiteParameter("0", BatchNo));
                        cmd.Parameters.Add(new SQLiteParameter("1", emp.r_code));
                        cmd.Parameters.Add(new SQLiteParameter("2", emp.r_cname));
                        cmd.Parameters.Add(new SQLiteParameter("3", emp.r_ename));
                        cmd.Parameters.Add(new SQLiteParameter("4", emp.r_dept));
                        cmd.Parameters.Add(new SQLiteParameter("5", emp.r_degress));
                        cmd.Parameters.Add(new SQLiteParameter("6", emp.r_cell_phone));
                        cmd.Parameters.Add(new SQLiteParameter("7", emp.r_birthday));
                        cmd.Parameters.Add(new SQLiteParameter("8", emp.r_sex));
                        cmd.Parameters.Add(new SQLiteParameter("9", emp.r_email));
                        cmd.Parameters.Add(new SQLiteParameter("10", emp.r_phone_ext));
                        cmd.Parameters.Add(new SQLiteParameter("11", emp.r_skype_id));
                        cmd.Parameters.Add(new SQLiteParameter("12", emp.r_online_date));
                        cmd.Parameters.Add(new SQLiteParameter("13", emp.r_online));
                        cmd.Parameters.Add(new SQLiteParameter("14", emp.r_offline_date));

                        int cnt = cmd.ExecuteNonQuery();
                    }


                    //                    sql = $@"
                    //INSERT INTO [HR_Batch]
                    //		([BatchNo])
                    //	VALUES
                    //		(:BatchNo)
                    //";
                    //                    cmd.CommandText = sql;
                    //                    cmd.Parameters.Clear();
                    //                    cmd.Parameters.Add(new SQLiteParameter("BatchNo", BatchNo));
                    //                    int cnt2 = cmd.ExecuteNonQuery();

                    //                    Console.WriteLine($"insert need {DateTime.Now.Subtract(time_start).TotalSeconds} sec(s).");

                    //                    // 這裡要再加上 刪除3天以前的資料
                    //                    string B4BatchNo = DateTime.Today.AddDays(-3).ToString("yyyy-MM-dd HH:mm:ss.fff");
                    //                    sql = $@"DELETE FROM [HR_Employee] WHERE [BatchNo] <= :BatchNo";
                    //                    cmd.Parameters.Clear();
                    //                    cmd.CommandText = sql;
                    //                    cmd.Parameters.Clear();
                    //                    cmd.Parameters.Add(new SQLiteParameter("BatchNo", B4BatchNo));
                    //                    cnt2 = cmd.ExecuteNonQuery();

                    //                    sql = $@"DELETE FROM [HR_Batch]  WHERE [BatchNo] <= :BatchNo";
                    //                    cmd.Parameters.Clear();
                    //                    cmd.CommandText = sql;
                    //                    cmd.Parameters.Clear();
                    //                    cmd.Parameters.Add(new SQLiteParameter("BatchNo", B4BatchNo));
                    //                    cnt2 = cmd.ExecuteNonQuery();

                    conn.Close();



                }
                catch (Exception exhr)
                {
                    log.WriteErrorLog($"InsertIntoHRTable 失敗:{exhr.Message}{exhr.InnerException}");
                }

            }
        }


        /// <summary>
        /// 寫入 supplier Table
        /// </summary>
        /// <param name="emps"></param>
        /// <param name="BatchNo"></param>
        void InsertIntoSupplierTable(List<SupplierModel> emps, string BatchNo)
        {
            db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
            if (!Directory.Exists(db_path))
            {
                Directory.CreateDirectory(db_path);
            }
            db_file = $"{db_path}OracleData.sqlite";
            using (SQLiteConnection conn = new SQLiteConnection(string.Format("Data Source={0}", db_file)))
            {
                conn.Open();
                SQLiteCommand cmd = new SQLiteCommand(conn);

                DateTime time_start = DateTime.Now;
                //string BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                try
                {
                    string sql = "";
                    sql = $@"
INSERT INTO [Supplier]
		([BatchNo]
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
		,[workplacename]
		,[idno]
		,[salaccountname]
		,[bankcode]
		,[bankbranch]
		,[salaccountid]
		,[jobstatus]

)
	VALUES (
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
	:14,
	:15,
	:16,
	:17,
	:18,
	:19,
	:20,
	:21,
	:22,
	:23,
	:24
)
";
                    cmd.CommandText = sql;
                    int total = 0;
                    foreach (SupplierModel emp in emps)
                    {
                        // 網家都是 8 開頭
                        if (!emp.RCode.StartsWith("8"))
                        {
                            continue;
                        }

                        cmd.Parameters.Clear();
                        cmd.Parameters.Add(new SQLiteParameter("0", BatchNo));
                        cmd.Parameters.Add(new SQLiteParameter("1", emp.RCode));
                        cmd.Parameters.Add(new SQLiteParameter("2", emp.RCname));
                        cmd.Parameters.Add(new SQLiteParameter("3", emp.REname));
                        cmd.Parameters.Add(new SQLiteParameter("4", emp.RDept));
                        cmd.Parameters.Add(new SQLiteParameter("5", emp.RDegress));
                        cmd.Parameters.Add(new SQLiteParameter("6", emp.RCellPhone));
                        cmd.Parameters.Add(new SQLiteParameter("7", emp.RBirthday));
                        cmd.Parameters.Add(new SQLiteParameter("8", emp.RSex));
                        cmd.Parameters.Add(new SQLiteParameter("9", emp.REmail));
                        cmd.Parameters.Add(new SQLiteParameter("10", emp.RPhoneExt));
                        cmd.Parameters.Add(new SQLiteParameter("11", emp.RSkypeId));
                        cmd.Parameters.Add(new SQLiteParameter("12", emp.ROnlineDate));
                        cmd.Parameters.Add(new SQLiteParameter("13", emp.ROnline));
                        cmd.Parameters.Add(new SQLiteParameter("14", emp.ROfflineDate));
                        cmd.Parameters.Add(new SQLiteParameter("15", emp.Departname));
                        cmd.Parameters.Add(new SQLiteParameter("16", emp.Dutyid));
                        cmd.Parameters.Add(new SQLiteParameter("17", emp.Workplaceid));
                        cmd.Parameters.Add(new SQLiteParameter("18", emp.Workplacename));
                        cmd.Parameters.Add(new SQLiteParameter("19", emp.Idno));
                        cmd.Parameters.Add(new SQLiteParameter("20", emp.Salaccountname));
                        cmd.Parameters.Add(new SQLiteParameter("21", emp.Bankcode));
                        cmd.Parameters.Add(new SQLiteParameter("22", emp.Bankbranch));
                        cmd.Parameters.Add(new SQLiteParameter("23", emp.Salaccountid));
                        cmd.Parameters.Add(new SQLiteParameter("24", emp.Jobstatus));

                        int cnt = cmd.ExecuteNonQuery();
                        total += cnt;
                    }
                    log.WriteLog("5", $"從飛騰同步了 {total} 筆回來");


                    //                    sql = $@"
                    //INSERT INTO [HR_Batch]
                    //		([BatchNo])
                    //	VALUES
                    //		(:BatchNo)
                    //";
                    //                    cmd.CommandText = sql;
                    //                    cmd.Parameters.Clear();
                    //                    cmd.Parameters.Add(new SQLiteParameter("BatchNo", BatchNo));
                    //                    int cnt2 = cmd.ExecuteNonQuery();

                    //                    Console.WriteLine($"insert need {DateTime.Now.Subtract(time_start).TotalSeconds} sec(s).");

                    //                    // 這裡要再加上 刪除3天以前的資料
                    //                    string B4BatchNo = DateTime.Today.AddDays(-3).ToString("yyyy-MM-dd HH:mm:ss.fff");
                    //                    sql = $@"DELETE FROM [HR_Employee] WHERE [BatchNo] <= :BatchNo";
                    //                    cmd.Parameters.Clear();
                    //                    cmd.CommandText = sql;
                    //                    cmd.Parameters.Clear();
                    //                    cmd.Parameters.Add(new SQLiteParameter("BatchNo", B4BatchNo));
                    //                    cnt2 = cmd.ExecuteNonQuery();

                    //                    sql = $@"DELETE FROM [HR_Batch]  WHERE [BatchNo] <= :BatchNo";
                    //                    cmd.Parameters.Clear();
                    //                    cmd.CommandText = sql;
                    //                    cmd.Parameters.Clear();
                    //                    cmd.Parameters.Add(new SQLiteParameter("BatchNo", B4BatchNo));
                    //                    cnt2 = cmd.ExecuteNonQuery();

                    conn.Close();



                }
                catch (Exception exhr)
                {
                    log.WriteErrorLog($"InsertIntoHRTable 失敗:{exhr.Message}{exhr.InnerException}");
                }

            }
        }


        string db_path;
        string db_file;
        string db_Batch_file;
        void CreateSQLiteFileIfNotExists_HR_Batch()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_HR_Batch";
            try
            {
                //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                //if (!Directory.Exists(db_path))
                //{
                //    Directory.CreateDirectory(db_path);
                //}
                //db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);

                string table_name = "HR_Batch";
                db_Batch_file = table_name;
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
                    sql = $"SELECT BatchNo FROM {table_name} ORDER BY BatchNo DESC";
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



        void CreateSQLiteFileIfNotExists_HR_Employee()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_HR_Employee";
            try
            {
                //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                //if (!Directory.Exists(db_path))
                //{
                //    Directory.CreateDirectory(db_path);
                //}
                //db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "HR_Employee";
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
        r_code nvarchar(100),
        r_cname nvarchar(100),
        r_dept nvarchar(300),
        r_degress nvarchar(100),
        r_cell_phone nvarchar(300),
        r_birthday nvarchar(20),
        r_sex nvarchar(20),
        r_email nvarchar(100),
        r_phone_ext nvarchar(100),
        r_skype_id nvarchar(300),
        r_online_date nvarchar(20),
        r_online nvarchar(20),
        r_offline_date nvarchar(20),
        r_ename nvarchar(100)
)";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (r_code);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM {table_name} ORDER BY BatchNo DESC";
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


        void CreateSQLiteFileIfNotExists_Supplier()
        {
            string fn_name = "CreateSQLiteFileIfNotExists_Supplier";
            try
            {
                //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                //if (!Directory.Exists(db_path))
                //{
                //    Directory.CreateDirectory(db_path);
                //}
                //db_file = $"{db_path}OracleData.sqlite";
                //SQLiteUtl sqlite = new SQLiteUtl(db_file);
                string table_name = "Supplier";
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
        r_code nvarchar(100),
        r_cname nvarchar(100),
        r_ename nvarchar(100),
        r_dept nvarchar(300),
        r_degress nvarchar(100),
        r_cell_phone nvarchar(300),
        r_birthday nvarchar(20),
        r_sex nvarchar(20),
        r_email nvarchar(100),
        r_phone_ext nvarchar(100),
        r_skype_id nvarchar(300),
        r_online_date nvarchar(20),
        r_online nvarchar(20),
        r_offline_date nvarchar(20),
        departname nvarchar(100),
        dutyid nvarchar(100),
        workplaceid nvarchar(100),
        Workplacename nvarchar(100),
        idno nvarchar(20),
        salaccountname nvarchar(100),
        bankcode nvarchar(100),
        bankbranch nvarchar(100),
        salaccountid nvarchar(100),
        jobstatus nvarchar(10)
)";

                    int cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index ON {table_name} (r_code);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);
                    sqlite_sql = $"CREATE INDEX  {table_name}_index1 ON {table_name} (BatchNo);";
                    cnt = sqlite.ExecuteByCmd(sqlite_sql, null);

                }
                else
                {
                    // 如果有 就保留 3 批
                    sql = $"SELECT BatchNo FROM {table_name} ORDER BY BatchNo DESC";
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

        public class WebRequestResult
        {
            public string ReturnData { get; set; }
            public HttpStatusCode StatusCode { get; set; }
            public string StatusDescription { get; set; }
            public string ErrorMessage { get; set; }
        }


    }
}
