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
    public class old_code
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
        ORSyncOracleData.SyncOracleDataBack sync_oracle;
        SQLiteUtl sqlite;
        string db_path;
        string db_file;
        string db_Batch_file;


        string ap83位置 = "";
        string iniPath = "";

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

        public void Old_ExecJob(Dictionary<string, string> datamap)
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


                // for test 產生壓縮檔給oracle匯入員工供應商
                //GET_SUPPLIER_FILES();
                //return;


                //for test
                //var job_level1 = GetJobLevelByEmpNo("801524");

                //要先取得信義上俊中的 oracle ap 位置, 不然打不出去
                sync_oracle.Oracle_AP = sync_oracle.GetOracleApDomainName();
                log.WriteLog("5", $"sync_oracle.Oracle_AP={sync_oracle.Oracle_AP}");



                // 先把ORACLE上的 user 同步回來
                // 方便等一下可以停用 oracle 上的帳號
                string BatchNo2 = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
                sync_oracle.SyncUsers(BatchNo2);



                //var nm = sync_oracle.GetAssignmentNumberByEmpNo("801524");
                //return;

                //string vsuccess = "";
                //string verrmsg = "";
                //var tmp_empno = "vickywei02";
                //sync_oracle.AddWorkerByMiddle(tmp_empno, tmp_empno, tmp_empno, tmp_empno,
                //    "vickywei@staff.pchome.com.tw", out vsuccess, out verrmsg);

                //sync_oracle.先檢查Oracle的UserName不一樣才更新(tmp_empno, tmp_empno);
                //return;



                //如果有指定 EmpNo 就只做 EmpNo
                //沒有就取全部 Oracle Employees 包含 assignments
                List<OracleEmployee2> allOracleEmps = null;
                if (!string.IsNullOrWhiteSpace(EmpNosBycomma))
                {
                    bool 取得所有oracleEmployee成功否 = false;
                    int try三次 = 1;
                    while (try三次 <= 3)
                    {
                        try三次++;
                        try
                        {
                            allOracleEmps = sync_oracle.GetSomeOracleEmployeesIncludeAssignments(out 取得所有oracleEmployee成功否, EmpNosBycomma);
                            取得所有oracleEmployee成功否 = true;
                            break;
                        }
                        catch (Exception ex)
                        {
                            log.WriteErrorLog($"{ex.Message}");
                            Thread.Sleep(3 * 1000);
                        }
                    }
                    if (!取得所有oracleEmployee成功否)
                    {
                        throw new Exception("取得 oracle 部份 employee 失敗!");
                    }

                    log.WriteLog("5", $"取得 oracle 部份 employee 數量 = {allOracleEmps.Count}");

                }
                else
                {


                    //-取全部Oracle Employees回來
                    //var all_oracle_emps = sync_oracle.GetAllOracleEmployees();
                    bool 取得所有oracleEmployee成功否 = false;
                    int try三次 = 1;
                    while (try三次 <= 3)
                    {
                        try三次++;
                        try
                        {
                            allOracleEmps = sync_oracle.GetAllOracleEmployeesIncludeAssignments(out 取得所有oracleEmployee成功否);
                            取得所有oracleEmployee成功否 = true;
                            break;
                        }
                        catch (Exception ex)
                        {
                            log.WriteErrorLog($"{ex.Message}");
                            Thread.Sleep(3 * 1000);
                        }
                    }
                    if (!取得所有oracleEmployee成功否)
                    {
                        throw new Exception("取得 oracle 所有 employee 失敗!");
                    }

                    log.WriteLog("5", $"all_oracle_emps count = {allOracleEmps.Count}");

                }

                //先準備 Job ID
                log.WriteLog("5", $"準備取得 Oracle Job ID..");
                sync_oracle.GetOracleJobLevelCollection();

                //-.取飛騰資料
                log.WriteLog("5", $"準備取得 飛騰 token");
                SCSUtils scs = new SCSUtils();
                var guid = scs.GetLoginToken();
                log.WriteLog("5", $"準備取得 飛騰 部門資料..");
                List<Hum0010300> depts = scs.GetDeptsDetails(guid).Where(
                    d => string.IsNullOrWhiteSpace(d.STOPDATE)).ToList();
                log.WriteLog("5", $"取得 飛騰 部門資料 {depts.Count} 筆");

                //var 有自訂欄位3的部門們 = depts.Where(
                //    d => !string.IsNullOrWhiteSpace(d.Selfdef3)).ToList();

                log.WriteLog("5", $"準備取得 飛騰 人員資料..");
                ReportBody[] emps = scs.GetEmployees(guid);
                log.WriteLog("5", $"取得 飛騰 人員資料 {emps.Length} 筆");

                if (string.IsNullOrWhiteSpace(不做停用))
                {
                    //先做停用
                    List<ReportBody> disable_emps = (from emp in emps
                                                     where emp.Employeeid.StartsWith("8")
                                                       && emp.Jobstatus == "5"
                                                     //&& emp.Employeeid == "805783" //for test
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
                                   where emp.Employeeid.StartsWith("8")
                                     && emp.Jobstatus != "5"
                                   //&& emp.Employeeid == "805783" //for test
                                   select emp).ToList();
                }
                else
                {
                    pchome_emps = (from emp in emps
                                   where emp.Jobstatus != "5"
                                   //&& emp.Employeeid == "805783" //for test
                                   select emp).ToList();

                    string[] 指定的員編們 = EmpNosBycomma.Split(',');
                    pchome_emps = pchome_emps.Where(item => 指定的員編們.Contains(item.Employeeid)).ToList();

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
                    var JobLevelName = "職員";
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
                        // 決定簽核主管的邏輯=如果部門的自訂欄位3有值(有指定簽核主管，例:805385A)
                        // 就以自訂欄位三為準，否則以自訂欄位2為準
                        int 迴圈保險 = 0; //避免無窮迴圈

                        var 部門代碼暫存 = DeptNo;
                        while (ManagerEmpNo == "")
                        {
                            迴圈保險++;
                            if (迴圈保險 > 20)
                            {
                                break;
                            }
                            var find_dept = (from dept in depts
                                             where dept.SysViewid == 部門代碼暫存
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
                                    log.WriteLog("5", $"{EmpNo} {EmpName} 的簽核主管是 (Selfdef3) {self3}");
                                    ManagerEmpNo = self3;
                                    break;
                                }
                                else
                                {
                                    var 飛騰裡的JobLevel = $"{find_dept.Selfdef2}".Trim();
                                    if (string.IsNullOrWhiteSpace(飛騰裡的JobLevel))
                                    {
                                        log.WriteLog("5", $"{find_dept.SysViewid} {find_dept.SysName} 的 Job Level(自訂欄位2) 空白, 無簽核權限，再找上階部門!");
                                    }
                                    else if ((飛騰裡的JobLevel == "1") || (飛騰裡的JobLevel == "2") ||
                                         (飛騰裡的JobLevel == "3") || (飛騰裡的JobLevel == "4"))
                                    {
                                        if (find_dept.TmpManagerid != EmpNo)
                                        {
                                            ManagerEmpNo = find_dept.TmpManagerid; //部門主管
                                            log.WriteLog("5", $"判斷 {EmpNo} {EmpName} 的簽核主管為={ManagerEmpNo}");
                                            break;//找到主管了可以離開迴圈了
                                        }
                                    }
                                }

                                //上階部門代碼
                                部門代碼暫存 = find_dept.TMP_PDEPARTID;
                            }
                            else
                            {
                                log.WriteErrorLog($"沒有找到飛騰的部門:{部門代碼暫存}");
                                break;
                            }
                        }
                        #endregion
                        var 總經理員編 = "800551"; //目前預設 蔡凱文
                        var 總經理員編file = $"{AppDomain.CurrentDomain.BaseDirectory}Jobs/總經理員編.txt";
                        if (File.Exists(總經理員編file))
                        {
                            var lines = File.ReadAllLines(總經理員編file);
                            if (lines.Length > 0)
                            {
                                總經理員編 = lines[0];
                            }
                        }
                        if (string.IsNullOrWhiteSpace(ManagerEmpNo))
                        {
                            if (EmpNo == 總經理員編)
                            {
                                // 如果是總經理就算了 總經理 沒有主管
                            }
                            else
                            {
                                log.WriteErrorLog($"怎麼回事?? {EmpNo} 的主管竟然是空的?? ");
                            }

                        }


                        //-.檢查ORACLE是否有此帳號
                        var oraEmp = (from oraemp in allOracleEmps
                                      where oraemp.LastName == EmpNo
                                      select oraemp).FirstOrDefault();

                        //oracle 的 employee
                        var ora_PersonNumber = "";
                        var ora_PersonID = "";
                        if (oraEmp == null)
                        {
                            //沒有就建立
                            success = "";
                            errmsg = "";
                            //log.WriteLog("5", $"Oracle沒有Employee {EmpNo}, 準備建立Worker");
                            sync_oracle.AddWorkerByMiddle2(PersonNumber, EmpNo, EmpName, EmpEngName, email, out success, out errmsg);
                            //sync_oracle.AddWorkerByMiddle(PersonNumber, EmpNo, EmpName, EmpEngName, email, out success, out errmsg);

                            //更新 username
                            //log.WriteLog("5", $"準備更新 User Name {EmpNo}");
                            sync_oracle.ModifyOracleUserNameWhenDifferent(EmpNo, EmpNo);

                            //Default Expense Account / Line Manager / Job Level
                            //log.WriteLog("5", $"準備更新 Default Expense Account / Line Manager / Job Level");
                            //這行因為最後一個參數應該要傳 JobLevel 才行，不想改舊程式，所以直接 mark 掉
                            //sync_oracle.Update_DefaultExpenseAccount_JobLevel_LineManager(EmpNo, DeptNo, ManagerEmpNo, JobLevelName);
                        }
                        else
                        {
                            //有就update UserName, Default Expanse Account, Job Level, Line Manager
                            ora_PersonNumber = $"{oraEmp.PersonNumber}";
                            ora_PersonID = $"{oraEmp.PersonId}";

                            //這裡先判斷，不一樣才更新
                            if ((oraEmp.UserName != EmpNo) || (oraEmp.WorkEmail != email))
                            {
                                //更新 username
                                //sync_oracle.UpdateUserNameByEmployeeApi(EmpNo, EmpNo);
                                log.WriteLog("5", $"準備更新 {EmpNo} 的 User Name");
                                sync_oracle.UpdateUserNameEMailByOracleEmployee2(oraEmp, EmpNo, email);
                            }

                            //Default Expense Account / Line Manager / Job Level
                            //sync_oracle.UpdateDeptNoByEmployeeApi(EmpNo, DeptNo, ManagerEmpNo, JobLevelName);


                            #region 更新 Job Level

                            #region 飛騰上的 JOB LEVEL
                            //主管的 Job Level 是依部門的自訂欄位2決定的
                            //值如下: 1 2 3 4
                            //JobLevelName
                            //表示這個 EmpNo 就是這個部門的主管
                            //那就可以查看他的 Job Level
                            //1 室/處長
                            //2 部長
                            //3 營運長
                            //4 執行長/總經理

                            var 這個員工當部門主管的所有部門 = from dp in depts
                                                 where dp.TmpManagerid == EmpNo
                                                        && string.IsNullOrWhiteSpace(dp.STOPDATE)
                                                 group dp by dp.Selfdef2 into g
                                                 select g.OrderByDescending(s => s.Selfdef2).FirstOrDefault();
                            var _file1 = "這個員工當部門主管的所有部門.txt";
                            foreach (var _部門 in 這個員工當部門主管的所有部門)
                            {
                                var str = $"[{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}] {EmpNo}\t{EmpName}\t{_部門.SysViewid}\t{_部門.SysName}\t自訂欄位2={_部門.Selfdef2}{Environment.NewLine}";
                                System.IO.File.AppendAllText(_file1, str);
                            }

                            if (這個員工當部門主管的所有部門 != null)
                            {
                                var 最大的部門的jobLevel = (from 部門 in 這個員工當部門主管的所有部門
                                                      group 部門 by 部門.Selfdef2 into g
                                                      orderby g.Key descending
                                                      select g.Key
                                                      ).FirstOrDefault();
                                //var 最大的部門的jobLevel = (from 部門 in 這個員工當部門主管的所有部門
                                //                      group 部門 by 部門.Selfdef2 into g
                                //                      select g.OrderByDescending(s => s.Selfdef2)).FirstOrDefault();

                                if (!string.IsNullOrWhiteSpace(最大的部門的jobLevel))
                                {
                                    JobLevelName = GetJobLevelName(最大的部門的jobLevel);
                                }
                            }

                            #endregion


                            var jobIdCollection = sync_oracle.GetOracleJobLevelCollection();
                            foreach (var item in jobIdCollection)
                            {
                                var msg = $"{item.Key}={item.Value}";
                                log.WriteLog("5", $"Job Level {msg}");
                            }
                            var jobId = jobIdCollection[JobLevelName];
                            //取得EmployeeAssignment自已的url
                            string EmployeeAssignmentsSelfUrl = (from asm in oraEmp.Assignments
                                                                 from link in asm.Links
                                                                 where link.Rel == "self" && link.Name == "assignments"
                                                                 select link.Href).FirstOrDefault();

                            var Oracle上EmpNo的JobLevel = oraEmp.Assignments[0].JobId;
                            var Oracle上EmpNo的JobLevelName = "";
                            try
                            {
                                foreach (var item in jobIdCollection)
                                {
                                    if (item.Value == Oracle上EmpNo的JobLevel)
                                    {
                                        Oracle上EmpNo的JobLevelName = item.Key;
                                        break;
                                    }
                                }
                            }
                            catch { }

                            log.WriteLog("5", $"飛騰上{EmpNo} 的 Job Level = {jobId} {JobLevelName}");
                            log.WriteLog("5", $"Oracle 上 {EmpNo} 的 Job Level = {Oracle上EmpNo的JobLevel} {Oracle上EmpNo的JobLevelName}");
                            if (jobId != Oracle上EmpNo的JobLevel)
                            {
                                log.WriteLog("5", $"準備更新 {EmpNo} 的 Job Level = {jobId} {JobLevelName}");
                                var content = new
                                {
                                    JobId = jobId
                                };
                                var mr = sync_oracle.HttpPatchFromOracleAP(EmployeeAssignmentsSelfUrl, content);
                                if (mr.StatusCode == "200")
                                {
                                    log.WriteLog("5", $"更新 JobId 成功!");
                                    //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                                    //string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                                }
                                else
                                {
                                    log.WriteErrorLog($"更新 JobId 失敗:{mr.ErrorMessage}{mr.ReturnData}");
                                }
                            }
                            else
                            {
                                log.WriteLog("5", $"Job ID 相同，不更新。");
                            }
                            #endregion


                            #region 更新 Default Expense Account
                            //正式環境User 的 Default Expense Account 的規則一樣比照Stage(DEV1), default account =6288099
                            //完整  0001.< 員工所屬profit center >.< 員工所屬Department > .6288099.0000.000000000.0000000.0000
                            //範例: 如 0001.POS.POS000000.6288099.0000.000000000.0000000.0000

                            var deptNo = DeptNo.Substring(0, 9);
                            var profitCenter = deptNo.Substring(0, 3);
                            var defaultExpenseAccount = $"0001.{profitCenter}.{deptNo}.6288099.0000.000000000.0000000.0000";
                            // 寫到這裡 比較一下 defaultExpenseAccount
                            if (defaultExpenseAccount != oraEmp.Assignments[0].DefaultExpenseAccount)
                            {
                                log.WriteLog("5", $"準備更新 {EmpNo} 的 Default Expense Account = {defaultExpenseAccount}");
                                var content = new
                                {
                                    DefaultExpenseAccount = defaultExpenseAccount
                                };
                                var mr = sync_oracle.HttpPatchFromOracleAP(EmployeeAssignmentsSelfUrl, content);
                                if (mr.StatusCode == "200")
                                {
                                    log.WriteLog("5", $"更新 DefaultExpenseAccount 成功!");
                                    //byte[] bs64_bytes = Convert.FromBase64String(mr2.ReturnData);
                                    //string desc_str = Encoding.UTF8.GetString(bs64_bytes);
                                }
                                else
                                {
                                    log.WriteErrorLog($"更新 DefaultExpenseAccount 失敗:{mr.ErrorMessage}{mr.ReturnData}");
                                }
                            }


                            #endregion


                            #region 更新 Line Manager
                            log.WriteLog("5", $"準備更新 Line Manager 為 {ManagerEmpNo}");
                            var manager = (from emp in allOracleEmps
                                           where emp.LastName == ManagerEmpNo
                                           select emp).FirstOrDefault();
                            if (ManagerEmpNo == EmpNo)
                            {
                                log.WriteLog("5", $"{EmpNo} 的主管等於自已，不更新(oracle 也不給更新)。");
                            }
                            else if (manager == null)
                            {
                                if (ManagerEmpNo.ToUpper().IndexOf('A') > 0)
                                {
                                    log.WriteLog("5", $"ManagerEmpNo 是自訂欄位3指定的 {ManagerEmpNo}, 需向Oracle查詢此 Employee");
                                    var self3_manager = sync_oracle.GetOracleEmployeeByEmpNo(ManagerEmpNo);
                                    if (self3_manager == null)
                                    {
                                        log.WriteErrorLog($"{EmpNo} {EmpName} 的自訂欄位3指定的簽核主管 {ManagerEmpNo} 在 Oracle 找不到! 現在建立!");

                                        var mEmpNo = ManagerEmpNo.ToUpper().Replace("A", "");
                                        var mPersonNumber = mEmpNo;
                                        var mEmpName = "";
                                        var mEmpEngName = "";
                                        var mEmail = "";
                                        var tmpgrp = (from tmpEmp in pchome_emps
                                                      where tmpEmp.Employeeid == mEmpNo
                                                      select tmpEmp).FirstOrDefault();
                                        if (tmpgrp != null)
                                        {
                                            mEmpName = $"{tmpgrp.Employeename}1";
                                            mEmpEngName = $"{tmpgrp.Employeeengname} 1";
                                            mEmail = tmpgrp.Psnemail;
                                            var succ = "";
                                            var errmsg3 = "";
                                            sync_oracle.AddWorkerByMiddle2(mPersonNumber, mEmpNo, mEmpName, mEmpEngName, mEmail,
                                                out succ, out errmsg3);

                                            if (succ == "true")
                                            {
                                                log.WriteLog("5", $"兼任主管:{mEmpNo} {mEmpName} Create Worker 成功!");

                                                // 準備更新 Line Manager
                                                var suc = false;
                                                var _errmsg = "";
                                                var _PersonID = "";
                                                var _AssignmentID = "";
                                                sync_oracle.GetEmployeePersonIDAssignmentIDByEmpNo(
                                                     ManagerEmpNo, out _PersonID, out _AssignmentID);
                                                sync_oracle.UpdateLineManager(EmployeeAssignmentsSelfUrl,
                                                     _PersonID, _AssignmentID,
                                                     out suc, out _errmsg);
                                                if (suc)
                                                {
                                                    log.WriteLog("5", $"更新 Line Manager 為 {ManagerEmpNo} 成功!");
                                                }
                                                if (!string.IsNullOrWhiteSpace(_errmsg))
                                                {
                                                    log.WriteErrorLog($"更新 Line Manager 失敗:{_errmsg}");
                                                }
                                            }
                                            else
                                            {
                                                log.WriteErrorLog($"兼任主管:{mEmpNo} {mEmpName} Create Worker 失敗! {errmsg3}");
                                            }
                                        }
                                        else
                                        {
                                            log.WriteErrorLog($"飛騰資料找不到 {mEmpNo} !");
                                        }
                                    }



                                }
                                else
                                {

                                    var pchome_manager_emps = from emp in emps
                                                              where emp.Employeeid == ManagerEmpNo
                                                              select emp;
                                    foreach (var man_emp in pchome_manager_emps)
                                    {
                                        log.WriteLog("5", $"準備建立 {EmpNo} 的主管 {ManagerEmpNo} {man_emp.Employeename}");
                                        sync_oracle.AddWorkerByMiddle(man_emp.Employeeid, man_emp.Employeeid,
                                            man_emp.Employeename, man_emp.Employeeengname,
                                            man_emp.Psnemail, out success, out errmsg);
                                        var suc = false;
                                        var _errmsg = "";
                                        var _PersonID = "";
                                        var _AssignmentID = "";
                                        sync_oracle.GetEmployeePersonIDAssignmentIDByEmpNo(
                                             ManagerEmpNo, out _PersonID, out _AssignmentID);
                                        sync_oracle.UpdateLineManager(EmployeeAssignmentsSelfUrl,
                                             _PersonID, _AssignmentID,
                                             out suc, out _errmsg);
                                        if (!string.IsNullOrWhiteSpace(_errmsg))
                                        {
                                            log.WriteErrorLog($"更新 Line Manager 失敗:{_errmsg}");
                                        }
                                    }

                                }
                            }
                            else if (oraEmp.Assignments[0].ManagerAssignmentId != manager.Assignments[0].AssignmentId)
                            {
                                log.WriteLog("5", $"準備更新 {EmpNo} 的 Line Manager 為 {ManagerEmpNo} ");
                                var suc = false;
                                var _errmsg = "";
                                sync_oracle.UpdateLineManager(EmployeeAssignmentsSelfUrl,
                                    manager.PersonId, manager.Assignments[0].AssignmentId,
                                    out suc, out _errmsg);
                                if (suc)
                                {
                                    log.WriteLog("5", $"更新 Line Manager 為 {ManagerEmpNo} 成功!");
                                }
                                else
                                {
                                    log.WriteErrorLog($"更新 Line Manager 為 {ManagerEmpNo} 失敗:{_errmsg}");
                                }
                                //var content1 = new
                                //{
                                //    ManagerAssignmentId = manager.Assignments[0].AssignmentId, //如果在Create Worker中已獲取並保存，
                                //                                                               //则此步可不用执行；
                                //                                                               //若未保存，则需要执行查询獲取。 
                                //    ActionCode = "MANAGER_CHANGE",
                                //    ManagerId = manager.PersonId, //ManagerId为Manager這個用戶所對應的PersonId
                                //    ManagerType = "LINE_MANAGER"
                                //};
                                //var mr3 = sync_oracle.HttpPatchFromOracleAP(EmployeeAssignmentsSelfUrl, content1);
                                //if (mr3.StatusCode == "200")
                                //{
                                //    log.WriteLog("5", $"更新 Line Manager 為 {ManagerEmpNo} 成功!");
                                //    //byte[] bs64_bytes1 = Convert.FromBase64String(mr2.ReturnData);
                                //    //string desc_str1 = Encoding.UTF8.GetString(bs64_bytes1);

                                //}
                                //else
                                //{
                                //    log.WriteErrorLog($"更新 Line Manager 為 {ManagerEmpNo} 失敗:{mr3.ErrorMessage}{mr3.ReturnData}");
                                //}
                            }
                            #endregion
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


            return;
            //test get employee ->OK
            //var assignment_number = sync_oracle.GetAssignmentNumberByEmpNo("PCA01524");

            // TEST 取 JOB LEVEL ID ->OK
            //var job_id = sync_oracle.GetJobIDByChinese("執行長/總經理");


            ////TEST CREATE帳號 OK!
            //var vEmpNo = "PCA01524";
            //var vPersonNumber = vEmpNo;
            //var vEmpName = "陳松柏";
            //var vEmpEngName = "Alvin";
            //var vemail = $"alvin@staff.pchome.com.tw";
            //var vsuccess = "";
            //var verrmsg = "";
            //sync_oracle.AddWorkerByMiddle(vPersonNumber, vEmpNo, vEmpName, vEmpEngName, vemail, out vsuccess, out verrmsg);
            //return;


            // 同步 ORACLE 的 BANK DATA 回來
            //這個同步成功了 OK!
            //先 mark
            //var bsuccess = sync_oracle.SyncOracleBankData();
            //return;


            //要試試 update 員工供應商的銀行資料
            //var rst = sync_oracle.GetSupplierDataByEmpNo("PCA01524");
            ////從飛騰抓資料
            //SCSUtils scs = new SCSUtils();
            //var guid = scs.GetLoginToken();

            //List<OREmpNoData> suppliers = new List<OREmpNoData>();
            //ReportBody[] emps = scs.GetEmployees(guid);
            //var pchome_emps = from emp in emps
            //                  where emp.Employeeid=="801524"
            //                  select emp;
            //Update_Supplier_Bank_Data();



            //抓舊資料
            //string empno = "801524";
            //string old_empno = "PCA01524";

            //取得飛騰資料 ok
            //SCSUtils scs = new SCSUtils();
            //var guid = scs.GetLoginToken();
            //ReportBody[] emps = scs.GetEmployees(guid);
            //var pchome_emps = from emp in emps
            //                  where emp.Employeeid == empno
            //                  select emp;
            //var ftemp = pchome_emps.FirstOrDefault();
            //string _EmpNo = old_empno; //, // 員編
            //string _EmpName = ftemp.Employeename; //, // 中文姓名
            //string _TW_ID = ftemp.Idno; //, // 帳戶身分證號
            //string _BankNumber = ftemp.Bankcode; //, // 銀行代碼
            //string _ACCOUNT_NUMBER = ftemp.Salaccountid; //, // 銀行收款帳號
            //string _ACCOUNT_NAME = ftemp.Salaccountname; //, // 銀行帳戶戶名
            //string _BranchNumber = ftemp.Bankbranch; //, // 銀行分行代碼
            //var _rst = sync_oracle.CreateSupplier(_EmpNo, _EmpName, _TW_ID, _BankNumber, _ACCOUNT_NUMBER,
            //      _ACCOUNT_NAME, _BranchNumber, "3,4,5");

            //var supplier1 = sync_oracle.GetSupplierDataByEmpNo(old_empno);



            ////取得 oracle 銀行資料 ok
            //var banks = sync_oracle.GetOracleBankData();
            ////OracleApiChkBankBranchesMd
            ////取得新銀行資料
            //var tmp = from bank in banks
            //          where bank.BANK_NUMBER == ftemp.Bankcode
            //          && bank.BRANCH_NUMBER == ftemp.Bankbranch
            //          select bank;
            //var bk = tmp.FirstOrDefault();

            //string _BankName = bk.BANK_NAME; //, // 銀行名稱
            //string _BranchName = bk.BANK_BRANCH_NAME; //, // 銀行分行名稱

            ////取得 oracle supplier 資料
            //var supplier = sync_oracle.GetSupplierDataByEmpNo(old_empno);
            //string _OldBankName = supplier.BankName; //,  //舊的銀行名稱
            //string _OldBranchName = supplier.BranchName; //,  //舊的分行名稱
            //string _OldAccountNumber = supplier.AccountNumber; // //舊的銀行帳號
            //foreach (var emp in pchome_emps)
            //{

            //    sync_oracle.UpdateSupplierBankDataByEmpNo(
            //            _EmpNo, // = old_empno; //, // 員編
            //            _EmpName, // = ftemp.Employeename; //, // 中文姓名
            //            _TW_ID, // = ftemp.Idno; //, // 帳戶身分證號
            //            _BankNumber, // = ftemp.Bankcode; //, // 銀行代碼
            //            _BankName, // = bk.BANK_NAME; //, // 銀行名稱
            //            _ACCOUNT_NUMBER, // = ftemp.Salaccountid; //, // 銀行收款帳號
            //            _ACCOUNT_NAME, // = ftemp.Salaccountname; //, // 銀行帳戶戶名
            //            _BranchNumber, // = ftemp.Bankbranch; //, // 銀行分行代碼
            //            _BranchName, // = bk.BANK_BRANCH_NAME; //, // 銀行分行名稱
            //            _OldBankName, // = supplier.BankName; //,  //舊的銀行名稱
            //            _OldBranchName, // = supplier.BranchName; //,  //舊的分行名稱
            //            _OldAccountNumber//  = supplier.AccountNumber; // //舊的銀行帳號
            //                        );
            //}





            //if (FUNC == "GET-SUPPLIER-FILES")
            //{
            //    GET_SUPPLIER_FILES();
            //    return;
            //}

            //if (FUNC == "CreateNewAcc")
            //{
            //    CreateWorkerNewOracleAccForTest("881524", "陳松柏", "alvin", "alvin@staff.pchome.com.tw");
            //    return;
            //}



            //string BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");



            //// FOR TEST
            //BatchNo = "2021-08-06 15:39:08.088";


            // 從 HR 同步回來 
            // for test 先 mark
            //SyncFromHRFT(BatchNo);

            // 從 oracle 同步回來
            // for test 先 mark
            //SyncFromOracle1(BatchNo);

            // for test
            //BatchNo = "2021-07-27 09:15:13.912";

            // 這個不需要了，因為下面的推送會自已判斷需不需要同步
            // 需要的會自行同步
            // 從 oracle 同步 supplier 回來 給下面推送使用
            // 下面在推送之前會先檢查是否 oracle 有此 supplier
            // 若沒有就會先同步
            //SyncSupplierFromOracle(BatchNo);

            // 這個不需要了
            //FOR TEST 先MARK
            // 這個其實可以不用，因為在 push supplier to oracle 時，
            // 會看若沒有這個 user 就會做 create 
            //CreateWorkerIfNotExists(BatchNo);

            // 推送 supplier to oracle 供應商另有 DLL JOB 在跑
            //PushSupplierToOracle(BatchNo);

            // 停用 oracle 上的帳號 這個 fn 停用
            //SuspendAccount(BatchNo);

            //這些暫時不用
            //            // 最後寫入 BatchNo
            //            string sql = $@"
            //INSERT INTO [HR_Batch]
            //		([BatchNo])
            //	VALUES
            //		(:0)
            //";

            //            //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
            //            //db_file = $"{db_path}OracleData.sqlite";
            //            //SQLiteUtl sqlite = new SQLiteUtl(db_file);
            //            sqlite.ExecuteByCmd(sql, BatchNo);

            //            // 這裡要再加上 刪除3天以前的資料
            //            string B4BatchNo = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            //            sql = $@"DELETE FROM [HR_Batch]    WHERE [BatchNo] <= :0";
            //            sqlite.ExecuteByCmd(sql, B4BatchNo);
            //            sql = $@"DELETE FROM [HR_Employee]    WHERE [BatchNo] <= :0";
            //            sqlite.ExecuteByCmd(sql, B4BatchNo);

            // for test
            return;


            //this.OldEmpNo = "PCZ002";
            //this.NewEmpNo = "PCZ003";
            //sync_oracle.UpdateUserName(OldEmpNo, NewEmpNo);


            // 然後開始檢查 有沒有要新增的員工
            // HR有啟用的 ORACLE 沒有啟用就啟用 或 沒有這個員工就新增
            // HR停用的 ORACLE 沒停用 就停用


            try
            {
                // boxman api for oracle 新增員工api ok 了嗎?
                // api/BoxmanOracleEmployee/BoxmanAddOracleWorker
                // OK 了  但是 host 位置在正式區要改 現在是用 config["host"] 的方式 不準


                // SyncOnDutyDay 這個參數用來控制要往前同步到多久以前到職的同仁
                //if (string.IsNullOrWhiteSpace(SyncOnDutyDay))
                //{
                //    SyncOnDutyDay = DateTime.Today.AddDays(-7).ToString("yyyy-MM-dd");
                //}
                //SyncOnDutyDay = SyncOnDutyDay.Replace('/', '-');

                // FOR TEST
                //SyncOnDutyDay = new DateTime(2000, 1, 1).ToString("yyyy-MM-dd");


                //db_path = $@"{AppDomain.CurrentDomain.BaseDirectory}Jobs\ORSyncOracleData\";
                //db_file = $"{db_path}OracleData.sqlite";
                //string sql = "";





                //                #region HR有啟用的 ORACLE 沒有啟用就啟用 或 沒有這個員工就新增
                //                /*
                //                 "r_code": "XXXX", 員編
                //"r_cname": "XXX", 中文姓名
                //"r_dept": "918", 部門
                //"r_degress": "XXX", 職級
                //"r_cell_phone": "XXXXXX", 手機號碼
                //"r_birthday": "XXXXX", 生日
                //"r_sex": "1", 1是男 2是女
                //"r_email": "XXXX", EMAIL
                //"r_phone_ext": "8984", 分機
                //"r_skype_id": "XXXX", SKYPE ID
                //"r_online_date": "2018-10-01", 到職日
                //"r_online": "Y", Y是 在職: N是非在職 無法分辨是否留職停薪
                //"r_offline_date": 離職日


                //{ 
                //"r_dept_code": "XXXX", 部門代號
                //"r_cname": "XXXXX", 部門名稱
                //"r_belong": "916", 上層部門
                //"leader": "PCA01234" 部門主管
                //}

                //                 */

                //                // 2021/04/29 ALVIN
                //                // 這裡要先查一下 有沒有對應到舊員編
                //                // 有的話要用舊員編處理
                //                // 沒有的話就用現在的員編處理


                //                // 先檢查有沒有對應到的舊員編
                //                //                sql = $@"select *
                //                //from (
                //                //SELECT 
                //                //h.Batchno
                //                //,h.r_code
                //                //,h.r_cname
                //                //,h.r_ename
                //                //,h.r_email
                //                //,m.[OEmpNo] 
                //                //,m.[PersonNumber]
                //                //,w.firstname
                //                //,w.lastname
                //                //,w.middlenames
                //                //	FROM hr_employee h
                //                //	left JOIN [OracleHRMap] m on h.r_code = m.NEmpNo 
                //                //	left join workernames w on h.r_code = w.lastname and h.batchno = w.batchno
                //                //	where h.batchno = (select max(batchno) from batch)
                //                //	and r_online = '在職'
                //                //	) 
                //                //	where LastName is null
                //                //	order by r_code
                //                //";

                //                //    sql = $@"
                //                //                select *
                //                //                from
                //                //                (
                //                //                select *
                //                //                from
                //                //                (
                //                //     select r_code, r_cname, r_ename, r_email, 
                //                //     O.OEmpNo, 
                //                //     CASE WHEN NEMPNO > ' ' THEN NEMPNO ELSE R_CODE END MAPPING_EMPNO
                //                //                from HR_Employee HE
                //                //                LEFT JOIN OracleHRMap O ON HE.R_CODE=O.NEmpNo
                //                //                where BatchNo = (select max(batchno) from HR_batch)	
                //                //                	and r_online = '在職'
                //                //                ) HR
                //                //                  LEFT JOIN  
                //                //                (
                //                //                select PersonNumber
                //                //                from OracleUsers
                //                //                where BatchNo = (select max(batchno) from batch)
                //                //                ) O ON HR.MAPPING_EMPNO = O.PersonNumber
                //                //) F
                //                //                WHERE PersonNumber IS NULL
                //                //                ";

                //                //DataTable tb = new DataTable();
                //                //foreach (DataRow row in sqlite_utl.QueryOkWithDataRows(out tb, sql, null))
                //                //{
                //                //    string PersonNumber = $"{row["PersonNumber"]}";
                //                //    string EmpNo = $"{row["r_code"]}";
                //                //    string EmpName = $"{row["r_cname"]}";
                //                //    string r_online_date = $"{row["r_online_date"]}";  //到職日
                //                //    if (string.IsNullOrWhiteSpace(PersonNumber))
                //                //    {
                //                //        // add user to oracle
                //                //        AddWorkerCommand worder = new AddWorkerCommand();
                //                //        worder.EmpNo = EmpNo;
                //                //        worder.EmpName = EmpName;
                //                //        //var ret = ApiOperation.CallApi<string>("api/BoxmanOracleEmployee/BoxmanGetUsersByPagesAsync", WebRequestMethods.Http.Post, par);
                //                //        DateTime dt_start = DateTime.Now;
                //                //        var ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                //                //        {
                //                //            MethodRoute = "api/BoxmanOracleEmployee/BoxmanAddOracleWorker",
                //                //            Data = worder,
                //                //            MethodType = "POST",
                //                //            TimeOut = 1000 * 60 * 5
                //                //        }
                //                //        );
                //                //    }
                //                //}


                //                DataTable tb = new DataTable();
                //                using (SQLiteConnection conn = new SQLiteConnection(string.Format("Data Source={0}", db_file)))
                //                {
                //                    conn.Open();
                //                    SQLiteCommand cmd = new SQLiteCommand(conn);
                //                    cmd.Parameters.Clear();
                //                    cmd.CommandText = sql;

                //                    SQLiteDataAdapter adapt = new SQLiteDataAdapter(cmd);
                //                    int cnt = adapt.Fill(tb);
                //                    if (tb != null && tb.Rows.Count > 0)
                //                    {
                //                        foreach (DataRow row in tb.Rows)
                //                        {
                //                            string PersonNumber = $"{row["PersonNumber"]}";
                //                            string EmpNo = $"{row["r_code"]}";
                //                            string EmpName = $"{row["r_cname"]}";
                //                            string EmpEngName = $"{row["r_ename"]}";
                //                            string email = $"{row["r_email"]}";

                //                            try
                //                            {

                //                                // string r_online_date = $"{row["r_online_date"]}";  //到職日
                //                                if (string.IsNullOrWhiteSpace(PersonNumber))
                //                                {
                //                                    string success = "false";
                //                                    string errmsg = "";
                //                                    //ORSyncOracleData.SyncOracleDataBack sync_oracle = new ORSyncOracleData.SyncOracleDataBack(conn, log);
                //                                    sync_oracle.AddWorkerByMiddle(EmpNo, EmpName, EmpEngName, email, out success, out errmsg);

                //                                    //// 在 ORACLE 中找不到 準備新增
                //                                    //// add user to oracle
                //                                    //AddWorkerCommand worder = new AddWorkerCommand();
                //                                    //worder.EmpNo = EmpNo;
                //                                    //worder.EmpName = EmpName;
                //                                    //worder.EngName = EmpEngName;
                //                                    //worder.Email = string.IsNullOrWhiteSpace(email) ? "" : $"{email}@staff.pchome.com.tw";
                //                                    ////var ret = ApiOperation.CallApi<string>("api/BoxmanOracleEmployee/BoxmanGetUsersByPagesAsync", WebRequestMethods.Http.Post, par);
                //                                    //DateTime dt_start = DateTime.Now;
                //                                    //ResponseData<string> ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                //                                    //{
                //                                    //    MethodRoute = "api/BoxmanOracleEmployee/BoxmanAddOracleWorker",
                //                                    //    Data = worder,
                //                                    //    MethodType = "POST",
                //                                    //    TimeOut = 1000 * 60 * 5
                //                                    //}
                //                                    //);

                //                                    //// 檢查 回傳值

                //                                    //if (ret.StatusCode == 200)
                //                                    //{
                //                                    //    string rtndata = ret.Data;
                //                                    //    string _str = Encoding.UTF8.GetString(Convert.FromBase64String(ret.Data));
                //                                    //    // 收到 ORACLE 的回傳值
                //                                    //    WebRequestResult ret2 = Newtonsoft.Json.JsonConvert.DeserializeObject<WebRequestResult>(_str);
                //                                    //    if (ret2.StatusCode == HttpStatusCode.Created)
                //                                    //    {
                //                                    //        log.WriteLog($"{EmpNo} {EmpName} Oracle帳號建立成功:{ret2.ReturnData}");


                //                                    //    }
                //                                    //    else
                //                                    //    {
                //                                    //        throw new Exception($"{ret2.ErrorMessage}");
                //                                    //    }
                //                                    //}
                //                                    //else
                //                                    //{
                //                                    //    throw new Exception($"{ret.ErrorMessage}");
                //                                    //}
                //                                }
                //                            }
                //                            catch (Exception exCreate)
                //                            {
                //                                log.WriteErrorLog($"{EmpNo} {EmpName} Oracle帳號建立失敗:{exCreate}");
                //                            }
                //                        }
                //                    }
                //                    else
                //                    {
                //                        log.WriteLog("5", $"沒有帳號需要建立!");
                //                    }
                //                }
                //                #endregion


                #region 接下來做停用
                string sql_disable = $@"
                select *
                from (
                select *
                from
                (
                select r_code, r_cname, r_online
                from HR_Employee
                where BatchNo = (select max(batchno) from HR_batch)	
                	and r_online != '在職'

                ) HR
                  LEFT JOIN 
                (
                select *
                from OracleUsers
                where BatchNo = (select max(batchno) from batch)
                and SuspendedFlag = 'False'
                ) O ON HR.R_CODE = O.PersonNumber
                ) where PersonNumber > ' '
                ";
                DataTable tb_disable = new DataTable();
                using (SQLiteConnection conn = new SQLiteConnection(string.Format("Data Source={0}", db_file)))
                {
                    conn.Open();
                    SQLiteCommand cmd = new SQLiteCommand(conn);
                    cmd.Parameters.Clear();
                    cmd.CommandText = sql_disable;

                    SQLiteDataAdapter adapt = new SQLiteDataAdapter(cmd);
                    int cnt = adapt.Fill(tb_disable);
                    if (tb_disable != null && tb_disable.Rows.Count > 0)
                    {
                        foreach (DataRow row in tb_disable.Rows)
                        {
                            string PersonNumber = $"{row["PersonNumber"]}";
                            string EmpNo = $"{row["r_code"]}";
                            string EmpName = $"{row["r_cname"]}";

                            try
                            {

                                // 在 ORACLE 中找不到 準備新增
                                // add user to oracle
                                EnableDisableUserModel model = new EnableDisableUserModel();
                                model.EmpNo = EmpNo;
                                model.Enable = false;
                                //var ret = ApiOperation.CallApi<string>("api/BoxmanOracleEmployee/BoxmanGetUsersByPagesAsync", WebRequestMethods.Http.Post, par);
                                DateTime dt_start = DateTime.Now;

                                ////  新的寫法
                                //ApiOperation.CallApi(
                                //    new ApiRequestSetting()
                                //    {
                                //        MethodRoute = "api/BoxmanOracleEmployee/BoxmanOracleEnableDisableUser",
                                //        Data = model,
                                //        MethodType = "POST",
                                //        TimeOut = 1000 * 60 * 5
                                //    })
                                //    .ResponseHandle(
                                //        (result) =>
                                //        {
                                //            if (result.StatusCode == 200)
                                //            {
                                //                WebRequestResult ret2 = Newtonsoft.Json.JsonConvert.DeserializeObject<WebRequestResult>(result.ResponseContent);
                                //                if (ret2.StatusCode == HttpStatusCode.Created)
                                //                {
                                //                    log.WriteLog($"{EmpNo} {EmpName} Oracle帳號停用成功:{ret2.ReturnData}");
                                //                }
                                //                else
                                //                {
                                //                    throw new Exception($"{ret2.ReturnData}");
                                //                }
                                //            }
                                //            else
                                //            {
                                //                throw new Exception($"{result.ResponseContent}");
                                //            }
                                //        },
                                //        (failure) =>
                                //        {
                                //            string errmsg = $"{failure.ErrorException.Message} {failure.ErrorException.InnerException} {failure.ResponseContent}";
                                //            throw new Exception(errmsg);
                                //        }
                                //    );


                                ResponseData<string> ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                                {
                                    MethodRoute = "api/BoxmanOracleEmployee/BoxmanOracleEnableDisableUser",
                                    Data = model,
                                    MethodType = "POST",
                                    TimeOut = 1000 * 60 * 5
                                }
                                );

                                // 檢查 回傳值
                                if (ret.StatusCode == 200)
                                {
                                    try
                                    {
                                        OracleEnableDisableUserByEmpNoReturnModel ret2 = Newtonsoft.Json.JsonConvert.DeserializeObject<OracleEnableDisableUserByEmpNoReturnModel>(ret.Data);
                                        string GetUserByPersonNumberReturnString = Encoding.UTF8.GetString(Convert.FromBase64String(ret2.GetUserByPersonNumberReturnString));
                                        log.WriteLog($"Get {EmpNo} {EmpName} User ID Return:{GetUserByPersonNumberReturnString}");
                                        string rtnstr = Encoding.UTF8.GetString(Convert.FromBase64String(ret2.EnableDisableUserReturnString));
                                        log.WriteLog($"{EmpNo} {EmpName} Oracle帳號停用成功:{rtnstr}");
                                        EnableDisableReturnModel ret3 = Newtonsoft.Json.JsonConvert.DeserializeObject<EnableDisableReturnModel>(rtnstr);
                                    }
                                    catch (Exception ex)
                                    {
                                        log.WriteLog($"{EmpNo} {EmpName} Oracle帳號停用成功:{ret.Data}");
                                        log.WriteLog("5", $"解base64失敗!{ex}{Environment.NewLine}{ret.Data}");
                                    }
                                }
                                else
                                {
                                    throw new Exception($"{ret.ErrorMessage}");
                                }


                            }
                            catch (Exception exCreate)
                            {
                                log.WriteErrorLog($"{EmpNo} {EmpName} Oracle帳號停用失敗:{exCreate}");
                            }
                        }
                    }
                    else
                    {
                        log.WriteLog("5", "沒有帳號需要停用!");
                    }
                }


                #endregion


                // 啟用 用不到  暫略
                //                #region 接下來做啟用
                //                string sql_enable = $@"
                //select *
                //from (
                //select *
                //from
                //(
                //select *
                //from OracleUsers
                //where BatchNo = (select max(batchno) from batch)
                //and SuspendedFlag = 'False'
                //) O 
                //  LEFT JOIN 
                //(
                //select r_code, r_cname, r_online
                //from HR_Employee
                //where BatchNo = (select max(batchno) from HR_batch)	
                //	and r_online = '在職'

                //) HR

                //ON  O.PersonNumber = HR.R_CODE 

                //) where R_CODE > ' '
                //";
                //                DataTable tb_enable = new DataTable();
                //                using (SQLiteConnection conn = new SQLiteConnection(string.Format("Data Source={0}", db_file)))
                //                {
                //                    conn.Open();
                //                    SQLiteCommand cmd = new SQLiteCommand(conn);
                //                    cmd.Parameters.Clear();
                //                    cmd.CommandText = sql_enable;

                //                    SQLiteDataAdapter adapt = new SQLiteDataAdapter(cmd);
                //                    int cnt = adapt.Fill(tb_enable);
                //                    if (tb_enable != null && tb_enable.Rows.Count > 0)
                //                    {
                //                        foreach (DataRow row in tb_enable.Rows)
                //                        {
                //                            string PersonNumber = $"{row["PersonNumber"]}";
                //                            string UserName = $"{row["UserName"]}";

                //                            try
                //                            {

                //                                // 在 ORACLE 中找不到 準備新增
                //                                // add user to oracle
                //                                EnableDisableUserModel model = new EnableDisableUserModel();
                //                                model.EmpNo = PersonNumber;
                //                                model.Enable = true;
                //                                //var ret = ApiOperation.CallApi<string>("api/BoxmanOracleEmployee/BoxmanGetUsersByPagesAsync", WebRequestMethods.Http.Post, par);
                //                                DateTime dt_start = DateTime.Now;

                //                                ////  新的寫法
                //                                //ApiOperation.CallApi(
                //                                //    new ApiRequestSetting()
                //                                //    {
                //                                //        MethodRoute = "api/BoxmanOracleEmployee/BoxmanOracleEnableDisableUser",
                //                                //        Data = model,
                //                                //        MethodType = "POST",
                //                                //        TimeOut = 1000 * 60 * 5
                //                                //    })
                //                                //    .ResponseHandle(
                //                                //        (result) =>
                //                                //        {
                //                                //            if (result.StatusCode == 200)
                //                                //            {
                //                                //                WebRequestResult ret2 = Newtonsoft.Json.JsonConvert.DeserializeObject<WebRequestResult>(result.ResponseContent);
                //                                //                if (ret2.StatusCode == HttpStatusCode.Created)
                //                                //                {
                //                                //                    log.WriteLog($"{EmpNo} {EmpName} Oracle帳號啟用成功:{ret2.ReturnData}");
                //                                //                }
                //                                //                else
                //                                //                {
                //                                //                    throw new Exception($"{ret2.ReturnData}");
                //                                //                }
                //                                //            }
                //                                //            else
                //                                //            {
                //                                //                throw new Exception($"{result.ResponseContent}");
                //                                //            }
                //                                //        },
                //                                //        (failure) =>
                //                                //        {
                //                                //            string errmsg = $"{failure.ErrorException.Message} {failure.ErrorException.InnerException} {failure.ResponseContent}";
                //                                //            throw new Exception(errmsg);
                //                                //        }
                //                                //    );


                //                                ResponseData<string> ret = ApiOperation.CallApi<string>(new ApiRequestSetting()
                //                                {
                //                                    MethodRoute = "api/BoxmanOracleEmployee/BoxmanOracleEnableDisableUser",
                //                                    Data = model,
                //                                    MethodType = "POST",
                //                                    TimeOut = 1000 * 60 * 5
                //                                }
                //                                );

                //                                // 檢查 回傳值
                //                                if (ret.StatusCode == 200)
                //                                {
                //                                    log.WriteLog($"{UserName} (PersonNumber={PersonNumber}) Oracle帳號啟用成功:{ret.Data}");

                //                                    //EnableDisableReturnModel ret2 = Newtonsoft.Json.JsonConvert.DeserializeObject<EnableDisableReturnModel>(ret.Data);
                //                                    //WebRequestResult ret2 = Newtonsoft.Json.JsonConvert.DeserializeObject<WebRequestResult>(ret.Data);
                //                                    //if (ret2.StatusCode == HttpStatusCode.Created)
                //                                    //{
                //                                    //    log.WriteLog($"{EmpNo} {EmpName} Oracle帳號啟用成功:{ret2.ReturnData}");
                //                                    //}
                //                                    //else
                //                                    //{
                //                                    //    throw new Exception($"{ret2.ErrorMessage}");
                //                                    //}
                //                                }
                //                                else
                //                                {
                //                                    throw new Exception($"{ret.ErrorMessage}");
                //                                }


                //                            }
                //                            catch (Exception exCreate)
                //                            {
                //                                log.WriteErrorLog($"{UserName} (PersonNumber={PersonNumber}) Oracle帳號啟用失敗:{exCreate}");
                //                            }
                //                        }
                //                    }
                //                    else
                //                    {
                //                        log.WriteLog("5", "沒有帳號需要啟用!");
                //                    }
                //                }

                //                #endregion
            }
            catch (Exception ex)
            {
                log.WriteErrorLog($"ExecJob失敗:{ex.Message}{ex.InnerException}");
            }
        }


    }
}
