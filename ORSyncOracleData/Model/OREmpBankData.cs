using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class OREmpBankData
    {
        /// <summary>
        /// 員工編號
        /// </summary>
        public string EmpNo { get; set; }
        /// <summary>
        /// 員工姓名
        /// </summary>
        public string EmpName { get; set; }
        /// <summary>
        /// 身分證字號
        /// </summary>
        public string TW_ID { get; set; }
        /// <summary>
        /// 銀行帳戶戶名
        /// </summary>
        public string ACCOUNT_NAME { get; set; }
        /// <summary>
        /// 銀行帳號
        /// </summary>
        public string ACCOUNT_NUMBER { get; set; }
        /// <summary>
        /// 銀行代碼
        /// </summary>
        public string BankNumber { get; set; }
        /// <summary>
        /// 分行代碼
        /// </summary>
        public string BranchNumber { get; set; }

        /// <summary>
        /// 是否要用執行某個類型(已做過不能再做)
        /// 全做: 1,2,3,3,4,5 逗號分隔
        /// 或單做: 3
        /// </summary>
        public string ReDoStep { get; set; }


        /// <summary>
        /// 執行的人，會打給阿中 我就固定寫 JOB 就好
        /// </summary>
        public string ApplyMan { get; set; }

        /// <summary>
        /// 舊的銀行帳號
        /// </summary>
        public string OldAccountNumber { get; set; }
        /// <summary>
        /// 舊的銀行名稱
        /// </summary>
        public string OldBankName { get; set; }
        /// <summary>
        /// 舊的分行名稱
        /// </summary>
        public string OldBranchName { get; set; }
        public string BankName { get; set; }
        public string BranchName { get; set; }



    }


}
