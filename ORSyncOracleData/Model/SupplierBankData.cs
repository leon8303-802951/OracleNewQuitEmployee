using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class SupplierBankData
    {
        /// <summary>
        /// 員工編號
        /// </summary>
        public string EmpNo { get; set; }
        /// <summary>
        /// 銀行帳號
        /// </summary>
        public string AccountNumber { get; set; }
        /// <summary>
        /// 銀行帳號戶名
        /// </summary>
        public string AccountName { get; set; }
        /// <summary>
        /// 銀行名稱
        /// </summary>
        public string BankName { get; set; }
        /// <summary>
        /// 分行名稱
        /// </summary>
        public string BranchName { get; set; }
    }
}
