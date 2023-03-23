using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class ResultObject<T>
    {
        public ResultObject()
        {
            this.ErrorMessages = new List<string>();
            this.Result = false;
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

        /// <summary>
        /// 回傳物件
        /// </summary>
        public T Data { get; set; }
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

}
