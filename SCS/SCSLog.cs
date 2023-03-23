using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace OracleNewQuitEmployee.SCS
{
    public class SCSLog
    {
        public SCSLog()
        {
            LogPath = "D:/SCSSite/";
            try
            {

                if (!Directory.Exists(LogPath))
                {
                    Directory.CreateDirectory(LogPath);
                }
            }
            catch (Exception ex)
            {
            }
        }
        public string LogPath { get; set; }

        /// <summary>
        /// 寫 log 到 D:/SCSSite/
        /// </summary>
        /// <param name="msg"></param>
        /// <param name="ChkName"></param>
        public void WriteLogAsync(string msg, string ChkName = "")
        {
            try
            {
                string _chkname = $"{ChkName}";
                if (string.IsNullOrWhiteSpace(_chkname))
                {
                    _chkname = _chkname.PadLeft(20, ' ');
                }
                string filepath = $"{LogPath}SCS.{DateTime.Today.ToString("yyyyMMdd")}.log";
                string _msg = $"[{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}] [{_chkname}] : {msg}{Environment.NewLine}";
                System.IO.File.AppendAllText(filepath, _msg);
            }
            catch (Exception)
            {
            }
        }
    }




}
