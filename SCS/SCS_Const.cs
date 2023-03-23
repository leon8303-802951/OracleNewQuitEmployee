using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.SCS
{
    public class SCS_Const
    {
        static SCS_Const()
        {
            //CompanyID = "SCS999";
            //ID = "api";
            //PW = "api1234";
            //AP01_Host = "ehrs.pchome.tw";
            //AP02_Host = "ap02.pchome.tw";

            DeptDetailType = "AIS.Define.TFindInputArgs, AIS.Define";
            ProgID = "HUM0010300";
            Depts_SelectFields = "SYS_VIEWID,SYS_NAME,SYS_ENGNAME";
            Depts_SystemFilterOptions = "Session, DataPermission, EmployeeLevel";

            Dept_SystemFilterOptions = "Session, DataPermission, EmployeeLevel";
        }
        //static public string CompanyID;
        //static public string ID;
        //static public string PW;
        //static public string AP01_Host;
        //static public string AP02_Host;
        static public string Dept_SystemFilterOptions;

        static public string DeptDetailType;
        static public string ProgID;
        static public string Depts_SelectFields;
        static public string Depts_SystemFilterOptions;
    }




}
