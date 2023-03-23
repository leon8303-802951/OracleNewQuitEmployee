using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.SCS
{
    public class ORHRPushDataModel
    {

        public List<staff_infoMd> staff_info { get; set; }


        public List<deptMd> dept { get; set; }

        //public Dept_Disable[] dept_disable { get; set; }
    }

    public class staff_infoMd
    {
        public string r_code { get; set; } // 員編
        public string r_cname { get; set; } // 中文姓名
        public string r_ename { get; set; } // 英文姓名
        public string r_dept { get; set; } // 部門
        public string r_degress { get; set; } // 職級
        public string r_cell_phone { get; set; } // 手機號碼
        public string r_birthday { get; set; } // 生日
        public string r_sex { get; set; } // 1是男  2是女
        public string r_email { get; set; } // EMAIL
        public string r_phone_ext { get; set; } // 分機
        public string r_skype_id { get; set; } // SKYPE ID
        public string r_online_date { get; set; } // 到職日
        public string r_online { get; set; } // Y是 在職: N是非在職 無法分辨是否留職停薪
        public string r_offline_date { get; set; } // 離職日
    }

    public class deptMd
    {
        public string r_dept_code { get; set; } // 部門代號
        public string r_cname { get; set; } // 部門名稱
        public string r_belong { get; set; } // 上層部門

    }




}
