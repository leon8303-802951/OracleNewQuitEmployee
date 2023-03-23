using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model.Model1
{
    public class OracleResultModel2<T>
    {
        public int itemsPerPage { get; set; }
        public int startIndex { get; set; }
        public List<T> Resources { get; set; }

    }
}
