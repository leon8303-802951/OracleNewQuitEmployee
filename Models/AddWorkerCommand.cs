using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee
{
    class AddWorkerCommand
    {
        public string Domain { get; set; }
        public string EmpName { get; set; }
        public string EmpNo { get; set; }
        public string EngName { get; set; }
        public string Email { get; set; }
        public string LegalEntityId { get; set; }
    }
}
