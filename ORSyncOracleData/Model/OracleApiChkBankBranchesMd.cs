using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
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

}
