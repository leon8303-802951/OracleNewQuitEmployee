using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class CreateWorkerModel2
    {
        public CreateWorkerModel2()
        {
            Names = new List<CreateWorkerName2>();
            Names.Add(new CreateWorkerName2());
            Names[0].LegislationCode = "TW";
            WorkRelationships = new List<CreateWorkRelationship2>();
            WorkRelationships.Add(new CreateWorkRelationship2());
            WorkRelationships[0].WorkerType = "E";
            WorkRelationships[0].LegalEmployerName = "網路家庭國際資訊股份有限公司";
            WorkRelationships[0].Assignments[0].ActionCode = "HIRE";
            WorkRelationships[0].Assignments[0].BusinessUnitName = "PChome";
            Emails = new List<CreateWorkEmail2>();
            Emails.Add(new CreateWorkEmail2());
        }

        [JsonProperty("names")]
        public List<CreateWorkerName2> Names { get; set; }

        [JsonProperty("PersonNumber")]
        public string PersonNumber { get; set; }

        [JsonProperty("emails")]
        public List<CreateWorkEmail2> Emails { get; set; }

        [JsonProperty("workRelationships")]
        public List<CreateWorkRelationship2> WorkRelationships { get; set; }
    }
    public class CreateWorkRelationship2
    {
        public CreateWorkRelationship2()
        {
            Assignments = new List<CreateAssignment2>();
            Assignments.Add(new CreateAssignment2());
        }
        [JsonProperty("LegalEmployerName")]
        public string LegalEmployerName { get; set; }

        [JsonProperty("WorkerType")]
        public string WorkerType { get; set; }

        [JsonProperty("assignments")]
        public List<CreateAssignment2> Assignments { get; set; }
    }

    public class CreateWorkerName2
    {
        /// <summary>
        /// 員工中文姓名
        /// </summary>
        [JsonProperty("FirstName")]
        public string FirstName { get; set; }

        /// <summary>
        /// 員工編號 EmpNo
        /// </summary>
        [JsonProperty("LastName")]
        public string LastName { get; set; }

        /// <summary>
        /// 員工英文姓名
        /// </summary>
        [JsonProperty("MiddleNames")]
        public string MiddleNames { get; set; }

        [JsonProperty("LegislationCode")]
        public string LegislationCode { get; set; }
    }
    public class CreateAssignment2
    {
        [JsonProperty("ActionCode")]
        public string ActionCode { get; set; }

        [JsonProperty("BusinessUnitName")]
        public string BusinessUnitName { get; set; }
    }

    public class CreateWorkEmail2
    {
        [JsonProperty("EmailAddress")]
        public string EmailAddress { get; set; }

        [JsonProperty("EmailType")]
        public string EmailType { get; set; }
    }

}
