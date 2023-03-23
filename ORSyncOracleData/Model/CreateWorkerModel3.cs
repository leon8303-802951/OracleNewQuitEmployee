using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    public class CreateWorkerModel3
    {
        public CreateWorkerModel3()
        {
            Names = new List<CreateWorkerName3>();
            Names.Add(new CreateWorkerName3());
            Names[0].LegislationCode = "TW";
            WorkRelationships = new List<CreateWorkRelationship3>();
            WorkRelationships.Add(new CreateWorkRelationship3());
            WorkRelationships[0].WorkerType = "E";
            WorkRelationships[0].LegalEmployerName = "網路家庭國際資訊股份有限公司";
            //WorkRelationships[0].LegalEntityId = 300000001503001;
            WorkRelationships[0].Assignments[0].ActionCode = "HIRE";
            WorkRelationships[0].Assignments[0].BusinessUnitName = "PChome";
            Emails = new List<CreateWorkEmail3>();
            Emails.Add(new CreateWorkEmail3());
        }

        [JsonProperty("names")]
        public List<CreateWorkerName3> Names { get; set; }

        [JsonProperty("PersonNumber")]
        public string PersonNumber { get; set; }

        [JsonProperty("emails")]
        public List<CreateWorkEmail3> Emails { get; set; }

        [JsonProperty("workRelationships")]
        public List<CreateWorkRelationship3> WorkRelationships { get; set; }
    }
    public class CreateWorkRelationship3
    {
        public CreateWorkRelationship3()
        {
            //LegalEmployerName = "網路家庭國際資訊股份有限公司";
            Assignments = new List<CreateAssignment3>();
            Assignments.Add(new CreateAssignment3());
        }
        [JsonProperty("LegalEmployerName")]
        public string LegalEmployerName { get; set; }

        [JsonProperty("WorkerType")]
        public string WorkerType { get; set; }

        [JsonProperty("assignments")]
        public List<CreateAssignment3> Assignments { get; set; }
    }

    public class CreateWorkerName3
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
    public class CreateAssignment3
    {
        [JsonProperty("ActionCode")]
        public string ActionCode { get; set; }

        [JsonProperty("BusinessUnitName")]
        public string BusinessUnitName { get; set; }
    }

    public class CreateWorkEmail3
    {
        [JsonProperty("EmailAddress")]
        public string EmailAddress { get; set; }

        [JsonProperty("EmailType")]
        public string EmailType { get; set; }
    }

}
