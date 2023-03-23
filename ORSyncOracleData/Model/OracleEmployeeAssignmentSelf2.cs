using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
    using Newtonsoft.Json;

    public class OracleEmployeeAssignmentSelf2
    {
        [JsonProperty("AssignmentName")]
        public string AssignmentName { get; set; }

        [JsonProperty("PersonTypeId")]
        public string PersonTypeId { get; set; }

        [JsonProperty("ProposedPersonTypeId")]
        public string ProposedPersonTypeId { get; set; }

        [JsonProperty("ProjectedStartDate")]
        public string ProjectedStartDate { get; set; }

        [JsonProperty("BusinessUnitId")]
        public string BusinessUnitId { get; set; }

        [JsonProperty("LocationId")]
        public string LocationId { get; set; }

        [JsonProperty("JobId")]
        public string JobId { get; set; }

        [JsonProperty("GradeId")]
        public string GradeId { get; set; }

        [JsonProperty("DepartmentId")]
        public string DepartmentId { get; set; }

        [JsonProperty("WorkerCategory")]
        public string WorkerCategory { get; set; }

        [JsonProperty("AssignmentCategory")]
        public string AssignmentCategory { get; set; }

        [JsonProperty("WorkingAtHome")]
        public string WorkingAtHome { get; set; }

        [JsonProperty("WorkingAsManager")]
        public string WorkingAsManager { get; set; }

        [JsonProperty("SalaryCode")]
        public string SalaryCode { get; set; }

        [JsonProperty("WorkingHours")]
        public string WorkingHours { get; set; }

        [JsonProperty("Frequency")]
        public string Frequency { get; set; }

        [JsonProperty("StartTime")]
        public string StartTime { get; set; }

        [JsonProperty("EndTime")]
        public string EndTime { get; set; }

        [JsonProperty("SalaryAmount")]
        public string SalaryAmount { get; set; }

        [JsonProperty("SalaryBasisId")]
        public string SalaryBasisId { get; set; }

        [JsonProperty("ActionCode")]
        public string ActionCode { get; set; }

        [JsonProperty("ActionReasonCode")]
        public string ActionReasonCode { get; set; }

        [JsonProperty("AssignmentStatus")]
        public string AssignmentStatus { get; set; }

        [JsonProperty("WorkTaxAddressId")]
        public string WorkTaxAddressId { get; set; }

        [JsonProperty("AssignmentId")]
        public string AssignmentId { get; set; }

        [JsonProperty("EffectiveStartDate")]
        public string EffectiveStartDate { get; set; }

        [JsonProperty("EffectiveEndDate")]
        public string EffectiveEndDate { get; set; }

        [JsonProperty("PositionId")]
        public string PositionId { get; set; }

        [JsonProperty("TermsEffectiveStartDate")]
        public string TermsEffectiveStartDate { get; set; }

        [JsonProperty("ManagerId")]
        public string ManagerId { get; set; }

        [JsonProperty("ManagerAssignmentId")]
        public string ManagerAssignmentId { get; set; }

        [JsonProperty("ManagerType")]
        public string ManagerType { get; set; }

        [JsonProperty("AssignmentNumber")]
        public string AssignmentNumber { get; set; }

        [JsonProperty("OriginalHireDate")]
        public string OriginalHireDate { get; set; }

        [JsonProperty("AssignmentStatusTypeId")]
        public string AssignmentStatusTypeId { get; set; }

        [JsonProperty("PrimaryAssignmentFlag")]
        public string PrimaryAssignmentFlag { get; set; }

        [JsonProperty("ProbationPeriodEndDate")]
        public string ProbationPeriodEndDate { get; set; }

        [JsonProperty("ProbationPeriodLength")]
        public string ProbationPeriodLength { get; set; }

        [JsonProperty("ProbationPeriodUnitOfMeasure")]
        public string ProbationPeriodUnitOfMeasure { get; set; }

        [JsonProperty("AssignmentProjectedEndDate")]
        public string AssignmentProjectedEndDate { get; set; }

        [JsonProperty("ActualTerminationDate")]
        public string ActualTerminationDate { get; set; }

        [JsonProperty("LegalEntityId")]
        public string LegalEntityId { get; set; }

        [JsonProperty("PrimaryWorkRelationFlag")]
        public string PrimaryWorkRelationFlag { get; set; }

        [JsonProperty("PrimaryWorkTermsFlag")]
        public string PrimaryWorkTermsFlag { get; set; }

        [JsonProperty("CreationDate")]
        public string CreationDate { get; set; }

        [JsonProperty("LastUpdateDate")]
        public string LastUpdateDate { get; set; }

        [JsonProperty("PeriodOfServiceId")]
        public string PeriodOfServiceId { get; set; }

        [JsonProperty("FullPartTime")]
        public string FullPartTime { get; set; }

        [JsonProperty("RegularTemporary")]
        public string RegularTemporary { get; set; }

        [JsonProperty("GradeLadderId")]
        public string GradeLadderId { get; set; }

        [JsonProperty("DefaultExpenseAccount")]
        public string DefaultExpenseAccount { get; set; }

        [JsonProperty("PeopleGroup")]
        public string PeopleGroup { get; set; }

        [JsonProperty("links")]
        public EmployeeAssignmentSelf2Link[] Links { get; set; }
    }

    public class EmployeeAssignmentSelf2Link
    {
        [JsonProperty("rel")]
        public string Rel { get; set; }

        [JsonProperty("href")]
        public Uri Href { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("kind")]
        public string Kind { get; set; }

        [JsonProperty("properties", NullValueHandling = NullValueHandling.Ignore)]
        public EmployeeAssignmentSelf2Properties Properties { get; set; }
    }

    public class EmployeeAssignmentSelf2Properties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }




}
