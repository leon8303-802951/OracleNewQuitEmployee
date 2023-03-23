using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{


    using System;
    using System.Collections.Generic;

    using System.Globalization;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public class OracleReturnWorkers
    {
        [JsonProperty("items")]
        public Worker2[] Items { get; set; }

        [JsonProperty("count")]
        public string Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public string Limit { get; set; }

        [JsonProperty("offset")]
        public string Offset { get; set; }

        [JsonProperty("links")]
        public Worker2Link[] Links { get; set; }
    }

    public class Worker2
    {
        [JsonProperty("PersonId")]
        public string PersonId { get; set; }

        [JsonProperty("PersonNumber")]
        public string PersonNumber { get; set; }

        [JsonProperty("CorrespondenceLanguage")]
        public object CorrespondenceLanguage { get; set; }

        [JsonProperty("BloodType")]
        public object BloodType { get; set; }

        [JsonProperty("DateOfBirth")]
        public object DateOfBirth { get; set; }

        [JsonProperty("DateOfDeath")]
        public object DateOfDeath { get; set; }

        [JsonProperty("CountryOfBirth")]
        public object CountryOfBirth { get; set; }

        [JsonProperty("RegionOfBirth")]
        public object RegionOfBirth { get; set; }

        [JsonProperty("TownOfBirth")]
        public object TownOfBirth { get; set; }

        [JsonProperty("ApplicantNumber")]
        public object ApplicantNumber { get; set; }

        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        [JsonProperty("CreationDate")]
        public string CreationDate { get; set; }

        [JsonProperty("LastUpdatedBy")]
        public string LastUpdatedBy { get; set; }

        [JsonProperty("LastUpdateDate")]
        public string LastUpdateDate { get; set; }

        [JsonProperty("workRelationships")]
        public WorkRelationship[] WorkRelationships { get; set; }

        [JsonProperty("links")]
        public Worker2Link[] Links { get; set; }
    }

    public class Worker2Link
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
        public Worker2Properties Properties { get; set; }
    }

    public class Worker2Properties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }

    public class WorkRelationship
    {
        [JsonProperty("PeriodOfServiceId")]
        public string PeriodOfServiceId { get; set; }

        [JsonProperty("LegislationCode")]
        public string LegislationCode { get; set; }

        [JsonProperty("LegalEntityId")]
        public string LegalEntityId { get; set; }

        [JsonProperty("LegalEmployerName")]
        public string LegalEmployerName { get; set; }

        [JsonProperty("WorkerType")]
        public string WorkerType { get; set; }

        [JsonProperty("PrimaryFlag")]
        public bool PrimaryFlag { get; set; }

        [JsonProperty("StartDate")]
        public string StartDate { get; set; }

        [JsonProperty("LegalEmployerSeniorityDate")]
        public string LegalEmployerSeniorityDate { get; set; }

        [JsonProperty("EnterpriseSeniorityDate")]
        public string EnterpriseSeniorityDate { get; set; }

        [JsonProperty("OnMilitaryServiceFlag")]
        public bool OnMilitaryServiceFlag { get; set; }

        [JsonProperty("WorkerNumber")]
        public object WorkerNumber { get; set; }

        [JsonProperty("ReadyToConvertFlag")]
        public object ReadyToConvertFlag { get; set; }

        [JsonProperty("TerminationDate")]
        public object TerminationDate { get; set; }

        [JsonProperty("NotificationDate")]
        public object NotificationDate { get; set; }

        [JsonProperty("LastWorkingDate")]
        public object LastWorkingDate { get; set; }

        [JsonProperty("RevokeUserAccess")]
        public object RevokeUserAccess { get; set; }

        [JsonProperty("RecommendedForRehire")]
        public string RecommendedForRehire { get; set; }

        [JsonProperty("RecommendationReason")]
        public object RecommendationReason { get; set; }

        [JsonProperty("RecommendationAuthorizedByPersonId")]
        public object RecommendationAuthorizedByPersonId { get; set; }

        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        [JsonProperty("CreationDate")]
        public string CreationDate { get; set; }

        [JsonProperty("LastUpdatedBy")]
        public string LastUpdatedBy { get; set; }

        [JsonProperty("LastUpdateDate")]
        public string LastUpdateDate { get; set; }

        [JsonProperty("ProjectedTerminationDate")]
        public object ProjectedTerminationDate { get; set; }

        [JsonProperty("assignments")]
        public Worker2Assignment[] Assignments { get; set; }

        [JsonProperty("links")]
        public Worker2Link[] Links { get; set; }
    }

    public class Worker2Assignment
    {
        [JsonProperty("AssignmentId")]
        public string AssignmentId { get; set; }

        [JsonProperty("AssignmentNumber")]
        public string AssignmentNumber { get; set; }

        [JsonProperty("AssignmentName")]
        public string AssignmentName { get; set; }

        [JsonProperty("ActionCode")]
        public string ActionCode { get; set; }

        [JsonProperty("ReasonCode")]
        public object ReasonCode { get; set; }

        [JsonProperty("EffectiveStartDate")]
        public string EffectiveStartDate { get; set; }

        [JsonProperty("EffectiveEndDate")]
        public string EffectiveEndDate { get; set; }

        [JsonProperty("EffectiveSequence")]
        public string EffectiveSequence { get; set; }

        [JsonProperty("EffectiveLatestChange")]
        public string EffectiveLatestChange { get; set; }

        [JsonProperty("BusinessUnitId")]
        public string BusinessUnitId { get; set; }

        [JsonProperty("BusinessUnitName")]
        public string BusinessUnitName { get; set; }

        [JsonProperty("AssignmentType")]
        public string AssignmentType { get; set; }

        [JsonProperty("AssignmentStatusTypeId")]
        public string AssignmentStatusTypeId { get; set; }

        [JsonProperty("AssignmentStatusTypeCode")]
        public string AssignmentStatusTypeCode { get; set; }

        [JsonProperty("AssignmentStatusType")]
        public string AssignmentStatusType { get; set; }

        [JsonProperty("SystemPersonType")]
        public string SystemPersonType { get; set; }

        [JsonProperty("UserPersonTypeId")]
        public string UserPersonTypeId { get; set; }

        [JsonProperty("UserPersonType")]
        public string UserPersonType { get; set; }

        [JsonProperty("ProposedUserPersonTypeId")]
        public object ProposedUserPersonTypeId { get; set; }

        [JsonProperty("ProposedUserPersonType")]
        public object ProposedUserPersonType { get; set; }

        [JsonProperty("ProjectedStartDate")]
        public object ProjectedStartDate { get; set; }

        [JsonProperty("ProjectedEndDate")]
        public object ProjectedEndDate { get; set; }

        [JsonProperty("PrimaryFlag")]
        public bool PrimaryFlag { get; set; }

        [JsonProperty("PrimaryAssignmentFlag")]
        public bool PrimaryAssignmentFlag { get; set; }

        [JsonProperty("PositionId")]
        public object PositionId { get; set; }

        [JsonProperty("PositionCode")]
        public object PositionCode { get; set; }

        [JsonProperty("SynchronizeFromPositionFlag")]
        public bool SynchronizeFromPositionFlag { get; set; }

        [JsonProperty("JobId")]
        public string JobId { get; set; }

        [JsonProperty("JobCode")]
        public string JobCode { get; set; }

        [JsonProperty("GradeId")]
        public object GradeId { get; set; }

        [JsonProperty("GradeCode")]
        public object GradeCode { get; set; }

        [JsonProperty("GradeLadderId")]
        public object GradeLadderId { get; set; }

        [JsonProperty("GradeLadderName")]
        public object GradeLadderName { get; set; }

        [JsonProperty("GradeStepEligibilityFlag")]
        public bool GradeStepEligibilityFlag { get; set; }

        [JsonProperty("GradeCeilingStepId")]
        public object GradeCeilingStepId { get; set; }

        [JsonProperty("GradeCeilingStep")]
        public object GradeCeilingStep { get; set; }

        [JsonProperty("DepartmentId")]
        public object DepartmentId { get; set; }

        [JsonProperty("DepartmentName")]
        public object DepartmentName { get; set; }

        [JsonProperty("ReportingEstablishmentId")]
        public object ReportingEstablishmentId { get; set; }

        [JsonProperty("ReportingEstablishmentName")]
        public object ReportingEstablishmentName { get; set; }

        [JsonProperty("LocationId")]
        public object LocationId { get; set; }

        [JsonProperty("LocationCode")]
        public object LocationCode { get; set; }

        [JsonProperty("WorkAtHomeFlag")]
        public bool WorkAtHomeFlag { get; set; }

        [JsonProperty("AssignmentCategory")]
        public object AssignmentCategory { get; set; }

        [JsonProperty("WorkerCategory")]
        public object WorkerCategory { get; set; }

        [JsonProperty("PermanentTemporary")]
        public object PermanentTemporary { get; set; }

        [JsonProperty("FullPartTime")]
        public object FullPartTime { get; set; }

        [JsonProperty("ManagerFlag")]
        public bool ManagerFlag { get; set; }

        [JsonProperty("HourlySalariedCode")]
        public object HourlySalariedCode { get; set; }

        [JsonProperty("NormalHours")]
        public object NormalHours { get; set; }

        [JsonProperty("Frequency")]
        public object Frequency { get; set; }

        [JsonProperty("StartTime")]
        public object StartTime { get; set; }

        [JsonProperty("EndTime")]
        public object EndTime { get; set; }

        [JsonProperty("SeniorityBasis")]
        public string SeniorityBasis { get; set; }

        [JsonProperty("ProbationPeriod")]
        public object ProbationPeriod { get; set; }

        [JsonProperty("ProbationPeriodUnit")]
        public object ProbationPeriodUnit { get; set; }

        [JsonProperty("ProbationEndDate")]
        public object ProbationEndDate { get; set; }

        [JsonProperty("NoticePeriod")]
        public object NoticePeriod { get; set; }

        [JsonProperty("NoticePeriodUOM")]
        public object NoticePeriodUom { get; set; }

        [JsonProperty("WorkTaxAddressId")]
        public object WorkTaxAddressId { get; set; }

        [JsonProperty("ExpenseCheckSendToAddress")]
        public object ExpenseCheckSendToAddress { get; set; }

        [JsonProperty("RetirementAge")]
        public object RetirementAge { get; set; }

        [JsonProperty("RetirementDate")]
        public object RetirementDate { get; set; }

        [JsonProperty("LabourUnionMemberFlag")]
        public object LabourUnionMemberFlag { get; set; }

        [JsonProperty("UnionId")]
        public object UnionId { get; set; }

        [JsonProperty("UnionName")]
        public object UnionName { get; set; }

        [JsonProperty("BargainingUnitCode")]
        public object BargainingUnitCode { get; set; }

        [JsonProperty("CollectiveAgreementId")]
        public object CollectiveAgreementId { get; set; }

        [JsonProperty("CollectiveAgreementName")]
        public object CollectiveAgreementName { get; set; }

        [JsonProperty("ContractId")]
        public object ContractId { get; set; }

        [JsonProperty("ContractNumber")]
        public object ContractNumber { get; set; }

        [JsonProperty("InternalBuilding")]
        public object InternalBuilding { get; set; }

        [JsonProperty("InternalFloor")]
        public object InternalFloor { get; set; }

        [JsonProperty("InternalOfficeNumber")]
        public object InternalOfficeNumber { get; set; }

        [JsonProperty("InternalMailstop")]
        public object InternalMailstop { get; set; }

        [JsonProperty("DefaultExpenseAccount")]
        public string DefaultExpenseAccount { get; set; }

        [JsonProperty("PeopleGroup")]
        public object PeopleGroup { get; set; }

        [JsonProperty("StandardWorkingHours")]
        public object StandardWorkingHours { get; set; }

        [JsonProperty("StandardFrequency")]
        public object StandardFrequency { get; set; }

        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        [JsonProperty("CreationDate")]
        public string CreationDate { get; set; }

        [JsonProperty("LastUpdatedBy")]
        public string LastUpdatedBy { get; set; }

        [JsonProperty("LastUpdateDate")]
        public string LastUpdateDate { get; set; }

        [JsonProperty("links")]
        public Worker2Link[] Links { get; set; }
    }




}
