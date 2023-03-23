using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model.EmpAsm
{

    using Newtonsoft.Json;

    public class OracleEmployeeIncludeAssignmentsCollection2
    {
        [JsonProperty("items")]
        public OracleEmployee2[] Items { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public int Limit { get; set; }

        [JsonProperty("offset")]
        public int Offset { get; set; }

        [JsonProperty("links")]
        public Link2[] Links { get; set; }
    }

    public class OracleEmployee2
    {
        [JsonProperty("Salutation")]
        public string Salutation { get; set; }

        /// <summary>
        /// 員工中文姓名
        /// </summary>
        [JsonProperty("FirstName")]
        public string FirstName { get; set; }

        [JsonProperty("MiddleName")]
        public string MiddleName { get; set; }

        /// <summary>
        /// 員工編號 EmpNo
        /// </summary>
        [JsonProperty("LastName")]
        public string LastName { get; set; }

        [JsonProperty("PreviousLastName")]
        public string PreviousLastName { get; set; }

        [JsonProperty("NameSuffix")]
        public string NameSuffix { get; set; }

        [JsonProperty("DisplayName")]
        public string DisplayName { get; set; }

        [JsonProperty("PreferredName")]
        public string PreferredName { get; set; }

        [JsonProperty("Honors")]
        public string Honors { get; set; }

        [JsonProperty("CorrespondenceLanguage")]
        public string CorrespondenceLanguage { get; set; }

        [JsonProperty("PersonNumber")]
        public string PersonNumber { get; set; }

        [JsonProperty("WorkPhoneCountryCode")]
        public string WorkPhoneCountryCode { get; set; }

        [JsonProperty("WorkPhoneAreaCode")]
        public string WorkPhoneAreaCode { get; set; }

        [JsonProperty("WorkPhoneNumber")]
        public string WorkPhoneNumber { get; set; }

        [JsonProperty("WorkPhoneExtension")]
        public string WorkPhoneExtension { get; set; }

        [JsonProperty("WorkPhoneLegislationCode")]
        public string WorkPhoneLegislationCode { get; set; }

        [JsonProperty("WorkFaxCountryCode")]
        public string WorkFaxCountryCode { get; set; }

        [JsonProperty("WorkFaxAreaCode")]
        public string WorkFaxAreaCode { get; set; }

        [JsonProperty("WorkFaxNumber")]
        public string WorkFaxNumber { get; set; }

        [JsonProperty("WorkFaxExtension")]
        public string WorkFaxExtension { get; set; }

        [JsonProperty("WorkFaxLegislationCode")]
        public string WorkFaxLegislationCode { get; set; }

        [JsonProperty("WorkMobilePhoneCountryCode")]
        public string WorkMobilePhoneCountryCode { get; set; }

        [JsonProperty("WorkMobilePhoneAreaCode")]
        public string WorkMobilePhoneAreaCode { get; set; }

        [JsonProperty("WorkMobilePhoneNumber")]
        public string WorkMobilePhoneNumber { get; set; }

        [JsonProperty("WorkMobilePhoneExtension")]
        public string WorkMobilePhoneExtension { get; set; }

        [JsonProperty("WorkMobilePhoneLegislationCode")]
        public string WorkMobilePhoneLegislationCode { get; set; }

        [JsonProperty("HomePhoneCountryCode")]
        public string HomePhoneCountryCode { get; set; }

        [JsonProperty("HomePhoneAreaCode")]
        public string HomePhoneAreaCode { get; set; }

        [JsonProperty("HomePhoneNumber")]
        public string HomePhoneNumber { get; set; }

        [JsonProperty("HomePhoneExtension")]
        public string HomePhoneExtension { get; set; }

        [JsonProperty("HomePhoneLegislationCode")]
        public string HomePhoneLegislationCode { get; set; }

        [JsonProperty("HomeFaxCountryCode")]
        public string HomeFaxCountryCode { get; set; }

        [JsonProperty("HomeFaxAreaCode")]
        public string HomeFaxAreaCode { get; set; }

        [JsonProperty("HomeFaxNumber")]
        public string HomeFaxNumber { get; set; }

        [JsonProperty("HomeFaxExtension")]
        public string HomeFaxExtension { get; set; }

        [JsonProperty("WorkEmail")]
        public string WorkEmail { get; set; }

        [JsonProperty("HomeFaxLegislationCode")]
        public string HomeFaxLegislationCode { get; set; }

        [JsonProperty("AddressLine1")]
        public string AddressLine1 { get; set; }

        [JsonProperty("AddressLine2")]
        public string AddressLine2 { get; set; }

        [JsonProperty("AddressLine3")]
        public string AddressLine3 { get; set; }

        [JsonProperty("City")]
        public string City { get; set; }

        [JsonProperty("Region")]
        public string Region { get; set; }

        [JsonProperty("Region2")]
        public string Region2 { get; set; }

        [JsonProperty("Country")]
        public string Country { get; set; }

        [JsonProperty("PostalCode")]
        public string PostalCode { get; set; }

        [JsonProperty("DateOfBirth")]
        public string DateOfBirth { get; set; }

        [JsonProperty("Ethnicity")]
        public string Ethnicity { get; set; }

        [JsonProperty("ProjectedTerminationDate")]
        public string ProjectedTerminationDate { get; set; }

        [JsonProperty("LegalEntityId")]
        public string LegalEntityId { get; set; }

        [JsonProperty("HireDate")]
        public string HireDate { get; set; }

        [JsonProperty("TerminationDate")]
        public string TerminationDate { get; set; }

        [JsonProperty("Gender")]
        public string Gender { get; set; }

        [JsonProperty("MaritalStatus")]
        public string MaritalStatus { get; set; }

        [JsonProperty("NationalIdType")]
        public string NationalIdType { get; set; }

        [JsonProperty("NationalId")]
        public string NationalId { get; set; }

        [JsonProperty("NationalIdCountry")]
        public string NationalIdCountry { get; set; }

        [JsonProperty("NationalIdExpirationDate")]
        public string NationalIdExpirationDate { get; set; }

        [JsonProperty("NationalIdPlaceOfIssue")]
        public string NationalIdPlaceOfIssue { get; set; }

        [JsonProperty("PersonId")]
        public string PersonId { get; set; }

        [JsonProperty("EffectiveStartDate")]
        public string EffectiveStartDate { get; set; }

        [JsonProperty("UserName")]
        public string UserName { get; set; }

        [JsonProperty("CitizenshipId")]
        public string CitizenshipId { get; set; }

        [JsonProperty("CitizenshipStatus")]
        public string CitizenshipStatus { get; set; }

        [JsonProperty("CitizenshipLegislationCode")]
        public string CitizenshipLegislationCode { get; set; }

        [JsonProperty("CitizenshipToDate")]
        public string CitizenshipToDate { get; set; }

        [JsonProperty("Religion")]
        public string Religion { get; set; }

        [JsonProperty("ReligionId")]
        public string ReligionId { get; set; }

        [JsonProperty("PassportIssueDate")]
        public string PassportIssueDate { get; set; }

        [JsonProperty("PassportNumber")]
        public string PassportNumber { get; set; }

        [JsonProperty("PassportIssuingCountry")]
        public string PassportIssuingCountry { get; set; }

        [JsonProperty("PassportId")]
        public string PassportId { get; set; }

        [JsonProperty("PassportExpirationDate")]
        public string PassportExpirationDate { get; set; }

        [JsonProperty("LicenseNumber")]
        public string LicenseNumber { get; set; }

        [JsonProperty("DriversLicenseExpirationDate")]
        public string DriversLicenseExpirationDate { get; set; }

        [JsonProperty("DriversLicenseIssuingCountry")]
        public string DriversLicenseIssuingCountry { get; set; }

        [JsonProperty("DriversLicenseId")]
        public string DriversLicenseId { get; set; }

        [JsonProperty("MilitaryVetStatus")]
        public string MilitaryVetStatus { get; set; }

        [JsonProperty("CreationDate")]
        public string CreationDate { get; set; }

        [JsonProperty("LastUpdateDate")]
        public string LastUpdateDate { get; set; }

        [JsonProperty("WorkerType")]
        public string WorkerType { get; set; }

        [JsonProperty("assignments")]
        public EmployeeAssignment2[] Assignments { get; set; }

        [JsonProperty("links")]
        public Link2[] Links { get; set; }
    }

    public class EmployeeAssignment2
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
        public bool PrimaryAssignmentFlag { get; set; }

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
        public bool PrimaryWorkRelationFlag { get; set; }

        [JsonProperty("PrimaryWorkTermsFlag")]
        public bool PrimaryWorkTermsFlag { get; set; }

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
        public Link2[] Links { get; set; }
    }

    public class Link2
    {
        [JsonProperty("rel")]
        public string Rel { get; set; }

        [JsonProperty("href")]
        public string Href { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("kind")]
        public string Kind { get; set; }

        [JsonProperty("properties", NullValueHandling = NullValueHandling.Ignore)]
        public EmployeeLinkProperties2 Properties { get; set; }
    }

    public class EmployeeLinkProperties2
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }



}
