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

    public class SupplierResponseModel
    {
        [JsonProperty("items")]
        public SupplierModel[] Items { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public int Limit { get; set; }

        [JsonProperty("offset")]
        public int Offset { get; set; }

        [JsonProperty("links")]
        public SupplierLink[] Links { get; set; }
    }

    public class SupplierModel
    {
        [JsonProperty("SupplierId")]
        public string SupplierId { get; set; }

        [JsonProperty("SupplierPartyId")]
        public string SupplierPartyId { get; set; }

        [JsonProperty("Supplier")]
        public string Supplier { get; set; }

        [JsonProperty("SupplierNumber")]
        public string SupplierNumber { get; set; }

        [JsonProperty("AlternateName")]
        public string AlternateName { get; set; }

        [JsonProperty("TaxOrganizationTypeCode")]
        public string TaxOrganizationTypeCode { get; set; }

        [JsonProperty("TaxOrganizationType")]
        public string TaxOrganizationType { get; set; }

        [JsonProperty("SupplierTypeCode")]
        public string SupplierTypeCode { get; set; }

        [JsonProperty("SupplierType")]
        public string SupplierType { get; set; }

        [JsonProperty("InactiveDate")]
        public string InactiveDate { get; set; }

        [JsonProperty("Status")]
        public string Status { get; set; }

        [JsonProperty("BusinessRelationshipCode")]
        public string BusinessRelationshipCode { get; set; }

        [JsonProperty("BusinessRelationship")]
        public string BusinessRelationship { get; set; }

        [JsonProperty("ParentSupplierId")]
        public string ParentSupplierId { get; set; }

        [JsonProperty("ParentSupplier")]
        public string ParentSupplier { get; set; }

        [JsonProperty("ParentSupplierNumber")]
        public string ParentSupplierNumber { get; set; }

        [JsonProperty("CreationDate")]
        public string CreationDate { get; set; }

        [JsonProperty("CreationSourceCode")]
        public string CreationSourceCode { get; set; }

        [JsonProperty("CreationSource")]
        public string CreationSource { get; set; }

        [JsonProperty("Alias")]
        public string Alias { get; set; }

        [JsonProperty("DUNSNumber")]
        public string DunsNumber { get; set; }

        [JsonProperty("OneTimeSupplierFlag")]
        public string OneTimeSupplierFlag { get; set; }

        [JsonProperty("RegistryId")]
        public string RegistryId { get; set; }

        [JsonProperty("CustomerNumber")]
        public string CustomerNumber { get; set; }

        [JsonProperty("StandardIndustryClass")]
        public string StandardIndustryClass { get; set; }

        [JsonProperty("NationalInsuranceNumber")]
        public string NationalInsuranceNumber { get; set; }

        [JsonProperty("NationalInsuranceNumberExistsFlag")]
        public bool NationalInsuranceNumberExistsFlag { get; set; }

        [JsonProperty("CorporateWebsite")]
        public string CorporateWebsite { get; set; }

        [JsonProperty("YearEstablished")]
        public string YearEstablished { get; set; }

        [JsonProperty("MissionStatement")]
        public string MissionStatement { get; set; }

        [JsonProperty("YearIncorporated")]
        public string YearIncorporated { get; set; }

        [JsonProperty("ChiefExecutiveTitle")]
        public string ChiefExecutiveTitle { get; set; }

        [JsonProperty("ChiefExecutiveName")]
        public string ChiefExecutiveName { get; set; }

        [JsonProperty("PrincipalTitle")]
        public string PrincipalTitle { get; set; }

        [JsonProperty("PrincipalName")]
        public string PrincipalName { get; set; }

        [JsonProperty("FiscalYearEndMonthCode")]
        public string FiscalYearEndMonthCode { get; set; }

        [JsonProperty("FiscalYearEndMonth")]
        public string FiscalYearEndMonth { get; set; }

        [JsonProperty("CurrentFiscalYearPotentialRevenue")]
        public string CurrentFiscalYearPotentialRevenue { get; set; }

        [JsonProperty("PreferredFunctionalCurrencyCode")]
        public string PreferredFunctionalCurrencyCode { get; set; }

        [JsonProperty("PreferredFunctionalCurrency")]
        public string PreferredFunctionalCurrency { get; set; }

        [JsonProperty("TaxRegistrationCountryCode")]
        public string TaxRegistrationCountryCode { get; set; }

        [JsonProperty("TaxRegistrationCountry")]
        public string TaxRegistrationCountry { get; set; }

        [JsonProperty("TaxRegistrationNumber")]
        public string TaxRegistrationNumber { get; set; }

        [JsonProperty("TaxpayerCountryCode")]
        public string TaxpayerCountryCode { get; set; }

        [JsonProperty("TaxpayerCountry")]
        public string TaxpayerCountry { get; set; }

        [JsonProperty("TaxpayerId")]
        public string TaxpayerId { get; set; }

        [JsonProperty("TaxpayerIdExistsFlag")]
        public bool TaxpayerIdExistsFlag { get; set; }

        [JsonProperty("FederalReportableFlag")]
        public bool FederalReportableFlag { get; set; }

        [JsonProperty("FederalIncomeTaxTypeCode")]
        public string FederalIncomeTaxTypeCode { get; set; }

        [JsonProperty("FederalIncomeTaxType")]
        public string FederalIncomeTaxType { get; set; }

        [JsonProperty("StateReportableFlag")]
        public string StateReportableFlag { get; set; }

        [JsonProperty("TaxReportingName")]
        public string TaxReportingName { get; set; }

        [JsonProperty("NameControl")]
        public string NameControl { get; set; }

        [JsonProperty("VerificationDate")]
        public string VerificationDate { get; set; }

        [JsonProperty("UseWithholdingTaxFlag")]
        public string UseWithholdingTaxFlag { get; set; }

        [JsonProperty("WithholdingTaxGroupId")]
        public string WithholdingTaxGroupId { get; set; }

        [JsonProperty("WithholdingTaxGroup")]
        public string WithholdingTaxGroup { get; set; }

        [JsonProperty("BusinessClassificationNotApplicableFlag")]
        public string BusinessClassificationNotApplicableFlag { get; set; }

        [JsonProperty("DataFoxId")]
        public string DataFoxId { get; set; }

        [JsonProperty("DataFoxScore")]
        public string DataFoxScore { get; set; }

        [JsonProperty("IndustryCategory")]
        public string IndustryCategory { get; set; }

        [JsonProperty("IndustrySubcategory")]
        public string IndustrySubcategory { get; set; }

        [JsonProperty("DataFoxCompanyName")]
        public string DataFoxCompanyName { get; set; }

        [JsonProperty("DataFoxLegalName")]
        public string DataFoxLegalName { get; set; }

        [JsonProperty("DataFoxCompanyPrimaryURL")]
        public string DataFoxCompanyPrimaryUrl { get; set; }

        [JsonProperty("DataFoxNAICSCode")]
        public string DataFoxNaicsCode { get; set; }

        [JsonProperty("DataFoxCountry")]
        public string DataFoxCountry { get; set; }

        [JsonProperty("DataFoxEIN")]
        public string DataFoxEin { get; set; }

        [JsonProperty("DataFoxLastSyncDate")]
        public string DataFoxLastSyncDate { get; set; }

        [JsonProperty("links")]
        public SupplierLink[] Links { get; set; }
    }

    public class SupplierLink
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
        public SupplierProperties Properties { get; set; }
    }

    public class SupplierProperties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }



}

