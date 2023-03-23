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

    public class SupplierSiteReturnModel
    {
        [JsonProperty("items")]
        public SupplierSiteModel[] Items { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public int Limit { get; set; }

        [JsonProperty("offset")]
        public int Offset { get; set; }

        [JsonProperty("links")]
        public SupplierSiteLink[] Links { get; set; }
    }

    public class SupplierSiteModel
    {
        [JsonProperty("SupplierSiteId")]
        public string SupplierSiteId { get; set; }

        [JsonProperty("SupplierSite")]
        public string SupplierSite { get; set; }

        [JsonProperty("ProcurementBUId")]
        public string ProcurementBuId { get; set; }

        [JsonProperty("ProcurementBU")]
        public string ProcurementBu { get; set; }

        [JsonProperty("SupplierAddressId")]
        public string SupplierAddressId { get; set; }

        [JsonProperty("SupplierAddressName")]
        public string SupplierAddressName { get; set; }

        [JsonProperty("InactiveDate")]
        public string InactiveDate { get; set; }

        [JsonProperty("Status")]
        public string Status { get; set; }

        [JsonProperty("SitePurposeSourcingOnlyFlag")]
        public string SitePurposeSourcingOnlyFlag { get; set; }

        [JsonProperty("SitePurposePurchasingFlag")]
        public bool SitePurposePurchasingFlag { get; set; }

        [JsonProperty("SitePurposeProcurementCardFlag")]
        public string SitePurposeProcurementCardFlag { get; set; }

        [JsonProperty("SitePurposePayFlag")]
        public bool SitePurposePayFlag { get; set; }

        [JsonProperty("SitePurposePrimaryPayFlag")]
        public string SitePurposePrimaryPayFlag { get; set; }

        [JsonProperty("IncomeTaxReportingSiteFlag")]
        public string IncomeTaxReportingSiteFlag { get; set; }

        [JsonProperty("AlternateSiteName")]
        public string AlternateSiteName { get; set; }

        [JsonProperty("CustomerNumber")]
        public string CustomerNumber { get; set; }

        [JsonProperty("B2BCommunicationMethodCode")]
        public string B2BCommunicationMethodCode { get; set; }

        [JsonProperty("B2BCommunicationMethod")]
        public string B2BCommunicationMethod { get; set; }

        [JsonProperty("B2BSupplierSiteCode")]
        public string B2BSupplierSiteCode { get; set; }

        [JsonProperty("CommunicationMethodCode")]
        public string CommunicationMethodCode { get; set; }

        [JsonProperty("CommunicationMethod")]
        public string CommunicationMethod { get; set; }

        [JsonProperty("Email")]
        public string Email { get; set; }

        [JsonProperty("FaxCountryCode")]
        public string FaxCountryCode { get; set; }

        [JsonProperty("FaxAreaCode")]
        public string FaxAreaCode { get; set; }

        [JsonProperty("FaxNumber")]
        public string FaxNumber { get; set; }

        [JsonProperty("HoldAllNewPurchasingDocumentsFlag")]
        public string HoldAllNewPurchasingDocumentsFlag { get; set; }

        [JsonProperty("PurchasingHoldReason")]
        public string PurchasingHoldReason { get; set; }

        [JsonProperty("AllNewPurchasingDocumentsHoldDate")]
        public string AllNewPurchasingDocumentsHoldDate { get; set; }

        [JsonProperty("AllNewPurchasingDocumentsHoldBy")]
        public string AllNewPurchasingDocumentsHoldBy { get; set; }

        [JsonProperty("RequiredAcknowledgmentCode")]
        public string RequiredAcknowledgmentCode { get; set; }

        [JsonProperty("RequiredAcknowledgment")]
        public string RequiredAcknowledgment { get; set; }

        [JsonProperty("AcknowledgmentWithinDays")]
        public string AcknowledgmentWithinDays { get; set; }

        [JsonProperty("CarrierId")]
        public string CarrierId { get; set; }

        [JsonProperty("Carrier")]
        public string Carrier { get; set; }

        [JsonProperty("ModeOfTransportCode")]
        public string ModeOfTransportCode { get; set; }

        [JsonProperty("ModeOfTransport")]
        public string ModeOfTransport { get; set; }

        [JsonProperty("ServiceLevelCode")]
        public string ServiceLevelCode { get; set; }

        [JsonProperty("ServiceLevel")]
        public string ServiceLevel { get; set; }

        [JsonProperty("FreightTermsCode")]
        public string FreightTermsCode { get; set; }

        [JsonProperty("FreightTerms")]
        public string FreightTerms { get; set; }

        [JsonProperty("PayOnReceiptFlag")]
        public string PayOnReceiptFlag { get; set; }

        [JsonProperty("FOBCode")]
        public string FobCode { get; set; }

        [JsonProperty("FOB")]
        public string Fob { get; set; }

        [JsonProperty("CountryOfOriginCode")]
        public string CountryOfOriginCode { get; set; }

        [JsonProperty("CountryOfOrigin")]
        public string CountryOfOrigin { get; set; }

        [JsonProperty("BuyerManagedTransportationCode")]
        public string BuyerManagedTransportationCode { get; set; }

        [JsonProperty("BuyerManagedTransportation")]
        public string BuyerManagedTransportation { get; set; }

        [JsonProperty("PayOnUseFlag")]
        public string PayOnUseFlag { get; set; }

        [JsonProperty("AgingOnsetPointCode")]
        public string AgingOnsetPointCode { get; set; }

        [JsonProperty("AgingOnsetPoint")]
        public string AgingOnsetPoint { get; set; }

        [JsonProperty("AgingPeriodDays")]
        public string AgingPeriodDays { get; set; }

        [JsonProperty("ConsumptionAdviceFrequencyCode")]
        public string ConsumptionAdviceFrequencyCode { get; set; }

        [JsonProperty("ConsumptionAdviceFrequency")]
        public string ConsumptionAdviceFrequency { get; set; }

        [JsonProperty("ConsumptionAdviceSummaryCode")]
        public string ConsumptionAdviceSummaryCode { get; set; }

        [JsonProperty("ConsumptionAdviceSummary")]
        public string ConsumptionAdviceSummary { get; set; }

        [JsonProperty("AlternatePaySiteId")]
        public string AlternatePaySiteId { get; set; }

        [JsonProperty("AlternatePaySite")]
        public string AlternatePaySite { get; set; }

        [JsonProperty("InvoiceSummaryLevelCode")]
        public string InvoiceSummaryLevelCode { get; set; }

        [JsonProperty("InvoiceSummaryLevel")]
        public string InvoiceSummaryLevel { get; set; }

        [JsonProperty("GaplessInvoiceNumberingFlag")]
        public string GaplessInvoiceNumberingFlag { get; set; }

        [JsonProperty("SellingCompanyIdentifier")]
        public string SellingCompanyIdentifier { get; set; }

        [JsonProperty("CreateDebitMemoFromReturnFlag")]
        public string CreateDebitMemoFromReturnFlag { get; set; }

        [JsonProperty("ShipToExceptionCode")]
        public string ShipToExceptionCode { get; set; }

        [JsonProperty("ShipToException")]
        public string ShipToException { get; set; }

        [JsonProperty("ReceiptRoutingId")]
        public string ReceiptRoutingId { get; set; }

        [JsonProperty("ReceiptRouting")]
        public string ReceiptRouting { get; set; }

        [JsonProperty("OverReceiptTolerance")]
        public string OverReceiptTolerance { get; set; }

        [JsonProperty("OverReceiptActionCode")]
        public string OverReceiptActionCode { get; set; }

        [JsonProperty("OverReceiptAction")]
        public string OverReceiptAction { get; set; }

        [JsonProperty("EarlyReceiptToleranceInDays")]
        public string EarlyReceiptToleranceInDays { get; set; }

        [JsonProperty("LateReceiptToleranceInDays")]
        public string LateReceiptToleranceInDays { get; set; }

        [JsonProperty("AllowSubstituteReceiptsCode")]
        public string AllowSubstituteReceiptsCode { get; set; }

        [JsonProperty("AllowSubstituteReceipts")]
        public string AllowSubstituteReceipts { get; set; }

        [JsonProperty("AllowUnorderedReceiptsFlag")]
        public string AllowUnorderedReceiptsFlag { get; set; }

        [JsonProperty("ReceiptDateExceptionCode")]
        public string ReceiptDateExceptionCode { get; set; }

        [JsonProperty("ReceiptDateException")]
        public string ReceiptDateException { get; set; }

        [JsonProperty("InvoiceCurrencyCode")]
        public string InvoiceCurrencyCode { get; set; }

        [JsonProperty("InvoiceCurrency")]
        public string InvoiceCurrency { get; set; }

        [JsonProperty("InvoiceAmountLimit")]
        public string InvoiceAmountLimit { get; set; }

        [JsonProperty("InvoiceMatchOptionCode")]
        public string InvoiceMatchOptionCode { get; set; }

        [JsonProperty("InvoiceMatchOption")]
        public string InvoiceMatchOption { get; set; }

        [JsonProperty("MatchApprovalLevelCode")]
        public string MatchApprovalLevelCode { get; set; }

        [JsonProperty("MatchApprovalLevel")]
        public string MatchApprovalLevel { get; set; }

        [JsonProperty("QuantityTolerancesId")]
        public string QuantityTolerancesId { get; set; }

        [JsonProperty("QuantityTolerances")]
        public string QuantityTolerances { get; set; }

        [JsonProperty("AmountTolerancesId")]
        public string AmountTolerancesId { get; set; }

        [JsonProperty("AmountTolerances")]
        public string AmountTolerances { get; set; }

        [JsonProperty("InvoiceChannelCode")]
        public string InvoiceChannelCode { get; set; }

        [JsonProperty("InvoiceChannel")]
        public string InvoiceChannel { get; set; }

        [JsonProperty("PaymentCurrencyCode")]
        public string PaymentCurrencyCode { get; set; }

        [JsonProperty("PaymentCurrency")]
        public string PaymentCurrency { get; set; }

        [JsonProperty("PaymentPriority")]
        public string PaymentPriority { get; set; }

        [JsonProperty("PayGroupCode")]
        public string PayGroupCode { get; set; }

        [JsonProperty("PayGroup")]
        public string PayGroup { get; set; }

        [JsonProperty("HoldAllInvoicesFlag")]
        public string HoldAllInvoicesFlag { get; set; }

        [JsonProperty("HoldUnmatchedInvoicesCode")]
        public string HoldUnmatchedInvoicesCode { get; set; }

        [JsonProperty("HoldUnmatchedInvoices")]
        public string HoldUnmatchedInvoices { get; set; }

        [JsonProperty("HoldUnvalidatedInvoicesFlag")]
        public string HoldUnvalidatedInvoicesFlag { get; set; }

        [JsonProperty("PaymentHoldDate")]
        public string PaymentHoldDate { get; set; }

        [JsonProperty("PaymentHoldReason")]
        public string PaymentHoldReason { get; set; }

        [JsonProperty("PaymentTermsId")]
        public string PaymentTermsId { get; set; }

        [JsonProperty("PaymentTerms")]
        public string PaymentTerms { get; set; }

        [JsonProperty("PaymentTermsDateBasisCode")]
        public string PaymentTermsDateBasisCode { get; set; }

        [JsonProperty("PaymentTermsDateBasis")]
        public string PaymentTermsDateBasis { get; set; }

        [JsonProperty("PayDateBasisCode")]
        public string PayDateBasisCode { get; set; }

        [JsonProperty("PayDateBasis")]
        public string PayDateBasis { get; set; }

        [JsonProperty("BankChargeDeductionTypeCode")]
        public string BankChargeDeductionTypeCode { get; set; }

        [JsonProperty("BankChargeDeductionType")]
        public string BankChargeDeductionType { get; set; }

        [JsonProperty("AlwaysTakeDiscountCode")]
        public string AlwaysTakeDiscountCode { get; set; }

        [JsonProperty("AlwaysTakeDiscount")]
        public string AlwaysTakeDiscount { get; set; }

        [JsonProperty("ExcludeFreightFromDiscountCode")]
        public string ExcludeFreightFromDiscountCode { get; set; }

        [JsonProperty("ExcludeFreightFromDiscount")]
        public string ExcludeFreightFromDiscount { get; set; }

        [JsonProperty("ExcludeTaxFromDiscountCode")]
        public string ExcludeTaxFromDiscountCode { get; set; }

        [JsonProperty("ExcludeTaxFromDiscount")]
        public string ExcludeTaxFromDiscount { get; set; }

        [JsonProperty("CreateInterestInvoicesCode")]
        public string CreateInterestInvoicesCode { get; set; }

        [JsonProperty("CreateInterestInvoices")]
        public string CreateInterestInvoices { get; set; }

        [JsonProperty("links")]
        public SupplierSiteLink[] Links { get; set; }
    }

    public class SupplierSiteLink
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
        public SupplierSiteProperties Properties { get; set; }
    }

    public class SupplierSiteProperties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }



}

