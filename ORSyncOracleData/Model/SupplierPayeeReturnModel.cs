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

    public class SupplierPayeeReturnModel
    {
        [JsonProperty("items")]
        public SupplierPayee[] Items { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public int Limit { get; set; }

        [JsonProperty("offset")]
        public int Offset { get; set; }

        [JsonProperty("links")]
        public SupplierPayeeLink[] Links { get; set; }
    }

    public class SupplierPayee
    {
        [JsonProperty("PayeeId")]
        public string PayeeId { get; set; }

        [JsonProperty("PaymentFunctionCode")]
        public string PaymentFunctionCode { get; set; }

        [JsonProperty("OrganizationIdentifier")]
        public string OrganizationIdentifier { get; set; }

        [JsonProperty("OrganizationName")]
        public string OrganizationName { get; set; }

        [JsonProperty("OrganizationType")]
        public string OrganizationType { get; set; }

        [JsonProperty("PayEachDocumentAloneOption")]
        public string PayEachDocumentAloneOption { get; set; }

        [JsonProperty("DefaultPaymentMethodCode")]
        public string DefaultPaymentMethodCode { get; set; }

        [JsonProperty("DefaultPaymentMethodName")]
        public string DefaultPaymentMethodName { get; set; }

        [JsonProperty("BankChargeBearerCode")]
        public string BankChargeBearerCode { get; set; }

        [JsonProperty("BankChargeBearerName")]
        public string BankChargeBearerName { get; set; }

        [JsonProperty("BankInstructionCodeOne")]
        public string BankInstructionCodeOne { get; set; }

        [JsonProperty("BankInstructionNameOne")]
        public string BankInstructionNameOne { get; set; }

        [JsonProperty("BankInstructionCodeTwo")]
        public string BankInstructionCodeTwo { get; set; }

        [JsonProperty("BankInstructionNameTwo")]
        public string BankInstructionNameTwo { get; set; }

        [JsonProperty("BankInstructionDetails")]
        public string BankInstructionDetails { get; set; }

        [JsonProperty("PaymentReasonCode")]
        public string PaymentReasonCode { get; set; }

        [JsonProperty("PaymentReasonName")]
        public string PaymentReasonName { get; set; }

        [JsonProperty("PaymentReasonComments")]
        public string PaymentReasonComments { get; set; }

        [JsonProperty("PaymentTextMessageOne")]
        public string PaymentTextMessageOne { get; set; }

        [JsonProperty("PaymentTextMessageTwo")]
        public string PaymentTextMessageTwo { get; set; }

        [JsonProperty("PaymentTextMessageThree")]
        public string PaymentTextMessageThree { get; set; }

        [JsonProperty("DeliveryChannelCode")]
        public string DeliveryChannelCode { get; set; }

        [JsonProperty("DeliveryChannelName")]
        public string DeliveryChannelName { get; set; }

        [JsonProperty("ServiceLevelCode")]
        public string ServiceLevelCode { get; set; }

        [JsonProperty("ServiceLevelName")]
        public string ServiceLevelName { get; set; }

        [JsonProperty("SettlementPriority")]
        public string SettlementPriority { get; set; }

        [JsonProperty("DeliveryMethod")]
        public string DeliveryMethod { get; set; }

        [JsonProperty("Email")]
        public string Email { get; set; }

        [JsonProperty("Fax")]
        public string Fax { get; set; }

        [JsonProperty("PayeePartyIdentifier")]
        public string PayeePartyIdentifier { get; set; }

        [JsonProperty("PartyName")]
        public string PartyName { get; set; }

        [JsonProperty("PayeePartyNumber")]
        public string PayeePartyNumber { get; set; }

        [JsonProperty("PayeePartySiteIdentifier")]
        public long? PayeePartySiteIdentifier { get; set; }

        [JsonProperty("SupplierNumber")]
        public string SupplierNumber { get; set; }

        [JsonProperty("SupplierSiteCode")]
        public string SupplierSiteCode { get; set; }

        [JsonProperty("SupplierSiteIdentifier")]
        public long? SupplierSiteIdentifier { get; set; }

        [JsonProperty("PayeePartySiteNumber")]
        public long? PayeePartySiteNumber { get; set; }

        [JsonProperty("PaymentFormatCode")]
        public string PaymentFormatCode { get; set; }

        [JsonProperty("PaymentFormatName")]
        public string PaymentFormatName { get; set; }

        [JsonProperty("PersonId")]
        public string PersonId { get; set; }

        [JsonProperty("Intent")]
        public string Intent { get; set; }

        [JsonProperty("links")]
        public SupplierPayeeLink[] Links { get; set; }
    }

    public class SupplierPayeeLink
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
        public SupplierPayeeProperties Properties { get; set; }
    }

    public class SupplierPayeeProperties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }




}

