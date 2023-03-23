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

    public partial class SupplierBankAccountReturnModel
    {
        [JsonProperty("items")]
        public SupplierBankAccount[] Items { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public int Limit { get; set; }

        [JsonProperty("offset")]
        public int Offset { get; set; }

        [JsonProperty("links")]
        public SupplierBankAccountLink[] Links { get; set; }
    }

    public partial class SupplierBankAccount
    {
        [JsonProperty("BankAccountId")]
        public string BankAccountId { get; set; }

        [JsonProperty("AccountNumber")]
        public string AccountNumber { get; set; }

        [JsonProperty("IBAN")]
        public string Iban { get; set; }

        [JsonProperty("AccountName")]
        public string AccountName { get; set; }

        [JsonProperty("AlternateAccountName")]
        public string AlternateAccountName { get; set; }

        [JsonProperty("AccountCountry")]
        public string AccountCountry { get; set; }

        [JsonProperty("CurrencyCode")]
        public string CurrencyCode { get; set; }

        [JsonProperty("AllowInternationalPaymentsIndicator")]
        public string AllowInternationalPaymentsIndicator { get; set; }

        [JsonProperty("BankName")]
        public string BankName { get; set; }

        [JsonProperty("BankCode")]
        public string BankCode { get; set; }

        [JsonProperty("BranchName")]
        public string BranchName { get; set; }

        [JsonProperty("BranchNumber")]
        public string BranchNumber { get; set; }

        [JsonProperty("SWIFTCode")]
        public string SwiftCode { get; set; }

        [JsonProperty("CheckDigits")]
        public string CheckDigits { get; set; }

        [JsonProperty("AccountSuffix")]
        public string AccountSuffix { get; set; }

        [JsonProperty("SecondaryAccountReference")]
        public string SecondaryAccountReference { get; set; }

        [JsonProperty("AgencyLocationCode")]
        public string AgencyLocationCode { get; set; }

        [JsonProperty("AccountType")]
        public string AccountType { get; set; }

        [JsonProperty("FactorAccountIndicator")]
        public string FactorAccountIndicator { get; set; }

        [JsonProperty("Description")]
        public string Description { get; set; }

        [JsonProperty("StartDate")]
        public string StartDate { get; set; }

        [JsonProperty("EndDate")]
        public string EndDate { get; set; }

        [JsonProperty("links")]
        public SupplierBankAccountLink[] Links { get; set; }
    }

    public partial class SupplierBankAccountLink
    {
        [JsonProperty("rel")]
        public string Rel { get; set; }

        [JsonProperty("href")]
        public Uri Href { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("kind")]
        public string Kind { get; set; }
    }



}

