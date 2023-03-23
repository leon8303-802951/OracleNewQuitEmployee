using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.ORSyncOracleData.Model
{
 


    using System.Globalization;
    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public class SupplierAddrReturnModel
    {
        [JsonProperty("items")]
        public SupplierAddr[] Items { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public int Limit { get; set; }

        [JsonProperty("offset")]
        public int Offset { get; set; }

        [JsonProperty("links")]
        public SupplierAddrLink[] Links { get; set; }
    }

    public class SupplierAddr
    {
        [JsonProperty("SupplierAddressId")]
        public string SupplierAddressId { get; set; }

        [JsonProperty("AddressName")]
        public string AddressName { get; set; }

        [JsonProperty("CountryCode")]
        public string CountryCode { get; set; }

        [JsonProperty("Country")]
        public string Country { get; set; }

        [JsonProperty("AddressLine1")]
        public string AddressLine1 { get; set; }

        [JsonProperty("AddressLine2")]
        public string AddressLine2 { get; set; }

        [JsonProperty("AddressLine3")]
        public string AddressLine3 { get; set; }

        [JsonProperty("AddressLine4")]
        public string AddressLine4 { get; set; }

        [JsonProperty("City")]
        public string City { get; set; }

        [JsonProperty("State")]
        public string State { get; set; }

        [JsonProperty("PostalCode")]
        public string PostalCode { get; set; }

        [JsonProperty("PostalCodeExtension")]
        public string PostalCodeExtension { get; set; }

        [JsonProperty("Province")]
        public string Province { get; set; }

        [JsonProperty("County")]
        public string County { get; set; }

        [JsonProperty("Building")]
        public string Building { get; set; }

        [JsonProperty("FloorNumber")]
        public string FloorNumber { get; set; }

        [JsonProperty("PhoneticAddress")]
        public string PhoneticAddress { get; set; }

        [JsonProperty("LanguageCode")]
        public string LanguageCode { get; set; }

        [JsonProperty("Language")]
        public string Language { get; set; }

        [JsonProperty("Addressee")]
        public string Addressee { get; set; }

        [JsonProperty("GlobalLocationNumber")]
        public string GlobalLocationNumber { get; set; }

        [JsonProperty("AdditionalAddressAttribute1")]
        public string AdditionalAddressAttribute1 { get; set; }

        [JsonProperty("AdditionalAddressAttribute2")]
        public string AdditionalAddressAttribute2 { get; set; }

        [JsonProperty("AdditionalAddressAttribute3")]
        public string AdditionalAddressAttribute3 { get; set; }

        [JsonProperty("AdditionalAddressAttribute4")]
        public string AdditionalAddressAttribute4 { get; set; }

        [JsonProperty("AdditionalAddressAttribute5")]
        public string AdditionalAddressAttribute5 { get; set; }

        [JsonProperty("AddressPurposeOrderingFlag")]
        public bool AddressPurposeOrderingFlag { get; set; }

        [JsonProperty("AddressPurposeRemitToFlag")]
        public bool AddressPurposeRemitToFlag { get; set; }

        [JsonProperty("AddressPurposeRFQOrBiddingFlag")]
        public bool AddressPurposeRfqOrBiddingFlag { get; set; }

        [JsonProperty("PhoneCountryCode")]
        public string PhoneCountryCode { get; set; }

        [JsonProperty("PhoneAreaCode")]
        public string PhoneAreaCode { get; set; }

        [JsonProperty("PhoneNumber")]
        public string PhoneNumber { get; set; }

        [JsonProperty("PhoneExtension")]
        public string PhoneExtension { get; set; }

        [JsonProperty("FaxCountryCode")]
        public string FaxCountryCode { get; set; }

        [JsonProperty("FaxAreaCode")]
        public string FaxAreaCode { get; set; }

        [JsonProperty("FaxNumber")]
        public string FaxNumber { get; set; }

        [JsonProperty("Email")]
        public string Email { get; set; }

        [JsonProperty("InactiveDate")]
        public string InactiveDate { get; set; }

        [JsonProperty("Status")]
        public string Status { get; set; }

        [JsonProperty("links")]
        public SupplierAddrLink[] Links { get; set; }
    }

    public class SupplierAddrLink
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
        public SupplierAddrProperties Properties { get; set; }
    }

    public class SupplierAddrProperties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }


}

