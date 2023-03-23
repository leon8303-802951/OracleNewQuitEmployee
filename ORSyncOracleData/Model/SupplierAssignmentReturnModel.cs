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

    public class SupplierAssignmentReturnModel
    {
        [JsonProperty("items")]
        public SupplierAssignment[] Items { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public int Limit { get; set; }

        [JsonProperty("offset")]
        public int Offset { get; set; }

        [JsonProperty("links")]
        public SupplierAssignmentLink[] Links { get; set; }
    }

    public class SupplierAssignment
    {
        [JsonProperty("AssignmentId")]
        public string AssignmentId { get; set; }

        [JsonProperty("ClientBUId")]
        public string ClientBuId { get; set; }

        [JsonProperty("ClientBU")]
        public string ClientBu { get; set; }

        [JsonProperty("BillToBUId")]
        public string BillToBuId { get; set; }

        [JsonProperty("BillToBU")]
        public string BillToBu { get; set; }

        [JsonProperty("ShipToLocationId")]
        public string ShipToLocationId { get; set; }

        [JsonProperty("ShipToLocation")]
        public string ShipToLocation { get; set; }

        [JsonProperty("ShipToLocationCode")]
        public string ShipToLocationCode { get; set; }

        [JsonProperty("BillToLocationId")]
        public string BillToLocationId { get; set; }

        [JsonProperty("BillToLocation")]
        public string BillToLocation { get; set; }

        [JsonProperty("BillToLocationCode")]
        public string BillToLocationCode { get; set; }

        [JsonProperty("UseWithholdingTaxFlag")]
        public bool UseWithholdingTaxFlag { get; set; }

        [JsonProperty("WithholdingTaxGroupId")]
        public string WithholdingTaxGroupId { get; set; }

        [JsonProperty("WithholdingTaxGroup")]
        public string WithholdingTaxGroup { get; set; }

        [JsonProperty("ChartOfAccountsId")]
        public string ChartOfAccountsId { get; set; }

        [JsonProperty("LiabilityDistributionId")]
        public string LiabilityDistributionId { get; set; }

        [JsonProperty("LiabilityDistribution")]
        public string LiabilityDistribution { get; set; }

        [JsonProperty("PrepaymentDistributionId")]
        public string PrepaymentDistributionId { get; set; }

        [JsonProperty("PrepaymentDistribution")]
        public string PrepaymentDistribution { get; set; }

        [JsonProperty("BillPayableDistributionId")]
        public string BillPayableDistributionId { get; set; }

        [JsonProperty("BillPayableDistribution")]
        public string BillPayableDistribution { get; set; }

        [JsonProperty("DistributionSetId")]
        public string DistributionSetId { get; set; }

        [JsonProperty("DistributionSet")]
        public string DistributionSet { get; set; }

        [JsonProperty("InactiveDate")]
        public string InactiveDate { get; set; }

        [JsonProperty("Status")]
        public string Status { get; set; }

        [JsonProperty("links")]
        public SupplierAssignmentLink[] Links { get; set; }
    }

    public class SupplierAssignmentLink
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
        public SupplierAssignmentProperties Properties { get; set; }
    }

    public class SupplierAssignmentProperties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }



}

