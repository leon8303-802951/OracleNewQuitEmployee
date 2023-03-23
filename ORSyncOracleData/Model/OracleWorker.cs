using System;
using Newtonsoft.Json;
namespace OracleNewQuitEmployee.ORSyncOracleData.Model.OracleWorker
{
    public class TmpReturnObj20230119_1
    {
        [JsonProperty("items")]
        public OracleWorker[] Items { get; set; }

        [JsonProperty("count")]
        public int Count { get; set; }

        [JsonProperty("hasMore")]
        public bool HasMore { get; set; }

        [JsonProperty("limit")]
        public int Limit { get; set; }

        [JsonProperty("offset")]
        public int Offset { get; set; }

        [JsonProperty("links")]
        public Link[] Links { get; set; }
    }

    public class OracleWorker
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

        [JsonProperty("emails")]
        public Email[] Emails { get; set; }

        [JsonProperty("names")]
        public Name[] Names { get; set; }

        [JsonProperty("links")]
        public Link[] Links { get; set; }
    }

    public class Email
    {
        [JsonProperty("EmailAddressId")]
        public string EmailAddressId { get; set; }

        [JsonProperty("EmailType")]
        public string EmailType { get; set; }

        [JsonProperty("EmailAddress")]
        public string EmailAddress { get; set; }

        [JsonProperty("FromDate")]
        public string FromDate { get; set; }

        [JsonProperty("ToDate")]
        public object ToDate { get; set; }

        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        [JsonProperty("CreationDate")]
        public string CreationDate { get; set; }

        [JsonProperty("LastUpdatedBy")]
        public string LastUpdatedBy { get; set; }

        [JsonProperty("LastUpdateDate")]
        public string LastUpdateDate { get; set; }

        [JsonProperty("PrimaryFlag")]
        public bool PrimaryFlag { get; set; }

        [JsonProperty("links")]
        public Link[] Links { get; set; }
    }

    public class Link
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
        public Properties Properties { get; set; }
    }

    public class Properties
    {
        [JsonProperty("changeIndicator")]
        public string ChangeIndicator { get; set; }
    }

    public class Name
    {
        [JsonProperty("PersonNameId")]
        public string PersonNameId { get; set; }

        [JsonProperty("EffectiveStartDate")]
        public string EffectiveStartDate { get; set; }

        [JsonProperty("EffectiveEndDate")]
        public string EffectiveEndDate { get; set; }

        [JsonProperty("LegislationCode")]
        public string LegislationCode { get; set; }

        [JsonProperty("LastName")]
        public string LastName { get; set; }

        [JsonProperty("FirstName")]
        public string FirstName { get; set; }

        [JsonProperty("Title")]
        public object Title { get; set; }

        [JsonProperty("PreNameAdjunct")]
        public object PreNameAdjunct { get; set; }

        [JsonProperty("Suffix")]
        public object Suffix { get; set; }

        [JsonProperty("MiddleNames")]
        public string MiddleNames { get; set; }

        [JsonProperty("Honors")]
        public object Honors { get; set; }

        [JsonProperty("KnownAs")]
        public object KnownAs { get; set; }

        [JsonProperty("PreviousLastName")]
        public object PreviousLastName { get; set; }

        [JsonProperty("DisplayName")]
        public string DisplayName { get; set; }

        [JsonProperty("OrderName")]
        public string OrderName { get; set; }

        [JsonProperty("ListName")]
        public string ListName { get; set; }

        [JsonProperty("FullName")]
        public string FullName { get; set; }

        [JsonProperty("MilitaryRank")]
        public object MilitaryRank { get; set; }

        [JsonProperty("NameLanguage")]
        public string NameLanguage { get; set; }

        [JsonProperty("NameInformation1")]
        public object NameInformation1 { get; set; }

        [JsonProperty("NameInformation2")]
        public object NameInformation2 { get; set; }

        [JsonProperty("NameInformation3")]
        public object NameInformation3 { get; set; }

        [JsonProperty("NameInformation4")]
        public object NameInformation4 { get; set; }

        [JsonProperty("NameInformation5")]
        public object NameInformation5 { get; set; }

        [JsonProperty("NameInformation6")]
        public object NameInformation6 { get; set; }

        [JsonProperty("NameInformation7")]
        public object NameInformation7 { get; set; }

        [JsonProperty("NameInformation8")]
        public object NameInformation8 { get; set; }

        [JsonProperty("NameInformation9")]
        public object NameInformation9 { get; set; }

        [JsonProperty("NameInformation10")]
        public object NameInformation10 { get; set; }

        [JsonProperty("NameInformation11")]
        public object NameInformation11 { get; set; }

        [JsonProperty("NameInformation12")]
        public object NameInformation12 { get; set; }

        [JsonProperty("NameInformation13")]
        public object NameInformation13 { get; set; }

        [JsonProperty("NameInformation14")]
        public object NameInformation14 { get; set; }

        [JsonProperty("NameInformation15")]
        public object NameInformation15 { get; set; }

        [JsonProperty("NameInformation16")]
        public object NameInformation16 { get; set; }

        [JsonProperty("NameInformation17")]
        public object NameInformation17 { get; set; }

        [JsonProperty("NameInformation18")]
        public object NameInformation18 { get; set; }

        [JsonProperty("NameInformation19")]
        public object NameInformation19 { get; set; }

        [JsonProperty("NameInformation20")]
        public object NameInformation20 { get; set; }

        [JsonProperty("NameInformation21")]
        public object NameInformation21 { get; set; }

        [JsonProperty("NameInformation22")]
        public object NameInformation22 { get; set; }

        [JsonProperty("NameInformation23")]
        public object NameInformation23 { get; set; }

        [JsonProperty("NameInformation24")]
        public object NameInformation24 { get; set; }

        [JsonProperty("NameInformation25")]
        public object NameInformation25 { get; set; }

        [JsonProperty("NameInformation26")]
        public object NameInformation26 { get; set; }

        [JsonProperty("NameInformation27")]
        public object NameInformation27 { get; set; }

        [JsonProperty("NameInformation28")]
        public object NameInformation28 { get; set; }

        [JsonProperty("NameInformation29")]
        public object NameInformation29 { get; set; }

        [JsonProperty("NameInformation30")]
        public object NameInformation30 { get; set; }

        [JsonProperty("CreatedBy")]
        public string CreatedBy { get; set; }

        [JsonProperty("CreationDate")]
        public string CreationDate { get; set; }

        [JsonProperty("LastUpdatedBy")]
        public string LastUpdatedBy { get; set; }

        [JsonProperty("LastUpdateDate")]
        public string LastUpdateDate { get; set; }

        [JsonProperty("LocalPersonNameId")]
        public string LocalPersonNameId { get; set; }

        [JsonProperty("LocalEffectiveStartDate")]
        public string LocalEffectiveStartDate { get; set; }

        [JsonProperty("LocalEffectiveEndDate")]
        public string LocalEffectiveEndDate { get; set; }

        [JsonProperty("LocalLegislationCode")]
        public string LocalLegislationCode { get; set; }

        [JsonProperty("LocalLastName")]
        public string LocalLastName { get; set; }

        [JsonProperty("LocalFirstName")]
        public string LocalFirstName { get; set; }

        [JsonProperty("LocalTitle")]
        public object LocalTitle { get; set; }

        [JsonProperty("LocalPreNameAdjunct")]
        public object LocalPreNameAdjunct { get; set; }

        [JsonProperty("LocalSuffix")]
        public object LocalSuffix { get; set; }

        [JsonProperty("LocalMiddleNames")]
        public string LocalMiddleNames { get; set; }

        [JsonProperty("LocalHonors")]
        public object LocalHonors { get; set; }

        [JsonProperty("LocalKnownAs")]
        public object LocalKnownAs { get; set; }

        [JsonProperty("LocalPreviousLastName")]
        public object LocalPreviousLastName { get; set; }

        [JsonProperty("LocalDisplayName")]
        public string LocalDisplayName { get; set; }

        [JsonProperty("LocalOrderName")]
        public string LocalOrderName { get; set; }

        [JsonProperty("LocalListName")]
        public string LocalListName { get; set; }

        [JsonProperty("LocalFullName")]
        public string LocalFullName { get; set; }

        [JsonProperty("LocalMilitaryRank")]
        public object LocalMilitaryRank { get; set; }

        [JsonProperty("LocalNameLanguage")]
        public string LocalNameLanguage { get; set; }

        [JsonProperty("LocalNameInformation1")]
        public object LocalNameInformation1 { get; set; }

        [JsonProperty("LocalNameInformation2")]
        public object LocalNameInformation2 { get; set; }

        [JsonProperty("LocalNameInformation3")]
        public object LocalNameInformation3 { get; set; }

        [JsonProperty("LocalNameInformation4")]
        public object LocalNameInformation4 { get; set; }

        [JsonProperty("LocalNameInformation5")]
        public object LocalNameInformation5 { get; set; }

        [JsonProperty("LocalNameInformation6")]
        public object LocalNameInformation6 { get; set; }

        [JsonProperty("LocalNameInformation7")]
        public object LocalNameInformation7 { get; set; }

        [JsonProperty("LocalNameInformation8")]
        public object LocalNameInformation8 { get; set; }

        [JsonProperty("LocalNameInformation9")]
        public object LocalNameInformation9 { get; set; }

        [JsonProperty("LocalNameInformation10")]
        public object LocalNameInformation10 { get; set; }

        [JsonProperty("LocalNameInformation11")]
        public object LocalNameInformation11 { get; set; }

        [JsonProperty("LocalNameInformation12")]
        public object LocalNameInformation12 { get; set; }

        [JsonProperty("LocalNameInformation13")]
        public object LocalNameInformation13 { get; set; }

        [JsonProperty("LocalNameInformation14")]
        public object LocalNameInformation14 { get; set; }

        [JsonProperty("LocalNameInformation15")]
        public object LocalNameInformation15 { get; set; }

        [JsonProperty("LocalNameInformation16")]
        public object LocalNameInformation16 { get; set; }

        [JsonProperty("LocalNameInformation17")]
        public object LocalNameInformation17 { get; set; }

        [JsonProperty("LocalNameInformation18")]
        public object LocalNameInformation18 { get; set; }

        [JsonProperty("LocalNameInformation19")]
        public object LocalNameInformation19 { get; set; }

        [JsonProperty("LocalNameInformation20")]
        public object LocalNameInformation20 { get; set; }

        [JsonProperty("LocalNameInformation21")]
        public object LocalNameInformation21 { get; set; }

        [JsonProperty("LocalNameInformation22")]
        public object LocalNameInformation22 { get; set; }

        [JsonProperty("LocalNameInformation23")]
        public object LocalNameInformation23 { get; set; }

        [JsonProperty("LocalNameInformation24")]
        public object LocalNameInformation24 { get; set; }

        [JsonProperty("LocalNameInformation25")]
        public object LocalNameInformation25 { get; set; }

        [JsonProperty("LocalNameInformation26")]
        public object LocalNameInformation26 { get; set; }

        [JsonProperty("LocalNameInformation27")]
        public object LocalNameInformation27 { get; set; }

        [JsonProperty("LocalNameInformation28")]
        public object LocalNameInformation28 { get; set; }

        [JsonProperty("LocalNameInformation29")]
        public object LocalNameInformation29 { get; set; }

        [JsonProperty("LocalNameInformation30")]
        public object LocalNameInformation30 { get; set; }

        [JsonProperty("LocalCreatedBy")]
        public string LocalCreatedBy { get; set; }

        [JsonProperty("LocalCreationDate")]
        public string LocalCreationDate { get; set; }

        [JsonProperty("LocalLastUpdatedBy")]
        public string LocalLastUpdatedBy { get; set; }

        [JsonProperty("LocalLastUpdateDate")]
        public string LocalLastUpdateDate { get; set; }

        [JsonProperty("links")]
        public Link[] Links { get; set; }
    }

}