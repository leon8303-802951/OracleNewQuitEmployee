using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OracleNewQuitEmployee.SCS
{

    using Newtonsoft.Json;
    using Newtonsoft.Json.Converters;

    public class ScsLoginReturnModel
    {
        [JsonProperty("SessionGuid")]
        public string SessionGuid { get; set; }

        [JsonProperty("SessionInfo")]
        public SessionInfo SessionInfo { get; set; }

        [JsonProperty("IsPasswordChange")]
        public bool IsPasswordChange { get; set; }

        [JsonProperty("PasswordChangeMessage")]
        public string PasswordChangeMessage { get; set; }

        [JsonProperty("CloudEnabled")]
        public bool CloudEnabled { get; set; }

        [JsonProperty("LoginRedirects")]
        public object LoginRedirects { get; set; }

        [JsonProperty("RedirectUrl")]
        public Uri RedirectUrl { get; set; }

        [JsonProperty("Result")]
        public bool Result { get; set; }

        [JsonProperty("Message")]
        public string Message { get; set; }
    }

    public class SessionInfo
    {
        [JsonProperty("SessionGuid")]
        public Guid SessionGuid { get; set; }

        [JsonProperty("AppName")]
        public string AppName { get; set; }

        [JsonProperty("DatabaseID")]
        public string DatabaseId { get; set; }

        [JsonProperty("DatabaseType")]
        public string DatabaseType { get; set; }

        [JsonProperty("CompanyID")]
        public string CompanyId { get; set; }

        [JsonProperty("DecCompanyID")]
        public string DecCompanyId { get; set; }

        [JsonProperty("CompanyName")]
        public string CompanyName { get; set; }

        [JsonProperty("CustomizedID")]
        public string CustomizedId { get; set; }

        [JsonProperty("CloudEnabled")]
        public bool CloudEnabled { get; set; }

        [JsonProperty("FlowRevoke")]
        public bool FlowRevoke { get; set; }

        [JsonProperty("LanguageID")]
        public string LanguageId { get; set; }

        [JsonProperty("ViewName")]
        public string ViewName { get; set; }

        [JsonProperty("UtcOffset")]
        public double UtcOffset { get; set; }

        [JsonProperty("UserID")]
        public string UserId { get; set; }

        [JsonProperty("UserName")]
        public string UserName { get; set; }

        [JsonProperty("SourceUserID")]
        public string SourceUserId { get; set; }

        [JsonProperty("SourceUserName")]
        public string SourceUserName { get; set; }

        [JsonProperty("UserHostAddress")]
        public string UserHostAddress { get; set; }

        [JsonProperty("MachineName")]
        public string MachineName { get; set; }

        [JsonProperty("Roles")]
        public string Roles { get; set; }

        [JsonProperty("IsAdmin")]
        public bool IsAdmin { get; set; }

        [JsonProperty("IsSuperAdmin")]
        public bool IsSuperAdmin { get; set; }

        [JsonProperty("IsHRAdmin")]
        public bool IsHrAdmin { get; set; }

        [JsonProperty("IsHRGroup")]
        public bool IsHrGroup { get; set; }

        [JsonProperty("IsITAdmin")]
        public bool IsItAdmin { get; set; }

        [JsonProperty("IsDataPermissionFilter")]
        public bool IsDataPermissionFilter { get; set; }

        [JsonProperty("IsPermissionSettings")]
        public bool IsPermissionSettings { get; set; }

        [JsonProperty("IsDataPermissionUserFields")]
        public bool IsDataPermissionUserFields { get; set; }

        [JsonProperty("IsDataPermissionSysDepartmentID")]
        public bool IsDataPermissionSysDepartmentId { get; set; }

        [JsonProperty("UserType")]
        public string UserType { get; set; }

        [JsonProperty("TargetID")]
        public string TargetId { get; set; }

        [JsonProperty("Employee")]
        public Dictionary<string, string> Employee { get; set; }

        [JsonProperty("Departments")]
        public SCSLoginReturnDepartment[] Departments { get; set; }

        [JsonProperty("LastTime")]
        public string LastTime { get; set; }

        [JsonProperty("Timeout")]
        public long Timeout { get; set; }

        [JsonProperty("ExtendedProperties")]
        public object[] ExtendedProperties { get; set; }
    }

    public class SCSLoginReturnDepartment
    {
        [JsonProperty("$type")]
        public string Type { get; set; }

        [JsonProperty("ID")]
        public string Id { get; set; }

        [JsonProperty("ViewID")]
        public string ViewId { get; set; }

        [JsonProperty("Name")]
        public string Name { get; set; }

        [JsonProperty("Level")]
        public string Level { get; set; }
    }







}
