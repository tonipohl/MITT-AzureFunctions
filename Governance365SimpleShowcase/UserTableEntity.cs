using System;
using Newtonsoft.Json;
using Microsoft.Azure.Cosmos.Table;

namespace Governance365SimpleShowcase
{
    public class UserTableEntity : TableEntity
    {
        // ReSharper disable once EmptyConstructor
        public UserTableEntity() { }
        [JsonProperty("id")]
        public Guid Id { get; set; }

        [JsonProperty("deletedDateTime")]
        public DateTimeOffset? DeletedDateTime { get; set; }

        [JsonProperty("accountEnabled")]
        public string AccountEnabled { get; set; }

        [JsonProperty("ageGroup")]
        public string AgeGroup { get; set; }

        [JsonProperty("city")]
        public string City { get; set; }

        [JsonProperty("createdDateTime")]
        public DateTimeOffset? CreatedDateTime { get; set; }

        [JsonProperty("companyName")]
        public string CompanyName { get; set; }

        [JsonProperty("country")]
        public string Country { get; set; }

        [JsonProperty("department")]
        public string Department { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("employeeId")]
        public string EmployeeId { get; set; }

        [JsonProperty("faxNumber")]
        public string FaxNumber { get; set; }

        [JsonProperty("givenName")]
        public string GivenName { get; set; }

        [JsonProperty("isResourceAccount")]
        public string IsResourceAccount { get; set; }

        [JsonProperty("jobTitle")]
        public string JobTitle { get; set; }

        [JsonProperty("mail")]
        public string Mail { get; set; }

        [JsonProperty("mailNickname")]
        public string MailNickname { get; set; }

        [JsonProperty("mobilePhone")]
        public string MobilePhone { get; set; }

        [JsonProperty("onPremisesDistinguishedName")]
        public string OnPremisesDistinguishedName { get; set; }

        [JsonProperty("officeLocation")]
        public string OfficeLocation { get; set; }

        [JsonProperty("onPremisesDomainName")]
        public string OnPremisesDomainName { get; set; }

        [JsonProperty("onPremisesImmutableId")]
        public string OnPremisesImmutableId { get; set; }

        [JsonProperty("onPremisesLastSyncDateTime")]
        public DateTimeOffset? OnPremisesLastSyncDateTime { get; set; }

        [JsonProperty("onPremisesSamAccountName")]
        public string OnPremisesSamAccountName { get; set; }

        [JsonProperty("onPremisesSyncEnabled")]
        public string OnPremisesSyncEnabled { get; set; }

        [JsonProperty("onPremisesUserPrincipalName")]
        public string OnPremisesUserPrincipalName { get; set; }

        [JsonProperty("passwordPolicies")]
        public string PasswordPolicies { get; set; }

        [JsonProperty("postalCode")]
        public string PostalCode { get; set; }

        [JsonProperty("preferredDataLocation")]
        public string PreferredDataLocation { get; set; }

        [JsonProperty("preferredLanguage")]
        public string PreferredLanguage { get; set; }

        [JsonProperty("refreshTokensValidFromDateTime")]
        public DateTimeOffset? RefreshTokensValidFromDateTime { get; set; }

        [JsonProperty("showInAddressList")]
        public string ShowInAddressList { get; set; }

        [JsonProperty("signInSessionsValidFromDateTime")]
        public DateTimeOffset? SignInSessionsValidFromDateTime { get; set; }

        [JsonProperty("state")]
        public string State { get; set; }

        [JsonProperty("streetAddress")]
        public string StreetAddress { get; set; }

        [JsonProperty("surname")]
        public string Surname { get; set; }

        [JsonProperty("usageLocation")]
        public string UsageLocation { get; set; }

        [JsonProperty("userPrincipalName")]
        public string UserPrincipalName { get; set; }

        [JsonProperty("externalUserState")]
        public string ExternalUserState { get; set; }

        [JsonProperty("externalUserStateChangeDateTime")]
        public DateTimeOffset? ExternalUserStateChangeDateTime { get; set; }

        [JsonProperty("userType")]
        public string UserType { get; set; }
    }
}
