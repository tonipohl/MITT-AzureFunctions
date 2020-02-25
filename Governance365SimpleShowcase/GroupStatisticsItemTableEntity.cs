using System;
using Microsoft.Azure.Cosmos.Table;

namespace Governance365SimpleShowcase
{
    internal class GroupStatisticsItemTableEntity : TableEntity
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string GroupType { get; set; }
        public bool DynamicMembership { get; set; }
        public string Classification { get; set; }
        public DateTimeOffset? CreatedDateTime { get; set; }
        public DateTimeOffset? ExpirationDateTime { get; set; }
        public DateTimeOffset? DeletedDateTime { get; set; }
        public string Mail { get; set; }
        public string MailEnabled { get; set; }
        public string MailNickname { get; set; }
        public string OnPremisesSyncEnabled { get; set; }
        public DateTimeOffset? RenewedDateTime { get; set; }
        public string SecurityEnabled { get; set; }
        public string Visibility { get; set; }
        public string Description { get; set; }
        public int Owners { get; set; }
        public int Members { get; set; }
        public int Guests { get; set; }

        // ReSharper disable once EmptyConstructor
        public GroupStatisticsItemTableEntity() { }
    }
}
