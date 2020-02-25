using Microsoft.Azure.Cosmos.Table;

namespace Governance365SimpleShowcase
{
    internal class GroupMemberTableEntity : TableEntity
    {
        public string GroupId { get; set; }
        public string GroupDisplayName { get; set; }
        public string GroupMailNickname { get; set; }
        public string Id { get; set; }
        public string UPN { get; set; }
        public string DisplayName { get; set; }
        public string AccountEnabled { get; set; }
    }
}
