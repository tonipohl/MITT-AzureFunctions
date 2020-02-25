
using Microsoft.Azure.Cosmos.Table;

namespace Governance365.Models
{
    internal class UserStatisticsTableEntity : TableEntity
    {
        public int Users { get; set; }
        public int InternalUsers { get; set; }
        public int GuestUsers { get; set; }
        public int DeactivatedUsers { get; set; }
        public int DeletedUsers { get; set; }
        public UserStatisticsTableEntity() { }
    }
}
