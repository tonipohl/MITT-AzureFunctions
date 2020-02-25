using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.WindowsAzure.Storage.Table;
using Newtonsoft.Json;
// ReSharper disable MissingXmlDoc

namespace Governance365.Models
{
    public class UserQueryObject
    {

        [JsonProperty("@odata.nextLink")]
        public Uri OdataNextLink { get; set; }

        [JsonProperty("value")]
        public List<Governance365SimpleShowcase.UserTableEntity> Value { get; set; }
    }
}
