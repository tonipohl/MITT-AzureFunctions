using System;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using System.Net.Http;
using Microsoft.Azure.Cosmos.Table;
using Newtonsoft.Json.Linq;

// Governance365SimpleShowcase
// Demo solution for working with Azure Functions and Microsoft 365
// Feb. 2020, atwork.at, Jörg Schoba, Toni Pohl
// Repository https://github.com/tonipohl

namespace Governance365SimpleShowcase
{
    // Start collecting groups data
    public static class StartGroupSync
    {
        private static readonly string AppId = Environment.GetEnvironmentVariable("AppId");
        private static readonly string AppSecret = Environment.GetEnvironmentVariable("AppSecret");
        private static readonly string TenantId = Environment.GetEnvironmentVariable("TenantId");
        private static readonly string StorageAccount = Environment.GetEnvironmentVariable("StorageConnectionString");
        private static readonly string[] Scopes = { "https://graph.microsoft.com/.default" };
        private const string groupsUrl = "https://graph.microsoft.com/v1.0/groups";
        const string groupsSelection = "?$select=Id,DisplayName,Classification,CreatedDateTime,Description,DeletedDateTime,GroupTypes,Mail,MailEnabled,MailNickname,OnPremisesSyncEnabled,RenewedDateTime,SecurityEnabled,Visibility,resourceProvisioningOptions,resourceBehaviorOptions,creationOptions";

        [FunctionName("StartGroupSync")]
        public static async Task Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {

            var storage = CloudStorageAccount.Parse(StorageAccount);
            var cloudTableClient = storage.CreateCloudTableClient();
            var groupsStatisticsTable = cloudTableClient.GetTableReference("groupsstatistics");
            var groupsOwnersTable = cloudTableClient.GetTableReference("groupowners");
            var groupsGuestsTable = cloudTableClient.GetTableReference("groupguests");
            await groupsStatisticsTable.CreateIfNotExistsAsync().ConfigureAwait(false);
            await groupsOwnersTable.CreateIfNotExistsAsync().ConfigureAwait(false);
            await groupsGuestsTable.CreateIfNotExistsAsync().ConfigureAwait(false);

            // Get a Bearer Token with the App
            var httpClient = new HttpClient();
            string token;
            try
            {
                var app = ConfidentialClientApplicationBuilder.Create(AppId)
                    .WithAuthority($"https://login.windows.net/{TenantId}/oauth2/token")
                    .WithRedirectUri("https://governance365function")
                    .WithClientSecret(AppSecret)
                    .Build();

                var authenticationProvider = new MsalAuthenticationProvider(app, Scopes);
                token = authenticationProvider.GetTokenAsync().Result;
                httpClient.DefaultRequestHeaders.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);
            }
            catch (Exception)
            {
                throw;
            }

            httpClient.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", token);

            JObject groupsPageJson = null;
            do
            {
                if (string.IsNullOrEmpty(groupsPageJson?["@odata.nextLink"]?.ToString()))
                {
                    // first group page
                    groupsPageJson = JObject.Parse(await (await httpClient.GetAsync(groupsUrl + groupsSelection).ConfigureAwait(false)).Content.ReadAsStringAsync().ConfigureAwait(false));
                }
                else
                {
                    // next group page
                    groupsPageJson = JObject.Parse(await (await httpClient.GetAsync(groupsPageJson?["@odata.nextLink"]?.ToString()).ConfigureAwait(false)).Content.ReadAsStringAsync().ConfigureAwait(false));
                }
                if (!(groupsPageJson["value"] is JArray groupsPage)) continue;

                // Batch operation to store pages of groups in Azure table storage
                var groupPageBatchOperation = new TableBatchOperation();
                foreach (var group in groupsPage)
                {
                    var memberCount = 0;
                    var ownerCount = 0;
                    var guestCount = 0;

                    #region PROCESS GROUP TYPE AND DYNAMIC MEMBERSHIP

                    var groupType = string.Empty;
                    var dynamicMemberShip = false;
                    var yammerProvisionedOffice365Group = false;

                    foreach (var type in (JArray)group["groupTypes"])
                    {
                        if (type.ToString().Equals("Unified"))
                        {
                            if (group["resourceProvisioningOptions"].ToString().Contains("Team"))
                            {
                                groupType = "Team";
                            }
                            else
                            {
                                groupType = "Office365 Group";
                            }

                            if (group["creationOptions"].ToString().Contains("YammerProvisioning"))
                            {
                                yammerProvisionedOffice365Group = true;
                            }
                        }

                        if (type.ToString().Equals("DynamicMembership"))
                        {
                            dynamicMemberShip = true;
                        }
                    }

                    if (groupType.Equals(string.Empty))
                    {
                        if (!group["mailEnabled"].ToString().ToLower().Equals("true"))
                        {
                            groupType = "Security Group";
                        }
                        else
                        {
                            if (!group["securityEnabled"].ToString().ToLower().Equals("true"))
                            {
                                groupType = "Distribution Group";
                            }
                            else
                            {
                                groupType = "Mail-enabled Security Group";
                            }
                        }
                    }
                    #endregion

                    const string groupMemberSelection = "?$select=id,userPrincipalName,displayName,userType,accountEnabled";
                    var groupMembersUrl = $"https://graph.microsoft.com/v1.0/groups/{group["id"]?.ToString()}/members{groupMemberSelection}";
                    var groupOwnersUrl = $"https://graph.microsoft.com/v1.0/groups/{group["id"]?.ToString()}/owners{groupMemberSelection}";

                    #region PROCESS GROUP MEMBERS AND GUESTS

                    var processMemberPageResult = await ProcessGroupMemberPageAsync(groupMembersUrl, group, httpClient, groupsGuestsTable, groupsOwnersTable, false).ConfigureAwait(false);
                    memberCount += processMemberPageResult.MemberCount;
                    guestCount += processMemberPageResult.GuestCount;

                    while (processMemberPageResult.NextLink != null)
                    {
                        processMemberPageResult = await ProcessGroupMemberPageAsync(processMemberPageResult.NextLink, group, httpClient, groupsGuestsTable, groupsOwnersTable, false).ConfigureAwait(false);
                        memberCount += processMemberPageResult.MemberCount;
                        guestCount += processMemberPageResult.GuestCount;
                    }
                    #endregion

                    #region PROCESS GROUP OWNERS

                    var processOwnerPageResult = await ProcessGroupMemberPageAsync(groupOwnersUrl, group, httpClient, groupsGuestsTable, groupsOwnersTable, true).ConfigureAwait(false);
                    ownerCount += processOwnerPageResult.OwnerCount;
                    while (processOwnerPageResult.NextLink != null)
                    {
                        processOwnerPageResult = await ProcessGroupMemberPageAsync(processOwnerPageResult.NextLink, group, httpClient, groupsGuestsTable, groupsOwnersTable, true).ConfigureAwait(false);
                        ownerCount += processOwnerPageResult.OwnerCount;
                    }
                    #endregion

                    var groupStatisticsEntity = new GroupStatisticsItemTableEntity()
                    {
                        PartitionKey = "group",
                        RowKey = group["id"]?.ToString(),
                        Classification = group["classification"]?.ToString(),
                        CreatedDateTime = group["createdDateTime"] != null &&
                                              !group["createdDateTime"].ToString().Equals(string.Empty)
                                ? DateTimeOffset.Parse(group["createdDateTime"].ToString())
                                : (DateTimeOffset?)null,
                        DeletedDateTime = group["deletedDateTime"] != null &&
                                              !group["deletedDateTime"].ToString().Equals(string.Empty)
                                ? DateTimeOffset.Parse(group["deletedDateTime"].ToString())
                                : (DateTimeOffset?)null,
                        Description = group["description"]?.ToString(),
                        DisplayName = group["displayName"]?.ToString(),
                        GroupType = groupType,
                        DynamicMembership = dynamicMemberShip,
                        Id = group["id"]?.ToString(),
                        Mail = group["mail"]?.ToString(),
                        MailEnabled = group["mailEnabled"]?.ToString(),
                        MailNickname = group["mailNickname"]?.ToString(),
                        OnPremisesSyncEnabled = group["onPremisesSyncEnabled"]?.ToString(),
                        RenewedDateTime = group["renewedDateTime"] != null &
                                              !group["renewedDateTime"].ToString().Equals(string.Empty)
                                ? DateTimeOffset.Parse(group["renewedDateTime"].ToString())
                                : (DateTimeOffset?)null,
                        SecurityEnabled = group["securityEnabled"]?.ToString(),
                        Visibility = group["visibility"]?.ToString(),
                        Guests = guestCount,
                        Owners = ownerCount,
                        Members = memberCount
                    };
                    groupPageBatchOperation.Add(TableOperation.InsertOrReplace(groupStatisticsEntity));
                }
                await groupsStatisticsTable.ExecuteBatchAsync(groupPageBatchOperation).ConfigureAwait(false);

            } while (!string.IsNullOrEmpty(groupsPageJson["@odata.nextLink"]?.ToString()));
            httpClient.Dispose();
        }

        // Processes a page of members (including guests) or owners for one group and stores owners and guests to Azure table storage
        private static async Task<GroupMemberPageResult> ProcessGroupMemberPageAsync(string getMembersUrl, JToken group, HttpClient httpClient,
            CloudTable groupGuestsTable, CloudTable groupOwnersTable, bool isOwners)
        {
            var result = new GroupMemberPageResult();
            HttpResponseMessage groupMembersJson;
            groupMembersJson = await httpClient.GetAsync(getMembersUrl).ConfigureAwait(false);

            var groupMemberJsonObject = JObject.Parse(groupMembersJson.Content.ReadAsStringAsync().Result);

            var groupGuestsBatchOperation = new TableBatchOperation();
            var groupOwnersBatchOperation = new TableBatchOperation();
            if (groupMemberJsonObject["value"] is JArray groupMembers)
                foreach (var groupMember in groupMembers)
                {
                    var groupMemberTableEntity = new GroupMemberTableEntity()
                    {
                        PartitionKey = group["id"]?.ToString(),
                        RowKey = Guid.NewGuid().ToString(),
                        GroupId = group["id"]?.ToString(),
                        GroupDisplayName = group["displayName"]?.ToString(),
                        GroupMailNickname = group["mailNickname"]?.ToString(),
                        Id = groupMember["id"]?.ToString(),
                        UPN = groupMember["userPrincipalName"]?.ToString(),
                        DisplayName = groupMember["displayName"]?.ToString(),
                        AccountEnabled = groupMember["accountEnabled"]?.ToString()
                    };
                    if (isOwners)
                    {
                        result.OwnerCount++;
                        groupOwnersBatchOperation.Add(TableOperation.InsertOrReplace(groupMemberTableEntity));
                    }
                    else
                    {
                        result.MemberCount++;
                        //if member is a guest
                        if (groupMember["userType"] != null &&
                            groupMember["userType"].ToString().ToLower().Equals("guest"))
                        {
                            result.GuestCount++;
                            groupGuestsBatchOperation.Add(TableOperation.InsertOrReplace(groupMemberTableEntity));
                        }
                    }
                }

            if (isOwners && groupOwnersBatchOperation.Count > 0)
            {
                await groupOwnersTable.ExecuteBatchAsync(groupOwnersBatchOperation)
                    .ConfigureAwait(false);
            }
            else if (groupGuestsBatchOperation.Count > 0)
            {
                await groupGuestsTable.ExecuteBatchAsync(groupGuestsBatchOperation)
                    .ConfigureAwait(false);
            }

            result.NextLink = groupMemberJsonObject["@odata.nextLink"]?.ToString();
            return result;
        }

        private struct GroupMemberPageResult
        {
            public int MemberCount, OwnerCount, GuestCount;
            public string NextLink;
        }
    }
}
