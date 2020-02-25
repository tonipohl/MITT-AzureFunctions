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
using Newtonsoft.Json;
using Governance365.Models;

// Governance365SimpleShowcase
// Demo solution for working with Azure Functions and Microsoft 365
// Feb. 2020, atwork.at, Jörg Schoba, Toni Pohl
// Repository https://github.com/tonipohl

namespace Governance365SimpleShowcase
{
    // Start collecting user data
    public static class StartUserSync
    {
        private static readonly string AppId = Environment.GetEnvironmentVariable("AppId");
        private static readonly string AppSecret = Environment.GetEnvironmentVariable("AppSecret");
        private static readonly string TenantId = Environment.GetEnvironmentVariable("TenantId");
        private static readonly string StorageAccount = Environment.GetEnvironmentVariable("StorageConnectionString");
        private static readonly string[] Scopes = { "https://graph.microsoft.com/.default" };
        private const string usersUrl = "https://graph.microsoft.com/v1.0/users";
        private const string UserSelectedProperties = "?$select=id,deletedDateTime,accountEnabled,ageGroup,city,createdDateTime,companyName," +
                                                       "country,department,displayName,employeeId,faxNumber," +
                                                       "givenName,isResourceAccount,jobTitle,mail,mailNickname," +
                                                       "mobilePhone,onPremisesDistinguishedName,officeLocation,onPremisesDomainName,onPremisesImmutableId," +
                                                       "onPremisesLastSyncDateTime,onPremisesSamAccountName,onPremisesSyncEnabled," +
                                                       "onPremisesUserPrincipalName,passwordPolicies,postalCode," +
                                                       "preferredLanguage,refreshTokensValidFromDateTime,showInAddressList,signInSessionsValidFromDateTime," +
                                                       "state,streetAddress,surname,usageLocation,userPrincipalName," +
                                                       "userType,assignedLicenses";

        [FunctionName("StartUserSync")]
        public static async Task Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            var storage = CloudStorageAccount.Parse(StorageAccount);
            var cloudTableClient = storage.CreateCloudTableClient();
            var usersTable = cloudTableClient.GetTableReference("users");
            var userStatisticsTable = cloudTableClient.GetTableReference("userstatistics");
            await usersTable.CreateIfNotExistsAsync().ConfigureAwait(false);
            await userStatisticsTable.CreateIfNotExistsAsync().ConfigureAwait(false);

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

            //initialize object for storing all user statistics
            UserQueryObject usersPage = null;
            var userStatistics = new UserStatisticsTableEntity()
            {
                PartitionKey = "userstatistics",
                RowKey = TenantId,
                DeactivatedUsers = 0,
                DeletedUsers = 0,
                GuestUsers = 0,
                InternalUsers = 0,
                Users = 0
            };

            //get all users (until nextlinnk is empty) and members/guests + sum up statistics
            do
            {
                if (string.IsNullOrEmpty(usersPage?.OdataNextLink?.ToString()))
                {
                    //first request
                    usersPage = JsonConvert.DeserializeObject<Governance365.Models.UserQueryObject>(await (await httpClient.GetAsync(usersUrl + UserSelectedProperties).ConfigureAwait(false)).Content.ReadAsStringAsync().ConfigureAwait(false));
                }
                else
                {
                    //next page request
                    usersPage = JsonConvert.DeserializeObject<Governance365.Models.UserQueryObject>(await (await httpClient.GetAsync(usersPage?.OdataNextLink?.ToString()).ConfigureAwait(false)).Content.ReadAsStringAsync().ConfigureAwait(false));
                }

                //batch request to store pages of users in Azure storage
                var userPageBatchOperation = new TableBatchOperation();
                foreach (var user in usersPage.Value)
                {
                    if (user.UserType != null)
                    {
                        userStatistics.Users++;
                        if (user.UserType.Equals("Member"))
                        {
                            userStatistics.InternalUsers++;
                        }
                        else if (user.UserType.Equals("Guest"))
                        {
                            userStatistics.GuestUsers++;
                        }
                        if (string.IsNullOrEmpty(user.AccountEnabled) && !bool.Parse(user.AccountEnabled))
                        {
                            userStatistics.DeactivatedUsers++;
                        }
                    }
                    user.PartitionKey = "user";
                    user.RowKey = user.Id.ToString();
                    //add user entity to batch operation
                    userPageBatchOperation.Add(TableOperation.InsertOrReplace(user));
                }
                //write user page to Azure tabel storage
                await usersTable.ExecuteBatchAsync(userPageBatchOperation).ConfigureAwait(false);

            } while (!string.IsNullOrEmpty(usersPage.OdataNextLink?.ToString()));

            //write user statistics to table "userstatistics" -> single value with overwrite
            var insertUserStatisticsOperation = TableOperation.InsertOrReplace(userStatistics);
            await userStatisticsTable.ExecuteAsync(insertUserStatisticsOperation).ConfigureAwait(false);
            httpClient.Dispose();
        }
    }
}
