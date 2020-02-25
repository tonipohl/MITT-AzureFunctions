# MITT-AzureFunctions

This Visual Studio project contains a version of **Governance365SimpleShowcase** to demo a solution for working with Azure Functions and Microsoft 365. This project has been developed by atwork.at by Jörg Schoba and Toni Pohl. We wanted to deliver the idea of a .NET Core 3.1 Function App including several functions to get data out of a tenant.

The Function App requires app settings for running through an Office 365 tenant. The data is saved to an Azure Storage. To run it locally, add the following keys in **local.settings.json**:

~~~~ json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "<Azure storage account connection>",
    "TenantId": "<the Office 365 thenant id>",
    "AppId": "<the App ID>",
    "AppSecret": "<the App Secret>",
    "StorageConnectionString": "<Azure storage account connection>"
  }
}
~~~~

More documentation will follow shortly.
