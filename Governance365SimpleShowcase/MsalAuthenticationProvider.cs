using System.Threading.Tasks;
using Microsoft.Identity.Client;

namespace Governance365SimpleShowcase
{
    class MsalAuthenticationProvider
    {
        private readonly IConfidentialClientApplication _clientApplication;
        private readonly string[] _scopes;

        public MsalAuthenticationProvider(IConfidentialClientApplication clientApplication, string[] scopes)
        {
            _clientApplication = clientApplication;
            _scopes = scopes;
        }

        public async Task<string> GetTokenAsync()
        {
            var authResult = await _clientApplication.AcquireTokenForClient(_scopes).ExecuteAsync().ConfigureAwait(false);
            return authResult.AccessToken;
        }
    }
}
