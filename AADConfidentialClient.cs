using Microsoft.Identity.Client;
using System;
using System.Threading.Tasks;

namespace EwsOAuth
{
    public class AADConfidentialClient
    {
        private IConfidentialClientApplication _confidentialClient;

        public AADConfidentialClient(string appId, string clientSecret, string tenantId)
        {
            _confidentialClient = ConfidentialClientApplicationBuilder
                .Create(appId)
                .WithClientSecret(clientSecret)
                .WithTenantId(tenantId)
                .Build();
        }

        public async Task<string> GetAuthToken(string[] scopes)
        {
            try
            {
                var authResult = await _confidentialClient.AcquireTokenForClient(scopes).ExecuteAsync();
                return authResult.AccessToken;
            }
            catch (MsalException ex)
            {
                Console.WriteLine("Error acquiring access token: {0}", ex);

                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: {0}", ex);

                return null;
            }
        }
    }
}