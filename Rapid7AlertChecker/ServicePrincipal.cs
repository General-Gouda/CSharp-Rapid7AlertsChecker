using System.Threading.Tasks;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using CredentialManagement;
using System.Configuration;
using System;

namespace Rapid7AlertChecker
{
    class ServicePrincipal
    {
        private static Boolean testing = Convert.ToBoolean(ConfigurationManager.AppSettings["Testing"]);
        public static string resource = ConfigurationManager.AppSettings["Resource"];
        public static string authorityUri = ConfigurationManager.AppSettings["AuthorityUri"];
        public static string client_id = ConfigurationManager.AppSettings["Client_ID"];
        public static string cred_target = "Rapid7Alerts";

        static public async Task<AuthenticationResult> GetS2SAccessTokenForProdMSAAsync()
        {
            if (testing)
            {
                authorityUri = ConfigurationManager.AppSettings["TestAuthorityUri"];
                client_id = ConfigurationManager.AppSettings["Test_Client_ID"];
                cred_target = "TestRapid7Alerts";
            }

            // Pulls credentials out of Windows Credential Manager.
            // Target is the name of the Generic Credential Object that holds the credentials needed
            Credential creds = new Credential
            {
                Target = cred_target,
                Type = CredentialType.Generic
            };

            creds.Load();

            UserPasswordCredential userCreds = new UserPasswordCredential(creds.Username, creds.Password);

            return await GetS2SAccessToken(authorityUri, resource, client_id, userCreds);
        }

        static async Task<AuthenticationResult> GetS2SAccessToken(string authority, string resource, string client_id, UserPasswordCredential userCredentials)
        {
            AuthenticationContext context = new AuthenticationContext(authority);
            AuthenticationResult authenticationResult = await context.AcquireTokenAsync(
                resource,  // the resource (app) we are going to access with the token
                client_id, // client id
                userCredentials);  // the user credentials
            return authenticationResult;
        }
    }
}
