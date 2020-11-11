/* This is a sample daemon desktop app that acquires tokens and calls MS Graph on behalf of itself with no signed-in user
 * Read more: MSAL authentication flows and app scenarios:
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/authentication-flows-app-scenarios
 */

/* 4 steps:
     * 1. Libraries -> Import the relevant libraries
     * 2. IAuthenticationProvider -> Create an instance of an IAuthenticationProvider
     * 3. GraphServiceClient -> Create an instance of an authenticated GraphServiceClient, passing in an instance of an IAuthenticationProvider
     * 4. Call Graph -> Make calls using the instance of the authenticated GraphServiceClient
 */

using System;
using System.Collections.Generic;
using System.Net.Http;
using Microsoft.Identity.Client; // for authenticating into AAD
using Microsoft.Graph; // for making calls to MS Graph
using Microsoft.Extensions.Configuration;
using Microsoft.Graph.Auth; // for providing implementations for IAuthenticationProvider and handling access token acquisition and storage

namespace ConsoleGraphTest
{
    class Program
    {
        private static GraphServiceClient _graphServiceClient;
        private static HttpClient _httpClient;

        static void Main(string[] args)
        {
            var config = LoadAppSettings();
            if (null == config)
            {
                Console.WriteLine("Missing or invalid appsettings.json file. Please see README.md for configuration instructions.");
                return;
            }

            // Query using Graph SDK (preferred when possible)
            GraphServiceClient graphClient = GetAuthenticatedGraphClient(config);

            // OData query options: https://docs.microsoft.com/en-us/graph/query-parameters
            List<QueryOption> options = new List<QueryOption>
            {
                new QueryOption("$top", "5"),
                new QueryOption("$orderby", "displayName desc")
            };

            // Call MS Graph
            var graphResult = graphClient
                                    .Users
                                    .Request(options)
                                    .GetAsync().Result;

            Console.WriteLine("** Tenant users **");
            Console.WriteLine("\n---Graph Service Client Result---");
            foreach (var user in graphResult)
            {
                Console.WriteLine(user.DisplayName);
            }

            // Direct query using HTTPClient (for beta endpoint calls or not available in Graph SDK)
            HttpClient httpClient = GetAuthenticatedHTTPClient(config);
            Uri Uri = new Uri("https://graph.microsoft.com/v1.0/users?$top=5&$select=displayName");
            var httpResult = httpClient.GetStringAsync(Uri).Result;

            Console.WriteLine("\n---HTTP Result---");
            Console.WriteLine(httpResult);
        }

        /// <summary>
        /// Creates an IAuthenticationProvider object.
        /// </summary>
        /// <param name="config"></param>
        /// <returns>A client credential provider instance of an IAuthenticationProvider.</returns>
        private static IAuthenticationProvider CreateAuthenticationProvider(IConfigurationRoot config)
        {
            // See other examples for creating IAuthenticationProvider objects:
            // https://github.com/microsoftgraph/msgraph-sdk-dotnet-auth#example

            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var tenantId = config["tenantId"];

            var cca = ConfidentialClientApplicationBuilder
                                                    .Create(clientId)
                                                    .WithTenantId(tenantId)
                                                    .WithClientSecret(clientSecret)
                                                    // The Authority is a required parameter when your application is configured
                                                    // to accept authentications only from the tenant where it is registered.
                                                    .Build();

            return new ClientCredentialProvider(cca);
        }

        /// <summary>
        /// Creates an authenticated Graph client.
        /// </summary>
        /// <param name="config"></param>
        /// <returns>An authenticated Graph client.</returns>
        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthenticationProvider(config);
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }

        /// <summary>
        /// Creates an authenticated HTTP client.
        /// </summary>
        /// <param name="config"></param>
        /// <returns></returns>
        private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthenticationProvider(config);
            _httpClient = new HttpClient(new AuthenticationHandler(authenticationProvider, new HttpClientHandler()));
            return _httpClient;
        }

        /// <summary>
        /// Loads application config. settings.
        /// </summary>
        /// <returns>Application config. settings.</returns>
        private static IConfigurationRoot LoadAppSettings()
        {
            try
            {
                var config = new ConfigurationBuilder()
                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", false, true)
                .Build();

                // Validate required settings
                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["applicationSecret"]) ||
                    string.IsNullOrEmpty(config["redirectUri"]) ||
                    string.IsNullOrEmpty(config["tenantId"]) ||
                    string.IsNullOrEmpty(config["domain"]))
                {
                    return null;
                }

                return config;
            }
            catch (System.IO.FileNotFoundException)
            {
                return null;
            }
        }
    }
}
