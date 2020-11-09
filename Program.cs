using System;
using System.Collections.Generic;
using System.Net.Http;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph.Auth;

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
            List<QueryOption> options = new List<QueryOption>
            {
                new QueryOption("$top", "5")
            };

            var graphResult = graphClient.Users.Request(options).GetAsync().Result;
            Console.WriteLine("** Tenant users **");
            foreach (var user in graphResult)
            {
                Console.WriteLine(user.DisplayName);
            }

            // Direct query using HTTPClient (for beta endpoint calls or not available in Graph SDK)
            HttpClient httpClient = GetAuthenticatedHTTPClient(config);
            Uri Uri = new Uri("https://graph.microsoft.com/v1.0/users?$top=1");
            var httpResult = httpClient.GetStringAsync(Uri).Result;

            Console.WriteLine("HTTP Result");
            Console.WriteLine(httpResult);
        }

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

        /// <summary>
        /// Creates an IAuthenticationProvider object.
        /// </summary>
        /// <param name="config"></param>
        /// <returns>A client credential provider instance of an IAuthenticationProvider.</returns>
        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            // See other examples of creating IAuthenticationProvider objects:
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

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphServiceClient = new GraphServiceClient(authenticationProvider);
            return _graphServiceClient;
        }

        private static HttpClient GetAuthenticatedHTTPClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _httpClient = new HttpClient(new AuthenticationHandler(authenticationProvider, new HttpClientHandler()));
            return _httpClient;
        }
    }
}
