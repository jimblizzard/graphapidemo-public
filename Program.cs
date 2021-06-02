using System;
using System.Collections.Generic;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Disclaimer: This code is for demonstration purposes only and is provided "as is" and is provided without warranty of any kind.   
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////// 

// from: https://docs.microsoft.com/en-us/learn/modules/optimize-data-usage/3-exercise-retrieve-control-information-returned-from-microsoft-graph
//  and: https://docs.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0 


namespace graphconsoleapp
{
    class Program
    {
        private static GraphServiceClient _graphClient;
        static void Main(string[] args)
        {
            var config = LoadAppSettings();
            if (config == null)
            {
                Console.WriteLine("Invalid appsettings.json file.");
                return;
            }

            var client = GetAuthenticatedGraphClient(config);

            //var graphRequest = client.Users.Request();

            // var graphRequest = client.Users
            //         .Request()
            //         .Select(u => new { u.UserPrincipalName, u.Id, u.Surname, u.DisplayName, u.Mail, u.ProxyAddresses });

            // var graphRequest = client.Users
            //         .Request()
            //         .Select(u => new { u.DisplayName, u.Mail })
            //         .Top(15);

            // var graphRequest = client.Users
            //         .Request()
            //         .Select(u => new { u.DisplayName, u.Mail })
            //         .Top(15)
            //         .OrderBy("DisplayName desc");

            // var graphRequest = client.Users
            //         .Request()
            //         .Select(u => new { u.DisplayName, u.Mail })
            //         .Top(15)
            //         .Filter("startsWith(surname,'A') or startsWith(surname,'B') or startsWith(surname,'C')");

            var graphRequest = client.Users
                    .Request()
                    .Select(u => new { u.DisplayName, u.Mail, u.ProxyAddresses })
                    .Top(15);
                    // .Filter("identities/any(c:c/issuerAssignedId eq 'j.smith@yahoo.com' and c/issuer eq 'contoso.onmicrosoft.com')");

            var results = graphRequest.GetAsync().Result;
            foreach(var user in results)
            {
                Console.WriteLine("<" + user.UserPrincipalName + "> <" + user.Id + "> <" + user.DisplayName + "> <" + user.Mail + ">");
                
                foreach(var proxyAddr in user.ProxyAddresses)
                {
                    Console.WriteLine("ProxyAddress <" + proxyAddr + ">");
                }
            }

            Console.WriteLine("\nGraph Request:");
            Console.WriteLine(graphRequest.GetHttpRequestMessage().RequestUri);
        }

        private static IConfigurationRoot LoadAppSettings()
        {
        try
            {
                var config = new ConfigurationBuilder()
                                .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                                .AddJsonFile("appsettings.json", false, true)
                                .Build();

                if (string.IsNullOrEmpty(config["applicationId"]) ||
                    string.IsNullOrEmpty(config["applicationSecret"]) ||
                    string.IsNullOrEmpty(config["redirectUri"]) ||
                    string.IsNullOrEmpty(config["tenantId"]))
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

        private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
        {
            var clientId = config["applicationId"];
            var clientSecret = config["applicationSecret"];
            var redirectUri = config["redirectUri"];
            var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

            List<string> scopes = new List<string>();
            scopes.Add("https://graph.microsoft.com/.default");

            var cca = ConfidentialClientApplicationBuilder.Create(clientId)
                                                    .WithAuthority(authority)
                                                    .WithRedirectUri(redirectUri)
                                                    .WithClientSecret(clientSecret)
                                                    .Build();
            return new MsalAuthenticationProvider(cca, scopes.ToArray());
        }

        private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
        {
            var authenticationProvider = CreateAuthorizationProvider(config);
            _graphClient = new GraphServiceClient(authenticationProvider);
            return _graphClient;
        }

    }
}
