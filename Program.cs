using Microsoft.Graph.Models;
using Microsoft.Graph;
using Microsoft.Kiota.Abstractions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Graph.Drives.Item.Items.Item.Invite;
using Azure.Identity;

namespace GraphV5Demo
{
    internal class Program
    {
        static async Task<InviteResponse> SendInvit(GraphServiceClient graphClient)
        {
            var driveId = "b!zUOEYnR_Xke8CJ-u1VysBsqGOv2CkS9Nn4Do2gKZ-_zNnOOmFpQsSJ25oqbiOafC";
            var itemId = "01Z7WTS2UPV63U6E5C2RGJPQQGHAZ3RFGD";//"01Z7WTS2U5VPEM44VB6ZDYKLN4SYUGGJIE";
            var email = "kevin@tenant.onmicrosoft.com";
            // Code snippets are only available for the latest version. Current version is 5.x

            var requestBody = new InvitePostRequestBody
            {
                Recipients = new List<DriveRecipient>
                {
                    new DriveRecipient
                    {
                        Email = email,
                    },
                },
                Message = "Here's the file that we're collaborating on.",
                RequireSignIn = true,
                SendInvitation = true,
                Roles = new List<string>
                {
                    "write",
                }
                //,Password = "password123",
                //ExpirationDateTime = "2023-08-15T14:00:00.000Z",
            };
            return await graphClient.Drives[driveId].Items[itemId].Invite.PostAsync(requestBody);

        }
        static void Main(string[] args)
        {
            var scopes = new[] { "Files.ReadWrite.All" };

            // Multi-tenant apps can use "common",  
            // single-tenant apps must use the tenant ID from the Azure portal  
            var tenantId = "46bcf48b-a004-4991-ad82-f7486bb988f4";

            // Value from app registration  
            var clientId = "6d4d417d-1409-467b-9755-98bef31d01dc";

            // using Azure.Identity;  
            var options = new TokenCredentialOptions
            {
                AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
            };

            var userName = "lee@tenant.onmicrosoft.com";
            var password = "password";

            // https://learn.microsoft.com/dotnet/api/azure.identity.usernamepasswordcredential  
            var userNamePasswordCredential = new UsernamePasswordCredential(
                userName, password, tenantId, clientId, options);

            var graphClient = new GraphServiceClient(userNamePasswordCredential, scopes);            

            //authenticatie with delegated permission Fiels.ReadWrite.All 
            var task =SendInvit(graphClient);
            var result=task.Result.Value;

            Console.WriteLine(result);

            Console.ReadKey();
        }
    }
}
