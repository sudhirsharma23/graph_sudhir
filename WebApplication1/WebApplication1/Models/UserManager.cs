using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;

namespace WebApplication1.Models
{
    public class UserManager : IUserManager
    {

        private readonly GraphServiceClient graphClient;
        private readonly UserSetting userSettings;


        public UserManager(IOptions<UserSetting> usersettins)
        {
            this.userSettings = usersettins.Value;

            var scopes = new string[] { "https://graph.microsoft.com/.default" };

            // Configure the MSAL client as a confidential client
            var confidentialClient = ConfidentialClientApplicationBuilder
               .Create(this.userSettings.appId)
                 .WithTenantId(this.userSettings.tenantId)
                 .WithClientSecret(this.userSettings.clientSecret)
                 .Build();

            GraphServiceClient graphServiceClient =
                   new GraphServiceClient(new DelegateAuthenticationProvider(async (requestMessage) =>
                   {

                       // Retrieve an access token for Microsoft Graph (gets a fresh token if needed).
                       var authResult = await confidentialClient
                              .AcquireTokenForClient(scopes)
                              .ExecuteAsync();

                       // Add the access token in the Authorization header of the API request.
                       requestMessage.Headers.Authorization =
                              new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                   })
                   );

            this.graphClient = graphServiceClient;
        }

        public async Task<IGraphServiceUsersCollectionPage> getUser()
        {
            // Read user list
            try
            {
                // Get user by sign-in name
                var result = await this.graphClient.Users
                    .Request()
                    //.Filter($"identities/any(c:c/issuerAssignedId eq '{email}' and c/issuer eq '{this.userSettings.tenant}')")
                    .Select(e => new
                    {
                        e.DisplayName,
                        e.Id,
                        e.Identities
                    })
                    .GetAsync();

                if (result != null)
                {
                    return result;
                }
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public async Task<IGraphServiceUsersCollectionPage> getUserName(string email)
        {
            // Read user list
            try
            {
                // Get user by sign-in name
                var result = await this.graphClient.Users
                    .Request()
                    .Filter($"identities/any(c:c/issuerAssignedId eq '{email}' and c/issuer eq '{this.userSettings.email}')")
                    .Select(e => new
                    {
                        e.DisplayName,
                        e.Id,
                        e.Identities
                    })
                    .GetAsync();

                if (result != null)
                {
                    return result;
                }
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public async Task<IMessageAttachmentsCollectionPage> getMailAttachments()
        {
            // Read user list
            try
            {
                var mailFolders = await graphClient.Users["ms365sudhir_admin@4xprkm.onmicrosoft.com"].MailFolders.Inbox.Messages
                    .Request()
                    //.Filter("hasAttachments eq true")
                    .Expand("attachments").GetAsync();

                for (int i = 0; i < mailFolders.Count; i++)
                {

                    foreach (var message in mailFolders)
                    {
                        Console.WriteLine("Id = " + message.Id);
                        Console.WriteLine("Subject = " + message.Subject);

                        if (message.HasAttachments == true)
                        {
                            IMessageAttachmentsCollectionPage attachments = await graphClient.Users["ms365sudhir_admin@4xprkm.onmicrosoft.com"].Messages[message.Id].Attachments.Request().GetAsync();
                            foreach (Attachment attachment in attachments)
                            {
                                if (attachment.ODataType == "#microsoft.graph.fileAttachment")
                                {
                                    FileAttachment fileAttachment = attachment as FileAttachment;
                                    byte[] contentBytes = fileAttachment.ContentBytes;

                                    using (FileStream fileStream = new FileStream("d:\\test\\" + fileAttachment.Name, FileMode.Create, FileAccess.Write))
                                    {
                                        fileStream.Write(contentBytes);
                                    }
                                }
                            }
                        }
                    }
                }
                if (mailFolders != null)
                {
                    return (IMessageAttachmentsCollectionPage)mailFolders;
                }
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}
