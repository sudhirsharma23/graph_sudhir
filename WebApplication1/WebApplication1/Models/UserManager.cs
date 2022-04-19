using Microsoft.Exchange.WebServices.Data;
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

            var scopes = new string[] { this.userSettings.scopes };

            // Configure the MSAL client as a confidential client
            var confidentialClient = ConfidentialClientApplicationBuilder
               .Create(this.userSettings.appId)
                 .WithTenantId(this.userSettings.tenantId)
                 .WithClientSecret(this.userSettings.ClientSecret)
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

            var queryOptions = new List<QueryOption>()
                    {
                        new QueryOption("$count", "true")
                    };
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
                         e.Identities,
                         e.BusinessPhones,
                         e.MobilePhone
                     })
                     .OrderBy("userPrincipalName")
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

            var queryOptions = new List<QueryOption>()
                    {
                        new QueryOption("$count", "true")
                    };

            // Read user list
            try
            {
                // Get user by sign-in name
                var result = await this.graphClient.Users
                    .Request(queryOptions)
                    //.Filter($"identities/any(c:c/issuerAssignedId eq '{email}' and c/issuer eq '{this.userSettings.email}')")
                    .Filter($"endswith(mail,'{email}')")
                    .Header("ConsistencyLevel", "eventual")
                    .Select(e => new
                    {
                        e.DisplayName,
                        e.Id,
                        e.Identities,
                        e.BusinessPhones,
                        e.MobilePhone
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

        public async Task<IMailFolderMessagesCollectionPage> getMailAttachments()
        {
            // Read user list
            try
            {
                var mailFolders = await graphClient.Users[this.userSettings.userId].MailFolders.Inbox.Messages
                    .Request()
                    //.Filter("hasAttachments eq true")
                    .Expand("attachments").GetAsync();

                foreach (var message in mailFolders)
                {
                    Console.WriteLine("Id = " + message.Id);
                    Console.WriteLine("Subject = " + message.Subject);

                    if (message.HasAttachments == true)
                    {
                        IMessageAttachmentsCollectionPage attachments = await graphClient.Users[this.userSettings.userId].Messages[message.Id].Attachments.Request().GetAsync();
                        foreach (Microsoft.Graph.Attachment attachment in attachments)
                        {
                            if (attachment.ODataType == "#microsoft.graph.fileAttachment")
                            {
                                Microsoft.Graph.FileAttachment fileAttachment = attachment as Microsoft.Graph.FileAttachment;
                                byte[] contentBytes = fileAttachment.ContentBytes;

                                using (FileStream fileStream = new FileStream("c:\\test\\" + fileAttachment.Name, FileMode.Create, FileAccess.Write))
                                {
                                    fileStream.Write(contentBytes);
                                }
                            }
                        }
                    }
                }

                if (mailFolders != null)
                {
                    return mailFolders;
                }
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }


        public async Task<List<string>> getEWSUser()
        {
            List<string> FoldersList = new List<string>();

            // Configure the MSAL client as a confidential client
            var confidentialClient = ConfidentialClientApplicationBuilder
               .Create(this.userSettings.appId)
                 .WithTenantId(this.userSettings.tenantId)
                 .WithClientSecret(this.userSettings.ClientSecret)
                 .Build();

            var ewsScopes = new string[] { "https://outlook.office365.com/.default" };

            try
            {
                // Get token
                var authResult = await confidentialClient.AcquireTokenForClient(ewsScopes)
                    .ExecuteAsync();

                // Configure the ExchangeService with the access token
                var ewsClient = new ExchangeService(ExchangeVersion.Exchange2016);
                ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);
                ewsClient.ImpersonatedUserId =
                    new ImpersonatedUserId(ConnectingIdType.SmtpAddress, this.userSettings.userId);

                //Include x-anchormailbox header
                ewsClient.HttpHeaders.Add("X-AnchorMailbox", this.userSettings.userId);

                // Make an EWS call to list folders on exhange online
                var folders = ewsClient.FindFolders(WellKnownFolderName.MsgFolderRoot, new Microsoft.Exchange.WebServices.Data.FolderView(10));

                List<string> folderList = new List<string>();
               
                foreach (var folder in folders.Result)
                {

                    folderList.Add($"Folder: {folder.DisplayName}");

                }

                // Make an EWS call to read 50 emails(last 5 days) from Inbox folder

                TimeSpan ts = new TimeSpan(-5, 0, 0, 0);
                DateTime date = DateTime.Now.Add(ts);
                SearchFilter.IsGreaterThanOrEqualTo filter = new SearchFilter.IsGreaterThanOrEqualTo(ItemSchema.DateTimeReceived, date);
                var findResults = ewsClient.FindItems(WellKnownFolderName.Inbox, filter, new ItemView(50));
                foreach (var mailItem in findResults.Result)
                {
                    folderList.Add($"Subject: {mailItem.Subject}");
                }

                if (folderList != null)
                {
                    return folderList;
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
