using Microsoft.Exchange.WebServices.Data;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models
{

    public interface IUserManager
    {
        Task<IGraphServiceUsersCollectionPage> getUser();

        Task<List<string>> getEWSUser();

        Task<IGraphServiceUsersCollectionPage> getUserName(string email);

        Task<IMailFolderMessagesCollectionPage> getMailAttachments();

    }
}

