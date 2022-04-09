using System.Collections.Generic;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class UserController : Controller
    {
        private IUserManager UserManager { get; set; }

        public UserController(IUserManager userManager)
        {
            this.UserManager = userManager;
        }



        [HttpGet]
        public async Task<ActionResult> Get()
        {
            try
            {
                var user = await this.UserManager.getUser();

                if (user == null)
                {
                    return this.NotFound("User do not exist");
                }


                return this.Ok(user);
            }
            catch (Exception ex)
            {
                return this.BadRequest("Could not get the user");
            }
        }

        [HttpGet("getEmailByEmailId")]
        public async Task<ActionResult> Get([FromQuery] string email)
        {
            try
            {
                var user = await this.UserManager.getUserName(email);

                if (user == null)
                {
                    return this.NotFound("User do not exist");
                }


                return this.Ok(user.CurrentPage[0].DisplayName);
            }
            catch (Exception ex)
            {
                return this.BadRequest("Could not get the user");
            }
        }
        [ActionName("getMailAttachments")]
        [HttpGet("getMailAttachments")]
        public async Task<ActionResult> getMailAttachments()
        {
            try
            {
                var mailFolders = await this.UserManager.getMailAttachments();

                if (mailFolders == null)
                {
                    return this.NotFound("User do not exist");
                }


                return this.Ok(mailFolders);
            }
            catch (Exception ex)
            {
                return this.BadRequest("Could not get the user");
            }
        }

        [ActionName("getEWSUser")]
        [HttpGet("getEWSUser")]
        public async Task<ActionResult> getEWSUser()
        {
            try
            {
                List<string> mailFolders = await this.UserManager.getEWSUser();

                if (mailFolders == null)
                {
                    return this.NotFound("User do not exist");
                }


                return this.Ok(mailFolders);
            }
            catch (Exception ex)
            {
                return this.BadRequest("Could not get the user");
            }
        }

    }
}
