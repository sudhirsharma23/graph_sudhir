using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication1.Models
{
    public class UserSetting
    {
        public string appId { get; set; }
        public string tenantId { get; set; }
        public string clientSecret { get; set; }

        public string email { get; set; }
        
    }
}
