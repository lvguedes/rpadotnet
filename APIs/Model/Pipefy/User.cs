using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Model.Pipefy
{
    using Id = String;
    public class User
    {
        public string AvatarUrl { get; set; }
        public bool ConfirmationTokenHasExpired { get; set; }
        public bool Confirmed { get; set; }
        public string CreatedAt { get; set; }
        public string CurrentOrganizationRole { get; set; }
        public string DepartmentKey { get; set; }
        public string DepartmentName { get; set; }
        public string DisplayName { get; set; }
        public string Email { get; set; }
        public bool HasNotifications { get; set; }
        public Id Id { get; set; }
        public string IntercomHash { get; set; }
        public Id IntercomId { get; set; }
        public bool Invited { get; set; }
        public string Locale { get; set; }
        public string Name { get; set; }
        //public UserPermissionsInternalGqlType Permissions { get; set; }
        public string Phone { get; set; }
        //public UserPreference Preferences { get; set; }
        public string SignupData { get; set; }
        public string Timezone { get; set; }
        public string Username { get; set; }
        public Id Uuid { get; set; }
    }
}
