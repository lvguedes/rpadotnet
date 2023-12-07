using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RpaLib.APIs.Pipefy.Model
{
    public class Organization
    {
        //public Billing Billing { get; set; }
        public DateTime CreatedAt { get; set; }
        public User CreatedBy { get; set; }
        public string CustomLogoUrl { get; set; }
        public bool Freemium { get; set; }
        public string Id { get; set; }
        //public Member[] Members { get; set; }
        public int MembersCount { get; set; }
        //public Role[] MembersCountByRole { get; set; }
        public string Name { get; set; }
        public bool OnlyAdminCanCreatePipes { get; set; }
        public bool OnlyAdminCanInviteUsers { get; set; }
        //public UserGQLInternalTypeConnection OrgMembers { get; set; }
        //public OrganizationPermissionsInternalGQL permissions { get; set; }
        public List<Pipe> Pipes { get; set; }
        public int PipesCount { get; set; }
        public string PlanName { get; set; }
        public string Role { get; set; }
        //public TableConnection Tables { get; set; }
        //public OrganizationUserMetadata UserMetadata { get; set; }
        public List<User> Users { get; set; }
        public string Uuid { get; set; }
        //public List<Webhook> Webhooks { get; set; }
    }
}
