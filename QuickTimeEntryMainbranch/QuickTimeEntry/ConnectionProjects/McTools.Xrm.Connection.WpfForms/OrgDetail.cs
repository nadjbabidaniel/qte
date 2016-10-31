using Microsoft.Xrm.Sdk.Discovery;

namespace McTools.Xrm.Connection.WpfForms
{
    public class OrgDetail
    {
        public OrganizationDetail Detail { get; set; }
        public override string ToString()
        {
            return Detail.FriendlyName;
        }
    }
}
