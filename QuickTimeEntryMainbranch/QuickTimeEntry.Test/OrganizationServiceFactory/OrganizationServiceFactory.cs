using System.Net;
using System.ServiceModel.Description;
using Microsoft.Xrm.Client;
using A24.Xrm;
using System.Configuration;


using System.Security;

using CCMA24Produktiv;


namespace A24.Xrm
{
    public class OrganizationServiceFactory : AbstractOrganizationServiceFactoryBase
    {
         protected override void SetConnectionData(string password = "" )
         {
             // works only for jan nebendahl user on jan nebendahl laptop
             //var encryptedConnectionString = "AQAAANCMnd8BFdERjHoAwE/Cl+sBAAAA447ztD7f5EG2J32UyuoPjwAAAAACAAAAAAADZgAAwAAAABAAAABdkxchFdGuPtqipwNYFSr8AAAAAASAAACgAAAAEAAAABCTPvZk9ICUGkbB5IOQlmnwAAAAv59MM6PylKHbHVRnzfOM5kf9WV9y9lQCnFgdUvfHobHly67BdRYVbZUE7HA2kQyXno4LFtRS0e3YEq1iSOS6WwwDJZlC/d6QxfTaYDLVjQdbjXax39pkl9ta2jcznxjCdnMTgXY3/Nss9zkUfNAXNRwkE6w23z9w5dYVH0cv2xnBXjS8BGw7KAh04/l2k8zg665d7OdJGItKRTFlWbJtp8enoXA9SdK482DWsTDd0k2qsXteyZObMSEq46qDqWVrYkItbs5rrfSuPEJI37fSN802JbliiOaJTuFKGUIkeGrNaQikWBLojxPjeA6lbMEnFAAAAChlVolN8F6ssJzTwsdU/Pn26h55";
             //var decryptedConnectionString = CryptoHelper.ToInsecureString(CryptoHelper.DecryptString(encryptedConnectionString));
             //crmCnn = new CrmConnection("GFASandbox");
             crmCnn = new CrmConnection("A24QteDev");

         }
         public XrmServiceContext GetOrganizationServiceContext(string impersonateId = "")
         {
             return new XrmServiceContext(GetOrganizationService(impersonateId));
         }
    }
}
