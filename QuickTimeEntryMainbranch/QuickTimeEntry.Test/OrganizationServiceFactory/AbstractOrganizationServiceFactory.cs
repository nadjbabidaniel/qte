using System;
using System.Diagnostics;
using System.ServiceModel.Description;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk;

using A24.Xrm;


namespace A24.Xrm
{
    public abstract class AbstractOrganizationServiceFactoryBase
    {
        protected string protocol { get; set; }
        protected string host{ get  ; set;  }
        protected int port{ get; set; }
        protected string orgname{ get; set; }
        protected string userName{ get; set; } 
        protected string passWord{ get; set; }
        protected string domain{ get; set; }
        private OrganizationServiceProxy _serviceProxy{ get; set; }
        private IOrganizationService _organizationService { get; set; }
        protected CrmConnection crmCnn { get; set; }

        protected AbstractOrganizationServiceFactoryBase(string password = "")
        {
            SetConnectionData(password);
        }

        private void GenerateServiceProxy()
        {
            var sw = new Stopwatch();
            Console.WriteLine("Generating OrganizationServiceProxy.");
            sw.Start();
            Uri OrganizationUri = GetOrganizationUri();
            Uri HomeRealmUri = null;
            ClientCredentials Credentials = GetClientCredentials();
            ClientCredentials DeviceCredentials = null;
            _serviceProxy = new OrganizationServiceProxy(OrganizationUri,
                                                HomeRealmUri,
                                                Credentials,
                                                DeviceCredentials);
            _serviceProxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());

            Console.WriteLine(string.Format("Generating OrganizationServiceProxy took {0}ms...", sw.ElapsedMilliseconds));
        }

        public IOrganizationService GetOrganizationService(string impersonateUserId = "")
        {
            var sw = new Stopwatch();
            sw.Start();

            IOrganizationService orgService;

            if (_serviceProxy == null && impersonateUserId == "" && crmCnn != null)
                orgService = GetNewOrganizationService();
            else
            {
                GenerateServiceProxy();
                if (impersonateUserId != "")
                {
                    var impersonateUserGuid = new Guid(impersonateUserId);
                    _serviceProxy.CallerId = impersonateUserGuid;
                }
                orgService = (IOrganizationService) _serviceProxy;
            }

            Console.WriteLine(string.Format("Creating OrganizationService took {0}ms...", sw.ElapsedMilliseconds));
            this._organizationService = orgService;
            return orgService;
         }

        public IOrganizationService GetNewOrganizationService(string impersonateUserId = "")
        {
            var sw = new Stopwatch();
            sw.Start();
            var orgService = new OrganizationService(crmCnn);
            this._organizationService = orgService;
            return orgService;
        }

        private Uri GetOrganizationUri()
        {
            var path = String.Format("{0}/{1}", orgname, Constants.WebserviceEndpoints.Soap.OrganizationServiceExtension);
            var organizationUriBuilder = new UriBuilder(protocol, host, port, path);
            return organizationUriBuilder.Uri;
        }

        private ClientCredentials GetClientCredentials()
        {
            var credentials = new ClientCredentials();

            // Necessary for IFD
            if (protocol == "https")
            {
                credentials.UserName.UserName = userName;
                credentials.UserName.Password = passWord;
            }

            var networkCredential = new System.Net.NetworkCredential(userName, passWord, domain);
            credentials.Windows.ClientCredential = networkCredential;
            return credentials;
        }
        protected abstract void SetConnectionData(string password = "");
    }
}
