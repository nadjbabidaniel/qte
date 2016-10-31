namespace A24.Xrm
{
    class Constants
    {
        public class Crypto
        {
            public const string CryptoPassPhrase = "A24XrmSyncCrypto";
            public const string CryptoSaltValue = "Creating Impact!";
            public const string CryptoInitVector = "cNyVwBKGheFZCLS0";
            public const string CryptoHashAlgorythm = "SHA1";
            public const int CryptoPasswordIterations = 2;
            public const int CryptoKeySize = 256;
        }

        public class WebserviceEndpoints
        {
            public class Soap
            {
                public const string DiscoveryServiceExtension = "XRMServices/2011/Discovery.svc";
                public const string OrganizationServiceExtension = "XRMServices/2011/Organization.svc";
            }

            public class Rest
            {
                public const string OrganizationDataExtenstion = "XRMServices/2011/OrganizationData.svc";
            }


        }

        public class XmlFileNames
        {
            public const string CampaignName = @"\Linked\50 Data\Campaign.list.xml";
            public const string AccountName1 = @"\Linked\50 Data\AccountName1.list.xml";
            public const string AccountName2 = @"\Linked\50 Data\AccountName2.list.xml";
            public const string AccountName3 = @"\Linked\50 Data\AccountName3.list.xml";
            public const string Country = @"\Linked\50 Data\Country.list.xml";
            public const string EMailSubject = @"\Linked\50 Data\EMail.Subject.xml";
            public const string FirstName = @"\Linked\50 Data\FirstName.list.xml";
            public const string LastName = @"\Linked\50 Data\LastName.list.xml";
            public const string StreetName = @"\Linked\50 Data\Street.list.xml";
            public const string IncidentTitle = @"\Linked\50 Data\IncidentTitle.list.xml";
            public const string TaskSubject = @"\Linked\50 Data\Task.Subject.xml";
            public const string articleTitle = @"\Linked\50 Data\articleTitle.list.xml";
            public const string Response = @"\Linked\50 Data\Response.xml";
            public const string City = @"\Linked\50 Data\City.list.xml";
            public const string OpportunityName = @"\Linked\50 Data\OpportunityName.xml";
            public const string Leads = @"\Linked\50 Data\Config.xml";
            public const string GeneratedCampaigns = @"\Linked\50 Data\Config.xml";
            public const string ProductName = @"\Linked\50 Data\ProductName.xml";
            public const string PriceLevelName = @"\Linked\50 Data\PriceLevelName.xml";
            public const string QuoteName = @"\Linked\50 Data\QuoteName.xml";
            public const string articleStatus = @"\Linked\50 Data\articleState.list.xml";
            public const string ContractTitle = @"\Linked\50 Data\ContractTitle.list.xml";
            public const string Random = @"\Linked\50 Data\Config.xml";
        }

        public const string LoggingFile =
            @".\dat\log.xml";
    }
}
