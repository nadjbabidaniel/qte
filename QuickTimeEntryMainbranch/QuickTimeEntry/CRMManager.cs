using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Configuration;
using System.Data;
using System.ServiceModel.Description;
using Microsoft.Crm.Sdk.Messages;

using CCMA24Produktiv;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using McTools.Xrm.Connection;

namespace QuickTimeEntry
{
    class CRMManager
    {
        //public static Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy GetServiceProxy()
        public static Microsoft.Xrm.Sdk.IOrganizationService GetServiceProxy()
        {

            // nicht mehr benötigt
            throw new NotImplementedException("GetServiceProxy über ConnectionManager einsetzen.");

            //System.ServiceModel.Description.ClientCredentials _Credentials = new ClientCredentials();

            //Boolean UseDefaultCredentials = false;
            //String _MSCRM_UseDefaultCredentials = System.Configuration.ConfigurationManager.AppSettings["MSCRM.UseDefaultCredentials"].ToString();
            //if (_MSCRM_UseDefaultCredentials != "") UseDefaultCredentials = Convert.ToBoolean(_MSCRM_UseDefaultCredentials);

            //Uri _uri = new System.Uri(System.Configuration.ConfigurationManager.AppSettings["MSCRM.ServiceURL"].ToString());

            //if (!UseDefaultCredentials)
            //{
            //    String _MSCRM_CredentialsUsername = System.Configuration.ConfigurationManager.AppSettings["MSCRM.CredentialsUsername"].ToString();
            //    String _MSCRM_CredentialsPassword = System.Configuration.ConfigurationManager.AppSettings["MSCRM.CredentialsPassword"].ToString();
            //    String _MSCRM_CredentialsDomain = System.Configuration.ConfigurationManager.AppSettings["MSCRM.CredentialsDomain"].ToString();

            //    if (_MSCRM_CredentialsDomain == "")
            //    {
            //        _Credentials.Windows.ClientCredential = new NetworkCredential(_MSCRM_CredentialsUsername, _MSCRM_CredentialsPassword);
            //    }
            //    else
            //    {
            //        _Credentials.Windows.ClientCredential = new NetworkCredential(_MSCRM_CredentialsUsername, _MSCRM_CredentialsPassword, _MSCRM_CredentialsDomain);
            //    }
            //}
            //else
            //{
            //    _Credentials.Windows.ClientCredential = CredentialCache.DefaultNetworkCredentials;
            //}

            //Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy _proxy = new Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy(_uri, null, _Credentials, null);

            //_proxy.EnableProxyTypes();

            //return _proxy;
        }

        /*
        public static MetadataService GetMetaService()
        {
            Boolean UseDefaultCredentials = false;
            String _MSCRM_UseDefaultCredentials = System.Configuration.ConfigurationManager.AppSettings["MSCRM.UseDefaultCredentials"].ToString();
            if (_MSCRM_UseDefaultCredentials != "") UseDefaultCredentials = Convert.ToBoolean(_MSCRM_UseDefaultCredentials);

            String _MSCRM_OrganizationName = System.Configuration.ConfigurationManager.AppSettings["MSCRM.OrganizationName"].ToString();

            CRMMeta1.CrmAuthenticationToken token = new CRMMeta1.CrmAuthenticationToken();
            token.OrganizationName = _MSCRM_OrganizationName;

            MetadataService metaService = new MetadataService();

            metaService.UseDefaultCredentials = UseDefaultCredentials;

            if (!UseDefaultCredentials)
            {
                String _MSCRM_CredentialsUsername = System.Configuration.ConfigurationManager.AppSettings["MSCRM.CredentialsUsername"].ToString();
                String _MSCRM_CredentialsPassword = System.Configuration.ConfigurationManager.AppSettings["MSCRM.CredentialsPassword"].ToString();
                String _MSCRM_CredentialsDomain = System.Configuration.ConfigurationManager.AppSettings["MSCRM.CredentialsDomain"].ToString();

                // Set Credentials
                metaService.Credentials = new NetworkCredential(_MSCRM_CredentialsUsername, _MSCRM_CredentialsPassword, _MSCRM_CredentialsDomain);
            }

            metaService.CrmAuthenticationTokenValue = token;

            return metaService;
        }


        public static CrmService GetService()
        {
            Boolean UseDefaultCredentials = false;
            String _MSCRM_UseDefaultCredentials = System.Configuration.ConfigurationManager.AppSettings["MSCRM.UseDefaultCredentials"].ToString();
            if (_MSCRM_UseDefaultCredentials != "") UseDefaultCredentials = Convert.ToBoolean(_MSCRM_UseDefaultCredentials);

            String _MSCRM_OrganizationName = System.Configuration.ConfigurationManager.AppSettings["MSCRM.OrganizationName"].ToString();

            CRMService1.CrmAuthenticationToken token = new CRMService1.CrmAuthenticationToken();
            token.OrganizationName = _MSCRM_OrganizationName;

            CrmService crmServ = new CrmService();

            crmServ.UseDefaultCredentials = UseDefaultCredentials;

            if (!UseDefaultCredentials)
            {
                String _MSCRM_CredentialsUsername = System.Configuration.ConfigurationManager.AppSettings["MSCRM.CredentialsUsername"].ToString();
                String _MSCRM_CredentialsPassword = System.Configuration.ConfigurationManager.AppSettings["MSCRM.CredentialsPassword"].ToString();
                String _MSCRM_CredentialsDomain = System.Configuration.ConfigurationManager.AppSettings["MSCRM.CredentialsDomain"].ToString();

                // Set Credentials
                crmServ.Credentials = new NetworkCredential(_MSCRM_CredentialsUsername, _MSCRM_CredentialsPassword, _MSCRM_CredentialsDomain);
                crmServ.AllowAutoRedirect = true;
                
            }

            crmServ.CrmAuthenticationTokenValue = token;

            return crmServ;
        }
        */

        private static DataRow GetCacheItemBySubType(DataTable _TimeEntryCache, String itemTYPE, String SubType)
        {
            DataRow _rV = null;

            DataRow[] _result = _TimeEntryCache.Select("[type]='" + itemTYPE + "' AND [subtype]='" + SubType + "'");
            if (_result.Length == 1)
            {
                _rV = _result[0];
            }

            return _rV;
        }

        public static object[] GetCurrentUser(IOrganizationService _Service, DataTable _TimeEntryCache, Boolean FromCache, ConnectionDetail cnnDetail)
        {
            // A24 JNE 24.9.2015
            if (_Service == null && FromCache == false)
                return null;

            Guid userGuid = Guid.Empty;
            string fullname = "";

            String _User = String.Empty;

            if (cnnDetail.IsCustomAuth)
            {
                if (cnnDetail.UserName != String.Empty)
                    _User = String.Format("{0}\\{1}", cnnDetail.UserDomain, cnnDetail.UserName);
                else
                    _User = cnnDetail.UserName;
            }
            else
            {
                _User = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString();
            }

            

            try
            {
                if (!FromCache)
                {
                    WhoAmIRequest userRequest = new WhoAmIRequest();
                    WhoAmIResponse user = (WhoAmIResponse)_Service.Execute(userRequest);
                    userGuid = user.UserId;
                    
                    Entity _user = _Service.Retrieve(SystemUser.EntityLogicalName, userGuid, new ColumnSet(true));
                    
                    if (_user.Contains("fullname")) 
                        fullname = _user.Attributes["fullname"].ToString();
                }
            }
            catch { }

            DataRow _dr = null;
            String _tmp_User = _User;
            _tmp_User = _tmp_User.Replace("\\", "_");
            _tmp_User = _tmp_User.Replace(".", "_");

            if (userGuid != Guid.Empty)
            {
                // Update Cache
                if ((_dr = GetCacheItemBySubType(_TimeEntryCache, "SETTING_" + _tmp_User, "CURRENT_MSCRM_USER")) == null)
                {
                    // Create New
                    _dr = _TimeEntryCache.NewRow();
                    _dr["id"] = Guid.NewGuid().ToString();
                    _dr["type"] = "SETTING_" + _tmp_User;
                    _dr["subtype"] = "CURRENT_MSCRM_USER";
                    _dr["mscrmid"] = "";
                    _dr["mscrmparentid"] = "";
                    _dr["description"] = "";
                    _dr["zusatzinfo1"] = fullname;
                    _dr["zusatzinfo2"] = userGuid.ToString();

                    _TimeEntryCache.Rows.Add(_dr);
                    _TimeEntryCache.AcceptChanges();
                }
                else
                {
                    // Update
                    _dr["zusatzinfo1"] = fullname;
                    _dr["zusatzinfo2"] = userGuid.ToString();
                    _TimeEntryCache.AcceptChanges();
                }
            }
            else
            {
                // Try get User from Cache
                if ((_dr = GetCacheItemBySubType(_TimeEntryCache, "SETTING_" + _tmp_User, "CURRENT_MSCRM_USER")) != null)
                {
                    userGuid = new Guid(_dr["zusatzinfo2"].ToString());
                    fullname = _dr["zusatzinfo1"].ToString();
                }
            }

            try
            {
                object[] or = new object[] { userGuid, fullname };
                return or;
            }
            catch (Exception ex)
            {
            }

            return null;
        }

        public static String GetSearchList(IOrganizationService _Service, String sSearchString, String sEntity, String sNameAttribute, String sIDAttribute)
        {
            String rV = "";

            if (_Service != null)
            {
                sSearchString = sSearchString.Replace("*", "%");
                if (sSearchString.EndsWith("%") == false) sSearchString = sSearchString + "%";

                FilterExpression filterExpression = new FilterExpression();
                ConditionExpression condition = new ConditionExpression(sNameAttribute, ConditionOperator.Like, sSearchString);
                ConditionExpression condition2 = new ConditionExpression("isdisabled", ConditionOperator.Equal, "false");

                filterExpression.FilterOperator = LogicalOperator.And;
                filterExpression.AddCondition(condition);
                filterExpression.AddCondition(condition2);

                QueryExpression query = new QueryExpression()
                {
                    EntityName = sEntity,
                    ColumnSet = new ColumnSet(new String[] { sNameAttribute, sIDAttribute }),
                    Criteria = filterExpression,
                    Distinct = false,
                    NoLock = true
                };

                RetrieveMultipleRequest multipleRequest = new RetrieveMultipleRequest();
                multipleRequest.Query = query;

                RetrieveMultipleResponse response = (RetrieveMultipleResponse)_Service.Execute(multipleRequest);

                foreach (Entity _e in response.EntityCollection.Entities)
                {
                    if (_e.Attributes[sNameAttribute] != null && _e.Attributes[sIDAttribute] != null)
                    {
                        if (rV != "") rV += ";";
                        rV += _e.Attributes[sNameAttribute].ToString();
                        rV += "|";
                        rV += _e.Attributes[sIDAttribute].ToString();
                    }
                }
            }

            return rV;
        }

    }
}
