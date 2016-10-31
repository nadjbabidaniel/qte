using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows;
using System.Web.Services.Protocols;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk;

namespace QuickTimeEntry
{
    class PicklistManager
    {
        /*
        public DataTable GetPicklist()
        {
            MetadataService metaService = CRMManager.GetMetaService();

            RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest();
            attributeRequest.EntityLogicalName = "new_timeentry";
            attributeRequest.LogicalName = "new_billing";
            attributeRequest.RetrieveAsIfPublished = false;

            RetrieveAttributeResponse response = (RetrieveAttributeResponse)metaService.Execute(attributeRequest);
            PicklistAttributeMetadata picklist = (PicklistAttributeMetadata)response.AttributeMetadata;

            DataTable dt = new DataTable();
            dt.Columns.Add("Desc");
            dt.Columns.Add("Data");

            dt.Rows.Add(new object[] { "", -1 });

            foreach (Option o in picklist.Options)
            {
                DataRow dr = dt.NewRow();
                dr["Desc"] = o.Label.UserLocLabel.Label;
                dr["Data"] = o.Value.Value;

                dt.Rows.Add(dr);
            }

            return dt;
        }
        */

        public String GetPicklistStr(IOrganizationService _Service, String EntityLogicalName, String LogicalName)
        {
            String rV = "";

            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = EntityLogicalName,
                LogicalName = LogicalName,
                RetrieveAsIfPublished = true
            };

            // Execute the request.
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)_Service.Execute(retrieveAttributeRequest);
            Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata retrievedPicklistAttributeMetadata = (Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
            // Get the current options list for the retrieved attribute.
            OptionMetadata[] optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();

            foreach (OptionMetadata oMD in optionList)
            {
                if (rV != "") rV += ";";
                rV += oMD.Label.UserLocalizedLabel.Label;
                rV += "|";
                rV += oMD.Value.ToString();
            }

            return rV;
        }

        public String GetStatuslistStr(IOrganizationService _Service, String EntityLogicalName, String LogicalName)
        {
            String rV = "";

            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = EntityLogicalName,
                LogicalName = LogicalName,
                RetrieveAsIfPublished = true
            };

            // Execute the request.
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)_Service.Execute(retrieveAttributeRequest);
            Microsoft.Xrm.Sdk.Metadata.StatusAttributeMetadata retrievedStatusAttributeMetadata = (Microsoft.Xrm.Sdk.Metadata.StatusAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
            // Get the current options list for the retrieved attribute.
            OptionMetadata[] optionList = retrievedStatusAttributeMetadata.OptionSet.Options.ToArray();

            foreach (OptionMetadata oMD in optionList)
            {
                if (rV != "") rV += ";";
                rV += oMD.Label.UserLocalizedLabel.Label;
                rV += "|";
                rV += oMD.Value.ToString();
            }

            return rV;
        }
    }
}
