using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using System.Data;
using CCMA24Produktiv;
using System.Reflection;
using System.Resources;

namespace QuickTimeEntry
{
    /// <summary>
    /// Interaktionslogik für EnitySearch.xaml
    /// </summary>
    public partial class EntitySearch : Window
    {
        public String Selected_ID = "";
        public String Selected_EntityType = "";

        DataTable TimeEntryCache;

        DataTable SearchResult;

        private static ResourceManager rm = new ResourceManager("QuickTimeEntry.EditTimeSpanItems", Assembly.GetExecutingAssembly());

        public EntitySearch(DataTable _TimeEntryCache, String _User)
        {
            InitializeComponent();

            searchInfo.Content = "";
            this.Title = "";
            TimeEntryCache = _TimeEntryCache;

            SearchResult = new DataTable("favorites");

            SearchResult.Columns.Add("gctype", typeof(String));
            SearchResult.Columns.Add("gcdescription", typeof(String));
            SearchResult.Columns.Add("gcinfo1", typeof(String));
            SearchResult.Columns.Add("gcinfo2", typeof(String));

            SearchGrid.ItemsSource = SearchResult;

            SearchGrid.IsFilterEnabled = true;
            SearchGrid.RefreshData();
            SearchGrid.GroupBy(SearchGrid.Columns["gctype"], DevExpress.Data.ColumnSortOrder.Ascending);

            SearchGrid.ExpandAllGroups();

            //Boolean _check_failed = false;

            //// Check marked time entries
            //Int32 _sync_count = 0;
            //for (int i = 0; i < _parent_tv.SelectedRows.Count; i++)
            //{
            //    System.Data.DataRowView _drv = (System.Data.DataRowView)_parent_tv.SelectedRows[i];
            //    DataRow _dr = _drv.Row;
            //    String _status = "";
            //    String _tekz = "";
            //    try
            //    {
            //        _tekz = _dr["tekz"].ToString().Trim().ToUpper();
            //        _status = _dr["status"].ToString().Trim().ToUpper();
            //    }
            //    catch
            //    {
            //    }

            //    if (_status == "NOSYNC")
            //    {
            //        _sync_count++;
            //        if (Sync_Entity_Type == "")
            //        {
            //            Sync_Entity_Type = _tekz;
            //        }
            //        else
            //        {
            //            if (Sync_Entity_Type != _tekz)
            //            {
            //                MessageBox.Show("Es können nur Zeiteinträge vom gleichen MSCRM Typ (AP, ANF, VC) zugewiesen werden!", "Wichtiger Hinweis", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            //                _check_failed = true;
            //                break;
            //            }
            //        }
            //    }
            //}

            //if (_sync_count == 0)
            //{
            //    MessageBox.Show("Bitte markieren Sie mindestens einen Zeiteintrag für die Übertragung nach MSCRM!", "Wichtiger Hinweis", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            //    _check_failed = true;
            //}

            //if (_check_failed) 
            //{
            //    this.Close();
            //} else 
            //{
            //    if (Sync_Entity_Type == "AP")
            //    {
            //        if (_sync_count == 1)
            //             this.Title = "Zeiteintrag zu MSCRM Arbeitspaket zuweisen ...";
            //        else this.Title = _sync_count.ToString()+" Zeiteinträge zu MSCRM Arbeitspaket zuweisen ...";
            //    }
            //    else if (Sync_Entity_Type == "ANF")
            //    {
            //        if (_sync_count == 1)
            //             this.Title = "Zeiteintrag zu MSCRM Anfrage zuweisen ...";
            //        else this.Title = _sync_count.ToString()+" Zeiteinträge zu MSCRM Anfrage zuweisen ...";
            //    }
            //    else if (Sync_Entity_Type == "VC")
            //    {
            //        if (_sync_count == 1)
            //            this.Title = "Zeiteintrag zu MSCRM Verkaufschance zuweisen ...";
            //        else this.Title = _sync_count.ToString() + " Zeiteinträge zu MSCRM Verkaufschance zuweisen ...";
            //    }
            //}
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            // Search

            String _searchText = tbEntitySearch.Text.Trim();
            if (_searchText.Length > 1)
            {
                SearchResult.Rows.Clear();

                String _filterExpression = "";

                _filterExpression += "[type]='WORKPACKAGE'";
                _filterExpression += " OR ";
                _filterExpression += "[type]='OPPORTUNITY'";
                _filterExpression += " OR ";
                _filterExpression += "[type]='INCIDENT'";
                _filterExpression += " OR ";
                _filterExpression += "[type]='NEW_VERSION'";
                _filterExpression += " OR ";
                _filterExpression += "[type]='CAMPAIGN'";

                DataRow[] _result = TimeEntryCache.Select("("+_filterExpression+") AND [description] LIKE '%" + EscapeLikeValue(_searchText) + "%'");
                foreach (DataRow _row in _result)
                {
                    DataRow _nr = SearchResult.NewRow();

                    String _EntityType = _row["type"].ToString();

                    if (_EntityType == "WORKPACKAGE") _nr["gctype"] = "AP";
                    else if (_EntityType == "OPPORTUNITY") _nr["gctype"] = "VC";
                    else if (_EntityType == "INCIDENT") _nr["gctype"] = "ANF";
                    else if (_EntityType == "NEW_VERSION") _nr["gctype"] = "REL";
                    else if (_EntityType == "CAMPAIGN") _nr["gctype"] = "KA";

                    _nr["gcdescription"] = _row["description"].ToString();
                    _nr["gcinfo1"] = "";
                    _nr["gcinfo2"] = "";

                    //  _nr["gcmscrmid"] = _dr["mscrmid"];

                    SearchResult.Rows.Add(_nr);
                }

                SearchGrid.ExpandAllGroups();
            }

        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (SearchGrid.GetFocusedRowCellValue("gcdescription") != null)
                {
                    String _desc = SearchGrid.GetFocusedRowCellValue("gcdescription").ToString().Trim();
                    String _entitytype = SearchGrid.GetFocusedRowCellValue("gctype").ToString().Trim();
                    if (_desc != "" && _entitytype != "")
                    {
                        Selected_EntityType = _entitytype;
                        Selected_ID = _desc;
                        this.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("btnCreate_Click: " + ex.Message, rm.GetString("ErrorMessage"), MessageBoxButton.OK, MessageBoxImage.Error);
            }

            // Create MSCRM 
            //if (lv.SelectedItem != null)
            //{
            //    ListViewItem _li = (ListViewItem)lv.SelectedItem;
            //    Int32 _count = 0;

            //    // Iterate selected grid rows
            //    for (int i = 0; i < _parent_tv.SelectedRows.Count; i++)
            //    {
            //        System.Data.DataRowView _drv = (System.Data.DataRowView)_parent_tv.SelectedRows[i];
            //        DataRow _dr = _drv.Row;
            //        String _status = "";

            //        try
            //        {
            //            _status = _dr["status"].ToString().Trim().ToUpper();
            //        } catch
            //        {
            //        }

            //        if (_status == "NOSYNC")
            //        {
            //            // MessageBox.Show("id=" + _dr["id"].ToString());

            //            // MessageBox.Show("GUID=" + _li.Tag.ToString());

            //            Guid gSelectedParent = new Guid(_li.Tag.ToString());
            //            Double dHours = Convert.ToDouble(_dr["dauer"].ToString());

            //            new_timeentry te = new new_timeentry();
            //            te.new_name = _dr["taetigkeit"].ToString();
            //            te.new_description = "";
            //            te.new_annotation = "";

            //            Lookup lParent = new Lookup();
            //            lParent.Value = gSelectedParent;

            //            if (Sync_Entity_Type == "AP")
            //            {
            //                lParent.type = EntityName.new_workpackage.ToString();
            //                te.new_workpackageid = lParent;
            //            }
            //            else if (Sync_Entity_Type == "ANF")
            //            {
            //                lParent.type = EntityName.incident.ToString();
            //                te.new_incidentid = lParent;
            //            }
            //            else if (Sync_Entity_Type == "VC")
            //            {
            //                lParent.type = EntityName.opportunity.ToString();
            //                te.new_opportunityid = lParent;
            //            }

            //            CrmDecimal dDuration = new CrmDecimal();
            //            dDuration.Value = (decimal)dHours;
            //            te.new_durationh = dDuration;

            //            DateTime _dt = Convert.ToDateTime(_dr["datum"].ToString());

            //            CrmDateTime cdt = new CrmDateTime();
            //            cdt.Value = _dt.Month.ToString() + "-" + _dt.Day.ToString() + "-" + _dt.Year.ToString();
            //            te.new_date = cdt;

            //            Lookup lUser = new Lookup();
            //            lUser.Value = new Guid(_dr["mscrmuserid"].ToString());
            //            lUser.type = EntityName.systemuser.ToString();
            //            te.new_executeuserid = lUser;

            //            cdt = new CrmDateTime();
            //            cdt.Value = _dr["beginn"].ToString();
            //            te.new_starttime = cdt;

            //            cdt = new CrmDateTime();
            //            cdt.Value = _dr["ende"].ToString();
            //            te.new_endtime = cdt;

            //            // te.new_duration_measured = _dr["timemeasurements"].ToString();  ??????????????

            //            String _zusartzzeit = _dr["zusatzzeit"].ToString();
            //            String[] _zz = _zusartzzeit.Split(':');
            //            Int32 _zz_hours = Convert.ToInt32(_zz[0]);
            //            Int32 _zz_minutes = Convert.ToInt32(_zz[1]);
            //            Double _zz_duration = _zz_hours + (_zz_minutes / 60.0);

            //            CrmDecimal dDurationAdditional = new CrmDecimal();
            //            dDurationAdditional.Value = (decimal)_zz_duration;
            //            te.new_duration_additional = dDurationAdditional;

            //            try
            //            {
            //                Picklist pBilling = new Picklist();
            //                int iValue = Convert.ToInt16(_dr["mscrmbillingtypeid"].ToString());
            //                if (iValue >= 0)
            //                {
            //                    pBilling.name = (string)_dr["mscrmbillingtype"].ToString();
            //                    pBilling.Value = iValue;
            //                    te.new_billing = pBilling;
            //                }
            //            }
            //            catch (Exception ex)
            //            {
            //            }

            //            Guid teGuid = Guid.Empty;

            //            try
            //            {
            //                CrmService service = CRMManager.GetService();
            //                teGuid = service.Create(te);
            //            }
            //            catch (Exception ex)
            //            {
            //                MessageBox.Show(ex.Message);
            //            }

            //            if (teGuid == Guid.Empty)
            //            {
            //                MessageBox.Show("Record Not Saved");
            //            }
            //            else
            //            {
            //                _dr["status"] = "LINKED";
            //                _dr["mscrmguid"] = teGuid.ToString();
            //                _count++;
            //            }
            //        }
            //    }

            //    if (_count > 0)
            //    {
            //        MessageBox.Show("Es wurden " + _count.ToString() + " Zeiteinträge erfolgreich nach MSCRM übertragen!", "Hinweis", MessageBoxButton.OK, MessageBoxImage.Information);
            //    }

            //    this.Close();
            //}
        }

        public void RunIncidentSearch(String sUserName)
        {
            //RunSearch(sUserName, EntityName.incident.ToString(), "title", "incidentid");
        }

        public void RunOpportunitySearch(String sUserName)
        {
            //RunSearch(sUserName, EntityName.opportunity.ToString(), "name", "opportunityid");
        }

        public void RunWorkPackageSearch(String sUserName)
        {
            //RunSearch(sUserName, EntityName.new_workpackage.ToString(), "new_name", "new_workpackageid");
        }

        //private DataRow GetCacheItemByMSCRMID(String itemTYPE, String itemGUID)
        //{
        //    DataRow _rV = null;

        //    DataRow[] _result = _parent_timeentrycache.Select("[type]='" + itemTYPE + "' AND [mscrmid]='" + itemGUID+"'");
        //    if (_result.Length == 1)
        //    {
        //        _rV = _result[0];
        //    }

        //    return _rV;
        //}

        private string EscapeLikeValue(string value)
        {
            StringBuilder sb = new StringBuilder(value.Length);
            for (int i = 0; i < value.Length; i++)
            {
                char c = value[i];
                switch (c)
                {
                    case ']':
                    case '[':
                    case '%':
                    case '*':
                        sb.Append("[").Append(c).Append("]");
                        break;
                    case '\'':
                        sb.Append("''");
                        break;
                    default:
                        sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        private void RunSearch(String sSearchString, String sEntity, String sNameAttribute, String sIDAttribute)
        {
            String _searchText = tbEntitySearch.Text.Trim();
            if (_searchText.Length > 1)
            {
                SearchResult.Rows.Clear();

                DataRow[] _result = TimeEntryCache.Select("[description] LIKE '%" + EscapeLikeValue(_searchText) + "%'");
                foreach (DataRow _row in _result)
                {
                    DataRow _nr = SearchResult.NewRow();

                    String _EntityType = _row["type"].ToString();

                    if (_EntityType == "WORKPACKAGE") _nr["gctype"] = "AP";
                    else if (_EntityType == "OPPORTUNITY") _nr["gctype"] = "VC";
                    else if (_EntityType == "INCIDENT") _nr["gctype"] = "ANF";
                    else if (_EntityType == "NEW_VERSION") _nr["gctype"] = "REL";
                    else if (_EntityType == "CAMPAIGN") _nr["gctype"] = "KA";
                    
                    _nr["gcdescription"] = _row["description"].ToString();
                    _nr["gcinfo1"] = "";
                    _nr["gcinfo2"] = "";
                    
                    //  _nr["gcmscrmid"] = _dr["mscrmid"];

                    SearchResult.Rows.Add(_nr);
                }
            }


            /*
            sSearchString = sSearchString.Replace("*", "%");

            if (sSearchString.EndsWith("%") == false)
                sSearchString = sSearchString + "%";

            sDescFeildName = sNameAttribute;
            sDataFeildName = sIDAttribute;
            sEntityName = sEntity;

            CrmService service = CRMManager.GetService();

            // Create the ColumnSet that indicates the properties to be retrieved.
            ColumnSet cols = new ColumnSet();

            // Set the properties of the ColumnSet.
            cols.Attributes = new string[] { sNameAttribute, sIDAttribute };

            // Create the ConditionExpression.
            ConditionExpression condition = new ConditionExpression();

            // Set the condition for the retrieval to be when the contact's address' city is Sammamish.
            condition.AttributeName = sNameAttribute;
            condition.Operator = ConditionOperator.Like;
            condition.Values = new string[] { sSearchString };

            // Create the FilterExpression.
            FilterExpression filter = new FilterExpression();

            // Set the properties of the filter.
            filter.FilterOperator = LogicalOperator.And;
            filter.Conditions = new ConditionExpression[] { condition };

            // Create the QueryExpression object.
            QueryExpression query = new QueryExpression();

            // Set the properties of the QueryExpression object.
            query.EntityName = sEntity;
            query.ColumnSet = cols;
            query.Criteria = filter;

            RetrieveMultipleRequest retrieve = new RetrieveMultipleRequest();
            retrieve.Query = query;
            retrieve.ReturnDynamicEntities = true;
            RetrieveMultipleResponse retrieved = (RetrieveMultipleResponse)service.Execute(retrieve);
            ResultItems = retrieved.BusinessEntityCollection;

            // lv.Items.Clear();

            String sDesc = "";
            Guid gData = Guid.Empty;

            foreach (DynamicEntity de in ResultItems.BusinessEntities)
            {
                sDesc = "";
                gData = Guid.Empty;

                foreach (Property p in de.Properties)
                {
                    if (p.GetType().Name.ToLower() == "StringProperty".ToLower() && p.Name.ToLower() == sDescFeildName)
                        sDesc = ((StringProperty)(p)).Value;

                    if (p.GetType().Name.ToLower() == "KeyProperty".ToLower() && p.Name.ToLower() == sDataFeildName)
                        gData = ((KeyProperty)p).Value.Value;

                }

                if (string.IsNullOrEmpty(sDesc) == false && gData != Guid.Empty)
                {
                    ListViewItem _li = new ListViewItem();
                    _li.Tag = gData.ToString();
                    _li.Content = sDesc;

                    // lv.Items.Add(_li);

                    String _type = "";
                    if (Sync_Entity_Type == "AP")
                    {
                        _type = "WORKPACKAGE";
                    }
                    else if (Sync_Entity_Type == "ANF")
                    {
                        _type = "INCIDENT";
                    }
                    else if (Sync_Entity_Type == "VC")
                    {
                        _type = "OPPORTUNITY";
                    }

                    //if (GetCacheItemByMSCRMID(_type, gData.ToString()) == null)
                    //{
                    //    DataRow _dr = _parent_timeentrycache.NewRow();
                    //    _dr["id"] = Guid.NewGuid().ToString();
                    //    _dr["type"] = _type;
                    //    _dr["mscrmid"] = gData.ToString();
                    //    _dr["description"] = sDesc;
                    //    _parent_timeentrycache.Rows.Add(_dr);
                    //    _parent_timeentrycache.AcceptChanges();
                    //}
                }                    
            }
            */

            /*
            if (lv.Items.Count == 0)
            {
                searchInfo.Content = "Leider keine Einträge gefunden.";
            }
            else
            {
                if (lv.Items.Count > 1)
                     searchInfo.Content = "Es wurden "+lv.Items.Count.ToString()+" Einträge gefunden.";
                else searchInfo.Content = "Es wurden ein Eintrag gefunden.";
            }
             * */
        }

        private void SearchGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Klick
            try
            {
                if (SearchGrid.GetFocusedRowCellValue("gcdescription") != null)
                {
                    String _desc = SearchGrid.GetFocusedRowCellValue("gcdescription").ToString().Trim();
                    String _entitytype = SearchGrid.GetFocusedRowCellValue("gctype").ToString().Trim();
                    if (_desc != "" && _entitytype != "")
                    {
                        Selected_EntityType = _entitytype;
                        Selected_ID = _desc;
                        this.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("SearchGrid_MouseDoubleClick: " + ex.Message, rm.GetString("ErrorMessage"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}
