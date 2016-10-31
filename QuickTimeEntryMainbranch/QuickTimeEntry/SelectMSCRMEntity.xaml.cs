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
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using System.Reflection;
using System.Resources;

namespace QuickTimeEntry
{
    /// <summary>
    /// Interaktionslogik für SelectMSCRMEntity.xaml
    /// </summary>
    public partial class SelectMSCRMEntity : Window
    {
        public EntityReference _lookup = new EntityReference();

        public String Select_EntityName = "";
        public String Select_EntityID = "";

        private DataTable _SearchResult = null;

        private String _Entity = "";
        private String _NameAttribute = "";
        private String _IDAttribute = "";

        private IOrganizationService _Service = null;

        private static ResourceManager rm = new ResourceManager("QuickTimeEntry.SelectMSCRMEntity", Assembly.GetExecutingAssembly());

        public SelectMSCRMEntity(IOrganizationService Service, String SearchStr, String Title, Boolean ShowAllAtStart, String Entity, String NameAttribute, String IDAttribute, Boolean AutoAcceptOneResult)
        {
            InitializeComponent();

            _Service = Service;

            _SearchResult = new DataTable("searchresult");
            _SearchResult.Columns.Add("id", typeof(String));
            _SearchResult.Columns.Add("name", typeof(String));

            SearchGrid.ItemsSource = _SearchResult;

            SearchGrid.IsFilterEnabled = true;
            SearchGrid.RefreshData();

            foreach (DevExpress.Xpf.Grid.GridColumn _gc in SearchGrid.Columns)
            {
                _gc.AllowAutoFilter = true;
                _gc.AutoFilterCondition = DevExpress.Xpf.Grid.AutoFilterCondition.Contains;
                _gc.AllowEditing = DevExpress.Utils.DefaultBoolean.False;
            }

            // SearchGrid.GroupBy(SearchGrid.Columns["gctype"], DevExpress.Data.ColumnSortOrder.Ascending);
            // SearchGrid.ExpandAllGroups();

            _Entity = Entity;
            _NameAttribute = NameAttribute;
            _IDAttribute = IDAttribute;

            if (SearchStr != "") tbEntitySearch.Text = SearchStr.Trim();

            if (ShowAllAtStart && tbEntitySearch.Text != "")
            {
                Int32 _count = GetSearchList(_Service, tbEntitySearch.Text.Trim(), _Entity, _NameAttribute, _IDAttribute);
                if (_count == 1 && AutoAcceptOneResult)
                {
                    Accept();
                }
            }

            if (Title != "") this.Title = Title;

            tbEntitySearch.Focus();
        }

        public Int32 GetSearchList(IOrganizationService _Service, String sSearchString, String sEntity, String sNameAttribute, String sIDAttribute)
        {
            Int32 rV = 0;
            this.Cursor = Cursors.Wait;

            btnAccept.IsEnabled = false;
            btnClose.IsEnabled = false;

            if (_Service != null)
            {
                if (sSearchString.IndexOf("%") == -1) sSearchString = "%" + sSearchString + "%";

                FilterExpression filterExpression = new FilterExpression();
                ConditionExpression condition = new ConditionExpression(sNameAttribute, ConditionOperator.Like, sSearchString);
                ConditionExpression condition2 = new ConditionExpression("isdisabled", ConditionOperator.Equal, "false");

                filterExpression.FilterOperator = LogicalOperator.And;

                if (sEntity == SystemUser.EntityLogicalName)
                {
                    filterExpression.AddCondition(condition);
                    filterExpression.AddCondition(condition2);
                }
                else
                {
                    filterExpression.AddCondition(condition);
                }

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

                _SearchResult.Rows.Clear();

                RetrieveMultipleResponse response = (RetrieveMultipleResponse)_Service.Execute(multipleRequest);

                foreach (Entity _e in response.EntityCollection.Entities)
                {
                    rV++;

                    if (_e.Attributes[sNameAttribute] != null && _e.Attributes[sIDAttribute] != null)
                    {
                        DataRow _dr = null;
                        if ((_dr = _SearchResult.NewRow()) != null)
                        {
                            _dr["id"] = _e.Attributes[sIDAttribute].ToString();
                            _dr["name"] = _e.Attributes[sNameAttribute].ToString();
                            _SearchResult.Rows.Add(_dr);
                        }
                    }
                }

                SearchGrid.RefreshData();
            }

            btnAccept.IsEnabled = true;
            btnClose.IsEnabled = true;

            this.Cursor = Cursors.Arrow;

            return rV;
        }

        private void btnSearch_Click(object sender, RoutedEventArgs e)
        {
            // Search Entity
            if (tbEntitySearch.Text.Trim() != "")
            {
                GetSearchList(_Service, tbEntitySearch.Text.Trim(), _Entity, _NameAttribute, _IDAttribute);
            }
        }

        private void Accept()
        {
            // Übernahmen
            try
            {
                if (SearchGrid.GetFocusedRowCellValue("id") != null)
                {
                    String _id = SearchGrid.GetFocusedRowCellValue("id").ToString().Trim();
                    String _name = SearchGrid.GetFocusedRowCellValue("name").ToString().Trim();

                    if (_id != "" && _name != "")
                    {
                        Select_EntityName = _name;
                        Select_EntityID = _id;

                        _lookup.LogicalName = _Entity;
                        _lookup.Name = _name;
                        _lookup.Id = new Guid(_id);

                        this.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("btnAccept_Click: " + ex.Message, rm.GetString("ErrorNotification"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnAccept_Click(object sender, RoutedEventArgs e)
        {
            Accept();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            // Abbrechen
            Select_EntityName = "";
            Select_EntityID = "";

            this.Close();
        }

        private void SearchGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Click Grid
            try
            {
                if (SearchGrid.GetFocusedRowCellValue("id") != null)
                {
                    String _id = SearchGrid.GetFocusedRowCellValue("id").ToString().Trim();
                    String _name = SearchGrid.GetFocusedRowCellValue("name").ToString().Trim();

                    if (_id != "" && _name != "")
                    {
                        Select_EntityName = _name;
                        Select_EntityID = _id;

                        _lookup.LogicalName = _Entity;
                        _lookup.Name = _name;
                        _lookup.Id = new Guid(_id);

                        this.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("btnAccept_Click: " + ex.Message, rm.GetString("ErrorNotification"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

    }
}
