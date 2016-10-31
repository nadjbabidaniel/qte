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

namespace QuickTimeEntry
{
    // A24 JNE 17.9.2015 - Add New Entity was removed after Definition of RHE 16.9.2015

    /// <summary>
    /// Interaktionslogik für CreateNewReleaseEntity.xaml
    /// </summary>
    public partial class CreateNewReleaseEntity : Window
    {
        //public String Selected_ID = "";
        //public String Selected_EntityType = "";
        //public Guid Selected_EntityGUID = Guid.Empty;

        //public New_Version _entity = new New_Version();

        //SortedDictionary<String, String> pl_new_sprache = null;
        //SortedDictionary<String, String> pl_statuscode = null;

        //Guid gDefaultUser = Guid.Empty;
        //string sDefaultUser = "";

        //PicklistManager _pm = new PicklistManager();
        //IOrganizationService _OrganizationService;

        //DataTable _tec = null;

        //private SortedDictionary<String, String> ResolvePickList(String EntityName, String LogicalName)
        //{
        //    SortedDictionary<String, String> rV = new SortedDictionary<String, String>();

        //    String _picklist = _pm.GetPicklistStr(_OrganizationService, EntityName, LogicalName);
        //    if (_picklist != "")
        //    {
        //        String[] _picklist_items = _picklist .Split(';');
        //        foreach (String _item in _picklist_items)
        //        {
        //            String[] _picklist_item = _item.Split('|');
        //            if (rV.ContainsKey(_picklist_item[0]))
        //            {
        //                rV.Add(_picklist_item[0] + " (2)", _picklist_item[1]);
        //            }
        //            else
        //            {
        //                rV.Add(_picklist_item[0], _picklist_item[1]);
        //            }
        //        }
        //    }

        //    return rV;
        //}

        //private SortedDictionary<String, String> ResolveStatusList(String EntityName, String LogicalName)
        //{
        //    SortedDictionary<String, String> rV = new SortedDictionary<String, String>();

        //    String _picklist = _pm.GetStatuslistStr(_OrganizationService, EntityName, LogicalName);
        //    if (_picklist != "")
        //    {
        //        String[] _picklist_items = _picklist.Split(';');
        //        foreach (String _item in _picklist_items)
        //        {
        //            String[] _picklist_item = _item.Split('|');
        //            if (rV.ContainsKey(_picklist_item[0]))
        //            {
        //                rV.Add(_picklist_item[0] + " (2)", _picklist_item[1]);
        //            }
        //            else
        //            {
        //                rV.Add(_picklist_item[0], _picklist_item[1]);
        //            }
        //        }
        //    }

        //    return rV;
        //}

        //public CreateNewReleaseEntity(IOrganizationService _Service, DataTable _TimeEntryCache)
        //{
        //    InitializeComponent();

        //    _OrganizationService = _Service;

        //    _tec = _TimeEntryCache;

        //    // Initialize Picklists
        //    pl_new_sprache = ResolvePickList(New_Version.EntityLogicalName, "new_sprache");
            
        //    cbSprache.Items.Clear();
        //    foreach (String _item in pl_new_sprache.Keys)
        //    {
        //        cbSprache.Items.Add(_item);
        //    }
        //    cbSprache.Text = "Deutsch";

        //    cbStatusgrund.Items.Clear();

        //    pl_statuscode = ResolveStatusList(New_Version.EntityLogicalName, "statuscode");
        //    foreach (String _item in pl_statuscode.Keys)
        //    {
        //        cbStatusgrund.Items.Add(_item);
        //    }
        //    cbStatusgrund.Text = "In Entwicklung";

        //    dpReleaseDatum.Text = DateTime.Now.ToLongDateString();
        //    teVersionsnummer.Text = "";
        //    teProdukt.Text = "";
        //    teProduktmanager.Text = "";

        //    object[] UserDtls = CRMManager.GetCurrentUser(_OrganizationService, _TimeEntryCache, false);
        //    gDefaultUser = (Guid)UserDtls[0];
        //    sDefaultUser = (string)UserDtls[1];

        //    teEntwickler.Text = sDefaultUser;

        //    _entity.New_EntwicklerId = new EntityReference(SystemUser.EntityLogicalName, new Guid(gDefaultUser.ToString()));
        //}

        //private void btnStoreTemplate_Click(object sender, RoutedEventArgs e)
        //{
        //    // Store as Template

        //}

        //private void btnAccept_Click(object sender, RoutedEventArgs e)
        //{
        //    // Accept

        //    try
        //    {
        //        if (teProdukt.Text.Trim() == "" || teVersionsnummer.Text.Trim() == "" || teVersionsname.Text.Trim() == "" || cbSprache.Text.Trim() == "" || teEntwickler.Text.Trim() == "" || teProduktmanager.Text.Trim() == "")
        //        {
        //            MessageBox.Show("Bitte füllen Sie mindestens alle ROT markierten Felder für die neue Anfrage aus!", "Wichtiger Hinweis", MessageBoxButton.OK, MessageBoxImage.Exclamation);
        //        }
        //        else
        //        {
        //            _entity.New_Sprache = new OptionSetValue(Convert.ToInt32(pl_new_sprache[cbSprache.Text]));
        //            _entity.StatusCode = new OptionSetValue(Convert.ToInt32(pl_statuscode[cbStatusgrund.Text]));

        //            _entity.New_versionsnummer = teVersionsname.Text;
        //            _entity.New_Versionsnr = teVersionsnummer.Text;
        //            _entity.New_aenderung = teReleasenotes.Text;

        //            _entity.New_releasedatum = ConvertDateTime(Convert.ToDateTime(dpReleaseDatum.Text));
        //            // _entity.new_releasedatum = ConvertToCRMDateTime(Convert.ToDateTime(dpReleaseDatum.Text));

        //            //ConvertToCRMDateTime(Convert.ToDateTime(dpReleaseDatum.Text));

        //            Selected_EntityGUID = _OrganizationService.Create(_entity);

        //            Selected_EntityType = "REL";
        //            Selected_ID = _entity.New_versionsnummer;

        //            this.Close();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //        if (ex.InnerException != null) MessageBox.Show(ex.InnerException.Message);
        //    }
        //}

        //static public DateTime ConvertDateTime(DateTime dateTime)
        //{
        //    DateTime _DateTime = new DateTime();

        //    TimeSpan offset = TimeZone.CurrentTimeZone.GetUtcOffset(dateTime);

        //    string sOffset = string.Empty;

        //    if (offset.Hours < 0)
        //    {
        //        sOffset = "-" + (offset.Hours * -1).ToString().PadLeft(2, '0');
        //    }

        //    {
        //        sOffset = "+" + offset.Hours.ToString().PadLeft(2, '0');
        //    }

        //    sOffset += offset.Minutes.ToString().PadLeft(2, '0');

        //    _DateTime = Convert.ToDateTime(dateTime.ToString(string.Format("yyyy-MM-ddTHH:mm:ss{0}", sOffset)));

        //    return _DateTime;
        //}     

        ///*
        //static public CrmDateTime ConvertToCRMDateTime(DateTime dateTime)
        //{
        //    CrmDateTime crmDateTime = new CrmDateTime();

        //    crmDateTime.date = dateTime.ToShortDateString();
        //    crmDateTime.time = dateTime.ToShortTimeString();

        //    TimeSpan offset = TimeZone.CurrentTimeZone.GetUtcOffset(dateTime);

        //    string sOffset = string.Empty;

        //    if (offset.Hours < 0)
        //    {
        //        sOffset = "-" + (offset.Hours * -1).ToString().PadLeft(2, '0');
        //    }

        //    {
        //        sOffset = "+" + offset.Hours.ToString().PadLeft(2, '0');
        //    }

        //    sOffset += offset.Minutes.ToString().PadLeft(2, '0');
        //    crmDateTime.Value = dateTime.ToString(string.Format("yyyy-MM-ddTHH:mm:ss{0}", sOffset));

        //    return crmDateTime;
        //}     
        // * */

        //private void btnClose_Click(object sender, RoutedEventArgs e)
        //{
        //    // Close
        //    Selected_ID = "";
        //    Selected_EntityType = "";
        //    Selected_EntityGUID = Guid.Empty;
        //    this.Close();
        //}

        //private String NewVersion(String OldVersion)
        //{
        //    String rV = "";
        //    Int32 anz_nks = 0;

        //    if (OldVersion.IndexOf(".") != -1) anz_nks = OldVersion.Length - OldVersion.IndexOf(".") - 1;

        //    Double _old_version = 0.0;

        //    try
        //    {
        //        _old_version = Convert.ToDouble(OldVersion.Replace('.', ','));
        //    }
        //    catch { }

        //    if (anz_nks > 0) _old_version += (1.0 / Math.Pow(10, anz_nks)); else _old_version += 1.0;

        //    rV = _old_version.ToString().Replace(',', '.');

        //    return rV;
        //}

        //private void btnSearchProdukt_Click(object sender, RoutedEventArgs e)
        //{
        //    // Search Produkt
        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teProdukt.Text.Trim(), "", true, Product.EntityLogicalName, "name", "productid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();

        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teProdukt.Text = _modal_dialog.Select_EntityName;
        //        _entity.New_AddonId = _modal_dialog._lookup;

        //        Product _p = (Product)_OrganizationService.Retrieve(Product.EntityLogicalName, new Guid(_modal_dialog.Select_EntityID), new ColumnSet(true));
        //        if (_p != null)
        //        {
        //            _entity.New_ProduktmanagerId = _p.New_ProduktmanagerId;
        //            teProduktmanager.Text = _p.New_ProduktmanagerId.Name;

        //            DataRow[] _result = _tec.Select("[type]='NEW_VERSION' AND [mscrmparentid]='" + _modal_dialog._lookup.Id.ToString() + "'");
        //            if (_result.Length > 0)
        //            {
        //                DateTime _latest = DateTime.MinValue;
        //                Guid _latest_guid = Guid.Empty;

        //                foreach (DataRow _r in _result)
        //                {
        //                    DateTime _dt = DateTime.MinValue;

        //                    try
        //                    {
        //                        _dt = Convert.ToDateTime(_r["zusatzinfo1"].ToString());
        //                    } catch { }

        //                    if (_dt > _latest)
        //                    {
        //                        _latest_guid = new Guid(_r["id"].ToString());
        //                    }
        //                }

        //                if (_latest_guid != Guid.Empty)
        //                {
        //                    DataRow[] _result_2 = _tec.Select("[type]='NEW_VERSION' AND [id]='" + _latest_guid.ToString() + "'");
        //                    if (_result_2.Length == 1)
        //                    {
        //                        String _new_version = NewVersion(_result_2[0]["zusatzinfo2"].ToString());
        //                        String _new_product_name = _result_2[0]["description"].ToString().Replace(_result_2[0]["zusatzinfo2"].ToString(), _new_version);

        //                        teVersionsname.Text = _new_product_name;
        //                        teVersionsnummer.Text = _new_version;
        //                    }
        //                }
        //                else
        //                {
        //                    teVersionsname.Text = _p.Name+" v1.00";
        //                    teVersionsnummer.Text = "1.00";
        //                }

        //            }

        //        }
        //    }
        //}

        //private void btnSearchEntwickler_Click(object sender, RoutedEventArgs e)
        //{
        //    // Search User
        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teEntwickler.Text.Trim(), "", true, SystemUser.EntityLogicalName, "fullname", "systemuserid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();

        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teEntwickler.Text = _modal_dialog.Select_EntityName;
        //        _entity.New_EntwicklerId = _modal_dialog._lookup;
        //    }
        //}

        //private void btnSearchProduktmanager_Click(object sender, RoutedEventArgs e)
        //{
        //    // Search User
        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teProduktmanager.Text.Trim(), "", true, SystemUser.EntityLogicalName, "fullname", "systemuserid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();

        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teProduktmanager.Text = _modal_dialog.Select_EntityName;
        //        _entity.New_ProduktmanagerId = _modal_dialog._lookup;
        //    }
        //}

        //private void teProdukt_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teProdukt.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        // Search Produkt
        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teProdukt.Text.Trim(), "", true, Product.EntityLogicalName, "name", "productid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teProdukt.Text = _modal_dialog.Select_EntityName;
        //            _entity.New_AddonId = _modal_dialog._lookup;

        //            Product _p = (Product)_OrganizationService.Retrieve(Product.EntityLogicalName, new Guid(_modal_dialog.Select_EntityID), new ColumnSet(true));
        //            if (_p != null)
        //            {
        //                _entity.New_ProduktmanagerId = _p.New_ProduktmanagerId;
        //                teProduktmanager.Text = _p.New_ProduktmanagerId.Name;

        //                DataRow[] _result = _tec.Select("[type]='NEW_VERSION' AND [mscrmparentid]='" + _modal_dialog._lookup.Id.ToString() + "'");
        //                if (_result.Length > 0)
        //                {
        //                    DateTime _latest = DateTime.MinValue;
        //                    Guid _latest_guid = Guid.Empty;

        //                    foreach (DataRow _r in _result)
        //                    {
        //                        DateTime _dt = DateTime.MinValue;

        //                        try
        //                        {
        //                            _dt = Convert.ToDateTime(_r["zusatzinfo1"].ToString());
        //                        }
        //                        catch { }

        //                        if (_dt > _latest)
        //                        {
        //                            _latest_guid = new Guid(_r["id"].ToString());
        //                        }
        //                    }

        //                    if (_latest_guid != Guid.Empty)
        //                    {
        //                        DataRow[] _result_2 = _tec.Select("[type]='NEW_VERSION' AND [id]='" + _latest_guid.ToString() + "'");
        //                        if (_result_2.Length == 1)
        //                        {
        //                            String _new_version = NewVersion(_result_2[0]["zusatzinfo2"].ToString());
        //                            String _new_product_name = _result_2[0]["description"].ToString().Replace(_result_2[0]["zusatzinfo2"].ToString(), _new_version);

        //                            teVersionsname.Text = _new_product_name;
        //                            teVersionsnummer.Text = _new_version;
        //                        }
        //                    }
        //                    else
        //                    {
        //                        teVersionsname.Text = _p.Name+ " v1.00";
        //                        teVersionsnummer.Text = "1.00";
        //                    }
        //                }
        //            }
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}

        //private void teEntwickler_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teEntwickler.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teEntwickler.Text.Trim(), "", true, SystemUser.EntityLogicalName, "fullname", "systemuserid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teEntwickler.Text = _modal_dialog.Select_EntityName;
        //            _entity.New_EntwicklerId = _modal_dialog._lookup;
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}

        //private void teProduktmanager_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teProduktmanager.Text.Trim() != "")
        //    {
        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teProduktmanager.Text.Trim(), "", true, SystemUser.EntityLogicalName, "fullname", "systemuserid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teProduktmanager.Text = _modal_dialog.Select_EntityName;
        //            _entity.New_ProduktmanagerId = _modal_dialog._lookup;
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}
    }
}
