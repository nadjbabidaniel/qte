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
using System.Web.Services.Protocols;

using CCMA24Produktiv;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace QuickTimeEntry
{
    // A24 JNE 17.9.2015 - Add New Entity was removed after Definition of RHE 16.9.2015

    /// <summary>
    /// Interaktionslogik für CreateNewIncidentEntity.xaml
    /// </summary>
    public partial class CreateNewIncidentEntity : Window
    {
        //public String Selected_ID = "";
        //public String Selected_EntityType = "";
        //public Guid Selected_EntityGUID = Guid.Empty;

        //SortedDictionary<String, String> pl_caseorigincode = null;
        //SortedDictionary<String, String> pl_casetypecode = null;
        //SortedDictionary<String, String> pl_new_geschaeftsbereich = null;
        //SortedDictionary<String, String> pl_new_abrechnungsweise = null;
        //SortedDictionary<String, String> pl_new_bestehendevertraege_a24 = null;
        //SortedDictionary<String, String> pl_new_rechnungdurch = null;
        //SortedDictionary<String, String> pl_new_interneverrechnung = null;
        //SortedDictionary<String, String> pl_incidentstagecode = null;
        //SortedDictionary<String, String> pl_statuscode = null;
        //SortedDictionary<String, String> pl_severitycode = null;
        //SortedDictionary<String, String> pl_new_dringlichkeit = null;
        //SortedDictionary<String, String> pl_new_gemakategorie = null;

        //private String _User = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString();


        //public Incident _entity = new Incident();

        //Guid gDefaultUser = Guid.Empty;
        //string sDefaultUser = "";
        
        //PicklistManager _pm = new PicklistManager();

        //DataTable _tec = null;
        //IOrganizationService _OrganizationService;

        //private DataRow GetCacheItemBySubType(String itemTYPE, String SubType)
        //{
        //    DataRow _rV = null;

        //    DataRow[] _result = _tec.Select("[type]='" + itemTYPE + "' AND [subtype]='" + SubType + "'");
        //    if (_result.Length == 1)
        //    {
        //        _rV = _result[0];
        //    }

        //    return _rV;
        //}
        
        //private SortedDictionary<String, String> ResolvePickList(String EntityName, String LogicalName)
        //{
        //    SortedDictionary<String, String> rV = new SortedDictionary<String, String>();

        //    String _picklist = _pm.GetPicklistStr(_OrganizationService, EntityName, LogicalName);
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
        //                rV.Add(_picklist_item[0]+" (2)", _picklist_item[1]);
        //            }
        //            else
        //            {
        //                rV.Add(_picklist_item[0], _picklist_item[1]);
        //            }
        //        }
        //    }

        //    return rV;
        //}

        //public CreateNewIncidentEntity(IOrganizationService _Service, DataTable _TimeEntryCache)
        //{
        //    InitializeComponent();

        //    _OrganizationService = _Service;

        //    _tec = _TimeEntryCache;

        //    // Initialize Picklists

        //    // _entity.caseorigincode 
        //    pl_caseorigincode = ResolvePickList(Incident.EntityLogicalName, "caseorigincode");
        //    cbAnfrageursprung.Items.Clear(); foreach (String _item in pl_caseorigincode.Keys) cbAnfrageursprung.Items.Add(_item); cbAnfrageursprung.Text = "";
            
        //    // _entity.casetypecode 
        //    pl_casetypecode = ResolvePickList(Incident.EntityLogicalName, "casetypecode");
        //    cbAnfragetyp.Items.Clear(); foreach (String _item in pl_casetypecode.Keys) cbAnfragetyp.Items.Add(_item); cbAnfragetyp.Text = "";

        //    // _entity.new_geschaeftsbereich 
        //    pl_new_geschaeftsbereich = ResolvePickList(Incident.EntityLogicalName, "new_geschaeftsbereich");
        //    cbBusinessSolutions.Items.Clear(); foreach (String _item in pl_new_geschaeftsbereich.Keys) cbBusinessSolutions.Items.Add(_item); cbBusinessSolutions.Text = "";

        //    // _entity.new_abrechnungsweise 
        //    pl_new_abrechnungsweise = ResolvePickList(Incident.EntityLogicalName, "new_abrechnungsweise");
        //    cbAbrechnungsweise.Items.Clear(); foreach (String _item in pl_new_abrechnungsweise.Keys) cbAbrechnungsweise.Items.Add(_item); cbAbrechnungsweise.Text = "";

        //    // _entity.new_bestehendevertraege_a24 
        //    pl_new_bestehendevertraege_a24 = ResolvePickList(Incident.EntityLogicalName, "new_bestehendevertraege_a24");
        //    cbBestehendeVertraege.Items.Clear(); foreach (String _item in pl_new_bestehendevertraege_a24.Keys) cbBestehendeVertraege.Items.Add(_item); cbBestehendeVertraege.Text = "";

        //    // _entity.new_rechnungdurch 
        //    pl_new_rechnungdurch = ResolvePickList(Incident.EntityLogicalName, "new_rechnungdurch");
        //    cbRechnungDurch.Items.Clear(); foreach (String _item in pl_new_rechnungdurch.Keys) cbRechnungDurch.Items.Add(_item); cbRechnungDurch.Text = "";

        //    // _entity.new_interneverrechnung 
        //    pl_new_interneverrechnung = ResolvePickList(Incident.EntityLogicalName, "new_interneverrechnung");
        //    cbInterneVerrechnung.Items.Clear(); foreach (String _item in pl_new_interneverrechnung.Keys) cbInterneVerrechnung.Items.Add(_item); cbInterneVerrechnung.Text = "";

        //    // _entity.incidentstagecode 
        //    pl_incidentstagecode = ResolvePickList(Incident.EntityLogicalName, "incidentstagecode");
        //    cbAnfragephase.Items.Clear(); foreach (String _item in pl_incidentstagecode.Keys) cbAnfragephase.Items.Add(_item); cbAnfragephase.Text = "Erfassung / Neu";

        //    // _entity.severitycode 
        //    pl_severitycode = ResolvePickList(Incident.EntityLogicalName, "severitycode");
        //    cbSchweregrad.Items.Clear(); foreach (String _item in pl_severitycode.Keys) cbSchweregrad.Items.Add(_item); cbSchweregrad.Text = "";

        //    // _entity.new_dringlichkeit 
        //    pl_new_dringlichkeit = ResolvePickList(Incident.EntityLogicalName, "new_dringlichkeit");
        //    cbDringlichkeit.Items.Clear(); foreach (String _item in pl_new_dringlichkeit.Keys) cbDringlichkeit.Items.Add(_item); cbDringlichkeit.Text = "Normal";

        //    // _entity.new_gemakategorie 
        //    pl_new_gemakategorie = ResolvePickList(Incident.EntityLogicalName, "new_gemakategorie");
        //    cbGEMAKategorie.Items.Clear(); foreach (String _item in pl_new_gemakategorie.Keys) cbGEMAKategorie.Items.Add(_item); cbGEMAKategorie.Text = "(nicht definiert)";

        //    // (Statuscode) _entity.statuscode 
        //    pl_statuscode = ResolveStatusList(Incident.EntityLogicalName, "statuscode");
        //    cbStatus.Items.Clear(); foreach (String _item in pl_statuscode.Keys) cbStatus.Items.Add(_item); cbStatus.Text = "";

        //    // Text:
        //    teTitel.Text = ""; _entity.Title = "";
        //    teBeschreibung.Text = ""; _entity.Description = "";

        //    try
        //    {
        //        object[] UserDtls = CRMManager.GetCurrentUser(_OrganizationService, _TimeEntryCache, false);
        //        gDefaultUser = (Guid)UserDtls[0];
        //        sDefaultUser = (string)UserDtls[1];
        //    }
        //    catch { }

        //    DataRow _dr = null;

        //    String _tmp_User = _User;
        //    _tmp_User = _tmp_User.Replace("\\", "_");
        //    _tmp_User = _tmp_User.Replace(".", "_");

        //    teMitarbeiter.Text = sDefaultUser;
        //    _entity.OwnerId = new EntityReference(SystemUser.EntityLogicalName, new Guid(gDefaultUser.ToString()));

        //    teAktuellerBearbeiter.Text = sDefaultUser;
        //    _entity.New_AktuellerBearbeiterId = new EntityReference(SystemUser.EntityLogicalName, new Guid(gDefaultUser.ToString()));

        //    // Check Template

        //    // Anfrageursprung		DD 	caseorigincode			MUST
        //    if ((_dr = GetCacheItemBySubType("TEMPLATE_" + _tmp_User, "INCIDENT_CASEORIGINCODE")) != null)
        //    {
        //        foreach (KeyValuePair<String, String> _kvp in pl_caseorigincode)
        //        {
        //            if (_kvp.Value == _dr["zusatzinfo1"].ToString())
        //            {
        //                cbAnfrageursprung.Text = _kvp.Key;
        //                break;
        //            }
        //        }
        //    }

        //    // Anfragetyp		DD	casetypecode			MUST
        //    if ((_dr = GetCacheItemBySubType("TEMPLATE_" + _tmp_User, "INCIDENT_CASETYPECODE")) != null)
        //    {
        //        foreach (KeyValuePair<String, String> _kvp in pl_casetypecode)
        //        {
        //            if (_kvp.Value == _dr["zusatzinfo1"].ToString())
        //            {
        //                cbAnfragetyp.Text = _kvp.Key;
        //                break;
        //            }
        //        }
        //    }

        //    // Business Solutions	DD	new_geschaeftsbereich		MUST
        //    if ((_dr = GetCacheItemBySubType("TEMPLATE_" + _tmp_User, "INCIDENT_NEW_GESCHAEFTSBEREICH")) != null)
        //    {
        //        foreach (KeyValuePair<String, String> _kvp in pl_new_geschaeftsbereich)
        //        {
        //            if (_kvp.Value == _dr["zusatzinfo1"].ToString())
        //            {
        //                cbBusinessSolutions.Text = _kvp.Key;
        //                break;
        //            }
        //        }
        //    }

        //    // Betreff			SEARCH	subjectid			MUST
        //    if ((_dr = GetCacheItemBySubType("TEMPLATE_" + _tmp_User, "INCIDENT_SUBJECT")) != null)
        //    {
        //        if (_dr["zusatzinfo1"].ToString() != "")
        //        {
        //            Subject _subject = (Subject)_OrganizationService.Retrieve(Subject.EntityLogicalName, new Guid(_dr["zusatzinfo1"].ToString()), new ColumnSet(true));
        //            if (_subject != null)
        //            {
        //                teBetreff.Text = _subject.Title;
        //                _entity.SubjectId = new EntityReference(Subject.EntityLogicalName, new Guid(_dr["zusatzinfo1"].ToString()));
        //            }
        //        }                    
        //    }

        //    dpDeadline.SelectedDate = DateTime.Now;
        //    dpNachverfolgung.SelectedDate = DateTime.Now;
        //}

        //private void btnStoreTemplate_Click(object sender, RoutedEventArgs e)
        //{
        //    // Store Template

        //    _entity.CaseOriginCode = null;
        //    _entity.CaseTypeCode = null;
        //    _entity.New_Geschaeftsbereich = null;

        //    //if ((_picklist = new Picklist()) != null) { _picklist.IsNull = true; _picklist.IsNullSpecified = true; _entity.caseorigincode = _picklist; }
        //    //if ((_picklist = new Picklist()) != null) { _picklist.IsNull = true; _picklist.IsNullSpecified = true; _entity.casetypecode = _picklist; }
        //    //if ((_picklist = new Picklist()) != null) { _picklist.IsNull = true; _picklist.IsNullSpecified = true; _entity.new_geschaeftsbereich = _picklist; }

        //    if (cbAnfrageursprung.Text.Trim() != "") _entity.CaseOriginCode = new OptionSetValue(Convert.ToInt32(pl_caseorigincode[cbAnfrageursprung.Text]));
        //    if (cbAnfragetyp.Text.Trim() != "") _entity.CaseTypeCode = new OptionSetValue(Convert.ToInt32(pl_casetypecode[cbAnfragetyp.Text]));
        //    if (cbBusinessSolutions.Text.Trim() != "") _entity.New_Geschaeftsbereich = new OptionSetValue(Convert.ToInt32(pl_new_geschaeftsbereich[cbBusinessSolutions.Text]));

        //    String _tmp_User = _User;
        //    _tmp_User = _tmp_User.Replace("\\", "_");
        //    _tmp_User = _tmp_User.Replace(".", "_");

        //    DataRow _dr = null;

        //    // Anfrageursprung		DD 	caseorigincode			MUST
        //    if ((_dr=GetCacheItemBySubType("TEMPLATE_" + _tmp_User, "INCIDENT_CASEORIGINCODE")) == null)
        //    {
        //        // Create New
        //        _dr = _tec.NewRow();
        //        _dr["id"] = Guid.NewGuid().ToString();
        //        _dr["type"] = "TEMPLATE_" + _tmp_User;
        //        _dr["subtype"] = "INCIDENT_CASEORIGINCODE";
        //        _dr["mscrmid"] = "";
        //        _dr["mscrmparentid"] = "";
        //        _dr["description"] = "";
        //        if (_entity.CaseOriginCode != null)
        //            _dr["zusatzinfo1"] = _entity.CaseOriginCode.Value.ToString();
        //        else _dr["zusatzinfo1"] = "";
        //        _dr["zusatzinfo2"] = "";
        //        _tec.Rows.Add(_dr);
        //        _tec.AcceptChanges();
        //    }
        //    else
        //    {
        //        // Update
        //        if (_entity.CaseOriginCode != null)
        //            _dr["zusatzinfo1"] = _entity.CaseOriginCode.Value.ToString();
        //        else _dr["zusatzinfo1"] = "";
        //        _tec.AcceptChanges();
        //    }

        //    // Anfragetyp		DD	casetypecode			MUST
        //    if ((_dr = GetCacheItemBySubType("TEMPLATE_" + _tmp_User, "INCIDENT_CASETYPECODE")) == null)
        //    {
        //        // Create New
        //        _dr = _tec.NewRow();
        //        _dr["id"] = Guid.NewGuid().ToString();
        //        _dr["type"] = "TEMPLATE_" + _tmp_User;
        //        _dr["subtype"] = "INCIDENT_CASETYPECODE";
        //        _dr["mscrmid"] = "";
        //        _dr["mscrmparentid"] = "";
        //        _dr["description"] = "";
        //        if (_entity.CaseTypeCode != null)
        //            _dr["zusatzinfo1"] = _entity.CaseTypeCode.Value.ToString();
        //        else _dr["zusatzinfo1"] = "";

        //        _dr["zusatzinfo2"] = "";
        //        _tec.Rows.Add(_dr);
        //        _tec.AcceptChanges();
        //    }
        //    else
        //    {
        //        // Update
        //        if (_entity.CaseTypeCode != null)
        //            _dr["zusatzinfo1"] = _entity.CaseTypeCode.Value.ToString();
        //        else _dr["zusatzinfo1"] = "";
        //        _tec.AcceptChanges();
        //    }

        //    // Business Solutions	DD	new_geschaeftsbereich		MUST
        //    if ((_dr = GetCacheItemBySubType("TEMPLATE_" + _tmp_User, "INCIDENT_NEW_GESCHAEFTSBEREICH")) == null)
        //    {
        //        // Create New
        //        _dr = _tec.NewRow();
        //        _dr["id"] = Guid.NewGuid().ToString();
        //        _dr["type"] = "TEMPLATE_" + _tmp_User;
        //        _dr["subtype"] = "INCIDENT_NEW_GESCHAEFTSBEREICH";
        //        _dr["mscrmid"] = "";
        //        _dr["mscrmparentid"] = "";
        //        _dr["description"] = "";
        //        if (_entity.New_Geschaeftsbereich != null)
        //            _dr["zusatzinfo1"] = _entity.New_Geschaeftsbereich.Value.ToString();
        //        else _dr["zusatzinfo1"] = "";
        //        _dr["zusatzinfo2"] = "";
        //        _tec.Rows.Add(_dr);
        //        _tec.AcceptChanges();
        //    }
        //    else
        //    {
        //        // Update
        //        if (_entity.New_Geschaeftsbereich != null)
        //            _dr["zusatzinfo1"] = _entity.New_Geschaeftsbereich.Value.ToString();
        //        else _dr["zusatzinfo1"] = "";
        //        _tec.AcceptChanges();
        //    }

        //    // Betreff			SEARCH	subjectid			MUST
        //    if ((_dr = GetCacheItemBySubType("TEMPLATE_" + _tmp_User, "INCIDENT_SUBJECT")) == null)
        //    {
        //        // Create New
        //        _dr = _tec.NewRow();
        //        _dr["id"] = Guid.NewGuid().ToString();
        //        _dr["type"] = "TEMPLATE_" + _tmp_User;
        //        _dr["subtype"] = "INCIDENT_SUBJECT";
        //        _dr["mscrmid"] = "";
        //        _dr["mscrmparentid"] = "";
        //        _dr["description"] = "";
        //        if (_entity.SubjectId != null)
        //            _dr["zusatzinfo1"] = _entity.SubjectId.Id.ToString();
        //        else _dr["zusatzinfo1"] = "";
        //        _dr["zusatzinfo2"] = "";
        //        _tec.Rows.Add(_dr);
        //        _tec.AcceptChanges();
        //    }
        //    else
        //    {
        //        // Update
        //        if (_entity.SubjectId != null)
        //            _dr["zusatzinfo1"] = _entity.SubjectId.Id.ToString();
        //        _dr["zusatzinfo2"] = "";
        //        _tec.AcceptChanges();
        //    }

        //    MessageBox.Show("Die Anfragefelder Anfrageursprung, Betreff, Anfragetyp und Business Solutions wurden als Vorlage gespeichert!", "Hinweis", MessageBoxButton.OK, MessageBoxImage.Information);

        //}

        //private void btnAccept_Click(object sender, RoutedEventArgs e)
        //{
        //    // Accept
        //    try
        //    {
        //        if (cbAnfrageursprung.Text.Trim() == "" || cbAnfragetyp.Text.Trim() == "" || cbBusinessSolutions.Text.Trim() == "" || teBetreff.Text.Trim() == "" || cbAbrechnungsweise.Text.Trim() == "" || cbAnfragephase.Text.Trim() == "" || teMitarbeiter.Text.Trim() == "")
        //        {
        //            MessageBox.Show("Bitte füllen Sie mindestens alle ROT markierten Felder für die neue Anfrage aus!", "Wichtiger Hinweis", MessageBoxButton.OK, MessageBoxImage.Exclamation);
        //        }
        //        else
        //        {
        //            if (cbAnfrageursprung.Text != "") _entity.CaseOriginCode = new OptionSetValue(Convert.ToInt32(pl_caseorigincode[cbAnfrageursprung.Text])); else _entity.CaseOriginCode = null;
        //            if (cbAnfragetyp.Text != "") _entity.CaseTypeCode = new OptionSetValue(Convert.ToInt32(pl_casetypecode[cbAnfragetyp.Text])); else _entity.CaseTypeCode = null;
        //            if (cbBusinessSolutions.Text.Trim() != "") _entity.New_Geschaeftsbereich = new OptionSetValue(Convert.ToInt32(pl_new_geschaeftsbereich[cbBusinessSolutions.Text])); else _entity.New_Geschaeftsbereich = null;
        //            if (cbAbrechnungsweise.Text.Trim() != "") _entity.New_Abrechnungsweise = new OptionSetValue(Convert.ToInt32(pl_new_abrechnungsweise[cbAbrechnungsweise.Text])); else _entity.New_Abrechnungsweise = null;
        //            if (cbBestehendeVertraege.Text.Trim() != "") _entity.New_BestehendeVertraege_A24 = new OptionSetValue(Convert.ToInt32(pl_new_bestehendevertraege_a24[cbBestehendeVertraege.Text])); else _entity.New_BestehendeVertraege_A24 = null;
        //            if (cbRechnungDurch.Text.Trim() != "") _entity.New_Rechnungdurch = new OptionSetValue(Convert.ToInt32(pl_new_rechnungdurch[cbRechnungDurch.Text])); else _entity.New_Rechnungdurch = null;
        //            if (cbInterneVerrechnung.Text.Trim() != "") _entity.New_InterneVerrechnung = new OptionSetValue(Convert.ToInt32(pl_new_interneverrechnung[cbInterneVerrechnung.Text])); else _entity.New_InterneVerrechnung = null;
        //            if (cbAnfragephase.Text.Trim() != "") _entity.IncidentStageCode = new OptionSetValue(Convert.ToInt32(pl_incidentstagecode[cbAnfragephase.Text])); else _entity.IncidentStageCode = null;
        //            if (cbSchweregrad.Text.Trim() != "") _entity.SeverityCode = new OptionSetValue(Convert.ToInt32(pl_severitycode[cbSchweregrad.Text])); else _entity.SeverityCode = null;
        //            if (cbDringlichkeit.Text.Trim() != "") _entity.New_Dringlichkeit = new OptionSetValue(Convert.ToInt32(pl_new_dringlichkeit[cbDringlichkeit.Text])); else _entity.New_Dringlichkeit = null;
        //            if (cbGEMAKategorie.Text.Trim() != "") _entity.New_GEMAKategorie = new OptionSetValue(Convert.ToInt32(pl_new_gemakategorie[cbGEMAKategorie.Text])); else _entity.New_GEMAKategorie = null;

        //            // set statuscode
        //            if (cbStatus.Text.Trim() != "")
        //            {
        //                _entity.StatusCode = new OptionSetValue(Convert.ToInt32(pl_statuscode[cbStatus.Text]));
        //            }
        //            else _entity.StatusCode = null;

        //            // _entity.new_scheduledend 
        //            // _entity.New_scheduledend = Convert.ToDateTime(dpDeadline.Text);
        //            _entity.New_scheduledend = ConvertDateTime(Convert.ToDateTime(dpDeadline.Text));

        //            // _entity.followupby 
        //            //_entity.FollowupBy = Convert.ToDateTime(dpNachverfolgung.Text);
        //            _entity.FollowupBy = ConvertDateTime(Convert.ToDateTime(dpNachverfolgung.Text));

        //            _entity.Title = teTitel.Text;
        //            _entity.Description = teBeschreibung.Text;

        //            Selected_EntityGUID = _OrganizationService.Create(_entity);

        //            Selected_EntityType = "ANF";
        //            Selected_ID = _entity.Title;
                    
        //            this.Close();

        //        }
        //    }
        //    catch (SoapException se)
        //    {
        //        if (se.Detail != null) if (se.Detail.InnerText != null)
        //                MessageBox.Show(se.Detail.InnerText);
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
        //*/

        //private void btnClose_Click(object sender, RoutedEventArgs e)
        //{
        //    // Close
        //    Selected_EntityType = "";
        //    Selected_EntityGUID = Guid.Empty;
        //    this.Close();
        //}

        //private void btnSearchKunde_Click(object sender, RoutedEventArgs e)
        //{
        //    // Search Kunde
        //    // (Customer) _entity.customerid 

        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teKunde.Text.Trim(), "", false, Account.EntityLogicalName, "name", "accountid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();

        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teKunde.Text= _modal_dialog.Select_EntityName;
        //        Account _account = (Account)_OrganizationService.Retrieve(Account.EntityLogicalName, new Guid(_modal_dialog.Select_EntityID), new ColumnSet(true));
        //        if (_account != null)
        //        {
        //            String[] _titelStr = teTitel.Text.Split(':');

        //            if (_titelStr.Length > 1)
        //            {
        //                teTitel.Text = _account.New_kundenkuerzela24 + ": " + _titelStr[1].Trim();
        //            }
        //            else
        //            {
        //                teTitel.Text = _account.New_kundenkuerzela24 + ": " + _titelStr[0].Trim();
        //            }

        //            if (_account.New_BestendeVertraege_A24 != null)
        //            {
        //                foreach (KeyValuePair<String, String> _kvp in pl_new_bestehendevertraege_a24)
        //                {
        //                    if (_kvp.Value == _account.New_BestendeVertraege_A24.Value.ToString())
        //                    {
        //                        cbBestehendeVertraege.Text = _kvp.Key;
        //                        break;
        //                    }
        //                }
        //            }
        //        }

        //        _entity.CustomerId = new EntityReference(Account.EntityLogicalName, new Guid(_modal_dialog.Select_EntityID));
        //    }
        //}

        //private void btnSearchProjekt_Click(object sender, RoutedEventArgs e)
        //{
        //    // Search Projekt
        //    // (Lockup) _entity.new_projectid 

        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teProjekt.Text.Trim(), "", false, New_project.EntityLogicalName, "new_name", "new_projectid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();

        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teProjekt.Text = _modal_dialog.Select_EntityName;
        //        _entity.New_projectId = _modal_dialog._lookup; 
        //    }
        //}

        //private void btnSearchArbeitspaket_Click(object sender, RoutedEventArgs e)
        //{
        //    // Search Arbeitspaket
        //    // (Lockup) _entity.new_zugeordneteanfragenid 

        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teArbeitspaket.Text.Trim(), "", false, New_workpackage.EntityLogicalName, "new_name", "new_workpackageid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();

        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teArbeitspaket.Text = _modal_dialog.Select_EntityName;
        //        _entity.New_ZugeordneteAnfragenId = _modal_dialog._lookup;
        //    }

        //}

        //private void btnSearchBetreff_Click(object sender, RoutedEventArgs e)
        //{
        //    // Search Betreff
        //    // (LOckup) _entity.subjectid 

        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teBetreff.Text.Trim(), "", false, Subject.EntityLogicalName, "title", "subjectid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();

        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teBetreff.Text = _modal_dialog.Select_EntityName;
        //        _entity.SubjectId = _modal_dialog._lookup;
        //    }
        //}

        //private void btnSearchVerantwortlicherKontakt_Click(object sender, RoutedEventArgs e)
        //{
        //    // Saerch Kontakt
        //    // (Loockup) _entity.responsiblecontactid 

        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teVerantwortlicherKontakt.Text.Trim(), "", false, Contact.EntityLogicalName, "fullname", "contactid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();

        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teVerantwortlicherKontakt.Text = _modal_dialog.Select_EntityName;
        //        _entity.ResponsibleContactId = _modal_dialog._lookup;
        //    }
        //}

        //private void btnSearchInternGemeldetVon_Click(object sender, RoutedEventArgs e)
        //{
        //    // Search gemeldet von  (kontakt)
        //    // (Lockup) _entity.new_interngemeldetvonid
        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teInternGemeldetVon.Text.Trim(), "", false, SystemUser.EntityLogicalName, "fullname", "systemuserid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();
        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teInternGemeldetVon.Text = _modal_dialog.Select_EntityName;
        //        _entity.New_InterngemeldetvonId = _modal_dialog._lookup;
        //    }
        //}

        //private void btnSearchVerkaufschance_Click(object sender, RoutedEventArgs e)
        //{
        //    // Search Verkaufschance
        //    // (Lockup) _entity.new_opportunity_incident_linkid 
        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teVerkaufschance.Text.Trim(), "", false, Opportunity.EntityLogicalName, "name", "opportunityid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();
        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teVerkaufschance.Text = _modal_dialog.Select_EntityName;
        //        _entity.New_opportunity_incident_linkId = _modal_dialog._lookup;
        //    }
        //}

        //private void btnSearchMitarbeiter_Click(object sender, RoutedEventArgs e)
        //{
        //    // SearchMitarbeiter
        //    // (Lockup) _entity.ownerid

        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teMitarbeiter.Text.Trim(), "", false, SystemUser.EntityLogicalName, "fullname", "systemuserid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();
        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        _entity.OwnerId = new EntityReference(SystemUser.EntityLogicalName, new Guid(_modal_dialog.Select_EntityID));
        //        teMitarbeiter.Text = _modal_dialog.Select_EntityName;
        //    }
        //}

        //private void btnSearchAktuellerBearbeiter_Click(object sender, RoutedEventArgs e)
        //{
        //    // Search Aktueller Bearbeiter
        //    // (Lockup) _entity.new_aktuellerbearbeiterid 
        //    SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teAktuellerBearbeiter.Text.Trim(), "", false, SystemUser.EntityLogicalName, "fullname", "systemuserid", false);
        //    _modal_dialog.Topmost = false;
        //    _modal_dialog.Owner = this;
        //    _modal_dialog.ShowDialog();

        //    if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //    {
        //        teAktuellerBearbeiter.Text = _modal_dialog.Select_EntityName;
        //        _entity.New_AktuellerBearbeiterId = _modal_dialog._lookup;
        //    }
        //}

        //private void teKunde_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teKunde.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teKunde.Text.Trim(), "", true, Account.EntityLogicalName, "name", "accountid", true);

        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teKunde.Text = _modal_dialog.Select_EntityName;
        //            Account _account = (Account)_OrganizationService.Retrieve(Account.EntityLogicalName, new Guid(_modal_dialog.Select_EntityID), new ColumnSet(true));
        //            if (_account != null)
        //            {
        //                String[] _titelStr = teTitel.Text.Split(':');

        //                if (_titelStr.Length > 1)
        //                {
        //                    teTitel.Text = _account.New_kundenkuerzela24 + ": " + _titelStr[1].Trim();
        //                }
        //                else 
        //                {
        //                    teTitel.Text = _account.New_kundenkuerzela24 + ": " + _titelStr[0].Trim();
        //                }

        //                if (_account.New_BestendeVertraege_A24 != null)
        //                {
        //                    foreach (KeyValuePair<String, String> _kvp in pl_new_bestehendevertraege_a24)
        //                    {
        //                        if (_kvp.Value == _account.New_BestendeVertraege_A24.Value.ToString())
        //                        {
        //                            cbBestehendeVertraege.Text = _kvp.Key;
        //                            break;
        //                        }
        //                    }
        //                }
        //            }

        //            _entity.CustomerId = new EntityReference(Account.EntityLogicalName, new Guid(_modal_dialog.Select_EntityID));
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}

        //private void teProjekt_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teProjekt.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teProjekt.Text.Trim(), "", true, New_project.EntityLogicalName, "new_name", "new_projectid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teProjekt.Text = _modal_dialog.Select_EntityName;
        //            _entity.New_projectId = _modal_dialog._lookup;
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}

        //private void teArbeitspaket_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teArbeitspaket.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teArbeitspaket.Text.Trim(), "", true, New_workpackage.EntityLogicalName, "new_name", "new_workpackageid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teArbeitspaket.Text = _modal_dialog.Select_EntityName;
        //            _entity.New_ZugeordneteAnfragenId = _modal_dialog._lookup;
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}

        //private void teMitarbeiter_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teMitarbeiter.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teMitarbeiter.Text.Trim(), "", true, SystemUser.EntityLogicalName, "fullname", "systemuserid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            _entity.OwnerId = new EntityReference(SystemUser.EntityLogicalName, new Guid(_modal_dialog.Select_EntityID));
        //            teMitarbeiter.Text = _modal_dialog.Select_EntityName;
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}

        //private void teAktuellerBearbeiter_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teAktuellerBearbeiter.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teAktuellerBearbeiter.Text.Trim(), "", true, SystemUser.EntityLogicalName, "fullname", "systemuserid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teAktuellerBearbeiter.Text = _modal_dialog.Select_EntityName;
        //            _entity.New_AktuellerBearbeiterId = _modal_dialog._lookup;
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}

        //private void teBetreff_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teBetreff.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teBetreff.Text.Trim(), "", true, Subject.EntityLogicalName, "title", "subjectid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teBetreff.Text = _modal_dialog.Select_EntityName;
        //            _entity.SubjectId = _modal_dialog._lookup;
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}

        //private void teVerantwortlicherKontakt_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teVerantwortlicherKontakt.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teVerantwortlicherKontakt.Text.Trim(), "", true, Contact.EntityLogicalName, "fullname", "contactid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teVerantwortlicherKontakt.Text = _modal_dialog.Select_EntityName;
        //            _entity.ResponsibleContactId = _modal_dialog._lookup;
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}

        //private void teInternGemeldetVon_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teInternGemeldetVon.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teInternGemeldetVon.Text.Trim(), "", true, SystemUser.EntityLogicalName, "fullname", "systemuserid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teInternGemeldetVon.Text = _modal_dialog.Select_EntityName;
        //            _entity.New_InterngemeldetvonId = _modal_dialog._lookup;
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}

        //private void teVerkaufschance_LostFocus(object sender, RoutedEventArgs e)
        //{
        //    if (teVerkaufschance.Text.Trim() != "")
        //    {
        //        // Lost Focus
        //        this.Cursor = Cursors.Wait;

        //        SelectMSCRMEntity _modal_dialog = new SelectMSCRMEntity(_OrganizationService, teVerkaufschance.Text.Trim(), "", true, Opportunity.EntityLogicalName, "name", "opportunityid", true);
        //        try
        //        {
        //            _modal_dialog.Topmost = false;
        //            _modal_dialog.Owner = this;
        //            _modal_dialog.ShowDialog();
        //        }
        //        catch { }

        //        if (_modal_dialog.Select_EntityID != "" && _modal_dialog.Select_EntityName != "")
        //        {
        //            teVerkaufschance.Text = _modal_dialog.Select_EntityName;
        //            _entity.New_opportunity_incident_linkId = _modal_dialog._lookup;
        //        }

        //        this.Cursor = Cursors.Arrow;
        //    }
        //}
    }

}
