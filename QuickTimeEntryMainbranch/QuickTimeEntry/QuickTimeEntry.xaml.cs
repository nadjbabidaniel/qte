using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Timers;
using System.Xml;
using System.Xml.Linq;
using System.Reflection;
using System.IO;
using System.Windows.Threading;

using System.Threading;
using DevExpress.Xpf.Grid;
using System.Security.Principal;
using System.Diagnostics;
using System.Web.Services.Protocols;
using System.Net;
using System.Net.NetworkInformation;

using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Configuration;
using DevExpress.Xpf.Editors;
using System.Globalization;
using System.Windows.Markup;
using DevExpress.Xpf.Data;

using CCMA24Produktiv;

using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk;

using System.ServiceModel;

using A24.Xrm;

using McTools.Xrm.Connection;
using McTools.Xrm.Connection.WpfForms;
using Microsoft.Win32;
using System.Resources;
using System.ComponentModel;
using Infralution.Localization.Wpf;

namespace QuickTimeEntry
{

    /// <summary>
    /// Interaktionslogik für MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        Boolean TimerMode = false;
        System.Diagnostics.Stopwatch stopWatch = null;
        DispatcherTimer _timer;

        String _timer_info_begin = "";
        String _timer_info_end = "";
        String _timer_info_te_guid = "";

        String xml_path = "";

        private string sDefaultUser = "";
        private Guid gDefaultUser = Guid.Empty;

        DispatcherTimer _mscrm_worker;
        Thread _workerSyncThread = null;

        enum MSCRM_ConectionStates { NotConnected = 0, RequestConnection = 1, Connected = 2, InSyncProgress = 3, Error = -1 };
        String MSCRM_ConState_ErrInfo = "";

        MSCRM_ConectionStates MSCRM_ConState = MSCRM_ConectionStates.NotConnected;

        String _MSCRM_LinkTemplate_TE = "";
        String _MSCRM_LinkTemplate_AP = "";
        String _MSCRM_LinkTemplate_ANF = "";
        String _MSCRM_LinkTemplate_VC = "";
        //String _MSCRM_LinkTemplate_REL = "";
        String _MSCRM_LinkTemplate_KA = "";
        String _MSCRM_LinkTemplate_AIL = "";

        static String[] _BezugValues;/* = new string[]{
                                            EmptyString              // ""
                                           , WorkPackageString        // "AP"
                                           , OpportunityString         // "VC"
                                           , IncidentString            // "ANF"               
                                           , CampaignString           // "KA"
                                           , ActionItemString         // "AIL"
                                            };                                           */
        // String _Billing_PicklistData = "";

        DataTable _TimeEntryCache = null;

        // new_ccm_quicktimeentry _ccm_quicktimeentry = null;

        // XML File Paths
        String TimeEntrySource_Xml_Path = "";
        String TimeEntryCache_Xml_Path = "";
        String TimeEntryGridSettings = "";
        String LastSyncDateTime_Path = "";

        // Current System User
        String _User = "";

        Boolean _ActivateTimeEntryActive = false;

        String _Billing_PicklistData = "";
        String _Statuscode_PicklistData = "";
        String _User_ListData = "";

        DateTime _LastSyncDateTime_WORKPACKAGE = Convert.ToDateTime("01/01/1900 00:00:00");
        DateTime _LastSyncDateTime_INCIDENT = Convert.ToDateTime("01/01/1900 00:00:00");
        DateTime _LastSyncDateTime_OPPORTUNITY = Convert.ToDateTime("01/01/1900 00:00:00");
        DateTime _LastSyncDateTime_CAMPAIGN = Convert.ToDateTime("01/01/1900 00:00:00");
        //DateTime _LastSyncDateTime_NEW_VERSION = Convert.ToDateTime("01/01/1900 00:00:00");
        DateTime _LastSyncDateTime_ACTIONITEM = Convert.ToDateTime("01/01/1900 00:00:00");

        //Microsoft.Xrm.Sdk.Client.OrganizationServiceProxy _OrganizationService = null;
        Microsoft.Xrm.Sdk.IOrganizationService _OrganizationService = null;

        // Connection properties
        /// <summary>
        /// Form to bring up the connection management
        /// </summary>
        private FormHelper _fHelper;

        /// <summary>
        /// Connection Manager instance to manage connections to Mscrm
        /// </summary>
        private ConnectionManager _cManager;

        /// <summary>
        /// Metadata of the current connection to Mscrm
        /// </summary>
        private ConnectionDetail _currentConnectionDetail;

        /// <summary>
        ///  Flag to trigger the Full SyncProcess after connection has been established.
        /// </summary>
        private bool _bFullSyncCompleted = false;

        /// <summary>
        /// Flag to prevent multiple connection attempts
        /// </summary>
        private bool _bIsConnecting = false;


        private bool IsCurrentUserProcess(int ProcessID)
        {
            string stringSID = String.Empty;
            string process = ProcessUserInfo.ExGetProcessInfoByPID(ProcessID, out stringSID);
            return (String.Compare(stringSID, this._user.User.Value, true) == 0);
        }

        private bool AlreadyRunning()
        {
            Process current = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(current.ProcessName);

            foreach (Process process in processes)
            {
                if (process.Id != current.Id)
                {
                    // der Parser escaped in der Replace Methode einen \.
                    // Da Muss ein Dopelter \ hin.
                    if (Assembly.GetExecutingAssembly().Location.Replace("/", "\\") == current.MainModule.FileName)
                    {
                        if (IsCurrentUserProcess(process.Id)) return true;
                    }
                }
            }

            return false;
        }

        private WindowsIdentity _user = WindowsIdentity.GetCurrent();

        //Multilanguage DN
        private static ResourceManager rm = new ResourceManager("QuickTimeEntry.QuickTimeEntry", Assembly.GetExecutingAssembly());
        Boolean Entity_OK = false;
        Boolean Entity_OK2 = false;

        GridSummaryItem synced = new GridSummaryItem()
        {
            FieldName = "synced",
            SummaryType = DevExpress.Data.SummaryItemType.Count,
            DisplayFormat = "{0} " + rm.GetString("syncedLabel")
        };
        GridSummaryItem mscrmuser = new GridSummaryItem()
        {
            FieldName = "mscrmuser",
            SummaryType = DevExpress.Data.SummaryItemType.Count,
            DisplayFormat = "{0} " + rm.GetString("TimeItems")
        };

        //References dropdowns values read from config file -DN
        public const string EmptyString = "";
        public static string WorkPackageString = String.Empty;
        public static string OpportunityString = String.Empty;
        public static string IncidentString = String.Empty;
        public static string CampaignString = String.Empty;
        public static string ActionItemString = String.Empty;

        static List<String> _BezugValuesStartupLanguage;
        private static XmlDocument doc;
        public static void LoadRegardingEntities()
        {
            doc = new XmlDocument();
            string Xml_Path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            Xml_Path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Xml_Path), "XmlData\\" + "RelatedEntities.xml");
            doc.Load(Xml_Path);

            string lang = CultureManager.UICulture.TwoLetterISOLanguageName.ToLower();
            string languageMark = string.Empty;

            if (lang == "de") languageMark = "lang_1031";
            if (lang == "en") languageMark = "lang_1033";

            XmlNodeList xnList = doc.SelectNodes("/Settings/RelatedEntities/RelatedEntity/@name");
            foreach (XmlNode item in xnList)
            {
                if (item.Value.Equals("workpackage")) WorkPackageString = doc.DocumentElement.SelectSingleNode("/Settings/RelatedEntities/RelatedEntity[@name='" + "workpackage" + "']/Shortcut").Attributes[languageMark].Value;
                if (item.Value.Equals("opportunity")) OpportunityString = doc.DocumentElement.SelectSingleNode("/Settings/RelatedEntities/RelatedEntity[@name='" + "opportunity" + "']/Shortcut").Attributes[languageMark].Value;
                if (item.Value.Equals("incident")) IncidentString = doc.DocumentElement.SelectSingleNode("/Settings/RelatedEntities/RelatedEntity[@name='" + "incident" + "']/Shortcut").Attributes[languageMark].Value;
                if (item.Value.Equals("campaign")) CampaignString = doc.DocumentElement.SelectSingleNode("/Settings/RelatedEntities/RelatedEntity[@name='" + "campaign" + "']/Shortcut").Attributes[languageMark].Value;
                if (item.Value.Equals("actionitem")) ActionItemString = doc.DocumentElement.SelectSingleNode("/Settings/RelatedEntities/RelatedEntity[@name='" + "actionitem" + "']/Shortcut").Attributes[languageMark].Value;
            }

            if (WorkPackageString.Equals(String.Empty)) WorkPackageString = "AP";
            if (OpportunityString.Equals(String.Empty)) OpportunityString = "VS";
            if (IncidentString.Equals(String.Empty)) IncidentString = "ANF";
            if (CampaignString.Equals(String.Empty)) CampaignString = "KA";
            if (ActionItemString.Equals(String.Empty)) ActionItemString = "AIL";

            _BezugValues = new string[]{
                                            EmptyString              // ""
                                           , WorkPackageString        // "AP"
                                           , OpportunityString         // "VC"
                                           , IncidentString            // "ANF"               
                                           , CampaignString           // "KA"
                                           , ActionItemString         // "AIL"
                                            };
        }

        public MainWindow()
        {
            try
            {
                InitializeComponent();

                // McTools Connection Control initialisieren
                ManageConnectionControl();

                CultureManager.UICulture = CultureManager.UICulture = new CultureInfo("en");

                if (AlreadyRunning())
                {
                    MessageBox.Show(rm.GetString("AlreadyRunning"), "Info", MessageBoxButton.OK, MessageBoxImage.Error);
                    Application.Current.Shutdown();
                }

                System.Net.ServicePointManager.ServerCertificateValidationCallback +=
                delegate (object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors sslError)
                {
                    bool validationResult = true;
                    return validationResult;
                };

                // remember startup language reference label values
                LoadRegardingEntities();
                _BezugValuesStartupLanguage = new List<String>(_BezugValues);

                // set the link templates from the Configuration
                SetLinkTemplates();

                // set the current user from the OS
                _User = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString();

                // set the filepaths from the current path of the executing assembly
                SetFilePaths();

                // Check, if Files exist
                CreateXmlFiles();

                // read Time entries from Xml
                SetupTimeEntryDataSource();

                // read Cache
                SetupCache();

                // set main windows properties
                InitializeMainLayout();

                // restore the Layout from XML and fills the Summary Fields
                InitializeTimeEntryGrid();
                SetDateForAppropriateLanguage();

                _timer = new DispatcherTimer();
                _timer.Interval = TimeSpan.FromMilliseconds(1000);
                _timer.Tick += new EventHandler(t_Elapsed);

                // initializes the 5 TimeEntryTabs
                InitializeTimeEntryTabs();

                // enable or disable buttons
                InitializeButtonEnabling();

                // read the last syncdates of the entities
                ReadLastSyncDates();

                // Mscrm Worker Thread to sync non-synced entries to mscrm and check the Connection State
                _mscrm_worker = new DispatcherTimer();
                _mscrm_worker.Interval = TimeSpan.FromMilliseconds(1000);
                _mscrm_worker.Tick += new EventHandler(mscrm_worker_Elapsed);
                _mscrm_worker.Start();

                this.Show();

                // A24 JNE 24.9.2015 this triggers the generation of the connection
                MSCRM_ConState = MSCRM_ConectionStates.RequestConnection;
            }
            catch (Exception ex)
            {
                FileLog("MainWindow", "Error: " + ex.Message);
                this.Close();
            }
        }

        /// <summary>
        /// Reads the last syncdates from the LastSyncDateTime_Path
        /// </summary>
        private void ReadLastSyncDates()
        {

            if (File.Exists(LastSyncDateTime_Path))
            {
                using (StreamReader reader = new StreamReader(LastSyncDateTime_Path))
                {
                    String _tmpStr = "";

                    _tmpStr = reader.ReadLine(); _LastSyncDateTime_WORKPACKAGE = Convert.ToDateTime(_tmpStr);
                    _tmpStr = reader.ReadLine(); _LastSyncDateTime_INCIDENT = Convert.ToDateTime(_tmpStr);
                    _tmpStr = reader.ReadLine(); _LastSyncDateTime_OPPORTUNITY = Convert.ToDateTime(_tmpStr);
                    _tmpStr = reader.ReadLine(); _LastSyncDateTime_CAMPAIGN = Convert.ToDateTime(_tmpStr);
                    //_tmpStr = reader.ReadLine(); _LastSyncDateTime_NEW_VERSION = Convert.ToDateTime(_tmpStr);
                    _tmpStr = reader.ReadLine(); _LastSyncDateTime_ACTIONITEM = Convert.ToDateTime(_tmpStr);

                    reader.Close();
                }
            }
        }

        /// <summary>
        /// Sets the basic button handling
        /// </summary>
        private void InitializeButtonEnabling()
        {
            bSyncMSCRMActMonth.IsEnabled = false;
            bSyncMSCRMPreMonth.IsEnabled = false;

            bSyncMSCRMEntityRefresh.IsEnabled = false;

            bGoOnline.IsEnabled = false;
            bGoOffline.IsEnabled = true;
        }

        /// <summary>
        /// Sets Windows position, width, fitmode, startuplocation and title
        /// </summary>
        private void InitializeMainLayout()
        {
            view.BestFitMode = DevExpress.Xpf.Core.BestFitMode.Smart;

            this.MinWidth = 800;
            this.MinHeight = 600;

            this.Width = 1300;
            // A24 JNE 24.9.2015 removed from 900
            this.Height = 750;

            this.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;

            this.Title = "Quick Time Entry for Dynamics CRM [v2.00]";
        }

        /// <summary>
        /// Setup the 5 TimeEntry Tabs and their Datetime values
        /// </summary>
        private void InitializeTimeEntryTabs()
        {

            TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Collapsed;

            tabTE_1.Tag = "";
            tabTE_2.Tag = "";
            tabTE_3.Tag = "";
            tabTE_4.Tag = "";
            tabTE_5.Tag = "";

            deDatum_1.MaskType = DevExpress.Xpf.Editors.MaskType.DateTime;
            deDatum_1.MaskCulture = CultureInfo.CurrentCulture;
            deDatum_1.MaskUseAsDisplayFormat = true;
            deDatum_1.Mask = (CultureInfo.GetCultureInfo(CultureInfo.CurrentCulture.ToString()).DateTimeFormat.ShortDatePattern);
            //deDatum_1.Mask = "dd.MM.yyyy";

            deDatum_2.MaskType = DevExpress.Xpf.Editors.MaskType.DateTime;
            deDatum_2.MaskCulture = CultureInfo.CurrentCulture;
            deDatum_2.MaskUseAsDisplayFormat = true;
            deDatum_2.Mask = (CultureInfo.GetCultureInfo(CultureInfo.CurrentCulture.ToString()).DateTimeFormat.ShortDatePattern);

            deDatum_3.MaskType = DevExpress.Xpf.Editors.MaskType.DateTime;
            deDatum_3.MaskCulture = CultureInfo.CurrentCulture;
            deDatum_3.MaskUseAsDisplayFormat = true;
            deDatum_3.Mask = (CultureInfo.GetCultureInfo(CultureInfo.CurrentCulture.ToString()).DateTimeFormat.ShortDatePattern);

            deDatum_4.MaskType = DevExpress.Xpf.Editors.MaskType.DateTime;
            deDatum_4.MaskCulture = CultureInfo.CurrentCulture;
            deDatum_4.MaskUseAsDisplayFormat = true;
            deDatum_4.Mask = (CultureInfo.GetCultureInfo(CultureInfo.CurrentCulture.ToString()).DateTimeFormat.ShortDatePattern);

            deDatum_5.MaskType = DevExpress.Xpf.Editors.MaskType.DateTime;
            deDatum_5.MaskCulture = CultureInfo.CurrentCulture;
            deDatum_5.MaskUseAsDisplayFormat = true;
            deDatum_5.Mask = (CultureInfo.GetCultureInfo(CultureInfo.CurrentCulture.ToString()).DateTimeFormat.ShortDatePattern);
        }

        /// <summary>
        /// Restores the Layout from XML and fills the Summary Fields
        /// </summary>
        private void InitializeTimeEntryGrid()
        {

            // Restore Grid Layout
            try
            {
                TimeEntryGrid.RestoreLayoutFromXml(TimeEntryGridSettings);
            }
            catch
            {
            }

            TimeEntryGrid.TotalSummary.Clear();
            TimeEntryGrid.GroupSummary.Clear();

            // Group Summary

            TimeEntryGrid.GroupSummary.Add(mscrmuser);

            TimeEntryGrid.GroupSummary.Add(new GridSummaryItem()
            {
                FieldName = "dauer",
                SummaryType = DevExpress.Data.SummaryItemType.Custom,
                DisplayFormat = "{0}"
            });

            TimeEntryGrid.GroupSummary.Add(new GridSummaryItem()
            {
                FieldName = "zusatzzeit",
                SummaryType = DevExpress.Data.SummaryItemType.Custom,
                DisplayFormat = "{0}"
            });

            TimeEntryGrid.GroupSummary.Add(new GridSummaryItem()
            {
                FieldName = "gesamtdauer",
                SummaryType = DevExpress.Data.SummaryItemType.Custom,
                DisplayFormat = "{0}"
            });


            // Total Summary

            TimeEntryGrid.TotalSummary.Add(synced);

            TimeEntryGrid.TotalSummary.Add(new GridSummaryItem()
            {
                FieldName = "dauer",
                SummaryType = DevExpress.Data.SummaryItemType.Custom,
                DisplayFormat = "{0}"
            });

            TimeEntryGrid.TotalSummary.Add(new GridSummaryItem()
            {
                FieldName = "zusatzzeit",
                SummaryType = DevExpress.Data.SummaryItemType.Custom,
                DisplayFormat = "{0}"
            });

            TimeEntryGrid.TotalSummary.Add(new GridSummaryItem()
            {
                FieldName = "gesamtdauer",
                SummaryType = DevExpress.Data.SummaryItemType.Custom,
                DisplayFormat = "{0}"
            });

            TimeEntryGrid.ExpandAllGroups();
            TimeEntryGrid.UpdateTotalSummary();
        }

        /// <summary>
        /// Read the Cache from the Time Entry Cache XML File
        /// </summary>
        private void SetupCache()
        {
            var ds = new DataSet();
            ds.ReadXml(TimeEntryCache_Xml_Path);

            var keys = new DataColumn[3];
            keys[0] = ds.Tables[0].Columns["id"];
            keys[1] = ds.Tables[0].Columns["type"];
            keys[1].Unique = false;
            keys[2] = ds.Tables[0].Columns["mscrmid"];

            ds.Tables[0].PrimaryKey = keys;

            _TimeEntryCache = ds.Tables[0];

            staticItemLine.Content = "";
        }

        /// <summary>
        /// Setup a new Dataset to contain the data for the TimeEntryGrid.
        /// Reads the Data from the TimeEntrySource Xml.
        /// </summary>
        private void SetupTimeEntryDataSource()
        {
            DataSet ds = new DataSet();
            ds.ReadXml(TimeEntrySource_Xml_Path);

            DataColumn[] keys = new DataColumn[1];
            keys[0] = ds.Tables[0].Columns["id"];

            ds.Tables[0].PrimaryKey = keys;

            // Clean empty row
            try
            {
                DataRow _dr = ds.Tables[0].Rows.Find("00000000-0000-0000-0000-0000000");
                ds.Tables[0].Rows.Remove(_dr);
            }
            catch
            {
            }

            // Add new XML Fields

            if (!ds.Tables[0].Columns.Contains("taetigkeiturl")) ds.Tables[0].Columns.Add("taetigkeiturl", Type.GetType("System.String"));
            if (!ds.Tables[0].Columns.Contains("mscrmmodifiedon")) ds.Tables[0].Columns.Add("mscrmmodifiedon", Type.GetType("System.String"));

            if (!ds.Tables[0].Columns.Contains("tekz2")) ds.Tables[0].Columns.Add("tekz2", Type.GetType("System.String"));
            if (!ds.Tables[0].Columns.Contains("mscrmbezug2")) ds.Tables[0].Columns.Add("mscrmbezug2", Type.GetType("System.String"));
            if (!ds.Tables[0].Columns.Contains("mscrmbezugid2")) ds.Tables[0].Columns.Add("mscrmbezugid2", Type.GetType("System.String"));
            if (!ds.Tables[0].Columns.Contains("mscrmlinkurl2")) ds.Tables[0].Columns.Add("mscrmlinkurl2", Type.GetType("System.String"));
            if (!ds.Tables[0].Columns.Contains("mscrmlinktext2")) ds.Tables[0].Columns.Add("mscrmlinktext2", Type.GetType("System.String"));
            if (!ds.Tables[0].Columns.Contains("kw")) ds.Tables[0].Columns.Add("kw", Type.GetType("System.String"));

            if (ds.Tables[0].Columns.Contains("marked")) ds.Tables[0].Columns.Remove("marked");
            ds.Tables[0].Columns.Add("marked", Type.GetType("System.Boolean"));

            // Update MSCRM Entity Link (taetigkeiturl)
            foreach (DataRow _dr in ds.Tables[0].Rows)
            {
                String _linkurl = _MSCRM_LinkTemplate_TE;

                _linkurl = _linkurl.Replace("{MSCRMID}", _dr["mscrmentityid"].ToString());
                _linkurl = _linkurl.Replace("{MSCRMPARENTID}", _dr["mscrmbezugid"].ToString());

                _dr["taetigkeiturl"] = _linkurl;
                _dr["marked"] = false;
            }

            TimeEntryGrid.ItemsSource = ds.Tables[0];

            TimeEntryGrid.IsFilterEnabled = true;
            view.ShowAutoFilterRow = true;

            foreach (DevExpress.Xpf.Grid.GridColumn _gc in TimeEntryGrid.Columns)
            {
                _gc.AllowAutoFilter = true;
                _gc.AutoFilterCondition = DevExpress.Xpf.Grid.AutoFilterCondition.Contains;
                if (_gc.Name != "gc_marked") _gc.AllowEditing = DevExpress.Utils.DefaultBoolean.False;
            }

            TimeEntryGrid.RefreshData();

            if (!System.IO.File.Exists(TimeEntryGridSettings))
            {
                TimeEntryGrid.FilterString = "[synced] != '*new*'";

                TimeEntryGrid.GroupBy(TimeEntryGrid.Columns["datum"], DevExpress.Data.ColumnSortOrder.Descending);

                // TimeEntryGrid.SortBy(TimeEntryGrid.Columns["datum"], DevExpress.Data.ColumnSortOrder.Descending);
                //GridSortInfo _gsi = new GridSortInfo("datum", System.ComponentModel.ListSortDirection.Descending);
                //TimeEntryGrid.SortInfo.Add(_gsi);
            }
        }

        /// <summary>
        /// Check for TimeEntrySource and TimeEntryCache XML files and create them if they don't exist
        /// </summary>
        private void CreateXmlFiles()
        {
            if (!System.IO.File.Exists(TimeEntrySource_Xml_Path))
            {
                TextWriter tw = new StreamWriter(TimeEntrySource_Xml_Path);
                tw.WriteLine("<?xml version=\"1.0\" standalone=\"yes\"?>");
                tw.WriteLine("<TimeEntries>");
                tw.WriteLine("<item>");
                tw.WriteLine("<id>00000000-0000-0000-0000-0000000</id>");
                tw.WriteLine("<tekz />");
                tw.WriteLine("<tekz2 />");
                tw.WriteLine("<datum />");
                tw.WriteLine("<beginn />");
                tw.WriteLine("<ende />");
                tw.WriteLine("<dauer />");
                tw.WriteLine("<zusatzzeit />");
                tw.WriteLine("<gesamtdauer />");
                tw.WriteLine("<taetigkeit />");
                tw.WriteLine("<taetigkeiturl />");
                tw.WriteLine("<synced />");
                tw.WriteLine("<mscrmentityid />");
                tw.WriteLine("<timemeasurements />");
                tw.WriteLine("<beschreibung />");
                tw.WriteLine("<anmerkung />");
                tw.WriteLine("<mscrmstatus />");
                tw.WriteLine("<mscrmstatusid />");
                tw.WriteLine("<fahrtkosten />");
                tw.WriteLine("<mscrmbezug />");
                tw.WriteLine("<mscrmbezugid />");
                tw.WriteLine("<mscrmbezug2 />");
                tw.WriteLine("<mscrmbezugid2 />");
                tw.WriteLine("<mscrmuser />");
                tw.WriteLine("<mscrmuserid />");
                tw.WriteLine("<mscrmbillingtype />");
                tw.WriteLine("<mscrmbillingtypeid />");
                tw.WriteLine("<mscrmmodifiedon />");
                tw.WriteLine("<lasterror />");
                tw.WriteLine("<mscrmlinkurl />");
                tw.WriteLine("<mscrmlinktext />");
                tw.WriteLine("<mscrmlinkurl2 />");
                tw.WriteLine("<mscrmlinktext2 />");
                tw.WriteLine("<syncfrommscrm />");
                tw.WriteLine("</item>");
                tw.WriteLine("</TimeEntries>");

                // close the stream
                tw.Close();
            }

            if (!System.IO.File.Exists(TimeEntryCache_Xml_Path))
            {
                // create empty file
                TextWriter tw = new StreamWriter(TimeEntryCache_Xml_Path);
                tw.WriteLine("<?xml version=\"1.0\" standalone=\"yes\"?>");
                tw.WriteLine("<CacheEntries>");
                tw.WriteLine("<item>");
                tw.WriteLine("<id>00000000-0000-0000-0000-0000000</id>");
                tw.WriteLine("<type />");
                tw.WriteLine("<subtype />");
                tw.WriteLine("<mscrmid />");
                tw.WriteLine("<mscrmparentid />");
                tw.WriteLine("<description />");
                tw.WriteLine("<zusatzinfo1 />");
                tw.WriteLine("<zusatzinfo2 />");
                tw.WriteLine("</item>");

                tw.WriteLine("</CacheEntries>");

                // close the stream
                tw.Close();
            }
        }

        private void SetFilePaths()
        {
            String _tmp_User = _User;
            _tmp_User = _tmp_User.Replace("\\", "_");
            _tmp_User = _tmp_User.Replace(".", "_");

            TimeEntrySource_Xml_Path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            TimeEntrySource_Xml_Path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(TimeEntrySource_Xml_Path), "XmlData\\" + _tmp_User + "_TimeEntrySource.xml");

            TimeEntryCache_Xml_Path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            TimeEntryCache_Xml_Path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(TimeEntryCache_Xml_Path), "XmlData\\TimeEntryCache.xml");

            TimeEntryGridSettings = System.Reflection.Assembly.GetExecutingAssembly().Location;
            TimeEntryGridSettings = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(TimeEntryGridSettings), "XmlData\\" + _tmp_User + "_TimeEntryGridSettings.xml");

            LastSyncDateTime_Path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            LastSyncDateTime_Path = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(LastSyncDateTime_Path), "XmlData\\LastSyncDateTime.info");
        }

        /// <summary>
        /// Sets the Link Templates from the configuration.
        /// </summary>
        private void SetLinkTemplates()
        {

            try
            {
                _MSCRM_LinkTemplate_TE = System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_TE"].ToString();
                _MSCRM_LinkTemplate_AP = System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_AP"].ToString();
                _MSCRM_LinkTemplate_ANF = System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_ANF"].ToString();
                _MSCRM_LinkTemplate_VC = System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_VC"].ToString();
                //_MSCRM_LinkTemplate_REL = System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_REL"].ToString();
                _MSCRM_LinkTemplate_KA = System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_KA"].ToString();
                _MSCRM_LinkTemplate_AIL = System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_AIL"].ToString();
            }
            catch
            {
                String ErrInfoStr = "";
                if (System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_TE"] == null) ErrInfoStr += "MSCRM.LinkTemplate_TE";
                if (System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_AP"] == null) { if (ErrInfoStr != "") ErrInfoStr += ", "; ErrInfoStr += "MSCRM.LinkTemplate_AP"; }
                if (System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_ANF"] == null) { if (ErrInfoStr != "") ErrInfoStr += ", "; ErrInfoStr += "MSCRM.LinkTemplate_ANF"; }
                if (System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_VC"] == null) { if (ErrInfoStr != "") ErrInfoStr += ", "; ErrInfoStr += "MSCRM.LinkTemplate_VC"; }
                //if (System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_REL"] == null) { if (ErrInfoStr != "") ErrInfoStr += ", ";  ErrInfoStr += "MSCRM.LinkTemplate_REL"; }
                if (System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_KA"] == null) { if (ErrInfoStr != "") ErrInfoStr += ", "; ErrInfoStr += "MSCRM.LinkTemplate_KA"; }
                if (System.Configuration.ConfigurationManager.AppSettings["MSCRM.LinkTemplate_AIL"] == null) { if (ErrInfoStr != "") ErrInfoStr += ", "; ErrInfoStr += "MSCRM.LinkTemplate_AIL"; }

                MessageBox.Show(rm.GetString("ErrInfoStr") + ErrInfoStr, rm.GetString("ErrInfoStrTitle"), MessageBoxButton.OK, MessageBoxImage.Error);
                Application.Current.Shutdown();
            }
        }

        private DataRow GetSourceItemByMSCRMID(String itemGUID)
        {
            DataRow _rV = null;

            DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;

            DataRow[] _result = _dt.Select("[mscrmentityid]='" + itemGUID + "'");
            if (_result.Length == 1)
            {
                _rV = _result[0];
            }

            return _rV;
        }

        private DataRow GetCacheItemByMSCRMID(String itemTYPE, String itemGUID)
        {
            DataRow _rV = null;

            DataRow[] _result = _TimeEntryCache.Select("[type]='" + itemTYPE + "' AND [mscrmid]='" + itemGUID + "'");
            if (_result.Length == 1)
            {
                _rV = _result[0];
            }

            return _rV;
        }

        private DataRow GetCacheItemByDESCRIPTION(String itemTYPE, String itemDescription)
        {
            DataRow _rV = null;

            DataRow[] _result = _TimeEntryCache.Select("[type]='" + itemTYPE + "'");
            if (_result.Length > 0)
            {
                foreach (DataRow _dr in _result)
                {
                    if (_dr["description"].ToString().Trim().ToLower() == itemDescription.Trim().ToLower())
                    { _rV = _dr; break; }
                }
            }

            return _rV;
        }

        public DateTime LastDayOfMonthFromDateTime(DateTime dateTime)
        {
            DateTime firstDayOfTheMonth = new DateTime(dateTime.Year, dateTime.Month, 1);
            return firstDayOfTheMonth.AddMonths(1).AddDays(-1);
        }

        public void LoadTimeEntitiesFromMSCRM(IOrganizationService _Service, Boolean PreMonth)
        {
            try
            {
                FileLog("LoadTimeEntitiesFromMSCRM:begin", "User=" + sDefaultUser);

                Cursor _old_cursor = this.Cursor;
                this.Cursor = Cursors.Wait;

                FilterExpression filterExpression = new FilterExpression();

                ConditionExpression condition = new ConditionExpression("createdby", ConditionOperator.Equal, gDefaultUser.ToString());

                Int32 _curr_year = DateTime.Now.Year;
                Int32 _curr_month = DateTime.Now.Month;

                if (PreMonth)
                {
                    // Vormonat
                    _curr_month--;
                    if (_curr_month == 0)
                    {
                        _curr_month = 12;
                        _curr_year--;
                    }
                }

                DateTime _newdayofmonth = new DateTime(_curr_year, _curr_month, 1);
                DateTime _lastdayofmonth = LastDayOfMonthFromDateTime(_newdayofmonth);

                ConditionExpression condition_date_from = new ConditionExpression("a24_date_dat", ConditionOperator.GreaterEqual, Convert.ToDateTime("01." + _curr_month.ToString("00") + "." + _curr_year.ToString() + " 00:00:00"));
                //var temp = _lastdayofmonth.Day.ToString("00") + "." + _curr_month.ToString("00") + "." + _curr_year.ToString();
                ConditionExpression condition_date_to = new ConditionExpression("a24_date_dat", ConditionOperator.LessEqual, Convert.ToDateTime(_lastdayofmonth.ToShortDateString() + " 23:59:59", CultureInfo.CurrentCulture.DateTimeFormat));

                filterExpression.FilterOperator = LogicalOperator.And;
                filterExpression.AddCondition(condition);
                filterExpression.AddCondition(condition_date_from);
                filterExpression.AddCondition(condition_date_to);

                QueryExpression query = new QueryExpression()
                {
                    EntityName = a24_timeentry.EntityLogicalName,
                    ColumnSet = new ColumnSet(true), // new String[] { sNameAttribute, sIDAttribute, "a24_workpackage_ref", "new_opportunityid", "new_incidentid", "new_versionid", Constants.FieldNames.TimeEntry.CampaignReference, "new_executeuserid" }),
                    Criteria = filterExpression,
                    Distinct = false,
                    NoLock = true
                };

                RetrieveMultipleRequest multipleRequest = new RetrieveMultipleRequest();
                multipleRequest.Query = query;

                RetrieveMultipleResponse response = (RetrieveMultipleResponse)_Service.Execute(multipleRequest);

                String sDesc = "";
                Guid gData = Guid.Empty;
                String Sync_Entity_Type = "";
                String Sync_Entity_Type2 = "";
                Int32 _c = 0;

                staticItemInfo.Content = "";

                foreach (a24_timeentry _new_timeentry in response.EntityCollection.Entities)
                {
                    sDesc = "";
                    gData = Guid.Empty;

                    Guid _entity_id = Guid.Empty;
                    Guid _entity_id_2 = Guid.Empty;

                    Sync_Entity_Type = "";
                    Sync_Entity_Type2 = "";

                    EntityReference _new_executeuserid = _new_timeentry.a24_executed_by_ref;
                    sDesc = _new_timeentry.a24_matchcode_str;
                    gData = _new_timeentry.a24_timeentryId.Value;

                    if (string.IsNullOrEmpty(sDesc) == false && gData != Guid.Empty)
                    {
                        // ColumnSet _all_cols = new ColumnSet(true);
                        // a24_timeentry _new_timeentry_2 = (a24_timeentry)_Service.Retrieve(a24_timeentry.EntityLogicalName, gData, new ColumnSet(true));



                        if (_new_timeentry.a24_working_package_ref != null)
                        {
                            EntityReference _lp = (EntityReference)_new_timeentry.Attributes["a24_working_package_ref"];

                            if (_entity_id == Guid.Empty)
                            {
                                _entity_id = _lp.Id;
                                Sync_Entity_Type = WorkPackageString;
                            }
                            else if (_entity_id_2 == Guid.Empty)
                            {
                                _entity_id_2 = _lp.Id;
                                Sync_Entity_Type2 = WorkPackageString;
                            }
                        }
                        if (_new_timeentry.a24_opportunity_ref != null)
                        {
                            EntityReference _lp = (EntityReference)_new_timeentry.Attributes[Constants.FieldNames.TimeEntry.OpportunityReference];

                            if (_entity_id == Guid.Empty)
                            {
                                _entity_id = _lp.Id;
                                Sync_Entity_Type = OpportunityString;
                            }
                            else if (_entity_id_2 == Guid.Empty)
                            {
                                _entity_id_2 = _lp.Id;
                                Sync_Entity_Type2 = OpportunityString;
                            }
                        }
                        if (_new_timeentry.a24_incident_ref != null)
                        {
                            EntityReference _lp = (EntityReference)_new_timeentry.Attributes[Constants.FieldNames.TimeEntry.IncidentReference];

                            if (_entity_id == Guid.Empty)
                            {
                                _entity_id = _lp.Id;
                                Sync_Entity_Type = IncidentString;
                            }
                            else if (_entity_id_2 == Guid.Empty)
                            {
                                _entity_id_2 = _lp.Id;
                                Sync_Entity_Type2 = IncidentString;
                            }
                        }
                        //if (_new_timeentry.New_VersionId != null)
                        //{
                        //    EntityReference _lp = (EntityReference)_new_timeentry.Attributes["new_versionid"];

                        //    if (_entity_id == Guid.Empty)
                        //    {
                        //        _entity_id = _lp.Id;
                        //        Sync_Entity_Type = "REL";
                        //    }
                        //    else if (_entity_id_2 == Guid.Empty)
                        //    {
                        //        _entity_id_2 = _lp.Id;
                        //        Sync_Entity_Type2 = "REL";
                        //    }
                        //}
                        if (_new_timeentry.a24_campaign_ref != null)
                        {
                            EntityReference _lp = (EntityReference)_new_timeentry.Attributes[Constants.FieldNames.TimeEntry.CampaignReference];

                            if (_entity_id == Guid.Empty)
                            {
                                _entity_id = _lp.Id;
                                Sync_Entity_Type = CampaignString;
                            }
                            else if (_entity_id_2 == Guid.Empty)
                            {
                                _entity_id_2 = _lp.Id;
                                Sync_Entity_Type2 = CampaignString;
                            }
                        }
                        if (_new_timeentry.a24_action_item_ref != null)
                        {
                            EntityReference _lp = (EntityReference)_new_timeentry.Attributes[Constants.FieldNames.TimeEntry.ActionItemReference];

                            if (_entity_id == Guid.Empty)
                            {
                                _entity_id = _lp.Id;
                                Sync_Entity_Type = ActionItemString;
                            }
                            else if (_entity_id_2 == Guid.Empty)
                            {
                                _entity_id_2 = _lp.Id;
                                Sync_Entity_Type2 = ActionItemString;
                            }
                        }

                        String _enitity_name = "";
                        String _parent_entity_id = "";
                        String _enitity_name_2 = "";
                        String _parent_entity_id_2 = "";

                        // DateTime _new_date = _new_timeentry.New_date.Value;

                        Boolean SyncItem = true;

                        if (SyncItem)
                        {
                            _c++;

                            // BEZUG 1

                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _entity_id, new ColumnSet(new String[] { Constants.FieldNames.WorkPackage.CustomerReference, Constants.FieldNames.TimeEntry.Name }));
                                _parent_entity_id = wp.a24_customer_ref.Id.ToString();
                                _enitity_name = wp.a24_matchcode_str;
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _entity_id, new ColumnSet(new String[] { "customerid", "title" }));
                                _parent_entity_id = _in.CustomerId.Id.ToString();
                                _enitity_name = _in.Title;
                            }
                            else if (Sync_Entity_Type == OpportunityString)
                            {
                                Opportunity _op = (Opportunity)_Service.Retrieve(Opportunity.EntityLogicalName, _entity_id, new ColumnSet(new String[] { "customerid", "name" }));
                                _parent_entity_id = _op.CustomerId.Id.ToString();
                                _enitity_name = _op.Name;
                            }
                            //else if (Sync_Entity_Type == "REL")
                            //{
                            //    New_Version _nv = (New_Version)_Service.Retrieve(New_Version.EntityLogicalName, _entity_id, new ColumnSet(new String[] { "new_versionsnummer" }));
                            //    _parent_entity_id = "";
                            //    _enitity_name = _nv.New_versionsnummer;
                            //}
                            else if (Sync_Entity_Type == CampaignString)
                            {
                                Campaign _camp = (Campaign)_Service.Retrieve(Campaign.EntityLogicalName, _entity_id, new ColumnSet(new String[] { "name" }));
                                _parent_entity_id = "";
                                _enitity_name = _camp.Name;
                            }

                            // BEZUG 2

                            if (Sync_Entity_Type2 == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _entity_id_2, new ColumnSet(new String[] { Constants.FieldNames.TimeEntry.AccountReference, Constants.FieldNames.TimeEntry.Name }));
                                _parent_entity_id_2 = wp.a24_customer_ref.Id.ToString();
                                _enitity_name_2 = wp.a24_matchcode_str;
                            }
                            else if (Sync_Entity_Type2 == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _entity_id_2, new ColumnSet(new String[] { "customerid", "title" }));
                                _parent_entity_id_2 = _in.CustomerId.Id.ToString();
                                _enitity_name_2 = _in.Title;
                            }
                            else if (Sync_Entity_Type2 == OpportunityString)
                            {
                                Opportunity _op = (Opportunity)_Service.Retrieve(Opportunity.EntityLogicalName, _entity_id_2, new ColumnSet(new String[] { "customerid", "name" }));
                                _parent_entity_id_2 = _op.CustomerId.Id.ToString();
                                _enitity_name_2 = _op.Name;
                            }
                            //else if (Sync_Entity_Type2 == "REL")
                            //{
                            //    New_Version _nv = (New_Version)_Service.Retrieve(New_Version.EntityLogicalName, _entity_id_2, new ColumnSet(new String[] { "new_versionsnummer" }));
                            //    _parent_entity_id_2 = "";
                            //    _enitity_name_2 = _nv.New_versionsnummer;
                            //}
                            else if (Sync_Entity_Type2 == CampaignString)
                            {
                                Campaign _camp = (Campaign)_Service.Retrieve(Campaign.EntityLogicalName, _entity_id_2, new ColumnSet(new String[] { "name" }));
                                _parent_entity_id_2 = "";
                                _enitity_name_2 = _camp.Name;
                            }

                            DataRow _dr = GetSourceItemByMSCRMID(gData.ToString());

                            DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;

                            Boolean _Add_New = false;

                            if (_dr == null)
                            {
                                _dr = _dt.NewRow();
                                _dr["id"] = Guid.NewGuid().ToString();
                                _Add_New = true;
                            }

                            // *********************************

                            _dr["tekz"] = Sync_Entity_Type;
                            _dr["tekz2"] = Sync_Entity_Type2;

                            _dr["taetigkeit"] = _new_timeentry.a24_matchcode_str;
                            _dr["taetigkeiturl"] = "";

                            Double _dauer = 0.0;

                            try
                            {
                                _dauer = Convert.ToDouble(_new_timeentry.a24_duration_in_h_dec.Value);
                            }
                            catch
                            {
                            }

                            _dr["gesamtdauer"] = _dauer.ToString("0.00");

                            if (_new_timeentry.a24_date_dat != null) _dr["datum"] = ConvertDateTime(_new_timeentry.a24_date_dat.Value).ToShortDateString();

                            if (_new_timeentry.a24_start_dat == null && _new_timeentry.a24_end_dat == null)
                            {
                                // $$$$$$$$$$$$$$$$$$$$$ TODO


                            }
                            else
                            {
                                if (_new_timeentry.a24_start_dat != null) _dr["beginn"] = _new_timeentry.a24_start_dat.Value.Hour.ToString("00") + ":" + _new_timeentry.a24_start_dat.Value.Minute.ToString("00"); else _dr["beginn"] = "00:00";
                                if (_new_timeentry.a24_end_dat != null) _dr["ende"] = _new_timeentry.a24_end_dat.Value.Hour.ToString("00") + ":" + _new_timeentry.a24_end_dat.Value.Minute.ToString("00"); else _dr["ende"] = "00:00";
                            }

                            Double _zz = 0.0;
                            try
                            {
                                _zz = Convert.ToDouble(_new_timeentry.a24_additional_time_dec.Value);
                            }
                            catch
                            {
                            }

                            Double _zz_minutes = _zz * 60.0;

                            TimeSpan _zz_tsp = new TimeSpan(0, (int)_zz_minutes, 0);

                            _dr["zusatzzeit"] = _zz_tsp.Hours.ToString("00") + ":" + _zz_tsp.Minutes.ToString("00");

                            _dr["synced"] = "yes";
                            _dr["mscrmentityid"] = gData.ToString();

                            // Status

                            if (_new_timeentry.statuscode != null) _dr["mscrmstatus"] = GetDescFromList(_Statuscode_PicklistData, _new_timeentry.statuscode.Value.ToString()); else _dr["mscrmstatus"] = "";
                            if (_new_timeentry.statuscode != null) _dr["mscrmstatusid"] = _new_timeentry.statuscode.Value.ToString(); else _dr["mscrmstatusid"] = "";

                            try
                            {
                                _dr["fahrtkosten"] = Convert.ToBoolean(_new_timeentry.a24_travel_expenses_opt.Value).ToString();
                            }
                            catch
                            {
                                _dr["fahrtkosten"] = "false";
                            }

                            String _te_timemeasurements = "";

                            String _new_duration_measured = "";
                            try
                            {
                                _new_duration_measured = _new_timeentry.a24_duration_measured_dec.ToString();
                            }
                            catch
                            {
                            }

                            if (_new_duration_measured != null)
                            {
                                Double _dm_duration = 0.0;

                                // daurer = gemessene Zeit
                                if (_new_duration_measured.IndexOf(':') > 0)
                                {
                                    String[] _dm = _new_duration_measured.Split(':');
                                    Int32 _dm_hours = 0;
                                    Int32 _dm_minutes = 0;
                                    try
                                    {
                                        _dm_hours = Convert.ToInt32(_dm[0]);
                                        _dm_minutes = Convert.ToInt32(_dm[1]);
                                    }
                                    catch
                                    {
                                    }
                                    _dm_duration = _dm_hours + (_dm_minutes / 60.0);
                                }
                                else
                                {
                                    try
                                    {
                                        _dm_duration = Convert.ToDouble(_new_duration_measured);
                                    }
                                    catch
                                    {
                                    }
                                }

                                _dr["dauer"] = _dm_duration.ToString("0.00");
                                Double _ms = _dm_duration * 60.0 * 60.0 * 1000.0;
                                if (_ms > 0.0)
                                {
                                    _te_timemeasurements += "***|***|";
                                    _te_timemeasurements += _ms.ToString();
                                }
                            }
                            else
                            {
                                // keine gemessene Zeit -> Gesamtdauer - ZZ = gemessene Zeit

                                _dr["dauer"] = (_dauer - _zz).ToString("0.00");
                                Double _ms = (_dauer - _zz) * 60.0 * 60.0 * 1000.0;
                                if (_ms > 0.0)
                                {
                                    _te_timemeasurements += "***|***|";
                                    _te_timemeasurements += _ms.ToString();
                                }
                            }

                            _dr["timemeasurements"] = _te_timemeasurements;
                            if (_new_timeentry.a24_description_txt != null) _dr["beschreibung"] = _new_timeentry.a24_description_txt; else _dr["beschreibung"] = "";
                            if (_new_timeentry.a24_annotation_txt != null) _dr["anmerkung"] = _new_timeentry.a24_annotation_txt; else _dr["anmerkung"] = "";

                            // Bezug 1
                            _dr["mscrmbezug"] = _enitity_name;
                            _dr["mscrmbezugid"] = _entity_id.ToString();

                            // Bezug 2
                            _dr["mscrmbezug2"] = _enitity_name_2;
                            _dr["mscrmbezugid2"] = _entity_id_2.ToString();

                            // Ausgeführt durch
                            _dr["mscrmuser"] = _new_executeuserid.Name;
                            _dr["mscrmuserid"] = _new_executeuserid.Id.ToString();

                            // Billing Type
                            if (_new_timeentry.a24_billing_opt != null)
                            {
                                _dr["mscrmbillingtype"] = GetDescFromList(_Billing_PicklistData, _new_timeentry.a24_billing_opt.Value.ToString());
                                _dr["mscrmbillingtypeid"] = _new_timeentry.a24_billing_opt.Value;
                            }

                            _dr["lasterror"] = "";

                            // Bezuglink 1
                            String _LinkTmp = "";

                            if (Sync_Entity_Type == WorkPackageString) _LinkTmp = _MSCRM_LinkTemplate_AP;
                            else if (Sync_Entity_Type == IncidentString) _LinkTmp = _MSCRM_LinkTemplate_ANF;
                            else if (Sync_Entity_Type == OpportunityString) _LinkTmp = _MSCRM_LinkTemplate_VC;
                            //else if (Sync_Entity_Type == "REL") _LinkTmp = _MSCRM_LinkTemplate_REL;
                            else if (Sync_Entity_Type == CampaignString) _LinkTmp = _MSCRM_LinkTemplate_KA;

                            _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid"].ToString());
                            if (_parent_entity_id != "") _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _parent_entity_id);

                            _dr["mscrmlinkurl"] = _LinkTmp;
                            _dr["mscrmlinktext"] = _enitity_name;

                            // Bezuglink 2
                            _LinkTmp = "";

                            if (Sync_Entity_Type2 == WorkPackageString) _LinkTmp = _MSCRM_LinkTemplate_AP;
                            else if (Sync_Entity_Type2 == IncidentString) _LinkTmp = _MSCRM_LinkTemplate_ANF;
                            else if (Sync_Entity_Type2 == OpportunityString) _LinkTmp = _MSCRM_LinkTemplate_VC;
                            //else if (Sync_Entity_Type2 == "REL") _LinkTmp = _MSCRM_LinkTemplate_REL;
                            else if (Sync_Entity_Type2 == CampaignString) _LinkTmp = _MSCRM_LinkTemplate_KA;

                            _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid2"].ToString());
                            if (_parent_entity_id != "") _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _parent_entity_id_2);

                            _dr["mscrmlinkurl2"] = _LinkTmp;
                            _dr["mscrmlinktext2"] = _enitity_name_2;

                            _LinkTmp = _MSCRM_LinkTemplate_TE;
                            _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmentityid"].ToString());
                            _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr["mscrmbezugid"].ToString());
                            _dr["taetigkeiturl"] = _LinkTmp;

                            _dr["syncfrommscrm"] = "true";

                            _dr["marked"] = false;

                            if (_Add_New) _dt.Rows.Add(_dr);

                            _dt.AcceptChanges();

                        }
                    }

                }

                this.Cursor = _old_cursor;

                TimeEntryGrid.RefreshData();
                view.BestFitColumns();

                // TimeEntryGrid.SortBy(TimeEntryGrid.Columns["datum"], DevExpress.Data.ColumnSortOrder.Descending);
                // TimeEntryGrid.GroupBy(TimeEntryGrid.Columns["datum"]);

                TimeEntryGrid.ExpandAllGroups();
                TimeEntryGrid.UpdateTotalSummary();

                staticItemInfo.Content = _c.ToString() + " " + rm.GetString("SuccessfullyTransferredCRM");

                FileLog("LoadTimeEntitiesFromMSCRM:end", "User=" + sDefaultUser);
            }
            catch (FaultException<Microsoft.Xrm.Sdk.OrganizationServiceFault> err)
            {
                FileLog("LoadTimeEntitiesFromMSCRM", "Error: " + err.Message);
            }
        }

        private void SyncEntity(IOrganizationService _Service, String EntityXmlID, String sEntity, String sNameAttribute, String sIDAttribute, String StateCodeValue, DateTime LastSyncDateTime)
        {
            try
            {
                string sDataFieldName;
                string sDescFieldName;
                string sEntityName;

                String sSearchString = "";
                sSearchString = sSearchString.Replace("*", "%");

                if (sSearchString.EndsWith("%") == false)
                    sSearchString = sSearchString + "%";

                sDescFieldName = sNameAttribute;
                sDataFieldName = sIDAttribute;
                sEntityName = sEntity;

                // Create the ColumnSet that indicates the properties to be retrieved.
                ColumnSet cols = new ColumnSet();


                // Set the properties of the ColumnSet.
                if (EntityXmlID == "WORKPACKAGE")
                {   // AP
                    cols = new ColumnSet(new String[] { sNameAttribute, sIDAttribute, "modifiedon", Constants.FieldNames.WorkPackage.BillingOptionSet, Constants.FieldNames.WorkPackage.CustomerReference });
                }
                else if (EntityXmlID == "INCIDENT")
                {
                    // ANF
                    cols = new ColumnSet(new String[] { sNameAttribute, sIDAttribute, "modifiedon", Constants.FieldNames.Incident.BillingOptionSet, "customerid" });
                }
                else if (EntityXmlID == "OPPORTUNITY")
                {
                    // VC
                    cols = new ColumnSet(new String[] { sNameAttribute, sIDAttribute, "modifiedon", "customerid" });
                }
                else if (EntityXmlID == "CAMPAIGN")
                {
                    // KA
                    cols = new ColumnSet(new String[] { sNameAttribute, sIDAttribute, "modifiedon" });
                }
                //else if (EntityXmlID == "NEW_VERSION")
                //{
                //    // REL
                //    cols = new ColumnSet(new String[] { sNameAttribute, sIDAttribute, "modifiedon", "new_addonid", "createdon", "new_versionsnr" });
                //}
                else if (EntityXmlID == "ACTIONITEM")
                {
                    // KA
                    cols = new ColumnSet(new String[] { sNameAttribute, sIDAttribute, "modifiedon" });
                }

                FilterExpression filterExpression = new FilterExpression();

                ConditionExpression condition = new ConditionExpression(sNameAttribute, ConditionOperator.Like, sSearchString);
                ConditionExpression condition2 = new ConditionExpression("statecode", ConditionOperator.Equal, StateCodeValue);

                ConditionExpression condition3 = new ConditionExpression("modifiedon", ConditionOperator.GreaterThan, LastSyncDateTime);

                filterExpression.FilterOperator = LogicalOperator.And;
                filterExpression.AddCondition(condition);
                filterExpression.AddCondition(condition2);
                filterExpression.AddCondition(condition3);

                QueryExpression query = new QueryExpression()
                {
                    EntityName = sEntity,
                    ColumnSet = cols,
                    Criteria = filterExpression,
                    Distinct = false,
                    NoLock = true
                };

                RetrieveMultipleRequest multipleRequest = new RetrieveMultipleRequest();
                multipleRequest.Query = query;

                RetrieveMultipleResponse response = (RetrieveMultipleResponse)_Service.Execute(multipleRequest);

                String sDesc = "";
                Guid gData = Guid.Empty;

                DateTime _dt = new DateTime();
                DateTime _tmp_dt = new DateTime();

                Int32 _Entity_Count = 0;
                Int32 _Entity_New = 0;
                Int32 _Entity_Update = 0;
                Int32 _Entity_BillTypeCode = 0;

                Guid _MscrmParentID = Guid.Empty;
                String new_versionsnr = "";

                DateTime _last_sync_datetime = LastSyncDateTime;

                foreach (Entity _e in response.EntityCollection.Entities)
                {
                    sDesc = "";
                    gData = Guid.Empty;
                    new_versionsnr = "";

                    sDesc = _e.Attributes[sDescFieldName].ToString();
                    gData = new Guid(_e.Attributes[sDataFieldName].ToString());

                    if (_e.Contains("createdon"))
                    {
                        _tmp_dt = Convert.ToDateTime(_e.Attributes["createdon"].ToString());
                        if (_tmp_dt > _last_sync_datetime) _last_sync_datetime = _tmp_dt;
                    }
                    if (_e.Contains("modifiedon"))
                    {
                        _dt = Convert.ToDateTime(_e.Attributes["modifiedon"].ToString());
                        if (_dt > _last_sync_datetime) _last_sync_datetime = _dt;
                    }

                    if (_e.Contains(Constants.FieldNames.WorkPackage.BillingOptionSet)) if (_e.Attributes[Constants.FieldNames.WorkPackage.BillingOptionSet] != null) _Entity_BillTypeCode = ((OptionSetValue)_e.Attributes[Constants.FieldNames.WorkPackage.BillingOptionSet]).Value;
                    if (_e.Contains(Constants.FieldNames.TimeEntry.AccountReference)) if (_e.Attributes[Constants.FieldNames.TimeEntry.AccountReference] != null) _MscrmParentID = ((EntityReference)_e.Attributes[Constants.FieldNames.TimeEntry.AccountReference]).Id;
                    if (_e.Contains(Constants.FieldNames.Incident.BillingOptionSet)) if (_e.Attributes[Constants.FieldNames.Incident.BillingOptionSet] != null) _Entity_BillTypeCode = ((OptionSetValue)_e.Attributes[Constants.FieldNames.Incident.BillingOptionSet]).Value;
                    if (_e.Contains("customerid")) if (_e.Attributes["customerid"] != null) _MscrmParentID = ((EntityReference)_e.Attributes["customerid"]).Id;
                    if (_e.Contains("new_addonid")) if (_e.Attributes["new_addonid"] != null) _MscrmParentID = ((EntityReference)_e.Attributes["new_addonid"]).Id;
                    //if (_e.Contains("new_versionsnr")) if (_e.Attributes["new_versionsnr"] != null) new_versionsnr = _e.Attributes["new_versionsnr"].ToString();

                    if (string.IsNullOrEmpty(sDesc) == false && gData != Guid.Empty)
                    {
                        DataRow _dr = null;
                        _Entity_Count++;

                        DateTime _datetime = _dt;

                        // FileLog(">>>>>", _datetime.ToShortDateString() + " " + _datetime.ToLongTimeString());

                        if ((_dr = GetCacheItemByMSCRMID(EntityXmlID, gData.ToString())) == null)
                        {
                            _Entity_New++;
                            _dr = _TimeEntryCache.NewRow();
                            _dr["id"] = Guid.NewGuid().ToString();
                            _dr["type"] = EntityXmlID;
                            if (EntityXmlID == "WORKPACKAGE" || EntityXmlID == "INCIDENT")
                            {
                                _dr["subtype"] = _Entity_BillTypeCode.ToString();
                            }
                            else _dr["subtype"] = "";

                            _dr["mscrmid"] = gData.ToString();
                            _dr["description"] = sDesc;
                            _dr["mscrmparentid"] = _MscrmParentID.ToString();
                            _dr["zusatzinfo1"] = "";
                            if (_tmp_dt != null) _dr["zusatzinfo1"] = _tmp_dt.ToShortDateString() + " " + _tmp_dt.ToShortTimeString();
                            _dr["zusatzinfo2"] = "";
                            //if (EntityXmlID == "NEW_VERSION") _dr["zusatzinfo2"] = new_versionsnr;
                            _TimeEntryCache.Rows.Add(_dr);
                            _TimeEntryCache.AcceptChanges();
                        }
                        else
                        {
                            _Entity_Update++;
                            _dr["type"] = EntityXmlID;
                            if (EntityXmlID == "WORKPACKAGE" || EntityXmlID == "INCIDENT")
                            {
                                _dr["subtype"] = _Entity_BillTypeCode.ToString();
                            }
                            else _dr["subtype"] = "";
                            _dr["mscrmid"] = gData.ToString();
                            _dr["mscrmparentid"] = _MscrmParentID.ToString();
                            _dr["description"] = sDesc;
                            _dr["zusatzinfo1"] = "";
                            if (_tmp_dt != null) _dr["zusatzinfo1"] = _tmp_dt.ToShortDateString() + " " + _tmp_dt.ToShortTimeString();
                            _dr["zusatzinfo2"] = "";
                            //if (EntityXmlID == "NEW_VERSION") _dr["zusatzinfo2"] = new_versionsnr;
                            _TimeEntryCache.AcceptChanges();
                        }
                    }
                }

                if (EntityXmlID == "WORKPACKAGE")
                {
                    _LastSyncDateTime_WORKPACKAGE = _last_sync_datetime;
                }
                else if (EntityXmlID == "INCIDENT")
                {
                    _LastSyncDateTime_INCIDENT = _last_sync_datetime;
                }
                else if (EntityXmlID == "OPPORTUNITY")
                {
                    _LastSyncDateTime_OPPORTUNITY = _last_sync_datetime;
                }
                else if (EntityXmlID == "CAMPAIGN")
                {
                    _LastSyncDateTime_CAMPAIGN = _last_sync_datetime;
                }
                //else if (EntityXmlID == "NEW_VERSION")
                //{
                //    _LastSyncDateTime_NEW_VERSION = _last_sync_datetime;
                //}
                else if (EntityXmlID == "ACTIONITEM")
                {
                    _LastSyncDateTime_ACTIONITEM = _last_sync_datetime;
                }

                WriteCache();

                FileLog("SyncEntity@" + sEntity + "/" + sNameAttribute + "/" + sIDAttribute, "Count=" + _Entity_Count.ToString() + "/New=" + _Entity_New.ToString() + "/Update=" + _Entity_Update.ToString() + "/_last_sync_datetime=" + _last_sync_datetime.ToLongDateString() + " " + _last_sync_datetime.ToLongTimeString());

            }
            catch (Exception ex)
            {
                FileLog("SyncEntity", "Error: " + ex.Message);
            }
        }

        /// <summary>
        /// The MSCRM Worker Thread checks the connection every second and sets the correct buttonstates.
        /// Gets triggered from the MainWindow Constructor and repeated every 1000 ms.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        [STAThreadAttribute]
        void mscrm_worker_Elapsed(object sender, EventArgs e)
        {
            try
            {
                ConnectionCheckState();

                SyncToMscrm();
            }
            catch (SoapException se)
            {
                if (se.Detail != null) if (se.Detail.InnerText != null)
                        FileLog("mscrm_worker_Elapsed", "Error: " + se.Detail.InnerText);

                MSCRM_ConState = MSCRM_ConectionStates.NotConnected;
            }
            catch (Exception ex)
            {
                FileLog("mscrm_worker_Elapsed", "Error: " + ex.Message);
                if (ex.InnerException != null) FileLog("mscrm_worker_Elapsed", "Error: " + ex.InnerException.Message);

                MSCRM_ConState = MSCRM_ConectionStates.NotConnected;
            }
        }

        /// <summary>
        /// Syncs the entries in the TimeEntryGrid marked with SYNCED = NO to MSCRM
        /// Gets triggered by the mscrm worker thread every 1000 ms.
        /// </summary>
        private void SyncToMscrm()
        {
            // Sync to MSCRM
            if (MSCRM_ConState == MSCRM_ConectionStates.Connected)
            {
                DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
                foreach (DataRow _dr in _dt.Rows)
                {
                    // Update KW
                    try
                    {
                        DateTime _DateTime = Convert.ToDateTime(_dr["datum"].ToString());
                        DateUtils.CalendarWeek _cw = DateUtils.GetGermanCalendarWeek(_DateTime);

                        if (_cw.Week.ToString() != _dr["kw"].ToString())
                        {
                            _dr["kw"] = _cw.Week;
                            _dr.AcceptChanges();
                        }
                    }
                    catch { }

                    if (_dr["synced"].ToString().Trim().ToLower() == "no")
                    {
                        int ViewRowHandle = 0;

                        for (Int32 _c = 0; _c < _dt.Rows.Count; _c++)
                        {
                            DataRowView _dv = (DataRowView)TimeEntryGrid.GetRowByListIndex(_c);

                            String _id = _dv.Row["id"].ToString();

                            if (_id == _dr["id"].ToString())
                            {
                                ViewRowHandle = TimeEntryGrid.GetRowHandleByListIndex(_c);
                                break;
                            }

                            // TimeEntryGrid.GetRowHandleByListIndex(listIndex++)
                            // TimeEntryGrid.GetRowByListIndex(_c)
                        }

                        view.ScrollIntoView(ViewRowHandle);

                        // TimeEntryGrid.ExpandGroupRow(groupRowHandle, true);

                        Guid gSelectedParent = Guid.Empty;
                        Guid gSelectedParent2 = Guid.Empty;

                        try
                        {
                            if (_dr["mscrmbezugid"].ToString() != string.Empty)
                                gSelectedParent = new Guid(_dr["mscrmbezugid"].ToString());
                        }
                        catch { }

                        try
                        {
                            // A24 JNE 24.9.2015 added check for empty string to prevent exception
                            if (_dr["mscrmbezugid2"].ToString() != string.Empty)
                                gSelectedParent2 = new Guid(_dr["mscrmbezugid2"].ToString());
                        }
                        catch { }

                        if (gSelectedParent != Guid.Empty)
                        {
                            a24_timeentry te = null;

                            String Sync_Entity_Type = _dr["tekz"].ToString().Trim().ToUpper();
                            String Sync_Entity_Type2 = _dr["tekz2"].ToString().Trim().ToUpper();

                            String _parent_entity_id = "";
                            String _parent_entity_id2 = "";

                            Boolean AddNew = false;

                            Guid teGuid = Guid.Empty;

                            if (_dr["mscrmentityid"].ToString().Trim() != "")
                            {
                                AddNew = false;
                                teGuid = new Guid(_dr["mscrmentityid"].ToString());
                                te = (a24_timeentry)_OrganizationService.Retrieve(a24_timeentry.EntityLogicalName, teGuid, new ColumnSet(true));
                                // service.Delete(EntityName.new_timeentry.ToString(), new Guid(_dr["mscrmentityid"].ToString()));
                            }
                            else
                            {
                                te = new a24_timeentry();
                                _dr["mscrmentityid"] = "";
                                AddNew = true;
                            }

                            te.a24_matchcode_str = _dr["taetigkeit"].ToString();


                            EntityReference lParent = new EntityReference();
                            lParent.Id = gSelectedParent;

                            EntityReference lParent2 = new EntityReference();
                            lParent2.Id = gSelectedParent2;

                            if (te.a24_working_package_ref != null) te.a24_working_package_ref = null;
                            if (te.a24_incident_ref != null) te.a24_incident_ref = null;
                            if (te.a24_opportunity_ref != null) te.a24_opportunity_ref = null;
                            //if (te.New_VersionId != null) te.New_VersionId = null;
                            if (te.a24_campaign_ref != null) te.a24_campaign_ref = null;
                            if (te.a24_action_item_ref != null) te.a24_action_item_ref = null;

                            // BEZUG 1

                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                lParent.LogicalName = a24_workpackage.EntityLogicalName;
                                te.a24_working_package_ref = lParent;

                                a24_workpackage wp = (a24_workpackage)_OrganizationService.Retrieve(a24_workpackage.EntityLogicalName, new Guid(_dr["mscrmbezugid"].ToString()), new ColumnSet(true));

                                if (wp.a24_customer_ref != null)
                                {
                                    _parent_entity_id = wp.a24_customer_ref.Id.ToString();

                                }

                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                lParent.LogicalName = Incident.EntityLogicalName;
                                te.a24_incident_ref = lParent;

                                Incident _in = (Incident)_OrganizationService.Retrieve(Incident.EntityLogicalName, new Guid(_dr["mscrmbezugid"].ToString()), new ColumnSet(true));
                                _parent_entity_id = _in.CustomerId.Id.ToString();
                            }
                            else if (Sync_Entity_Type == OpportunityString)
                            {

                                lParent.LogicalName = Opportunity.EntityLogicalName;
                                te.a24_opportunity_ref = lParent;

                                Opportunity _op = (Opportunity)_OrganizationService.Retrieve(Opportunity.EntityLogicalName, new Guid(_dr["mscrmbezugid"].ToString()), new ColumnSet(true));
                                _parent_entity_id = _op.CustomerId.Id.ToString();
                            }
                            //else if (Sync_Entity_Type == "REL")
                            //{
                            //    lParent.LogicalName = New_Version.EntityLogicalName;
                            //    te.New_VersionId = lParent;

                            //    _parent_entity_id = "";
                            //}
                            else if (Sync_Entity_Type == CampaignString)
                            {
                                lParent.LogicalName = Campaign.EntityLogicalName;
                                te.a24_campaign_ref = lParent;

                                _parent_entity_id = "";
                            }
                            else if (Sync_Entity_Type == ActionItemString)
                            {
                                lParent.LogicalName = a24_action_item.EntityLogicalName;
                                te.a24_action_item_ref = lParent;

                                _parent_entity_id = "";
                            }

                            // BEZUG 2

                            if (Sync_Entity_Type2 == WorkPackageString)
                            {
                                lParent.LogicalName = a24_workpackage.EntityLogicalName;
                                if (te.a24_working_package_ref == null) te.a24_working_package_ref = lParent2;
                                try
                                {
                                    a24_workpackage wp = (a24_workpackage)_OrganizationService.Retrieve(a24_workpackage.EntityLogicalName, new Guid(_dr["mscrmbezugid2"].ToString()), new ColumnSet(true));
                                    _parent_entity_id = wp.a24_customer_ref.Id.ToString();
                                }
                                catch { }
                            }
                            else if (Sync_Entity_Type2 == IncidentString)
                            {
                                lParent.LogicalName = Incident.EntityLogicalName;
                                if (te.a24_incident_ref == null) te.a24_incident_ref = lParent2;
                                try
                                {
                                    Incident _in = (Incident)_OrganizationService.Retrieve(Incident.EntityLogicalName, new Guid(_dr["mscrmbezugid2"].ToString()), new ColumnSet(true));
                                    _parent_entity_id = _in.CustomerId.Id.ToString();
                                }
                                catch { }
                            }
                            else if (Sync_Entity_Type2 == OpportunityString)
                            {
                                lParent.LogicalName = Opportunity.EntityLogicalName;
                                if (te.a24_opportunity_ref == null) te.a24_opportunity_ref = lParent2;

                                try
                                {
                                    Opportunity _op = (Opportunity)_OrganizationService.Retrieve(Opportunity.EntityLogicalName, new Guid(_dr["mscrmbezugid2"].ToString()), new ColumnSet(true));
                                    _parent_entity_id = _op.CustomerId.Id.ToString();
                                }
                                catch { }
                            }
                            //else if (Sync_Entity_Type2 == "REL")
                            //{
                            //    lParent.LogicalName = New_Version.EntityLogicalName;
                            //    if (te.New_VersionId == null) te.New_VersionId = lParent2;

                            //    _parent_entity_id = "";
                            //}
                            else if (Sync_Entity_Type2 == CampaignString)
                            {
                                lParent.LogicalName = Campaign.EntityLogicalName;
                                if (te.a24_campaign_ref == null) te.a24_campaign_ref = lParent2;

                                _parent_entity_id = "";
                            }
                            else if (Sync_Entity_Type == ActionItemString)
                            {
                                lParent.LogicalName = a24_action_item.EntityLogicalName;
                                te.a24_action_item_ref = lParent;

                                _parent_entity_id = "";
                            }

                            Double dHours = Convert.ToDouble(_dr["gesamtdauer"].ToString());

                            //CrmDecimal dDuration = new CrmDecimal();
                            //dDuration.Value = (decimal)dHours;


                            // TODO: Warum Indexerfehler hier?
                            //te.a24_duration_in_h_dec.Value = Convert.ToDecimal(dHours);
                            
                            DateTime _dt_temp = Convert.ToDateTime(_dr["datum"].ToString(), CultureInfo.CurrentCulture.DateTimeFormat);

                            //DateTime cdt = new DateTime();
                            //cdt.Value = _dt_temp.Month.ToString() + "-" + _dt_temp.Day.ToString() + "-" + _dt_temp.Year.ToString();

                            //string temp = _dt_temp.Day.ToString() + "." + _dt_temp.Month.ToString() + "." + _dt_temp.Year.ToString();
                            te.a24_date_dat = Convert.ToDateTime(_dt_temp.ToShortDateString(), CultureInfo.CurrentCulture.DateTimeFormat);

                            EntityReference lUser = new EntityReference();
                            lUser.Id = new Guid(_dr["mscrmuserid"].ToString());
                            lUser.LogicalName = SystemUser.EntityLogicalName;
                            te.a24_executed_by_ref = lUser;

                            DateTime cdt = Convert.ToDateTime(_dr["beginn"].ToString());
                            te.a24_start_dat = cdt;

                            cdt = Convert.ToDateTime(_dr["ende"].ToString());
                            te.a24_end_dat = cdt;

                            String _zusartzzeit = _dr["zusatzzeit"].ToString();

                            Double _zz_duration = 0.0;

                            if (_zusartzzeit.IndexOf(':') > 0)
                            {
                                String[] _zz = _zusartzzeit.Split(':');
                                Int32 _zz_hours = 0;
                                Int32 _zz_minutes = 0;
                                try
                                {
                                    _zz_hours = Convert.ToInt32(_zz[0]);
                                    _zz_minutes = Convert.ToInt32(_zz[1]);
                                }
                                catch
                                {
                                }
                                _zz_duration = _zz_hours + (_zz_minutes / 60.0);
                            }
                            else
                            {
                                try
                                {
                                    _zz_duration = Convert.ToDouble(_zusartzzeit);
                                }
                                catch
                                {
                                }
                            }

                            // CrmDecimal dDurationAdditional = new CrmDecimal();
                            // dDurationAdditional.Value = (decimal)_zz_duration;

                            te.a24_additional_time_dec = (decimal)_zz_duration;

                            // CrmBoolean btravelexpenses = new CrmBoolean();
                            // btravelexpenses.Value = Convert.ToBoolean(_dr["fahrtkosten"].ToString());
                            te.a24_travel_expenses_opt = Convert.ToBoolean(_dr["fahrtkosten"].ToString());

                            te.a24_description_txt = (string)_dr["beschreibung"];
                            te.a24_annotation_txt = (string)_dr["anmerkung"];

                            Double _dauer = Convert.ToDouble(_dr["dauer"].ToString());

                            _dauer = _dauer * 60.0;
                            TimeSpan _dauer_gemessen = new TimeSpan(0, (int)_dauer, 0);

                            te.a24_duration_measured_dec = _dauer_gemessen.Hours + _dauer_gemessen.Minutes / 60; // _dauer_gemessen.Hours.ToString("00") + ":" + _dauer_gemessen.Minutes.ToString("00");

                            try
                            {
                                te.a24_billing_opt = new OptionSetValue(Convert.ToInt32(_dr["mscrmbillingtypeid"].ToString()));

                                /*Picklist pBilling = new Picklist();
                                int iValue = Convert.ToInt32(_dr["mscrmbillingtypeid"].ToString());
                                if (iValue >= 0)
                                {
                                    pBilling.name = (string)_dr["mscrmbillingtype"].ToString();
                                    pBilling.Value = iValue;
                                    te.new_billing = pBilling;
                                }*/
                            }
                            catch
                            {
                            }

                            try
                            {
                                // A24 JNE check for number to prevent format exception
                                int n;
                                if (int.TryParse(_dr["mscrmstatusid"].ToString(), out n))
                                    te.statuscode = new OptionSetValue(Convert.ToInt16(_dr["mscrmstatusid"].ToString()));
                                /*Status _new_status = new Status();
                                int iValue = Convert.ToInt32(_dr["mscrmstatusid"].ToString());
                                if (iValue >= 0)
                                {
                                    _new_status.Value = iValue;
                                    _new_status.name = _dr["mscrmstatus"].ToString();
                                    te.statuscode = _new_status;
                                }*/
                            }
                            catch
                            {
                            }

                            try
                            {
                                if (AddNew) teGuid = _OrganizationService.Create(te); else _OrganizationService.Update(te);

                                if (teGuid != Guid.Empty)
                                {
                                    _dr["synced"] = "yes";
                                    _dr["syncfrommscrm"] = "false";
                                    _dr["lasterror"] = "";
                                    _dr["mscrmentityid"] = teGuid.ToString();

                                    _dr["mscrmlinktext"] = _dr["mscrmbezug"];
                                    _dr["mscrmlinktext2"] = _dr["mscrmbezug2"];

                                    String _LinkTmp = "";

                                    // BEZUG 1

                                    if (Sync_Entity_Type == WorkPackageString) _LinkTmp = _MSCRM_LinkTemplate_AP;
                                    else if (Sync_Entity_Type == IncidentString) _LinkTmp = _MSCRM_LinkTemplate_ANF;
                                    else if (Sync_Entity_Type == OpportunityString) _LinkTmp = _MSCRM_LinkTemplate_VC;
                                    //else if (Sync_Entity_Type == "REL") _LinkTmp = _MSCRM_LinkTemplate_REL;
                                    else if (Sync_Entity_Type == CampaignString) _LinkTmp = _MSCRM_LinkTemplate_KA;
                                    else
                                    {
                                        _dr["mscrmlinkurl"] = "";
                                        _dr["mscrmlinktext"] = "";
                                    }

                                    if (_LinkTmp != "")
                                    {
                                        _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid"].ToString());
                                        if (_parent_entity_id != "") _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _parent_entity_id);
                                        _dr["mscrmlinkurl"] = _LinkTmp;
                                    }

                                    // BEZUG 2

                                    _LinkTmp = "";
                                    if (Sync_Entity_Type2 == WorkPackageString) _LinkTmp = _MSCRM_LinkTemplate_AP;
                                    else if (Sync_Entity_Type2 == IncidentString) _LinkTmp = _MSCRM_LinkTemplate_ANF;
                                    else if (Sync_Entity_Type2 == OpportunityString) _LinkTmp = _MSCRM_LinkTemplate_VC;
                                    //else if (Sync_Entity_Type2 == "REL") _LinkTmp = _MSCRM_LinkTemplate_REL;
                                    else if (Sync_Entity_Type2 == CampaignString) _LinkTmp = _MSCRM_LinkTemplate_KA;
                                    else
                                    {
                                        _dr["mscrmlinkurl2"] = "";
                                        _dr["mscrmlinktext2"] = "";
                                    }

                                    if (_LinkTmp != "")
                                    {
                                        _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid2"].ToString());
                                        if (_parent_entity_id2 != "") _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _parent_entity_id2);
                                        _dr["mscrmlinkurl2"] = _LinkTmp;
                                    }

                                    // TÄTIGKEIT

                                    _LinkTmp = _MSCRM_LinkTemplate_TE;
                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmentityid"].ToString());
                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr["mscrmbezugid"].ToString());
                                    _dr["taetigkeiturl"] = _LinkTmp;

                                    _dr.AcceptChanges();
                                }
                            }
                            catch (Exception ex)
                            {
                                _dr["synced"] = "error";
                                _dr["lasterror"] = ex.Message;
                                _dr.AcceptChanges();

                                FileLog("mscrm_worker_Elapsed@SyncTimeentry", "Error: " + ex.Message);
                            }
                        }
                    }
                }

                ConnectionCheckState();
            }
        }

        /// <summary>
        /// Checks the Connection State and sets the Button Enabling and status text accordingly
        /// </summary>
        private void ConnectionCheckState()
        {
            switch (MSCRM_ConState)
            {
                case MSCRM_ConectionStates.Connected:
                    staticItemLine.Content = "ONLINE (" + sDefaultUser + ")";
                    bSyncMSCRMActMonth.IsEnabled = true;
                    bSyncMSCRMPreMonth.IsEnabled = true;
                    bSyncMSCRMEntityRefresh.IsEnabled = true;
                    bGoOnline.IsEnabled = false;
                    bGoOffline.IsEnabled = true;

                    // set status text
                    if (_currentConnectionDetail != null)
                        staticItemInfo.Content = _currentConnectionDetail.OrganizationServiceUrl;
                    else
                    {
                        staticItemInfo.Content = rm.GetString("NoConnection");
                    }

                    break;
                case MSCRM_ConectionStates.Error:
                    staticItemLine.Content = "OFFLINE";
                    bSyncMSCRMActMonth.IsEnabled = false;
                    bSyncMSCRMPreMonth.IsEnabled = false;
                    bSyncMSCRMEntityRefresh.IsEnabled = true;
                    bGoOnline.IsEnabled = true;
                    bGoOffline.IsEnabled = false;

                    break;
                case MSCRM_ConectionStates.NotConnected:
                    staticItemLine.Content = "OFFLINE";
                    bSyncMSCRMActMonth.IsEnabled = false;
                    bSyncMSCRMPreMonth.IsEnabled = false;
                    bSyncMSCRMEntityRefresh.IsEnabled = true;
                    bGoOnline.IsEnabled = true;
                    bGoOffline.IsEnabled = false;

                    break;
                case MSCRM_ConectionStates.InSyncProgress:
                    staticItemLine.Content = "SYNCING...";
                    bSyncMSCRMEntityRefresh.IsEnabled = false;
                    bGoOnline.IsEnabled = false;
                    bGoOffline.IsEnabled = false;

                    break;
                // this state is used to connect to mscrm
                // the thread gets triggered every 1000 ms.
                case MSCRM_ConectionStates.RequestConnection:
                    staticItemLine.Content = "CONNECTING...";
                    bSyncMSCRMActMonth.IsEnabled = false;
                    bSyncMSCRMPreMonth.IsEnabled = false;
                    bSyncMSCRMEntityRefresh.IsEnabled = false;

                    // set status text
                    if (_currentConnectionDetail != null)
                        staticItemInfo.Content = _currentConnectionDetail.OrganizationServiceUrl;
                    else
                    {
                        staticItemInfo.Content = rm.GetString("NoConnection");
                    }

                    // check whether program is currently attempting to establish a connection
                    if (!_bIsConnecting)
                    {
                        _bIsConnecting = true;

                        // if connection details exist, connect.
                        if (_currentConnectionDetail != null)
                        {
                            _cManager.ConnectToServer(_currentConnectionDetail, null);
                            bGoOnline.IsEnabled = false;
                            bGoOffline.IsEnabled = true;
                        }
                        else
                        {
                            // if saved connections exist, take the first one
                            if (_cManager.ConnectionsList.Connections.Count > 0)
                            {
                                _currentConnectionDetail = _cManager.ConnectionsList.Connections[0];
                                if (string.IsNullOrEmpty(_currentConnectionDetail.UserPassword))
                                {
                                    _currentConnectionDetail = null;
                                    bGoOnline.IsEnabled = true;
                                    bGoOffline.IsEnabled = false;

                                }
                                else
                                {
                                    _cManager.ConnectToServer(_currentConnectionDetail, null);
                                    bGoOnline.IsEnabled = false;
                                    bGoOffline.IsEnabled = true;
                                }
                            }
                            else
                            {
                                // if neither exists, open window
                                _fHelper.AskForConnection("ApplyConnectionToTabs");
                            }
                        }
                    }


                    //System.Configuration.ConfigurationManager.AppSettings["MSCRM.ServiceURL"].ToString();



                    break;
            }
        }

        private String GetDescFromList(String List, String ID)
        {
            String _rV = "";

            String[] _list_items = List.Split(';');

            foreach (String _item in _list_items)
            {
                String[] _value_item = _item.Split('|');
                if (_value_item[1] == ID) { _rV = _value_item[0]; break; }
            }

            return _rV;
        }

        private Boolean PingUri(String _uri)
        {
            Boolean _rV = false;

            try
            {
                System.Uri uri = new System.Uri(_uri);

                Ping ping = new Ping();
                PingReply pingreply = ping.Send(uri.Host);

                if (pingreply.Status == IPStatus.Success) _rV = true;
            }
            catch (PingException pe)
            {
                _rV = false; // MessageBox.Show(pe.Message);
            }

            return _rV;
        }

        static public DateTime ConvertDateTime(DateTime dateTime)
        {
            DateTime _DateTime = new DateTime();

            TimeSpan offset = TimeZone.CurrentTimeZone.GetUtcOffset(dateTime);

            string sOffset = string.Empty;

            if (offset.Hours < 0)
            {
                sOffset = "-" + (offset.Hours * -1).ToString().PadLeft(2, '0');
            }

            {
                sOffset = "+" + offset.Hours.ToString().PadLeft(2, '0');
            }

            sOffset += offset.Minutes.ToString().PadLeft(2, '0');

            _DateTime = Convert.ToDateTime(dateTime.ToString(string.Format("yyyy-MM-ddTHH:mm:ss{0}", sOffset)));

            return _DateTime;
        }

        [STAThreadAttribute]
        public void workerSyncThreadDoWork()
        {
            try
            {
                if (MSCRM_ConState == MSCRM_ConectionStates.RequestConnection || MSCRM_ConState == MSCRM_ConectionStates.InSyncProgress) return;

                Boolean ConnectionSuccess = false;

                if (MSCRM_ConState == MSCRM_ConectionStates.NotConnected || MSCRM_ConState == MSCRM_ConectionStates.Error)
                {
                    // A24 JNE 24.9.2015 connection establishing is now triggered via the Request Connection state.
                    MSCRM_ConState = MSCRM_ConectionStates.RequestConnection;
                    //_OrganizationService = CRMManager.GetServiceProxy();
                }
                else ConnectionSuccess = true;

                PicklistManager _pm = new PicklistManager();

                if (ConnectionSuccess)
                {
                    MSCRM_ConState = MSCRM_ConectionStates.InSyncProgress;

                    // Caches
                    try
                    {
                        _Billing_PicklistData = _pm.GetPicklistStr(_OrganizationService, a24_timeentry.EntityLogicalName, "a24_billing_opt");

                        // Update PICKLIST.BILLING Cache
                        if (_Billing_PicklistData != "")
                        {
                            // Clear PICKLIST.BILLING Cache
                            DataRow[] _result = _TimeEntryCache.Select("[type]='PICKLIST.BILLING'");
                            foreach (DataRow _row in _result)
                            {
                                _TimeEntryCache.Rows.Remove(_row);
                            }

                            String[] _picklist_items = _Billing_PicklistData.Split(';');
                            foreach (String _item in _picklist_items)
                            {
                                String[] _picklist_item = _item.Split('|');

                                DataRow _dr = _TimeEntryCache.NewRow();
                                _dr["id"] = Guid.NewGuid().ToString();
                                _dr["type"] = "PICKLIST.BILLING";
                                _dr["subtype"] = "";
                                _dr["mscrmid"] = _picklist_item[1];
                                _dr["mscrmparentid"] = "";
                                _dr["description"] = _picklist_item[0];
                                _TimeEntryCache.Rows.Add(_dr);
                                _TimeEntryCache.AcceptChanges();
                            }
                        }
                    }
                    catch (SoapException se)
                    {
                        if (se.Detail != null) if (se.Detail.InnerText != null)
                                FileLog("worker_Elapsed@PICKLIST.BILLING", "Error: " + se.Detail.InnerText);
                    }
                    catch (Exception ex)
                    {
                        FileLog("worker_Elapsed@PICKLIST.BILLING", "Error: " + ex.Message);
                        if (ex.InnerException != null) FileLog("worker_Elapsed@PICKLIST.BILLING", "Error: " + ex.InnerException.Message);
                    }

                    // Caches
                    try
                    {
                        _Statuscode_PicklistData = _pm.GetStatuslistStr(_OrganizationService, a24_timeentry.EntityLogicalName, "statuscode");

                        // Update STATUSLIST.STATUSCODE Cache
                        if (_Statuscode_PicklistData != "")
                        {
                            // Clear PICKLIST.BILLING Cache
                            DataRow[] _result = _TimeEntryCache.Select("[type]='STATUSLIST.STATUSCODE'");
                            foreach (DataRow _row in _result)
                            {
                                _TimeEntryCache.Rows.Remove(_row);
                            }

                            String[] _statuslist_items = _Statuscode_PicklistData.Split(';');

                            foreach (String _item in _statuslist_items)
                            {
                                String[] _picklist_item = _item.Split('|');

                                DataRow _dr = _TimeEntryCache.NewRow();
                                _dr["id"] = Guid.NewGuid().ToString();
                                _dr["type"] = "STATUSLIST.STATUSCODE";
                                _dr["subtype"] = "";
                                _dr["mscrmid"] = _picklist_item[1];
                                _dr["mscrmparentid"] = "";
                                _dr["description"] = _picklist_item[0];
                                _TimeEntryCache.Rows.Add(_dr);
                                _TimeEntryCache.AcceptChanges();
                            }
                        }
                    }
                    catch (SoapException se)
                    {
                        if (se.Detail != null) if (se.Detail.InnerText != null)
                                FileLog("worker_Elapsed@STATUSLIST.STATUSCODE", "Error: " + se.Detail.InnerText);
                    }
                    catch (Exception ex)
                    {
                        FileLog("worker_Elapsed@STATUSLIST.STATUSCODE", "Error: " + ex.Message);
                        if (ex.InnerException != null) FileLog("worker_Elapsed@STATUSLIST.STATUSCODE", "Error: " + ex.InnerException.Message);
                    }

                    try
                    {
                        _User_ListData = CRMManager.GetSearchList(_OrganizationService, "", SystemUser.EntityLogicalName, "fullname", "systemuserid");

                        // Update USER.LIST Cache
                        if (_User_ListData != "")
                        {
                            // Clear PICKLIST.BILLING Cache
                            DataRow[] _result = _TimeEntryCache.Select("[type]='USER.LIST'");
                            foreach (DataRow _row in _result)
                            {
                                _TimeEntryCache.Rows.Remove(_row);
                            }

                            String[] _statuslist_items = _User_ListData.Split(';');

                            foreach (String _item in _statuslist_items)
                            {
                                String[] _picklist_item = _item.Split('|');

                                if (_picklist_item[0].Trim().ToUpper() != "SYSTEM" && _picklist_item[0].Trim().ToUpper() != "INTEGRATION")
                                {
                                    DataRow _dr = _TimeEntryCache.NewRow();
                                    _dr["id"] = Guid.NewGuid().ToString();
                                    _dr["type"] = "USER.LIST";
                                    _dr["subtype"] = "";
                                    _dr["mscrmid"] = _picklist_item[1];
                                    _dr["mscrmparentid"] = "";
                                    _dr["description"] = _picklist_item[0];
                                    _TimeEntryCache.Rows.Add(_dr);
                                    _TimeEntryCache.AcceptChanges();
                                }
                            }
                        }
                    }
                    catch (SoapException se)
                    {
                        if (se.Detail != null) if (se.Detail.InnerText != null)
                                FileLog("worker_Elapsed@USER.LIST", "Error: " + se.Detail.InnerText);
                    }
                    catch (Exception ex)
                    {
                        FileLog("worker_Elapsed@USER.LIST", "Error: " + ex.Message);
                        if (ex.InnerException != null) FileLog("worker_Elapsed@USER.LIST", "Error: " + ex.InnerException.Message);
                    }

                    try
                    {
                        FileLog("worker_Elapsed@SyncEntities", "Sync Dynamics CRM WORKPACKAGE Entities ...");
                        SyncEntity(_OrganizationService, "WORKPACKAGE", a24_workpackage.EntityLogicalName, Constants.FieldNames.TimeEntry.Name, Constants.FieldNames.WorkPackage.Id, "Active", _LastSyncDateTime_WORKPACKAGE);

                        FileLog("worker_Elapsed@SyncEntities", "Sync Dynamics CRM INCIDENT Entities ...");
                        SyncEntity(_OrganizationService, "INCIDENT", Incident.EntityLogicalName, "title", "incidentid", "Active", _LastSyncDateTime_INCIDENT);

                        FileLog("worker_Elapsed@SyncEntities", "Sync Dynamics CRM OPPORTUNITY Entities ...");
                        SyncEntity(_OrganizationService, "OPPORTUNITY", Opportunity.EntityLogicalName, "name", "opportunityid", "Open", _LastSyncDateTime_OPPORTUNITY);

                        FileLog("worker_Elapsed@SyncEntities", "Sync Dynamics CRM CAMPAIGN Entities ...");
                        SyncEntity(_OrganizationService, "CAMPAIGN", Campaign.EntityLogicalName, "name", "campaignid", "Active", _LastSyncDateTime_CAMPAIGN);

                        //FileLog("worker_Elapsed@SyncEntities", "Sync Dynamics CRM NEW_VERSION Entities ...");
                        //SyncEntity(_OrganizationService, "NEW_VERSION", New_Version.EntityLogicalName, "new_versionsnummer", "new_versionid", "Active", _LastSyncDateTime_NEW_VERSION);

                        FileLog("worker_Elapsed@SyncEntities", "Sync Dynamics CRM ACTIONITEM Entities ...");
                        SyncEntity(_OrganizationService, "ACTIONITEM", a24_action_item.EntityLogicalName, Constants.FieldNames.ActionItem.Name, Constants.FieldNames.ActionItem.Id, "Active", _LastSyncDateTime_ACTIONITEM);

                        FileLog("worker_Elapsed", "Sync completed!");
                    }
                    catch (SoapException se)
                    {
                        if (se.Detail != null) if (se.Detail.InnerText != null)
                                FileLog("worker_Elapsed@SyncEntities", "Error: " + se.Detail.InnerText);
                    }
                    catch (Exception ex)
                    {
                        FileLog("worker_Elapsed@SyncEntities", "Error: " + ex.Message);
                        if (ex.InnerException != null) FileLog("worker_Elapsed@SyncEntities", "Error: " + ex.InnerException.Message);
                    }

                    MSCRM_ConState = MSCRM_ConectionStates.Connected;

                    WriteCache();

                }
            }
            catch (SoapException se)
            {
                if (se.Detail != null) if (se.Detail.InnerText != null)
                        FileLog("worker_Elapsed", "Error: " + se.Detail.InnerText);
            }
            catch (Exception ex)
            {
                FileLog("worker_Elapsed", "Error: " + ex.Message);
                if (ex.InnerException != null) FileLog("worker_Elapsed", "Error: " + ex.InnerException.Message);
            }
        }

        public String RightStr(string param, Int32 length)
        {
            try
            {
                string result = param.Substring(param.Length - length, length);
                return result;
            }
            catch
            {
                return "";
            }
        }

        private void FileLog(String Source, String Message)
        {
            try
            {
                String logFilePath = System.IO.Directory.GetCurrentDirectory() + "\\logs";

                if (!System.IO.Directory.Exists(logFilePath)) System.IO.Directory.CreateDirectory(logFilePath);

                logFilePath += "\\logs" + DateTime.Now.Year.ToString() + RightStr("00" + DateTime.Now.Month.ToString(), 2) + RightStr("00" + DateTime.Now.Day.ToString(), 2) + ".txt";

                TextWriter tw = new StreamWriter(logFilePath, true);

                // write a line of text to the file
                tw.WriteLine(DateTime.Now.ToLongTimeString() + "." + DateTime.Now.Millisecond.ToString() + ";" + Source + ";" + System.Environment.MachineName + ";" + Message);

                // close the stream
                tw.Close();

                // MessageBox.Show(Message);
            }
            catch
            {

            }
        }

        private void OnNewItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            ActivateTimeEntry(Guid.Empty.ToString(), false, DateTime.Now, "", "");
        }

        private void OnCopyItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            CopyTimeEntry();
        }

        private void OnEditItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // activate selected time entry ...

            try
            {
                if (tabTE_1.Tag.ToString() == "" || tabTE_2.Tag.ToString() == "" || tabTE_3.Tag.ToString() == "" || tabTE_4.Tag.ToString() == "" || tabTE_5.Tag.ToString() == "")
                {
                    Int32 _Count_Marked = 0;


                    DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
                    foreach (DataRow _dr in _dt.Rows)
                    {
                        if (_dr["marked"] != DBNull.Value)
                        {
                            if (Convert.ToBoolean(_dr["marked"].ToString()))
                            {
                                _Count_Marked++;
                                String _te_guid = _dr["id"].ToString();
                                if (_te_guid != "")
                                {
                                    if (!ActivateTimeEntry(_te_guid, false, DateTime.Now, "", "")) break;
                                }

                            }
                        }
                    }

                    staticItemInfo.Content = rm.GetString("TimeEntriesOpen1") + " " + _Count_Marked.ToString() + " " + rm.GetString("TimeEntriesOpen2");
                }
                else
                {
                    staticItemInfo.Content = rm.GetString("EntriesOpened");
                }
            }
            catch (Exception ex)
            {
                FileLog("OnEditItemClick", "Error: " + ex.Message);
                // MessageBox.Show("OnEditItemClick: " + ex.Message, "Fehlermeldung", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OnDeleteItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // delete selected time entry ... 

            try
            {
                Int32 _Count_Marked = 0;
                Boolean ActiveTEtoDelete = false;

                DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
                foreach (DataRow _dr in _dt.Rows)
                {
                    if (_dr["marked"] != DBNull.Value)
                    {
                        if (Convert.ToBoolean(_dr["marked"].ToString()))
                        {
                            _Count_Marked++;
                            String _te_guid = _dr["id"].ToString();
                            if (_te_guid != "")
                            {
                                if (tabTE_1.Tag.ToString() == _te_guid || tabTE_2.Tag.ToString() == _te_guid || tabTE_3.Tag.ToString() == _te_guid || tabTE_4.Tag.ToString() == _te_guid || tabTE_5.Tag.ToString() == _te_guid)
                                {
                                    ActiveTEtoDelete = true;
                                    break;
                                }
                            }

                        }
                    }
                }

                if (ActiveTEtoDelete)
                {
                    staticItemInfo.Content = rm.GetString("TemporaryContract");
                }
                else
                {
                    if (_Count_Marked > 0)
                    {
                        Int32 _Count_Removed = 0;

                        Boolean _RemoveItems = false;
                        if (_Count_Marked == 1)
                        {
                            if (MessageBox.Show(rm.GetString("ActiveTEtoDelete"), rm.GetString("ActiveTEtoDeleteTitle"), MessageBoxButton.YesNoCancel, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes) _RemoveItems = true;
                        }
                        else
                        {
                            if (MessageBox.Show(rm.GetString("ActiveTEtoDeleteMany1") + " " + _Count_Marked.ToString() + " " + rm.GetString("ActiveTEtoDeleteMany2"), rm.GetString("ActiveTEtoDeleteManyTitle"), MessageBoxButton.YesNoCancel, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes) _RemoveItems = true;
                        }

                        if (_RemoveItems)
                        {
                            Cursor _old_cursor = this.Cursor;
                            this.Cursor = Cursors.Wait;

                            Int32 _Marked_ZE = _Count_Marked;

                            try
                            {
                                _dt = (DataTable)TimeEntryGrid.ItemsSource;
                                while (_Count_Marked > 0)
                                {
                                    foreach (DataRow _dr in _dt.Rows)
                                    {
                                        if (_dr["marked"] != DBNull.Value)
                                        {
                                            if (Convert.ToBoolean(_dr["marked"].ToString()))
                                            {
                                                if (_dr["synced"].ToString().ToLower() == "yes")
                                                {
                                                    Guid _mscrmentityid = new Guid(_dr["mscrmentityid"].ToString());

                                                    a24_timeentry _new_timeentry = (a24_timeentry)_OrganizationService.Retrieve(a24_timeentry.EntityLogicalName, _mscrmentityid, new ColumnSet(true));
                                                    if (_new_timeentry != null)
                                                    {
                                                        try
                                                        {
                                                            _OrganizationService.Delete(a24_timeentry.EntityLogicalName, _mscrmentityid);
                                                            _Count_Removed++;
                                                        }
                                                        catch { }
                                                    }
                                                }

                                                _Count_Marked--;
                                                _dt.Rows.Remove(_dr);

                                                break;
                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            {
                            }

                            this.Cursor = _old_cursor;

                            String _InfoStr = "";

                            if (_Count_Removed > 0)
                            {
                                if (_Count_Removed == 1)
                                {
                                    _InfoStr = rm.GetString("DeletedDynamicsCRM");
                                }
                                else
                                {
                                    _InfoStr = rm.GetString("DynamicsCRMdeleted1") + " " + _Count_Removed.ToString() + " " + rm.GetString("DynamicsCRMdeleted2");
                                }
                            }

                            if (_Count_Marked > 0)
                            {
                                if (_InfoStr != "") _InfoStr += " ";
                                if (_Count_Marked == 1)
                                {
                                    _InfoStr += rm.GetString("DeletedTimeEntry");
                                }
                                else
                                {
                                    _InfoStr += rm.GetString("DeletedFromGrid1") + " " + _Count_Marked.ToString() + " " + rm.GetString("DeletedFromGrid2");
                                }
                            }

                            staticItemInfo.Content = _InfoStr;

                            TimeEntryGrid.RefreshData();
                        }
                    }
                    else
                    {
                        staticItemInfo.Content = rm.GetString("TimeEntryForDeletion");
                    }
                }
            }
            catch (Exception ex)
            {
                FileLog("OnDeleteItemClick", "Error: " + ex.Message);
            }
        }

        static internal ImageSource doGetImageSourceFromResource(string psAssemblyName, string psResourceName)
        {
            Uri oUri = new Uri("pack://application:,,,/" + psAssemblyName + ";component/" + psResourceName, UriKind.RelativeOrAbsolute);
            return BitmapFrame.Create(oUri);
        }

        private void TimerCtrl_StartTimer()
        {
            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;
            if (_ti != null)
            {
                _timer.Start();
                if (stopWatch != null) stopWatch = null;
                stopWatch = new System.Diagnostics.Stopwatch();
                stopWatch.Start();
                DateTime _start_time = DateTime.Now;

                TimerMode = true;

                _timer_info_begin = _start_time.ToShortDateString() + " " + _start_time.ToLongTimeString();
                _timer_info_end = "";

                _timer_info_te_guid = _ti.Tag.ToString();

                if (_timer_info_te_guid == tabTE_1.Tag.ToString())
                {
                    lblTimerCtrlInfo_1.Content = "00:00:00";
                    btnTimerCtrlImage_1.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/stop.png");
                }
                else if (_timer_info_te_guid == tabTE_2.Tag.ToString())
                {
                    lblTimerCtrlInfo_2.Content = "00:00:00";
                    btnTimerCtrlImage_2.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/stop.png");
                }
                else if (_timer_info_te_guid == tabTE_3.Tag.ToString())
                {
                    lblTimerCtrlInfo_3.Content = "00:00:00";
                    btnTimerCtrlImage_3.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/stop.png");
                }
                else if (_timer_info_te_guid == tabTE_4.Tag.ToString())
                {
                    lblTimerCtrlInfo_4.Content = "00:00:00";
                    btnTimerCtrlImage_4.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/stop.png");
                }
                else if (_timer_info_te_guid == tabTE_5.Tag.ToString())
                {
                    lblTimerCtrlInfo_5.Content = "00:00:00";
                    btnTimerCtrlImage_5.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/stop.png");
                }
            }
        }

        private void UpdateTimeMeasurements()
        {
            // Update TimeMeasurements

            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;
            if (_ti != null)
            {
                if (_ti.Tag != null)
                {
                    String _timer_info_te_guid = _ti.Tag.ToString();

                    DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
                    DataRow _dr = _dt.Rows.Find(_timer_info_te_guid);
                    if (_dr != null)
                    {
                        if (_timer_info_te_guid == tabTE_1.Tag.ToString())
                        {
                            CalcTimeValues(_dr, 1);
                        }
                        else if (_timer_info_te_guid == tabTE_2.Tag.ToString())
                        {
                            CalcTimeValues(_dr, 2);
                        }
                        else if (_timer_info_te_guid == tabTE_3.Tag.ToString())
                        {
                            CalcTimeValues(_dr, 3);
                        }
                        else if (_timer_info_te_guid == tabTE_4.Tag.ToString())
                        {
                            CalcTimeValues(_dr, 4);
                        }
                        else if (_timer_info_te_guid == tabTE_5.Tag.ToString())
                        {
                            CalcTimeValues(_dr, 5);
                        }

                        _dr.AcceptChanges();
                    }
                }
            }
        }

        private void TimerCtrl_StopTimer()
        {
            _timer.Stop();
            stopWatch.Stop();
            DateTime _stop_time = DateTime.Now;

            TimerMode = false;

            _timer_info_end = _stop_time.ToShortDateString() + " " + _stop_time.ToLongTimeString();

            // Add time measurement to active time entry
            DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
            DataRow _dr = _dt.Rows.Find(_timer_info_te_guid);
            if (_dr != null)
            {
                String _te_timemeasurements = _dr["timemeasurements"].ToString();

                if (_te_timemeasurements != "") _te_timemeasurements += ";";

                _te_timemeasurements += _timer_info_begin + "|";
                _te_timemeasurements += _timer_info_end + "|";
                _te_timemeasurements += stopWatch.ElapsedMilliseconds.ToString();

                _dr["timemeasurements"] = _te_timemeasurements;

                if (_timer_info_te_guid == tabTE_1.Tag.ToString())
                {
                    CalcTimeValues(_dr, 1);
                    lblTimerCtrlInfo_1.Content = "00:00:00";
                    btnTimerCtrlImage_1.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/media_play.png");
                }
                else if (_timer_info_te_guid == tabTE_2.Tag.ToString())
                {
                    CalcTimeValues(_dr, 2);
                    lblTimerCtrlInfo_2.Content = "00:00:00";
                    btnTimerCtrlImage_2.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/media_play.png");
                }
                else if (_timer_info_te_guid == tabTE_3.Tag.ToString())
                {
                    CalcTimeValues(_dr, 3);
                    lblTimerCtrlInfo_3.Content = "00:00:00";
                    btnTimerCtrlImage_3.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/media_play.png");
                }
                else if (_timer_info_te_guid == tabTE_4.Tag.ToString())
                {
                    CalcTimeValues(_dr, 4);
                    lblTimerCtrlInfo_4.Content = "00:00:00";
                    btnTimerCtrlImage_4.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/media_play.png");
                }
                else if (_timer_info_te_guid == tabTE_5.Tag.ToString())
                {
                    CalcTimeValues(_dr, 5);
                    lblTimerCtrlInfo_5.Content = "00:00:00";
                    btnTimerCtrlImage_5.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/media_play.png");
                }

                _dr.AcceptChanges();
            }

            _timer_info_te_guid = "";
        }

        private void btnTimerCtrl_Click(object sender, RoutedEventArgs e)
        {
            String new_timer_info_te_guid = "";

            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;
            if (_ti != null)
            {
                new_timer_info_te_guid = _ti.Tag.ToString();
                if (new_timer_info_te_guid != _timer_info_te_guid && _timer_info_te_guid != "") TimerCtrl_StopTimer();
            }

            if (!TimerMode)
            {
                // Start Timer
                TimerCtrl_StartTimer();
            }
            else
            {
                // Stop Timer
                TimerCtrl_StopTimer();
            }
        }

        void t_Elapsed(object sender, EventArgs e)
        {
            String _time_info = "";

            _time_info += TimeSpan.FromMilliseconds(stopWatch.ElapsedMilliseconds).Hours.ToString("00");
            _time_info += ":";
            _time_info += TimeSpan.FromMilliseconds(stopWatch.ElapsedMilliseconds).Minutes.ToString("00");
            _time_info += ":";
            _time_info += TimeSpan.FromMilliseconds(stopWatch.ElapsedMilliseconds).Seconds.ToString("00");

            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;
            if (_ti != null)
            {
                if (_timer_info_te_guid == tabTE_1.Tag.ToString())
                {
                    lblTimerCtrlInfo_1.Content = _time_info;
                }
                else if (_timer_info_te_guid == tabTE_2.Tag.ToString())
                {
                    lblTimerCtrlInfo_2.Content = _time_info;
                }
                else if (_timer_info_te_guid == tabTE_3.Tag.ToString())
                {
                    lblTimerCtrlInfo_3.Content = _time_info;
                }
                else if (_timer_info_te_guid == tabTE_4.Tag.ToString())
                {
                    lblTimerCtrlInfo_4.Content = _time_info;
                }
                else if (_timer_info_te_guid == tabTE_5.Tag.ToString())
                {
                    lblTimerCtrlInfo_5.Content = _time_info;
                }
            }

            if (DateTime.Now.Hour == 23 &&
                DateTime.Now.Minute == 59 &&
                DateTime.Now.Second == 59)
            {
                // Stop Timer
                TimerCtrl_StopTimer();
            }
        }

        private void TimeEntryGrid_Loaded(object sender, RoutedEventArgs e)
        {
            view.BestFitColumns();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                //Return back values in DataTable
                DataTable _dt2 = (DataTable)TimeEntryGrid.ItemsSource;
                foreach (DataRow _dr in _dt2.Rows)
                {
                    if (_BezugValues.ToList().Contains(_dr["tekz"]))
                    {
                        int i = _BezugValues.ToList().IndexOf(_dr["tekz"].ToString());
                        string value = _BezugValuesStartupLanguage[i];
                        _dr["tekz"] = value;
                    }
                }

                FileLog("Window_Closing", "[start]");

                // Timer beenden...
                FileLog("Window_Closing", "TimerCtrl_StopTimer()");
                if (TimerMode) TimerCtrl_StopTimer();

                // Quit worker
                FileLog("Window_Closing", "_mscrm_worker.Stop()");
                if (_mscrm_worker != null) _mscrm_worker.Stop();

                FileLog("Window_Closing", "_workerThread.Abort()");
                if (_workerSyncThread != null) _workerSyncThread.Abort();

                // Änderungen speichern...
                if (TimeEntryGrid.ItemsSource != null)
                {
                    // Abort all open TimeEntries ...
                    foreach (TabItem _ti in TimeEntryTabCtrl.Items)
                    {
                        CancelTimeEntry(_ti);
                    }

                    // Add empty row ...
                    DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
                    DataRow _dr = _dt.NewRow();

                    _dr["id"] = "00000000-0000-0000-0000-0000000";
                    _dr["tekz"] = "";
                    _dr["tekz2"] = "";
                    _dr["datum"] = "";
                    _dr["beginn"] = "";
                    _dr["ende"] = "";
                    _dr["dauer"] = "";
                    _dr["zusatzzeit"] = "";
                    _dr["gesamtdauer"] = "";
                    _dr["taetigkeit"] = "";
                    _dr["taetigkeiturl"] = "";
                    _dr["synced"] = "";
                    _dr["mscrmentityid"] = "";
                    _dr["mscrmstatus"] = "";
                    _dr["mscrmstatusid"] = "";
                    _dr["fahrtkosten"] = "";
                    _dr["timemeasurements"] = "";
                    _dr["beschreibung"] = "";
                    _dr["anmerkung"] = "";
                    _dr["mscrmbezug"] = "";
                    _dr["mscrmbezugid"] = "";
                    _dr["mscrmbezug2"] = "";
                    _dr["mscrmbezugid2"] = "";
                    _dr["mscrmuser"] = "";
                    _dr["mscrmuserid"] = "";
                    _dr["mscrmbillingtype"] = "";
                    _dr["mscrmbillingtypeid"] = "";
                    _dr["mscrmlinktext"] = "";
                    _dr["mscrmlinkurl"] = "";
                    _dr["mscrmlinktext2"] = "";
                    _dr["mscrmlinkurl2"] = "";
                    _dr["mscrmmodifiedon"] = "";
                    _dr["syncfrommscrm"] = "";
                    _dr["lasterror"] = "";

                    _dt.Rows.Add(_dr);
                    _dt.AcceptChanges();

                    FileLog("Window_Closing", "DataSet.WriteXml(TimeEntrySource_Xml_Path)");

                    WriteGrid();
                    WriteCache();
                }

                FileLog("Window_Closing", "TimeEntryGrid.SaveLayoutToXml(TimeEntryGridSettings)");

                WriteLayout();

                FileLog("Window_Closing", "[end]");
            }
            catch (Exception ex)
            {
                FileLog("Window_Closing", "Error: " + ex.Message);
            }
        }

        private void WriteLayout()
        {
            try
            {
                TimeEntryGrid.SaveLayoutToXml(TimeEntryGridSettings);
            }
            catch (Exception ex)
            {
                FileLog("WriteLayout", "Error: " + ex.Message);
            }
        }

        private void WriteGrid()
        {
            try
            {
                ((DataTable)TimeEntryGrid.ItemsSource).DataSet.WriteXml(TimeEntrySource_Xml_Path);
            }
            catch (Exception ex)
            {
                FileLog("WriteGrid", "Error: " + ex.Message);
            }
        }

        private void WriteCache()
        {
            try
            {
                _TimeEntryCache.WriteXml(TimeEntryCache_Xml_Path);
                using (StreamWriter writer = new StreamWriter(LastSyncDateTime_Path, false))
                {
                    writer.WriteLine(this._LastSyncDateTime_WORKPACKAGE.ToShortDateString() + " " + this._LastSyncDateTime_WORKPACKAGE.ToLongTimeString());
                    writer.WriteLine(this._LastSyncDateTime_INCIDENT.ToShortDateString() + " " + this._LastSyncDateTime_INCIDENT.ToLongTimeString());
                    writer.WriteLine(this._LastSyncDateTime_OPPORTUNITY.ToShortDateString() + " " + this._LastSyncDateTime_OPPORTUNITY.ToLongTimeString());
                    writer.WriteLine(this._LastSyncDateTime_CAMPAIGN.ToShortDateString() + " " + this._LastSyncDateTime_CAMPAIGN.ToLongTimeString());
                    //writer.WriteLine(this._LastSyncDateTime_NEW_VERSION.ToShortDateString() + " " + this._LastSyncDateTime_NEW_VERSION.ToLongTimeString());
                    writer.WriteLine(this._LastSyncDateTime_ACTIONITEM.ToShortDateString() + " " + this._LastSyncDateTime_ACTIONITEM.ToLongTimeString());

                    writer.Close();
                }
            }
            catch (Exception ex)
            {
                FileLog("WriteCache", "Error: " + ex.Message);
            }
        }

        private void btnSaveTimeentry_Click(object sender, RoutedEventArgs e)
        {
            // save timeentry changes 
            SaveActiveTimeEntry(false, false);
        }

        private void SaveActiveTimeEntry(Boolean DoClose, Boolean DoOpenNew)
        {
            try
            {
                DateTime _SetDateTime = DateTime.Now;

                if (TimeEntryTabCtrl.SelectedItem != null)
                {
                    TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;
                    Boolean _save_success = false;

                    DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
                    DataRow _dr = _dt.Rows.Find(_ti.Tag.ToString());

                    DateTime _tmp_time;
                    String _LinkTmp = "";

                    if (_dr != null)
                    {
                        if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
                        {
                            if ((cbMSCRMEntity_1.SelectedItem != null && CheckSelectedEntity(_OrganizationService, 1, false)) || cbMSCRMEntity_1.Text.Trim() == "")
                            {
                                if (cbTaetigkeit_1.Text.Trim() != "")
                                {
                                    if (cbAusgefuehrtDurch_1.SelectedItem != null)
                                    {
                                        if (cbAbrechnungsweise_1.SelectedItem != null)
                                        {
                                            _dr["taetigkeit"] = cbTaetigkeit_1.Text.Trim();

                                            if (cbMSCRMEntity_1.SelectedItem != null)
                                            {
                                                _dr["tekz"] = cbBezugEntity_1.Text.Trim().ToUpper();

                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_1.SelectedItem;
                                                _dr["mscrmbezug"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbezugid"] = _li.Tag.ToString();

                                                DataRow _dr2 = null;
                                                String EntityXmlID = "";

                                                if (_dr["tekz"].ToString() == WorkPackageString) { _LinkTmp = _MSCRM_LinkTemplate_AP; EntityXmlID = "WORKPACKAGE"; }
                                                else if (_dr["tekz"].ToString() == IncidentString) { _LinkTmp = _MSCRM_LinkTemplate_ANF; EntityXmlID = "OPPORTUNITY"; }
                                                else if (_dr["tekz"].ToString() == OpportunityString) { _LinkTmp = _MSCRM_LinkTemplate_VC; EntityXmlID = "INCIDENT"; }
                                                ////else if (_dr["tekz"].ToString() == "REL") { _LinkTmp = _MSCRM_LinkTemplate_REL; EntityXmlID = "NEW_VERSION"; }
                                                else if (_dr["tekz"].ToString() == CampaignString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "CAMPAIGN"; }
                                                else if (_dr["tekz"].ToString() == ActionItemString) { _LinkTmp = _MSCRM_LinkTemplate_AIL; EntityXmlID = "ACTIONITEM"; }

                                                if ((_dr2 = GetCacheItemByMSCRMID(EntityXmlID, _li.Tag.ToString())) != null)
                                                {
                                                    // Bezuglink
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid"].ToString());
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr2["mscrmparentid"].ToString());

                                                    _dr["mscrmlinkurl"] = _LinkTmp;
                                                    _dr["mscrmlinktext"] = _dr2["description"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    _dr["mscrmlinktext"] = "";
                                                    _dr["mscrmlinkurl"] = "";
                                                }
                                            }
                                            else
                                            {
                                                _dr["mscrmbezug"] = "";
                                                _dr["mscrmbezugid"] = "";
                                                _dr["mscrmlinktext"] = "";
                                                _dr["mscrmlinkurl"] = "";
                                                _dr["tekz"] = "";
                                            }

                                            if (cbMSCRMEntity2_1.SelectedItem != null)
                                            {
                                                _dr["tekz2"] = cbBezugEntity2_1.Text.Trim().ToUpper();

                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_1.SelectedItem;
                                                _dr["mscrmbezug2"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbezugid2"] = _li.Tag.ToString();

                                                DataRow _dr2 = null;
                                                String EntityXmlID = "";

                                                if (_dr["tekz2"].ToString() == WorkPackageString) { _LinkTmp = _MSCRM_LinkTemplate_AP; EntityXmlID = "WORKPACKAGE"; }
                                                else if (_dr["tekz2"].ToString() == IncidentString) { _LinkTmp = _MSCRM_LinkTemplate_ANF; EntityXmlID = "OPPORTUNITY"; }
                                                else if (_dr["tekz2"].ToString() == OpportunityString) { _LinkTmp = _MSCRM_LinkTemplate_VC; EntityXmlID = "INCIDENT"; }
                                                //else if (_dr["tekz2"].ToString() == "REL") { _LinkTmp = _MSCRM_LinkTemplate_REL; EntityXmlID = "NEW_VERSION"; }
                                                else if (_dr["tekz2"].ToString() == CampaignString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "CAMPAIGN"; }
                                                else if (_dr["tekz2"].ToString() == ActionItemString) { _LinkTmp = _MSCRM_LinkTemplate_AIL; EntityXmlID = "ACTIONITEM"; }

                                                if ((_dr2 = GetCacheItemByMSCRMID(EntityXmlID, _li.Tag.ToString())) != null)
                                                {
                                                    // Bezuglink
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid2"].ToString());
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr2["mscrmparentid"].ToString());

                                                    _dr["mscrmlinkurl2"] = _LinkTmp;
                                                    _dr["mscrmlinktext2"] = _dr2["description"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    _dr["mscrmlinktext2"] = "";
                                                    _dr["mscrmlinkurl2"] = "";
                                                }
                                            }
                                            else
                                            {
                                                _dr["mscrmbezug2"] = "";
                                                _dr["mscrmbezugid2"] = "";
                                                _dr["mscrmlinktext2"] = "";
                                                _dr["mscrmlinkurl2"] = "";
                                                _dr["tekz2"] = "";
                                            }

                                            _dr["datum"] = deDatum_1.DateTime.ToShortDateString();
                                            _SetDateTime = deDatum_1.DateTime;

                                            try
                                            {
                                                if (tbBeginn_1.Text.Trim() != "")
                                                {
                                                    _tmp_time = Convert.ToDateTime(tbBeginn_1.EditValue);
                                                    _dr["beginn"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                                }
                                                else
                                                {
                                                    _dr["beginn"] = "00:00";
                                                }
                                            }
                                            catch { }

                                            try
                                            {
                                                _tmp_time = Convert.ToDateTime(tbEnde_1.EditValue);
                                                _dr["ende"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                            }
                                            catch { }

                                            _dr["dauer"] = lblDauer_1.Content;

                                            _tmp_time = Convert.ToDateTime(tbZusatzzeit_1.EditValue);
                                            _dr["zusatzzeit"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                            _dr["gesamtdauer"] = lblGesamtDauer_1.Content.ToString().Replace(",", "."); 

                                            _dr["beschreibung"] = teBeschreibung_1.Text.Trim();
                                            _dr["anmerkung"] = teAnmerkung_1.Text.Trim();
                                            _dr["fahrtkosten"] = ceFahrtkosten_1.IsChecked.ToString();

                                            if (cbStatus_1.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbStatus_1.SelectedItem;
                                                _dr["mscrmstatus"] = _li.Content.ToString().Trim();
                                                _dr["mscrmstatusid"] = _li.Tag;
                                            }

                                            if (cbAusgefuehrtDurch_1.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbAusgefuehrtDurch_1.SelectedItem;
                                                _dr["mscrmuser"] = _li.Content.ToString().Trim();
                                                _dr["mscrmuserid"] = _li.Tag;
                                            }

                                            if (cbAbrechnungsweise_1.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbAbrechnungsweise_1.SelectedItem;
                                                _dr["mscrmbillingtype"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbillingtypeid"] = _li.Tag.ToString();
                                            }

                                            _dr["syncfrommscrm"] = "false";
                                            _dr["synced"] = "no";

                                            _dr.AcceptChanges();
                                            _save_success = true;
                                        }
                                        else
                                        {
                                            MessageBox.Show(rm.GetString("BillingType"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show(rm.GetString("DynamicsCRM"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show(rm.GetString("ActivityNote"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                }
                            }
                            else
                            {
                                MessageBox.Show(rm.GetString("BezugNote"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                            }
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
                        {
                            if ((cbMSCRMEntity_2.SelectedItem != null && CheckSelectedEntity(_OrganizationService, 2, false)) || cbMSCRMEntity_2.Text.Trim() == "")
                            {
                                if (cbTaetigkeit_2.Text.Trim() != "")
                                {
                                    if (cbAusgefuehrtDurch_2.SelectedItem != null)
                                    {
                                        if (cbAbrechnungsweise_2.SelectedItem != null)
                                        {
                                            _dr["taetigkeit"] = cbTaetigkeit_2.Text.Trim();

                                            if (cbMSCRMEntity_2.SelectedItem != null)
                                            {
                                                _dr["tekz"] = cbBezugEntity_2.Text.Trim().ToUpper();

                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_2.SelectedItem;
                                                _dr["mscrmbezug"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbezugid"] = _li.Tag.ToString();

                                                DataRow _dr2 = null;
                                                String EntityXmlID = "";

                                                if (_dr["tekz"].ToString() == WorkPackageString) { _LinkTmp = _MSCRM_LinkTemplate_AP; EntityXmlID = "WORKPACKAGE"; }
                                                else if (_dr["tekz"].ToString() == IncidentString) { _LinkTmp = _MSCRM_LinkTemplate_ANF; EntityXmlID = "OPPORTUNITY"; }
                                                else if (_dr["tekz"].ToString() == OpportunityString) { _LinkTmp = _MSCRM_LinkTemplate_VC; EntityXmlID = "INCIDENT"; }
                                                //else if (_dr["tekz"].ToString() == "REL") { _LinkTmp = _MSCRM_LinkTemplate_REL; EntityXmlID = "NEW_VERSION"; }
                                                else if (_dr["tekz"].ToString() == CampaignString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "CAMPAIGN"; }
                                                else if (_dr["tekz"].ToString() == ActionItemString) { _LinkTmp = _MSCRM_LinkTemplate_AIL; EntityXmlID = "ACTIONITEM"; }

                                                if ((_dr2 = GetCacheItemByMSCRMID(EntityXmlID, _li.Tag.ToString())) != null)
                                                {
                                                    // Bezuglink
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid"].ToString());
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr2["mscrmparentid"].ToString());

                                                    _dr["mscrmlinkurl"] = _LinkTmp;
                                                    _dr["mscrmlinktext"] = _dr2["description"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    _dr["mscrmlinktext"] = "";
                                                    _dr["mscrmlinkurl"] = "";
                                                }
                                            }
                                            else
                                            {
                                                _dr["mscrmbezug"] = "";
                                                _dr["mscrmbezugid"] = "";
                                                _dr["mscrmlinktext"] = "";
                                                _dr["mscrmlinkurl"] = "";
                                                _dr["tekz"] = "";
                                            }

                                            if (cbMSCRMEntity2_2.SelectedItem != null)
                                            {
                                                _dr["tekz2"] = cbBezugEntity2_2.Text.Trim().ToUpper();

                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_2.SelectedItem;
                                                _dr["mscrmbezug2"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbezugid2"] = _li.Tag.ToString();

                                                DataRow _dr2 = null;
                                                String EntityXmlID = "";

                                                if (_dr["tekz2"].ToString() == WorkPackageString) { _LinkTmp = _MSCRM_LinkTemplate_AP; EntityXmlID = "WORKPACKAGE"; }
                                                else if (_dr["tekz2"].ToString() == IncidentString) { _LinkTmp = _MSCRM_LinkTemplate_ANF; EntityXmlID = "OPPORTUNITY"; }
                                                else if (_dr["tekz2"].ToString() == OpportunityString) { _LinkTmp = _MSCRM_LinkTemplate_VC; EntityXmlID = "INCIDENT"; }
                                                //else if (_dr["tekz2"].ToString() == "REL") { _LinkTmp = _MSCRM_LinkTemplate_REL; EntityXmlID = "NEW_VERSION"; }
                                                else if (_dr["tekz2"].ToString() == CampaignString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "CAMPAIGN"; }
                                                else if (_dr["tekz2"].ToString() == ActionItemString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "ACTIONITEM"; }

                                                if ((_dr2 = GetCacheItemByMSCRMID(EntityXmlID, _li.Tag.ToString())) != null)
                                                {
                                                    // Bezuglink
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid2"].ToString());
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr2["mscrmparentid"].ToString());

                                                    _dr["mscrmlinkurl2"] = _LinkTmp;
                                                    _dr["mscrmlinktext2"] = _dr2["description"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    _dr["mscrmlinktext2"] = "";
                                                    _dr["mscrmlinkurl2"] = "";
                                                }
                                            }
                                            else
                                            {
                                                _dr["mscrmbezug2"] = "";
                                                _dr["mscrmbezugid2"] = "";
                                                _dr["mscrmlinktext2"] = "";
                                                _dr["mscrmlinkurl2"] = "";
                                                _dr["tekz2"] = "";
                                            }

                                            _dr["datum"] = deDatum_2.DateTime.ToShortDateString();
                                            _SetDateTime = deDatum_2.DateTime;

                                            try
                                            {
                                                if (tbBeginn_2.Text.Trim() != "")
                                                {
                                                    _tmp_time = Convert.ToDateTime(tbBeginn_2.EditValue);
                                                    _dr["beginn"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                                }
                                                else
                                                {
                                                    _dr["beginn"] = "00:00";
                                                }
                                            }
                                            catch { }

                                            try
                                            {
                                                _tmp_time = Convert.ToDateTime(tbEnde_2.EditValue);
                                                _dr["ende"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                            }
                                            catch { }

                                            _dr["dauer"] = lblDauer_2.Content;

                                            _tmp_time = Convert.ToDateTime(tbZusatzzeit_2.EditValue);
                                            _dr["zusatzzeit"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                            _dr["gesamtdauer"] = lblGesamtDauer_2.Content.ToString().Replace(",", ".");

                                            _dr["beschreibung"] = teBeschreibung_2.Text.Trim();
                                            _dr["anmerkung"] = teAnmerkung_2.Text.Trim();
                                            _dr["fahrtkosten"] = ceFahrtkosten_2.IsChecked.ToString();

                                            if (cbStatus_2.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbStatus_2.SelectedItem;
                                                _dr["mscrmstatus"] = _li.Content.ToString().Trim();
                                                _dr["mscrmstatusid"] = _li.Tag;
                                            }

                                            if (cbAusgefuehrtDurch_2.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbAusgefuehrtDurch_2.SelectedItem;
                                                _dr["mscrmuser"] = _li.Content.ToString().Trim();
                                                _dr["mscrmuserid"] = _li.Tag;
                                            }

                                            if (cbAbrechnungsweise_2.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbAbrechnungsweise_2.SelectedItem;
                                                _dr["mscrmbillingtype"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbillingtypeid"] = _li.Tag.ToString();
                                            }

                                            _dr["syncfrommscrm"] = "false";
                                            _dr["synced"] = "no";

                                            _dr.AcceptChanges();
                                            _save_success = true;
                                        }
                                        else
                                        {
                                            MessageBox.Show(rm.GetString("BillingType"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show(rm.GetString("DynamicsCRM"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show(rm.GetString("ActivityNote"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                }
                            }
                            else
                            {
                                MessageBox.Show(rm.GetString("BezugNote"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                            }
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
                        {
                            if ((cbMSCRMEntity_3.SelectedItem != null && CheckSelectedEntity(_OrganizationService, 3, false)) || cbMSCRMEntity_3.Text.Trim() == "")
                            {
                                if (cbTaetigkeit_3.Text.Trim() != "")
                                {
                                    if (cbAusgefuehrtDurch_3.SelectedItem != null)
                                    {
                                        if (cbAbrechnungsweise_3.SelectedItem != null)
                                        {
                                            _dr["taetigkeit"] = cbTaetigkeit_3.Text.Trim();

                                            if (cbMSCRMEntity_3.SelectedItem != null)
                                            {
                                                _dr["tekz"] = cbBezugEntity_3.Text.Trim().ToUpper();

                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_3.SelectedItem;
                                                _dr["mscrmbezug"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbezugid"] = _li.Tag.ToString();

                                                DataRow _dr2 = null;
                                                String EntityXmlID = "";

                                                if (_dr["tekz"].ToString() == WorkPackageString) { _LinkTmp = _MSCRM_LinkTemplate_AP; EntityXmlID = "WORKPACKAGE"; }
                                                else if (_dr["tekz"].ToString() == IncidentString) { _LinkTmp = _MSCRM_LinkTemplate_ANF; EntityXmlID = "OPPORTUNITY"; }
                                                else if (_dr["tekz"].ToString() == OpportunityString) { _LinkTmp = _MSCRM_LinkTemplate_VC; EntityXmlID = "INCIDENT"; }
                                                //else if (_dr["tekz"].ToString() == "REL") { _LinkTmp = _MSCRM_LinkTemplate_REL; EntityXmlID = "NEW_VERSION"; }
                                                else if (_dr["tekz"].ToString() == CampaignString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "CAMPAIGN"; }
                                                else if (_dr["tekz"].ToString() == ActionItemString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "ACTIONITEM"; }

                                                if ((_dr2 = GetCacheItemByMSCRMID(EntityXmlID, _li.Tag.ToString())) != null)
                                                {
                                                    // Bezuglink
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid"].ToString());
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr2["mscrmparentid"].ToString());

                                                    _dr["mscrmlinkurl"] = _LinkTmp;
                                                    _dr["mscrmlinktext"] = _dr2["description"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    _dr["mscrmlinktext"] = "";
                                                    _dr["mscrmlinkurl"] = "";
                                                }
                                            }
                                            else
                                            {
                                                _dr["mscrmbezug"] = "";
                                                _dr["mscrmbezugid"] = "";
                                                _dr["mscrmlinktext"] = "";
                                                _dr["mscrmlinkurl"] = "";
                                                _dr["tekz"] = "";
                                            }

                                            if (cbMSCRMEntity2_3.SelectedItem != null)
                                            {
                                                _dr["tekz2"] = cbBezugEntity2_3.Text.Trim().ToUpper();

                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_3.SelectedItem;
                                                _dr["mscrmbezug2"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbezugid2"] = _li.Tag.ToString();

                                                DataRow _dr2 = null;
                                                String EntityXmlID = "";

                                                if (_dr["tekz2"].ToString() == WorkPackageString) { _LinkTmp = _MSCRM_LinkTemplate_AP; EntityXmlID = "WORKPACKAGE"; }
                                                else if (_dr["tekz2"].ToString() == IncidentString) { _LinkTmp = _MSCRM_LinkTemplate_ANF; EntityXmlID = "OPPORTUNITY"; }
                                                else if (_dr["tekz2"].ToString() == OpportunityString) { _LinkTmp = _MSCRM_LinkTemplate_VC; EntityXmlID = "INCIDENT"; }
                                                //else if (_dr["tekz2"].ToString() == "REL") { _LinkTmp = _MSCRM_LinkTemplate_REL; EntityXmlID = "NEW_VERSION"; }
                                                else if (_dr["tekz2"].ToString() == CampaignString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "CAMPAIGN"; }
                                                else if (_dr["tekz2"].ToString() == ActionItemString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "ACTIONITEM"; }

                                                if ((_dr2 = GetCacheItemByMSCRMID(EntityXmlID, _li.Tag.ToString())) != null)
                                                {
                                                    // Bezuglink
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid2"].ToString());
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr2["mscrmparentid"].ToString());

                                                    _dr["mscrmlinkurl2"] = _LinkTmp;
                                                    _dr["mscrmlinktext2"] = _dr2["description"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    _dr["mscrmbezug2"] = "";
                                                    _dr["mscrmbezugid2"] = "";
                                                    _dr["mscrmlinktext2"] = "";
                                                    _dr["mscrmlinkurl2"] = "";
                                                    _dr["tekz2"] = "";
                                                }
                                            }
                                            else
                                            {
                                                _dr["mscrmlinktext2"] = "";
                                                _dr["mscrmlinkurl2"] = "";
                                            }

                                            _dr["datum"] = deDatum_3.DateTime.ToShortDateString();
                                            _SetDateTime = deDatum_3.DateTime;

                                            try
                                            {
                                                if (tbBeginn_3.Text.Trim() != "")
                                                {
                                                    _tmp_time = Convert.ToDateTime(tbBeginn_3.EditValue);
                                                    _dr["beginn"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                                }
                                                else
                                                {
                                                    _dr["beginn"] = "00:00";
                                                }
                                            }
                                            catch { }

                                            try
                                            {
                                                _tmp_time = Convert.ToDateTime(tbEnde_3.EditValue);
                                                _dr["ende"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                            }
                                            catch { }

                                            _dr["dauer"] = lblDauer_3.Content;

                                            _tmp_time = Convert.ToDateTime(tbZusatzzeit_3.EditValue);
                                            _dr["zusatzzeit"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                            _dr["gesamtdauer"] = lblGesamtDauer_3.Content.ToString().Replace(",", ".");

                                            _dr["beschreibung"] = teBeschreibung_3.Text.Trim();
                                            _dr["anmerkung"] = teAnmerkung_3.Text.Trim();
                                            _dr["fahrtkosten"] = ceFahrtkosten_3.IsChecked.ToString();

                                            if (cbStatus_3.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbStatus_3.SelectedItem;
                                                _dr["mscrmstatus"] = _li.Content.ToString().Trim();
                                                _dr["mscrmstatusid"] = _li.Tag;
                                            }

                                            if (cbAusgefuehrtDurch_3.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbAusgefuehrtDurch_3.SelectedItem;
                                                _dr["mscrmuser"] = _li.Content.ToString().Trim();
                                                _dr["mscrmuserid"] = _li.Tag;
                                            }

                                            if (cbAbrechnungsweise_3.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbAbrechnungsweise_3.SelectedItem;
                                                _dr["mscrmbillingtype"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbillingtypeid"] = _li.Tag.ToString();
                                            }

                                            _dr["syncfrommscrm"] = "false";
                                            _dr["synced"] = "no";

                                            _dr.AcceptChanges();
                                            _save_success = true;

                                        }
                                        else
                                        {
                                            MessageBox.Show(rm.GetString("BillingType"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show(rm.GetString("DynamicsCRM"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show(rm.GetString("ActivityNote"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                }
                            }
                            else
                            {
                                MessageBox.Show(rm.GetString("BezugNote"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                            }
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
                        {
                            if ((cbMSCRMEntity_4.SelectedItem != null && CheckSelectedEntity(_OrganizationService, 4, false)) || cbMSCRMEntity_4.Text.Trim() == "")
                            {
                                if (cbTaetigkeit_4.Text.Trim() != "")
                                {
                                    if (cbAusgefuehrtDurch_4.SelectedItem != null)
                                    {
                                        if (cbAbrechnungsweise_4.SelectedItem != null)
                                        {
                                            _dr["taetigkeit"] = cbTaetigkeit_4.Text.Trim();

                                            if (cbMSCRMEntity_4.SelectedItem != null)
                                            {
                                                _dr["tekz"] = cbBezugEntity_4.Text.Trim().ToUpper();

                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_4.SelectedItem;
                                                _dr["mscrmbezug"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbezugid"] = _li.Tag.ToString();

                                                DataRow _dr2 = null;
                                                String EntityXmlID = "";

                                                if (_dr["tekz"].ToString() == WorkPackageString) { _LinkTmp = _MSCRM_LinkTemplate_AP; EntityXmlID = "WORKPACKAGE"; }
                                                else if (_dr["tekz"].ToString() == IncidentString) { _LinkTmp = _MSCRM_LinkTemplate_ANF; EntityXmlID = "OPPORTUNITY"; }
                                                else if (_dr["tekz"].ToString() == OpportunityString) { _LinkTmp = _MSCRM_LinkTemplate_VC; EntityXmlID = "INCIDENT"; }
                                                //else if (_dr["tekz"].ToString() == "REL") { _LinkTmp = _MSCRM_LinkTemplate_REL; EntityXmlID = "NEW_VERSION"; }
                                                else if (_dr["tekz"].ToString() == CampaignString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "CAMPAIGN"; }
                                                else if (_dr["tekz"].ToString() == ActionItemString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "ACTION ITEM"; }

                                                if ((_dr2 = GetCacheItemByMSCRMID(EntityXmlID, _li.Tag.ToString())) != null)
                                                {
                                                    // Bezuglink
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid"].ToString());
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr2["mscrmparentid"].ToString());

                                                    _dr["mscrmlinkurl"] = _LinkTmp;
                                                    _dr["mscrmlinktext"] = _dr2["description"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    _dr["mscrmlinktext"] = "";
                                                    _dr["mscrmlinkurl"] = "";
                                                }
                                            }
                                            else
                                            {
                                                _dr["mscrmbezug"] = "";
                                                _dr["mscrmbezugid"] = "";
                                                _dr["mscrmlinktext"] = "";
                                                _dr["mscrmlinkurl"] = "";
                                                _dr["tekz"] = "";
                                            }

                                            if (cbMSCRMEntity2_4.SelectedItem != null)
                                            {
                                                _dr["tekz2"] = cbBezugEntity2_4.Text.Trim().ToUpper();

                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_4.SelectedItem;
                                                _dr["mscrmbezug2"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbezugid2"] = _li.Tag.ToString();

                                                DataRow _dr2 = null;
                                                String EntityXmlID = "";

                                                if (_dr["tekz2"].ToString() == WorkPackageString) { _LinkTmp = _MSCRM_LinkTemplate_AP; EntityXmlID = "WORKPACKAGE"; }
                                                else if (_dr["tekz2"].ToString() == IncidentString) { _LinkTmp = _MSCRM_LinkTemplate_ANF; EntityXmlID = "OPPORTUNITY"; }
                                                else if (_dr["tekz2"].ToString() == OpportunityString) { _LinkTmp = _MSCRM_LinkTemplate_VC; EntityXmlID = "INCIDENT"; }
                                                //else if (_dr["tekz2"].ToString() == "REL") { _LinkTmp = _MSCRM_LinkTemplate_REL; EntityXmlID = "NEW_VERSION"; }
                                                else if (_dr["tekz2"].ToString() == CampaignString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "CAMPAIGN"; }
                                                else if (_dr["tekz2"].ToString() == ActionItemString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "ACTIONITEM"; }

                                                if ((_dr2 = GetCacheItemByMSCRMID(EntityXmlID, _li.Tag.ToString())) != null)
                                                {
                                                    // Bezuglink
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid2"].ToString());
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr2["mscrmparentid"].ToString());

                                                    _dr["mscrmlinkurl2"] = _LinkTmp;
                                                    _dr["mscrmlinktext2"] = _dr2["description"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    _dr["mscrmlinktext2"] = "";
                                                    _dr["mscrmlinkurl2"] = "";
                                                }
                                            }
                                            else
                                            {
                                                _dr["mscrmbezug2"] = "";
                                                _dr["mscrmbezugid2"] = "";
                                                _dr["mscrmlinktext2"] = "";
                                                _dr["mscrmlinkurl2"] = "";
                                                _dr["tekz2"] = "";
                                            }

                                            _dr["datum"] = deDatum_4.DateTime.ToShortDateString();
                                            _SetDateTime = deDatum_4.DateTime;

                                            try
                                            {
                                                if (tbBeginn_4.Text.Trim() != "")
                                                {
                                                    _tmp_time = Convert.ToDateTime(tbBeginn_4.EditValue);
                                                    _dr["beginn"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                                }
                                                else
                                                {
                                                    _dr["beginn"] = "00:00";
                                                }
                                            }
                                            catch { }

                                            try
                                            {
                                                _tmp_time = Convert.ToDateTime(tbEnde_4.EditValue);
                                                _dr["ende"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                            }
                                            catch { }

                                            _dr["dauer"] = lblDauer_4.Content;

                                            _tmp_time = Convert.ToDateTime(tbZusatzzeit_4.EditValue);
                                            _dr["zusatzzeit"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                            _dr["gesamtdauer"] = lblGesamtDauer_4.Content.ToString().Replace(",", ".");

                                            _dr["beschreibung"] = teBeschreibung_4.Text.Trim();
                                            _dr["anmerkung"] = teAnmerkung_4.Text.Trim();
                                            _dr["fahrtkosten"] = ceFahrtkosten_4.IsChecked.ToString();

                                            if (cbStatus_4.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbStatus_4.SelectedItem;
                                                _dr["mscrmstatus"] = _li.Content.ToString().Trim();
                                                _dr["mscrmstatusid"] = _li.Tag;
                                            }

                                            if (cbAusgefuehrtDurch_4.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbAusgefuehrtDurch_4.SelectedItem;
                                                _dr["mscrmuser"] = _li.Content.ToString().Trim();
                                                _dr["mscrmuserid"] = _li.Tag;
                                            }

                                            if (cbAbrechnungsweise_4.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbAbrechnungsweise_4.SelectedItem;
                                                _dr["mscrmbillingtype"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbillingtypeid"] = _li.Tag.ToString();
                                            }

                                            _dr["syncfrommscrm"] = "false";
                                            _dr["synced"] = "no";

                                            _dr.AcceptChanges();
                                            _save_success = true;

                                        }
                                        else
                                        {
                                            MessageBox.Show(rm.GetString("BillingType"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show(rm.GetString("DynamicsCRM"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show(rm.GetString("ActivityNote"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                }
                            }
                            else
                            {
                                MessageBox.Show(rm.GetString("BezugNote"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                            }
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
                        {
                            if ((cbMSCRMEntity_5.SelectedItem != null && CheckSelectedEntity(_OrganizationService, 5, false)) || cbMSCRMEntity_5.Text.Trim() == "")
                            {
                                if (cbTaetigkeit_5.Text.Trim() != "")
                                {
                                    if (cbAusgefuehrtDurch_5.SelectedItem != null)
                                    {
                                        if (cbAbrechnungsweise_5.SelectedItem != null)
                                        {
                                            _dr["taetigkeit"] = cbTaetigkeit_5.Text.Trim();

                                            if (cbMSCRMEntity_5.SelectedItem != null)
                                            {
                                                _dr["tekz"] = cbBezugEntity_5.Text.Trim().ToUpper();

                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_5.SelectedItem;
                                                _dr["mscrmbezug"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbezugid"] = _li.Tag.ToString();

                                                DataRow _dr2 = null;
                                                String EntityXmlID = "";

                                                if (_dr["tekz"].ToString() == WorkPackageString) { _LinkTmp = _MSCRM_LinkTemplate_AP; EntityXmlID = "WORKPACKAGE"; }
                                                else if (_dr["tekz"].ToString() == IncidentString) { _LinkTmp = _MSCRM_LinkTemplate_ANF; EntityXmlID = "OPPORTUNITY"; }
                                                else if (_dr["tekz"].ToString() == OpportunityString) { _LinkTmp = _MSCRM_LinkTemplate_VC; EntityXmlID = "INCIDENT"; }
                                                //else if (_dr["tekz"].ToString() == "REL") { _LinkTmp = _MSCRM_LinkTemplate_REL; EntityXmlID = "NEW_VERSION"; }
                                                else if (_dr["tekz"].ToString() == CampaignString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "CAMPAIGN"; }
                                                else if (_dr["tekz"].ToString() == ActionItemString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "ACTIONITEM"; }

                                                if ((_dr2 = GetCacheItemByMSCRMID(EntityXmlID, _li.Tag.ToString())) != null)
                                                {
                                                    // Bezuglink
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid"].ToString());
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr2["mscrmparentid"].ToString());

                                                    _dr["mscrmlinkurl"] = _LinkTmp;
                                                    _dr["mscrmlinktext"] = _dr2["description"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    _dr["mscrmlinktext"] = "";
                                                    _dr["mscrmlinkurl"] = "";
                                                }
                                            }
                                            else
                                            {
                                                _dr["mscrmbezug"] = "";
                                                _dr["mscrmbezugid"] = "";
                                                _dr["mscrmlinktext"] = "";
                                                _dr["mscrmlinkurl"] = "";
                                                _dr["tekz"] = "";
                                            }

                                            if (cbMSCRMEntity2_5.SelectedItem != null)
                                            {
                                                _dr["tekz2"] = cbBezugEntity2_5.Text.Trim().ToUpper();

                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_5.SelectedItem;
                                                _dr["mscrmbezug2"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbezugid2"] = _li.Tag.ToString();

                                                DataRow _dr2 = null;
                                                String EntityXmlID = "";

                                                if (_dr["tekz2"].ToString() == WorkPackageString) { _LinkTmp = _MSCRM_LinkTemplate_AP; EntityXmlID = "WORKPACKAGE"; }
                                                else if (_dr["tekz2"].ToString() == IncidentString) { _LinkTmp = _MSCRM_LinkTemplate_ANF; EntityXmlID = "OPPORTUNITY"; }
                                                else if (_dr["tekz2"].ToString() == OpportunityString) { _LinkTmp = _MSCRM_LinkTemplate_VC; EntityXmlID = "INCIDENT"; }
                                                //else if (_dr["tekz2"].ToString() == "REL") { _LinkTmp = _MSCRM_LinkTemplate_REL; EntityXmlID = "NEW_VERSION"; }
                                                else if (_dr["tekz2"].ToString() == CampaignString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "CAMPAIGN"; }
                                                else if (_dr["tekz2"].ToString() == ActionItemString) { _LinkTmp = _MSCRM_LinkTemplate_KA; EntityXmlID = "ACTIONITEM"; }

                                                if ((_dr2 = GetCacheItemByMSCRMID(EntityXmlID, _li.Tag.ToString())) != null)
                                                {
                                                    // Bezuglink
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMID}", _dr["mscrmbezugid2"].ToString());
                                                    _LinkTmp = _LinkTmp.Replace("{MSCRMPARENTID}", _dr2["mscrmparentid"].ToString());

                                                    _dr["mscrmlinkurl2"] = _LinkTmp;
                                                    _dr["mscrmlinktext2"] = _dr2["description"].ToString().Trim();
                                                }
                                                else
                                                {
                                                    _dr["mscrmlinktext2"] = "";
                                                    _dr["mscrmlinkurl2"] = "";
                                                }
                                            }
                                            else
                                            {
                                                _dr["mscrmbezug2"] = "";
                                                _dr["mscrmbezugid2"] = "";
                                                _dr["mscrmlinktext2"] = "";
                                                _dr["mscrmlinkurl2"] = "";
                                                _dr["tekz2"] = "";
                                            }

                                            _dr["datum"] = deDatum_5.DateTime.ToShortDateString();
                                            _SetDateTime = deDatum_5.DateTime;

                                            try
                                            {
                                                if (tbBeginn_5.Text.Trim() != "")
                                                {
                                                    _tmp_time = Convert.ToDateTime(tbBeginn_5.EditValue);
                                                    _dr["beginn"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                                }
                                                else
                                                {
                                                    _dr["beginn"] = "00:00";
                                                }
                                            }
                                            catch { }

                                            try
                                            {
                                                _tmp_time = Convert.ToDateTime(tbEnde_5.EditValue);
                                                _dr["ende"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                            }
                                            catch { }

                                            _dr["dauer"] = lblDauer_5.Content;

                                            _tmp_time = Convert.ToDateTime(tbZusatzzeit_5.EditValue);
                                            _dr["zusatzzeit"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                            _dr["gesamtdauer"] = lblGesamtDauer_5.Content.ToString().Replace(",", ".");

                                            _dr["beschreibung"] = teBeschreibung_5.Text.Trim();
                                            _dr["anmerkung"] = teAnmerkung_5.Text.Trim();
                                            _dr["fahrtkosten"] = ceFahrtkosten_5.IsChecked.ToString();

                                            if (cbStatus_5.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbStatus_5.SelectedItem;
                                                _dr["mscrmstatus"] = _li.Content.ToString().Trim();
                                                _dr["mscrmstatusid"] = _li.Tag;
                                            }

                                            if (cbAusgefuehrtDurch_5.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbAusgefuehrtDurch_5.SelectedItem;
                                                _dr["mscrmuser"] = _li.Content.ToString().Trim();
                                                _dr["mscrmuserid"] = _li.Tag;
                                            }

                                            if (cbAbrechnungsweise_5.SelectedItem != null)
                                            {
                                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbAbrechnungsweise_5.SelectedItem;
                                                _dr["mscrmbillingtype"] = _li.Content.ToString().Trim();
                                                _dr["mscrmbillingtypeid"] = _li.Tag.ToString();
                                            }

                                            _dr["syncfrommscrm"] = "false";
                                            _dr["synced"] = "no";

                                            _dr.AcceptChanges();
                                            _save_success = true;

                                        }
                                        else
                                        {
                                            MessageBox.Show(rm.GetString("BillingType"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show(rm.GetString("DynamicsCRM"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show(rm.GetString("ActivityNote"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                                }
                            }
                            else
                            {
                                MessageBox.Show(rm.GetString("BezugNote"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                            }
                        }

                        if (_save_success && DoClose)
                        {
                            CheckAndStopTimer();

                            _ti.Visibility = System.Windows.Visibility.Collapsed;
                            _ti.Tag = "";

                            if (tabTE_1.Tag.ToString() != "") TimeEntryTabCtrl.SelectedItem = tabTE_1;
                            else if (tabTE_2.Tag.ToString() != "") TimeEntryTabCtrl.SelectedItem = tabTE_2;
                            else if (tabTE_3.Tag.ToString() != "") TimeEntryTabCtrl.SelectedItem = tabTE_3;
                            else if (tabTE_4.Tag.ToString() != "") TimeEntryTabCtrl.SelectedItem = tabTE_4;
                            else if (tabTE_5.Tag.ToString() != "") TimeEntryTabCtrl.SelectedItem = tabTE_5;

                            if (tabTE_1.Visibility == System.Windows.Visibility.Collapsed && tabTE_2.Visibility == System.Windows.Visibility.Collapsed && tabTE_3.Visibility == System.Windows.Visibility.Collapsed && tabTE_4.Visibility == System.Windows.Visibility.Collapsed && tabTE_5.Visibility == System.Windows.Visibility.Collapsed)
                            {
                                TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Collapsed;
                            }
                        }

                        // Speichern und neu
                        if (_save_success && DoOpenNew)
                        {
                            ActivateTimeEntry(Guid.Empty.ToString(), false, _SetDateTime, "", "");
                        }
                    }

                    WriteGrid();
                }
            }
            catch (SoapException ex)
            {
                FileLog("SaveActiveTimeEntry", "SOAP Exception.Message: " + ex.Message);
                FileLog("SaveActiveTimeEntry", "SOAP Exception.Detail: " + ex.Detail.InnerText);
            }
            catch (Exception ex)
            {
                FileLog("SaveActiveTimeEntry", "Exception: " + ex.Message);
            }
        }

        private void CancelTimeEntry(TabItem _TimeEntry)
        {
            try
            {
                // Cancel action
                if (_TimeEntry != null)
                {
                    TabItem _ti = _TimeEntry;
                    _ti.Visibility = System.Windows.Visibility.Collapsed;

                    CheckAndStopTimer();

                    DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
                    DataRow _dr = _dt.Rows.Find(_ti.Tag);
                    if (_dr != null)
                    {
                        if (_dr["synced"].ToString().Trim().ToLower() == "*new*") _dt.Rows.Remove(_dr);
                    }

                    _ti.Tag = "";

                    if (tabTE_1.Tag.ToString() != "") TimeEntryTabCtrl.SelectedItem = tabTE_1;
                    else if (tabTE_2.Tag.ToString() != "") TimeEntryTabCtrl.SelectedItem = tabTE_2;
                    else if (tabTE_3.Tag.ToString() != "") TimeEntryTabCtrl.SelectedItem = tabTE_3;
                    else if (tabTE_4.Tag.ToString() != "") TimeEntryTabCtrl.SelectedItem = tabTE_4;
                    else if (tabTE_5.Tag.ToString() != "") TimeEntryTabCtrl.SelectedItem = tabTE_5;

                    if (tabTE_1.Visibility == System.Windows.Visibility.Collapsed && tabTE_2.Visibility == System.Windows.Visibility.Collapsed && tabTE_3.Visibility == System.Windows.Visibility.Collapsed && tabTE_4.Visibility == System.Windows.Visibility.Collapsed && tabTE_5.Visibility == System.Windows.Visibility.Collapsed)
                    {
                        TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Collapsed;
                    }

                }
            }
            catch (Exception ex)
            {
                FileLog("CancelTimeEntry", "Exception: " + ex.Message);
            }
        }

        private void btnCancelTimeentry_Click(object sender, RoutedEventArgs e)
        {
            if (TimeEntryTabCtrl.SelectedItem != null)
            {
                if (MessageBox.Show(rm.GetString("CancelTimeentry"), rm.GetString("CancelTimeentryTitle"), MessageBoxButton.YesNoCancel, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    // Cancel action
                    CancelTimeEntry((TabItem)TimeEntryTabCtrl.SelectedItem);
                }
            }
        }

        private void TimeEntryGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // activate selected time entry ...

            try
            {
                String _te_guid = TimeEntryGrid.GetFocusedRowCellValue("id").ToString().Trim();
                if (_te_guid != "") ActivateTimeEntry(_te_guid, false, DateTime.Now, "", "");
            }
            catch (Exception ex)
            {
                FileLog("TimeEntryGrid_MouseDoubleClick", "Error: " + ex.Message);
                // MessageBox.Show("TimeEntryGrid_MouseDoubleClick: " + ex.Message, "Fehlermeldung", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CalcTimeValues(DataRow _dr, Int32 UpdateTE)
        {
            try
            {

                DateTime _begin = Convert.ToDateTime(_dr["datum"].ToString() + " " + _dr["beginn"].ToString());
                DateTime _add_time = Convert.ToDateTime(_dr["zusatzzeit"].ToString() + ":00");

                DateTime _start = _begin;
                DateTime _start_dauer = _begin;

                String _te_timemeasurements = _dr["timemeasurements"].ToString().Trim();
                String[] _te_measurements = _te_timemeasurements.Split(';');
                if (_te_timemeasurements != "") foreach (String _item in _te_measurements)
                    {
                        String[] _te_parts = _item.Split('|');
                        _start = _start.AddMilliseconds(Convert.ToDouble(_te_parts[2]));
                    }

                // Dauer, gemessen
                TimeSpan ts = new TimeSpan();
                ts = _start.Subtract(_begin);

                _dr["dauer"] = Convert.ToDecimal(ts.TotalHours).ToString("0.00");


                // + Zusatzzeit
                _start = _start.AddHours(_add_time.Hour);
                _start = _start.AddMinutes(_add_time.Minute);

                ts = new TimeSpan();
                ts = _start.Subtract(_begin);

                _dr["ende"] = _start.Hour.ToString("00") + ":" + _start.Minute.ToString("00");


                // The round number, here is a quarter...
                int Round = 15;

                // Count of round number in this total minutes...
                double CountRound = (ts.TotalMinutes / Round);

                double test = CountRound - (int)Math.Truncate(CountRound);

                // The main formula to calculate round time...
                int Min = (int)Math.Truncate(CountRound) * Round;
                if (test > 0.0) Min += Round;

                Double _delay = (Min / 15) * 0.25;

                _dr["gesamtdauer"] = _delay.ToString("0.00").Replace(",", ".");

                if (UpdateTE == 1)
                {
                    lblDauer_1.Content = _dr["dauer"].ToString();
                    tbEnde_1.Text = _dr["ende"].ToString();
                    lblGesamtDauer_1.Content = _dr["gesamtdauer"].ToString().Replace(",","."); 
                }
                else if (UpdateTE == 2)
                {
                    lblDauer_2.Content = _dr["dauer"].ToString();
                    tbEnde_2.Text = _dr["ende"].ToString();
                    lblGesamtDauer_2.Content = _dr["gesamtdauer"].ToString().Replace(",", ".");
                }
                else if (UpdateTE == 3)
                {
                    lblDauer_3.Content = _dr["dauer"].ToString();
                    tbEnde_3.Text = _dr["ende"].ToString();
                    lblGesamtDauer_3.Content = _dr["gesamtdauer"].ToString().Replace(",", ".");
                }
                else if (UpdateTE == 4)
                {
                    lblDauer_4.Content = _dr["dauer"].ToString();
                    tbEnde_4.Text = _dr["ende"].ToString();
                    lblGesamtDauer_4.Content = _dr["gesamtdauer"].ToString().Replace(",", ".");
                }
                else if (UpdateTE == 5)
                {
                    lblDauer_5.Content = _dr["dauer"].ToString();
                    tbEnde_5.Text = _dr["ende"].ToString();
                    lblGesamtDauer_5.Content = _dr["gesamtdauer"].ToString().Replace(",", ".");
                }
            }
            catch (Exception ex)
            {
                FileLog("CalcTimeValues", "Error: " + ex.Message);
            }
        }

        private void FillEntityList(String TypeOfEntity, Int32 BezugNo, Int32 TabNo)
        {
            if (TypeOfEntity != "")
            {
                String _TOE_Name = "";

                if (TypeOfEntity == WorkPackageString) _TOE_Name = "WORKPACKAGE";
                else if (TypeOfEntity == OpportunityString) _TOE_Name = "OPPORTUNITY";
                else if (TypeOfEntity == IncidentString) _TOE_Name = "INCIDENT";
                //else if (TypeOfEntity == "REL") _TOE_Name = "NEW_VERSION";
                else if (TypeOfEntity == CampaignString) _TOE_Name = "CAMPAIGN";
                else if (TypeOfEntity == ActionItemString) _TOE_Name = "ACTIONITEM";

                if (BezugNo == 1)
                {
                    switch (TabNo)
                    {
                        case 1:
                            // Fill DropDown
                            DataRow[] _result1 = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                            if (_result1 != null)
                            {
                                cbMSCRMEntity_1.Items.Clear();
                                foreach (DataRow _row in _result1)
                                {
                                    DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                    _li.Content = _row["description"].ToString();
                                    _li.Tag = _row["mscrmid"].ToString();
                                    cbMSCRMEntity_1.Items.Add(_li);
                                }
                            }
                            break;
                        case 2:
                            // Fill DropDown
                            DataRow[] _result2 = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                            if (_result2 != null)
                            {
                                cbMSCRMEntity_2.Items.Clear();
                                foreach (DataRow _row in _result2)
                                {
                                    DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                    _li.Content = _row["description"].ToString();
                                    _li.Tag = _row["mscrmid"].ToString();
                                    cbMSCRMEntity_2.Items.Add(_li);
                                }
                            }
                            break;
                        case 3:
                            // Fill DropDown
                            DataRow[] _result3 = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                            if (_result3 != null)
                            {
                                cbMSCRMEntity_3.Items.Clear();
                                foreach (DataRow _row in _result3)
                                {
                                    DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                    _li.Content = _row["description"].ToString();
                                    _li.Tag = _row["mscrmid"].ToString();
                                    cbMSCRMEntity_3.Items.Add(_li);
                                }
                            }
                            break;
                        case 4:
                            // Fill DropDown
                            DataRow[] _result4 = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                            if (_result4 != null)
                            {
                                cbMSCRMEntity_4.Items.Clear();
                                foreach (DataRow _row in _result4)
                                {
                                    DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                    _li.Content = _row["description"].ToString();
                                    _li.Tag = _row["mscrmid"].ToString();
                                    cbMSCRMEntity_4.Items.Add(_li);
                                }
                            }
                            break;
                        case 5:
                            // Fill DropDown
                            DataRow[] _result5 = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                            if (_result5 != null)
                            {
                                cbMSCRMEntity_5.Items.Clear();
                                foreach (DataRow _row in _result5)
                                {
                                    DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                    _li.Content = _row["description"].ToString();
                                    _li.Tag = _row["mscrmid"].ToString();
                                    cbMSCRMEntity_5.Items.Add(_li);
                                }
                            }
                            break;
                    }
                }
                else
                {
                    switch (TabNo)
                    {
                        case 1:
                            // Fill DropDown
                            DataRow[] _result1 = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                            if (_result1 != null)
                            {
                                cbMSCRMEntity2_1.Items.Clear();
                                foreach (DataRow _row in _result1)
                                {
                                    DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                    _li.Content = _row["description"].ToString();
                                    _li.Tag = _row["mscrmid"].ToString();
                                    cbMSCRMEntity2_1.Items.Add(_li);
                                }
                            }
                            break;
                        case 2:
                            // Fill DropDown
                            DataRow[] _result2 = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                            if (_result2 != null)
                            {
                                cbMSCRMEntity2_2.Items.Clear();
                                foreach (DataRow _row in _result2)
                                {
                                    DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                    _li.Content = _row["description"].ToString();
                                    _li.Tag = _row["mscrmid"].ToString();
                                    cbMSCRMEntity2_2.Items.Add(_li);
                                }
                            }
                            break;
                        case 3:
                            // Fill DropDown
                            DataRow[] _result3 = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                            if (_result3 != null)
                            {
                                cbMSCRMEntity2_3.Items.Clear();
                                foreach (DataRow _row in _result3)
                                {
                                    DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                    _li.Content = _row["description"].ToString();
                                    _li.Tag = _row["mscrmid"].ToString();
                                    cbMSCRMEntity2_3.Items.Add(_li);
                                }
                            }
                            break;
                        case 4:
                            // Fill DropDown
                            DataRow[] _result4 = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                            if (_result4 != null)
                            {
                                cbMSCRMEntity2_4.Items.Clear();
                                foreach (DataRow _row in _result4)
                                {
                                    DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                    _li.Content = _row["description"].ToString();
                                    _li.Tag = _row["mscrmid"].ToString();
                                    cbMSCRMEntity2_4.Items.Add(_li);
                                }
                            }
                            break;
                        case 5:
                            // Fill DropDown
                            DataRow[] _result5 = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                            if (_result5 != null)
                            {
                                cbMSCRMEntity2_5.Items.Clear();
                                foreach (DataRow _row in _result5)
                                {
                                    DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                    _li.Content = _row["description"].ToString();
                                    _li.Tag = _row["mscrmid"].ToString();
                                    cbMSCRMEntity2_5.Items.Add(_li);
                                }
                            }
                            break;
                    }
                }
            }
        }

        private Boolean ActivateTimeEntry(String _te_guid, Boolean DoCopyTimeEntry, DateTime _SetDateTime, String _Dauer, String _Zusatzeit)
        {
            Boolean _rV = false;

            _ActivateTimeEntryActive = true;

            try
            {
                DevExpress.Xpf.Editors.ComboBoxEditItem _selected_li = null;
                DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
                Boolean _te_is_active = false;

                // Test, if time entry is already activated
                if (tabTE_1.Tag.ToString() == _te_guid)
                {
                    _te_is_active = true;
                    TimeEntryTabCtrl.SelectedItem = tabTE_1;
                    if (TimeEntryTabCtrl.Visibility == System.Windows.Visibility.Collapsed) TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Visible;
                }
                else if (tabTE_2.Tag.ToString() == _te_guid)
                {
                    _te_is_active = true;
                    TimeEntryTabCtrl.SelectedItem = tabTE_2;
                    if (TimeEntryTabCtrl.Visibility == System.Windows.Visibility.Collapsed) TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Visible;
                }
                else if (tabTE_3.Tag.ToString() == _te_guid)
                {
                    _te_is_active = true;
                    TimeEntryTabCtrl.SelectedItem = tabTE_3;
                    if (TimeEntryTabCtrl.Visibility == System.Windows.Visibility.Collapsed) TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Visible;
                }
                else if (tabTE_4.Tag.ToString() == _te_guid)
                {
                    _te_is_active = true;
                    TimeEntryTabCtrl.SelectedItem = tabTE_4;
                    if (TimeEntryTabCtrl.Visibility == System.Windows.Visibility.Collapsed) TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Visible;
                }
                else if (tabTE_5.Tag.ToString() == _te_guid)
                {
                    _te_is_active = true;
                    TimeEntryTabCtrl.SelectedItem = tabTE_5;
                    if (TimeEntryTabCtrl.Visibility == System.Windows.Visibility.Collapsed) TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Visible;
                }


                // Try to activate free tab for selected time entry
                if (!_te_is_active)
                {
                    String _save_te_guid = _te_guid; // Zu kopierende GUID merken

                    String _new_te_guid = "";

                    if (DoCopyTimeEntry) _te_guid = Guid.Empty.ToString(); // GUID = Empty -> Create New

                    if (_te_guid.ToString() == Guid.Empty.ToString())
                    {
                        if (tabTE_1.Visibility == System.Windows.Visibility.Collapsed || tabTE_2.Visibility == System.Windows.Visibility.Collapsed || tabTE_3.Visibility == System.Windows.Visibility.Collapsed || tabTE_4.Visibility == System.Windows.Visibility.Collapsed || tabTE_5.Visibility == System.Windows.Visibility.Collapsed)
                        {
                            // New Timeentry 
                            _te_guid = Guid.NewGuid().ToString();

                            // Add new data row ...
                            DataRow _dr = _dt.NewRow();

                            if (gDefaultUser == Guid.Empty)
                            {
                                object[] UserDtls = CRMManager.GetCurrentUser(_OrganizationService, _TimeEntryCache, !(MSCRM_ConState == MSCRM_ConectionStates.Connected), _currentConnectionDetail);
                                gDefaultUser = (Guid)UserDtls[0];
                                sDefaultUser = (string)UserDtls[1];
                            }

                            _dr["id"] = _te_guid.ToString();
                            _dr["marked"] = false;
                            _dr["tekz"] = "";
                            _dr["tekz2"] = "";

                            DateTime _DateTime = _SetDateTime;

                            _dr["datum"] = _DateTime.ToShortDateString();

                            _dr["beginn"] = _DateTime.Hour.ToString("00") + ":" + DateTime.Now.Minute.ToString("00");
                            _dr["ende"] = _DateTime.Hour.ToString("00") + ":" + DateTime.Now.Minute.ToString("00");

                            _dr["dauer"] = "0,00";
                            _dr["zusatzzeit"] = "00:00";
                            _dr["gesamtdauer"] = "0,00";
                            _dr["taetigkeit"] = "";
                            _dr["taetigkeiturl"] = "";
                            _dr["synced"] = "*new*";
                            _dr["mscrmentityid"] = "";
                            _dr["mscrmstatus"] = "Erfasst";
                            _dr["mscrmstatusid"] = "";
                            _dr["fahrtkosten"] = false.ToString();
                            _dr["timemeasurements"] = "";
                            _dr["beschreibung"] = "";
                            _dr["anmerkung"] = "";
                            _dr["mscrmbezug"] = "";
                            _dr["mscrmbezugid"] = "";
                            _dr["mscrmbezug2"] = "";
                            _dr["mscrmbezugid2"] = "";
                            _dr["mscrmuser"] = sDefaultUser;
                            _dr["mscrmuserid"] = gDefaultUser;
                            _dr["mscrmbillingtype"] = "";
                            _dr["mscrmbillingtypeid"] = "";
                            _dr["mscrmmodifiedon"] = "";
                            _dr["lasterror"] = "";
                            _dr["mscrmlinkurl"] = "";
                            _dr["mscrmlinktext"] = "";
                            _dr["mscrmlinkurl2"] = "";
                            _dr["mscrmlinktext2"] = "";
                            _dr["syncfrommscrm"] = "";

                            _dt.Rows.Add(_dr);
                            _dt.AcceptChanges();
                        }
                    }

                    if (DoCopyTimeEntry)
                    {
                        _new_te_guid = _te_guid;
                        _te_guid = _save_te_guid;
                    }

                    if (tabTE_1.Visibility == System.Windows.Visibility.Collapsed)
                    {
                        DataRow _dr = null;

                        _dr = _dt.Rows.Find(_te_guid.ToString());

                        if (_dr != null)
                        {
                            tabTE_1.Visibility = System.Windows.Visibility.Visible;
                            if (!DoCopyTimeEntry) tabTE_1.Tag = _te_guid.ToString(); else tabTE_1.Tag = _new_te_guid.ToString();
                            tabTE_1.Header = rm.GetString("tabTE_1");

                            //btAddNewEntity_1.IsEnabled = false;
                            //btAddNewEntity2_1.IsEnabled = false;

                            cbBezugEntity_1.Items.Clear();

                            foreach (var item in _BezugValues)
                            {
                                cbBezugEntity_1.Items.Add(item);
                            }


                            cbBezugEntity_1.Text = _dr["tekz"].ToString().Trim().ToUpper();
                            cbBezugEntity2_1.Text = _dr["tekz2"].ToString().Trim().ToUpper();

                            cbMSCRMEntity_1.Text = _dr["mscrmbezug"].ToString();
                            cbMSCRMEntity2_1.Text = _dr["mscrmbezug2"].ToString();


                            if (DoCopyTimeEntry)
                            {
                                //var temp = _SetDateTime.Day.ToString("00") + "." + _SetDateTime.Month.ToString("00") + "." + _SetDateTime.Year.ToString() + " 00:00:00";
                                deDatum_1.EditValue = Convert.ToDateTime(_SetDateTime, CultureInfo.CurrentCulture.DateTimeFormat);

                                if (_Dauer != "")
                                {
                                    DataRow _dr_new = null;
                                    if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                    {
                                        String[] _dm = _Dauer.Split(':');
                                        Int32 _dm_hours = 0;
                                        Int32 _dm_minutes = 0;
                                        try
                                        {
                                            _dm_hours = Convert.ToInt32(_dm[0]);
                                            _dm_minutes = Convert.ToInt32(_dm[1]);
                                        }
                                        catch
                                        {
                                        }

                                        if (_dm_hours != 0 || _dm_minutes != 0)
                                        {
                                            String _timemeasurements = "";

                                            _timemeasurements = _SetDateTime.ToShortDateString() + " " + _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00") + ":00|";

                                            _dr_new["beginn"] = _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00");

                                            _SetDateTime = _SetDateTime.AddHours(_dm_hours);
                                            _SetDateTime = _SetDateTime.AddMinutes(_dm_minutes);

                                            _dr_new["ende"] = _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00");

                                            _timemeasurements += _SetDateTime.ToShortDateString() + " " + _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00") + ":00|";

                                            tbBeginn_1.Text = _dr_new["beginn"].ToString();
                                            tbEnde_1.Text = _dr_new["ende"].ToString();

                                            // new timemeasurements
                                            long startTick = Convert.ToDateTime(tbBeginn_1.EditValue).Ticks;
                                            long endTick = Convert.ToDateTime(tbEnde_1.EditValue).Ticks;
                                            long tick = endTick - startTick;
                                            long seconds = tick / TimeSpan.TicksPerSecond;
                                            long milliseconds = tick / TimeSpan.TicksPerMillisecond;

                                            _timemeasurements += milliseconds.ToString();

                                            _dr_new["timemeasurements"] = _timemeasurements;

                                            _dr_new.AcceptChanges();
                                        }
                                    }
                                }
                                else
                                {
                                    tbBeginn_1.Text = _dr["beginn"].ToString();
                                    tbEnde_1.Text = _dr["ende"].ToString();
                                }
                                if (_Zusatzeit != "")
                                {
                                    DataRow _dr_new = null;
                                    if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                    {
                                        _dr_new["zusatzzeit"] = _Zusatzeit;
                                        tbZusatzzeit_1.Text = _Zusatzeit;

                                        _dr_new.AcceptChanges();
                                    }
                                }
                                else tbZusatzzeit_1.Text = _dr["zusatzzeit"].ToString();
                            }
                            else
                            {
                                deDatum_1.EditValue = Convert.ToDateTime(_dr["datum"].ToString() + " 00:00:00");
                                lblDauer_1.Content = _dr["dauer"].ToString();
                                tbBeginn_1.Text = _dr["beginn"].ToString();
                                tbEnde_1.Text = _dr["ende"].ToString();
                                tbZusatzzeit_1.Text = _dr["zusatzzeit"].ToString();
                            }

                            cbTaetigkeit_1.Items.Clear();

                            String _tmp_User = _User;
                            _tmp_User = _tmp_User.Replace("\\", "_");
                            _tmp_User = _tmp_User.Replace(".", "_");

                            DataRow[] _result = _TimeEntryCache.Select("[type]='" + "FAVORITE_ACTIVITY_" + _tmp_User + "'");
                            if (_result.Length > 0)
                            {
                                foreach (DataRow _current_dr in _result)
                                {
                                    cbTaetigkeit_1.Items.Add(_current_dr["description"].ToString());
                                }
                            }
                            cbTaetigkeit_1.Text = _dr["taetigkeit"].ToString();

                            cbAusgefuehrtDurch_1.Items.Clear();
                            cbAusgefuehrtDurch_1.Text = "";
                            _result = _TimeEntryCache.Select("[type]='USER.LIST'");
                            _selected_li = null;

                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmuser"].ToString()) _selected_li = _li;
                                cbAusgefuehrtDurch_1.Items.Add(_li);
                            }

                            if (_selected_li != null) cbAusgefuehrtDurch_1.SelectedItem = _selected_li;

                            cbAbrechnungsweise_1.Items.Clear();
                            cbAbrechnungsweise_1.Text = "";
                            cbAbrechnungsweise_1.Tag = "";
                            _result = _TimeEntryCache.Select("[type]='PICKLIST.BILLING'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmbillingtype"].ToString()) _selected_li = _li;
                                cbAbrechnungsweise_1.Items.Add(_li);
                            }
                            if (_selected_li != null) cbAbrechnungsweise_1.SelectedItem = _selected_li;

                            // STATUSLIST.STATUSCODE
                            cbStatus_1.Items.Clear();
                            cbStatus_1.Text = "";
                            _result = _TimeEntryCache.Select("[type]='STATUSLIST.STATUSCODE'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmstatus"].ToString()) _selected_li = _li;
                                cbStatus_1.Items.Add(_li);
                            }
                            if (_selected_li != null) cbStatus_1.SelectedItem = _selected_li;

                            teBeschreibung_1.Text = _dr["beschreibung"].ToString();
                            teAnmerkung_1.Text = _dr["anmerkung"].ToString();
                            ceFahrtkosten_1.IsChecked = Convert.ToBoolean(_dr["fahrtkosten"].ToString());

                            // Calc Time Values
                            if (!DoCopyTimeEntry)
                            {
                                CalcTimeValues(_dr, 1);
                            }
                            else
                            {
                                DataRow _dr_new = null;
                                if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                {
                                    CalcTimeValues(_dr_new, 1);

                                    _dr_new.AcceptChanges();
                                }
                            }

                            btnTimerCtrl_1.IsEnabled = true;
                            lblTimerCtrlInfo_1.Content = "00:00:00";
                            btnShowQuickTimeentryList_1.IsEnabled = true;

                            if (TimeEntryTabCtrl.Visibility == System.Windows.Visibility.Collapsed) TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Visible;
                            TimeEntryTabCtrl.SelectedItem = tabTE_1;

                            CheckSelectedEntity(_OrganizationService, 1, false);
                            CheckSelectedEntity2(_OrganizationService, 1, false);

                            if (cbBezugEntity_1.Text.Trim() == "")
                            {
                                cbBezugEntity2_1.Items.Clear();

                                cbBezugEntity2_1.IsEnabled = false;
                                cbMSCRMEntity2_1.IsEnabled = false;
                                btnMSCRMEntityInfoImage2_1.IsEnabled = false;
                                btnExtendedSearch2_1.IsEnabled = false;
                                btnFavorites2_1.IsEnabled = false;
                                btnAddToFavorites2_1.IsEnabled = false;
                            }
                            else
                            {
                                cbBezugEntity2_1.Items.Clear();

                                foreach (var item in _BezugValues)
                                {
                                    if (cbBezugEntity_1.Text.Trim().ToUpper() != item) cbBezugEntity2_1.Items.Add(item);
                                }

                                cbBezugEntity2_1.IsEnabled = true;
                                cbMSCRMEntity2_1.IsEnabled = true;
                                btnMSCRMEntityInfoImage2_1.IsEnabled = true;
                                btnExtendedSearch2_1.IsEnabled = true;
                                btnFavorites2_1.IsEnabled = true;
                                btnAddToFavorites2_1.IsEnabled = true;
                            }

                            cbTaetigkeit_1.Focus();
                        }
                    }
                    else if (tabTE_2.Visibility == System.Windows.Visibility.Collapsed)
                    {
                        DataRow _dr = _dt.Rows.Find(_te_guid.ToString());
                        if (_dr != null)
                        {
                            tabTE_2.Visibility = System.Windows.Visibility.Visible;
                            if (!DoCopyTimeEntry) tabTE_2.Tag = _te_guid.ToString(); else tabTE_2.Tag = _new_te_guid.ToString();
                            tabTE_2.Header = rm.GetString("tabTE_2");

                            //btAddNewEntity_2.IsEnabled = false;
                            //btAddNewEntity2_2.IsEnabled = false;

                            cbBezugEntity_2.Items.Clear();

                            foreach (var item in _BezugValues)
                            {
                                cbBezugEntity_2.Items.Add(item);
                            }

                            cbBezugEntity_2.Text = _dr["tekz"].ToString().Trim().ToUpper();
                            cbBezugEntity2_2.Text = _dr["tekz2"].ToString().Trim().ToUpper();
                            cbMSCRMEntity_2.Text = _dr["mscrmbezug"].ToString();
                            cbMSCRMEntity2_2.Text = _dr["mscrmbezug2"].ToString();

                            if (DoCopyTimeEntry)
                            {
                                deDatum_2.EditValue = Convert.ToDateTime(_SetDateTime, CultureInfo.CurrentCulture.DateTimeFormat);

                                if (_Dauer != "")
                                {
                                    DataRow _dr_new = null;
                                    if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                    {
                                        String[] _dm = _Dauer.Split(':');
                                        Int32 _dm_hours = 0;
                                        Int32 _dm_minutes = 0;
                                        try
                                        {
                                            _dm_hours = Convert.ToInt32(_dm[0]);
                                            _dm_minutes = Convert.ToInt32(_dm[1]);
                                        }
                                        catch
                                        {
                                        }

                                        if (_dm_hours != 0 || _dm_minutes != 0)
                                        {
                                            String _timemeasurements = "";

                                            _timemeasurements = _SetDateTime.ToShortDateString() + " " + _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00") + ":00|";

                                            _dr_new["beginn"] = _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00");

                                            _SetDateTime = _SetDateTime.AddHours(_dm_hours);
                                            _SetDateTime = _SetDateTime.AddMinutes(_dm_minutes);

                                            _dr_new["ende"] = _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00");

                                            _timemeasurements += _SetDateTime.ToShortDateString() + " " + _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00") + ":00|";

                                            tbBeginn_2.Text = _dr_new["beginn"].ToString();
                                            tbEnde_2.Text = _dr_new["ende"].ToString();

                                            // new timemeasurements
                                            long startTick = Convert.ToDateTime(tbBeginn_2.EditValue).Ticks;
                                            long endTick = Convert.ToDateTime(tbEnde_2.EditValue).Ticks;
                                            long tick = endTick - startTick;
                                            long seconds = tick / TimeSpan.TicksPerSecond;
                                            long milliseconds = tick / TimeSpan.TicksPerMillisecond;

                                            _timemeasurements += milliseconds.ToString();

                                            _dr_new["timemeasurements"] = _timemeasurements;

                                            _dr_new.AcceptChanges();
                                        }
                                    }
                                }
                                else
                                {
                                    tbBeginn_2.Text = _dr["beginn"].ToString();
                                    tbEnde_2.Text = _dr["ende"].ToString();
                                }
                                if (_Zusatzeit != "") tbZusatzzeit_2.Text = _Zusatzeit; else tbZusatzzeit_2.Text = _dr["zusatzzeit"].ToString();
                            }
                            else
                            {
                                deDatum_2.EditValue = Convert.ToDateTime(_dr["datum"].ToString() + " 00:00:00");
                                lblDauer_2.Content = _dr["dauer"].ToString();
                                tbBeginn_2.Text = _dr["beginn"].ToString();
                                tbEnde_2.Text = _dr["ende"].ToString();
                                tbZusatzzeit_2.Text = _dr["zusatzzeit"].ToString();
                            }

                            cbTaetigkeit_2.Items.Clear();
                            String _tmp_User = _User;
                            _tmp_User = _tmp_User.Replace("\\", "_");
                            _tmp_User = _tmp_User.Replace(".", "_");

                            DataRow[] _result = _TimeEntryCache.Select("[type]='" + "FAVORITE_ACTIVITY_" + _tmp_User + "'");
                            if (_result.Length > 0)
                            {
                                foreach (DataRow _current_dr in _result)
                                {
                                    cbTaetigkeit_2.Items.Add(_current_dr["description"].ToString());
                                }
                            }
                            cbTaetigkeit_2.Text = _dr["taetigkeit"].ToString();

                            cbAusgefuehrtDurch_2.Items.Clear();
                            cbAusgefuehrtDurch_2.Text = "";
                            _result = _TimeEntryCache.Select("[type]='USER.LIST'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmuser"].ToString()) _selected_li = _li;
                                cbAusgefuehrtDurch_2.Items.Add(_li);
                            }
                            if (_selected_li != null) cbAusgefuehrtDurch_2.SelectedItem = _selected_li;

                            cbAbrechnungsweise_2.Items.Clear();
                            cbAbrechnungsweise_2.Text = "";
                            cbAbrechnungsweise_2.Tag = "";
                            _result = _TimeEntryCache.Select("[type]='PICKLIST.BILLING'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmbillingtype"].ToString()) _selected_li = _li;
                                cbAbrechnungsweise_2.Items.Add(_li);
                            }
                            if (_selected_li != null) cbAbrechnungsweise_2.SelectedItem = _selected_li;

                            // STATUSLIST.STATUSCODE
                            cbStatus_2.Items.Clear();
                            cbStatus_2.Text = "";
                            _result = _TimeEntryCache.Select("[type]='STATUSLIST.STATUSCODE'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmstatus"].ToString()) _selected_li = _li;
                                cbStatus_2.Items.Add(_li);
                            }
                            if (_selected_li != null) cbStatus_2.SelectedItem = _selected_li;

                            teBeschreibung_2.Text = _dr["beschreibung"].ToString();
                            teAnmerkung_2.Text = _dr["anmerkung"].ToString();
                            ceFahrtkosten_2.IsChecked = Convert.ToBoolean(_dr["fahrtkosten"].ToString());

                            // Calc Time Values
                            if (!DoCopyTimeEntry)
                            {
                                CalcTimeValues(_dr, 2);
                            }
                            else
                            {
                                DataRow _dr_new = null;
                                if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                {
                                    CalcTimeValues(_dr_new, 2);

                                    _dr_new.AcceptChanges();
                                }
                            }

                            btnTimerCtrl_2.IsEnabled = true;
                            lblTimerCtrlInfo_2.Content = "00:00:00";
                            btnShowQuickTimeentryList_2.IsEnabled = true;

                            if (TimeEntryTabCtrl.Visibility == System.Windows.Visibility.Collapsed) TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Visible;
                            TimeEntryTabCtrl.SelectedItem = tabTE_2;

                            CheckSelectedEntity(_OrganizationService, 2, false);
                            CheckSelectedEntity2(_OrganizationService, 2, false);

                            if (cbBezugEntity_2.Text.Trim() == "")
                            {
                                cbBezugEntity2_2.Items.Clear();

                                cbBezugEntity2_2.IsEnabled = false;
                                cbMSCRMEntity2_2.IsEnabled = false;
                                btnMSCRMEntityInfoImage2_2.IsEnabled = false;
                                btnExtendedSearch2_2.IsEnabled = false;
                                btnFavorites2_2.IsEnabled = false;
                                btnAddToFavorites2_2.IsEnabled = false;
                            }
                            else
                            {
                                cbBezugEntity2_2.Items.Clear();

                                cbBezugEntity2_2.Items.Add("");

                                foreach (var item in _BezugValues)
                                {
                                    if (cbBezugEntity_1.Text.Trim().ToUpper() != item) cbBezugEntity2_1.Items.Add(item);
                                }



                                cbBezugEntity2_2.IsEnabled = true;
                                cbMSCRMEntity2_2.IsEnabled = true;
                                btnMSCRMEntityInfoImage2_2.IsEnabled = true;
                                btnExtendedSearch2_2.IsEnabled = true;
                                btnFavorites2_2.IsEnabled = true;
                                btnAddToFavorites2_2.IsEnabled = true;
                            }

                        }

                        cbTaetigkeit_2.Focus();
                    }
                    else if (tabTE_3.Visibility == System.Windows.Visibility.Collapsed)
                    {
                        DataRow _dr = _dt.Rows.Find(_te_guid.ToString());
                        if (_dr != null)
                        {
                            tabTE_3.Visibility = System.Windows.Visibility.Visible;
                            if (!DoCopyTimeEntry) tabTE_3.Tag = _te_guid.ToString(); else tabTE_3.Tag = _new_te_guid.ToString();
                            tabTE_3.Header = rm.GetString("tabTE_3");

                            //btAddNewEntity_3.IsEnabled = false;
                            //btAddNewEntity2_3.IsEnabled = false;

                            cbBezugEntity_3.Items.Clear();

                            foreach (var item in _BezugValues)
                            {
                                cbBezugEntity_3.Items.Add(item);
                            }

                            cbBezugEntity_3.Text = _dr["tekz"].ToString().Trim().ToUpper();
                            cbBezugEntity2_3.Text = _dr["tekz2"].ToString().Trim().ToUpper();
                            cbMSCRMEntity_3.Text = _dr["mscrmbezug"].ToString();
                            cbMSCRMEntity2_3.Text = _dr["mscrmbezug2"].ToString();

                            if (DoCopyTimeEntry)
                            {
                                deDatum_3.EditValue = Convert.ToDateTime(_SetDateTime, CultureInfo.CurrentCulture.DateTimeFormat);

                                if (_Dauer != "")
                                {
                                    DataRow _dr_new = null;
                                    if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                    {
                                        String[] _dm = _Dauer.Split(':');
                                        Int32 _dm_hours = 0;
                                        Int32 _dm_minutes = 0;
                                        try
                                        {
                                            _dm_hours = Convert.ToInt32(_dm[0]);
                                            _dm_minutes = Convert.ToInt32(_dm[1]);
                                        }
                                        catch
                                        {
                                        }

                                        if (_dm_hours != 0 || _dm_minutes != 0)
                                        {
                                            String _timemeasurements = "";

                                            _timemeasurements = _SetDateTime.ToShortDateString() + " " + _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00") + ":00|";

                                            _dr_new["beginn"] = _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00");

                                            _SetDateTime = _SetDateTime.AddHours(_dm_hours);
                                            _SetDateTime = _SetDateTime.AddMinutes(_dm_minutes);

                                            _dr_new["ende"] = _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00");

                                            _timemeasurements += _SetDateTime.ToShortDateString() + " " + _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00") + ":00|";

                                            tbBeginn_3.Text = _dr_new["beginn"].ToString();
                                            tbEnde_3.Text = _dr_new["ende"].ToString();

                                            // new timemeasurements
                                            long startTick = Convert.ToDateTime(tbBeginn_3.EditValue).Ticks;
                                            long endTick = Convert.ToDateTime(tbEnde_3.EditValue).Ticks;
                                            long tick = endTick - startTick;
                                            long seconds = tick / TimeSpan.TicksPerSecond;
                                            long milliseconds = tick / TimeSpan.TicksPerMillisecond;

                                            _timemeasurements += milliseconds.ToString();

                                            _dr_new["timemeasurements"] = _timemeasurements;

                                            _dr_new.AcceptChanges();
                                        }
                                    }
                                }
                                else
                                {
                                    tbBeginn_3.Text = _dr["beginn"].ToString();
                                    tbEnde_3.Text = _dr["ende"].ToString();
                                }
                                if (_Zusatzeit != "") tbZusatzzeit_3.Text = _Zusatzeit; else tbZusatzzeit_3.Text = _dr["zusatzzeit"].ToString();
                            }
                            else
                            {
                                deDatum_3.EditValue = Convert.ToDateTime(_dr["datum"].ToString() + " 00:00:00");
                                lblDauer_3.Content = _dr["dauer"].ToString();
                                tbBeginn_3.Text = _dr["beginn"].ToString();
                                tbEnde_3.Text = _dr["ende"].ToString();
                                tbZusatzzeit_3.Text = _dr["zusatzzeit"].ToString();
                            }

                            cbTaetigkeit_3.Items.Clear();
                            String _tmp_User = _User;
                            _tmp_User = _tmp_User.Replace("\\", "_");
                            _tmp_User = _tmp_User.Replace(".", "_");

                            DataRow[] _result = _TimeEntryCache.Select("[type]='" + "FAVORITE_ACTIVITY_" + _tmp_User + "'");
                            if (_result.Length > 0)
                            {
                                foreach (DataRow _current_dr in _result)
                                {
                                    cbTaetigkeit_3.Items.Add(_current_dr["description"].ToString());
                                }
                            }
                            cbTaetigkeit_3.Text = _dr["taetigkeit"].ToString();

                            cbAusgefuehrtDurch_3.Items.Clear();
                            cbAusgefuehrtDurch_3.Text = "";
                            _result = _TimeEntryCache.Select("[type]='USER.LIST'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmuser"].ToString()) _selected_li = _li;
                                cbAusgefuehrtDurch_3.Items.Add(_li);
                            }
                            if (_selected_li != null) cbAusgefuehrtDurch_3.SelectedItem = _selected_li;

                            cbAbrechnungsweise_3.Items.Clear();
                            cbAbrechnungsweise_3.Text = "";
                            cbAbrechnungsweise_3.Tag = "";
                            _result = _TimeEntryCache.Select("[type]='PICKLIST.BILLING'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmbillingtype"].ToString()) _selected_li = _li;
                                cbAbrechnungsweise_3.Items.Add(_li);
                            }
                            if (_selected_li != null) cbAbrechnungsweise_3.SelectedItem = _selected_li;

                            // STATUSLIST.STATUSCODE
                            cbStatus_3.Items.Clear();
                            cbStatus_3.Text = "";
                            _result = _TimeEntryCache.Select("[type]='STATUSLIST.STATUSCODE'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmstatus"].ToString()) _selected_li = _li;
                                cbStatus_3.Items.Add(_li);
                            }
                            if (_selected_li != null) cbStatus_3.SelectedItem = _selected_li;

                            teBeschreibung_3.Text = _dr["beschreibung"].ToString();
                            teAnmerkung_3.Text = _dr["anmerkung"].ToString();
                            ceFahrtkosten_3.IsChecked = Convert.ToBoolean(_dr["fahrtkosten"].ToString());

                            // Calc Time Values
                            if (!DoCopyTimeEntry)
                            {
                                CalcTimeValues(_dr, 3);
                            }
                            else
                            {
                                DataRow _dr_new = null;
                                if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                {
                                    CalcTimeValues(_dr_new, 3);

                                    _dr_new.AcceptChanges();
                                }
                            }

                            btnTimerCtrl_3.IsEnabled = true;
                            lblTimerCtrlInfo_3.Content = "00:00:00";
                            btnShowQuickTimeentryList_3.IsEnabled = true;

                            if (TimeEntryTabCtrl.Visibility == System.Windows.Visibility.Collapsed) TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Visible;
                            TimeEntryTabCtrl.SelectedItem = tabTE_3;

                            CheckSelectedEntity(_OrganizationService, 3, false);
                            CheckSelectedEntity2(_OrganizationService, 3, false);

                            if (cbBezugEntity_3.Text.Trim() == "")
                            {
                                cbBezugEntity2_3.Items.Clear();

                                cbBezugEntity2_3.IsEnabled = false;
                                cbMSCRMEntity2_3.IsEnabled = false;
                                btnMSCRMEntityInfoImage2_3.IsEnabled = false;
                                btnExtendedSearch2_3.IsEnabled = false;
                                btnFavorites2_3.IsEnabled = false;
                                btnAddToFavorites2_3.IsEnabled = false;
                            }
                            else
                            {
                                cbBezugEntity2_3.Items.Clear();
                                cbBezugEntity2_3.Items.Add("");

                                foreach (var item in _BezugValues)
                                {
                                    if (cbBezugEntity_3.Text.Trim().ToUpper() != item) cbBezugEntity2_3.Items.Add(item);
                                }

                                cbBezugEntity2_3.IsEnabled = true;
                                cbMSCRMEntity2_3.IsEnabled = true;
                                btnMSCRMEntityInfoImage2_3.IsEnabled = true;
                                btnExtendedSearch2_3.IsEnabled = true;
                                btnFavorites2_3.IsEnabled = true;
                                btnAddToFavorites2_3.IsEnabled = true;
                            }


                        }

                        cbTaetigkeit_3.Focus();
                    }
                    else if (tabTE_4.Visibility == System.Windows.Visibility.Collapsed)
                    {
                        DataRow _dr = _dt.Rows.Find(_te_guid.ToString());
                        if (_dr != null)
                        {
                            tabTE_4.Visibility = System.Windows.Visibility.Visible;
                            if (!DoCopyTimeEntry) tabTE_4.Tag = _te_guid.ToString(); else tabTE_4.Tag = _new_te_guid.ToString();
                            tabTE_4.Header = rm.GetString("tabTE_4");

                            //btAddNewEntity_4.IsEnabled = false;
                            //btAddNewEntity2_4.IsEnabled = false;

                            cbBezugEntity_4.Items.Clear();

                            foreach (var item in _BezugValues)
                            {
                                cbBezugEntity_4.Items.Add(item);
                            }

                            cbBezugEntity_4.Text = _dr["tekz"].ToString().Trim().ToUpper();
                            cbBezugEntity2_4.Text = _dr["tekz2"].ToString().Trim().ToUpper();
                            cbMSCRMEntity_4.Text = _dr["mscrmbezug"].ToString();
                            cbMSCRMEntity2_4.Text = _dr["mscrmbezug2"].ToString();

                            if (DoCopyTimeEntry)
                            {
                                deDatum_4.EditValue = Convert.ToDateTime(_SetDateTime, CultureInfo.CurrentCulture.DateTimeFormat);

                                if (_Dauer != "")
                                {
                                    DataRow _dr_new = null;
                                    if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                    {
                                        String[] _dm = _Dauer.Split(':');
                                        Int32 _dm_hours = 0;
                                        Int32 _dm_minutes = 0;
                                        try
                                        {
                                            _dm_hours = Convert.ToInt32(_dm[0]);
                                            _dm_minutes = Convert.ToInt32(_dm[1]);
                                        }
                                        catch
                                        {
                                        }

                                        if (_dm_hours != 0 || _dm_minutes != 0)
                                        {
                                            String _timemeasurements = "";

                                            _timemeasurements = _SetDateTime.ToShortDateString() + " " + _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00") + ":00|";

                                            _dr_new["beginn"] = _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00");

                                            _SetDateTime = _SetDateTime.AddHours(_dm_hours);
                                            _SetDateTime = _SetDateTime.AddMinutes(_dm_minutes);

                                            _dr_new["ende"] = _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00");

                                            _timemeasurements += _SetDateTime.ToShortDateString() + " " + _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00") + ":00|";

                                            tbBeginn_4.Text = _dr_new["beginn"].ToString();
                                            tbEnde_4.Text = _dr_new["ende"].ToString();

                                            // new timemeasurements
                                            long startTick = Convert.ToDateTime(tbBeginn_4.EditValue).Ticks;
                                            long endTick = Convert.ToDateTime(tbEnde_4.EditValue).Ticks;
                                            long tick = endTick - startTick;
                                            long seconds = tick / TimeSpan.TicksPerSecond;
                                            long milliseconds = tick / TimeSpan.TicksPerMillisecond;

                                            _timemeasurements += milliseconds.ToString();

                                            _dr_new["timemeasurements"] = _timemeasurements;

                                            _dr_new.AcceptChanges();
                                        }
                                    }
                                }
                                else
                                {
                                    tbBeginn_4.Text = _dr["beginn"].ToString();
                                    tbEnde_4.Text = _dr["ende"].ToString();
                                }
                                if (_Zusatzeit != "") tbZusatzzeit_4.Text = _Zusatzeit; else tbZusatzzeit_4.Text = _dr["zusatzzeit"].ToString();
                            }
                            else
                            {
                                deDatum_4.EditValue = Convert.ToDateTime(_dr["datum"].ToString() + " 00:00:00");
                                lblDauer_4.Content = _dr["dauer"].ToString();
                                tbBeginn_4.Text = _dr["beginn"].ToString();
                                tbEnde_4.Text = _dr["ende"].ToString();
                                tbZusatzzeit_4.Text = _dr["zusatzzeit"].ToString();
                            }

                            cbTaetigkeit_4.Items.Clear();
                            String _tmp_User = _User;
                            _tmp_User = _tmp_User.Replace("\\", "_");
                            _tmp_User = _tmp_User.Replace(".", "_");

                            DataRow[] _result = _TimeEntryCache.Select("[type]='" + "FAVORITE_ACTIVITY_" + _tmp_User + "'");
                            if (_result.Length > 0)
                            {
                                foreach (DataRow _current_dr in _result)
                                {
                                    cbTaetigkeit_4.Items.Add(_current_dr["description"].ToString());
                                }
                            }
                            cbTaetigkeit_4.Text = _dr["taetigkeit"].ToString();

                            cbAusgefuehrtDurch_4.Items.Clear();
                            cbAusgefuehrtDurch_4.Text = "";
                            _result = _TimeEntryCache.Select("[type]='USER.LIST'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmuser"].ToString()) _selected_li = _li;
                                cbAusgefuehrtDurch_4.Items.Add(_li);
                            }
                            if (_selected_li != null) cbAusgefuehrtDurch_4.SelectedItem = _selected_li;

                            cbAbrechnungsweise_4.Items.Clear();
                            cbAbrechnungsweise_4.Text = "";
                            cbAbrechnungsweise_4.Tag = "";
                            _result = _TimeEntryCache.Select("[type]='PICKLIST.BILLING'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmbillingtype"].ToString()) _selected_li = _li;
                                cbAbrechnungsweise_4.Items.Add(_li);
                            }
                            if (_selected_li != null) cbAbrechnungsweise_4.SelectedItem = _selected_li;

                            // STATUSLIST.STATUSCODE
                            cbStatus_4.Items.Clear();
                            cbStatus_4.Text = "";
                            _result = _TimeEntryCache.Select("[type]='STATUSLIST.STATUSCODE'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmstatus"].ToString()) _selected_li = _li;
                                cbStatus_4.Items.Add(_li);
                            }
                            if (_selected_li != null) cbStatus_4.SelectedItem = _selected_li;

                            teBeschreibung_4.Text = _dr["beschreibung"].ToString();
                            teAnmerkung_4.Text = _dr["anmerkung"].ToString();
                            ceFahrtkosten_4.IsChecked = Convert.ToBoolean(_dr["fahrtkosten"].ToString());

                            // Calc Time Values
                            if (!DoCopyTimeEntry)
                            {
                                CalcTimeValues(_dr, 4);
                            }
                            else
                            {
                                DataRow _dr_new = null;
                                if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                {
                                    CalcTimeValues(_dr_new, 4);

                                    _dr_new.AcceptChanges();
                                }
                            }

                            btnTimerCtrl_4.IsEnabled = true;
                            lblTimerCtrlInfo_4.Content = "00:00:00";
                            btnShowQuickTimeentryList_4.IsEnabled = true;

                            if (TimeEntryTabCtrl.Visibility == System.Windows.Visibility.Collapsed) TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Visible;
                            TimeEntryTabCtrl.SelectedItem = tabTE_4;

                            CheckSelectedEntity(_OrganizationService, 4, false);
                            CheckSelectedEntity2(_OrganizationService, 4, false);

                            if (cbBezugEntity_4.Text.Trim() == "")
                            {
                                cbBezugEntity2_4.Items.Clear();

                                cbBezugEntity2_4.IsEnabled = false;
                                cbMSCRMEntity2_4.IsEnabled = false;
                                btnMSCRMEntityInfoImage2_4.IsEnabled = false;
                                btnExtendedSearch2_4.IsEnabled = false;
                                btnFavorites2_4.IsEnabled = false;
                                btnAddToFavorites2_4.IsEnabled = false;
                            }
                            else
                            {
                                cbBezugEntity2_4.Items.Clear();

                                cbBezugEntity2_4.Items.Add("");

                                foreach (var item in _BezugValues)
                                {
                                    if (cbBezugEntity_4.Text.Trim().ToUpper() != item) cbBezugEntity2_4.Items.Add(item);
                                }


                                cbBezugEntity2_4.IsEnabled = true;
                                cbMSCRMEntity2_4.IsEnabled = true;
                                btnMSCRMEntityInfoImage2_4.IsEnabled = true;
                                btnExtendedSearch2_4.IsEnabled = true;
                                btnFavorites2_4.IsEnabled = true;
                                btnAddToFavorites2_4.IsEnabled = true;
                            }

                        }

                        cbTaetigkeit_4.Focus();
                    }
                    else if (tabTE_5.Visibility == System.Windows.Visibility.Collapsed)
                    {
                        DataRow _dr = _dt.Rows.Find(_te_guid.ToString());
                        if (_dr != null)
                        {
                            tabTE_5.Visibility = System.Windows.Visibility.Visible;
                            if (!DoCopyTimeEntry) tabTE_5.Tag = _te_guid.ToString(); else tabTE_5.Tag = _new_te_guid.ToString();
                            tabTE_5.Header = rm.GetString("tabTE_5");

                            //btAddNewEntity_5.IsEnabled = false;
                            //btAddNewEntity2_5.IsEnabled = false;

                            cbBezugEntity_5.Items.Clear();

                            foreach (var item in _BezugValues)
                            {
                                cbBezugEntity_5.Items.Add(item);
                            }

                            cbBezugEntity_5.Text = _dr["tekz"].ToString().Trim().ToUpper();
                            cbBezugEntity2_5.Text = _dr["tekz2"].ToString().Trim().ToUpper();
                            cbMSCRMEntity_5.Text = _dr["mscrmbezug"].ToString();
                            cbMSCRMEntity2_5.Text = _dr["mscrmbezug2"].ToString();

                            if (DoCopyTimeEntry)
                            {
                                deDatum_5.EditValue = Convert.ToDateTime(_SetDateTime, CultureInfo.CurrentCulture.DateTimeFormat);

                                if (_Dauer != "")
                                {
                                    DataRow _dr_new = null;
                                    if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                    {
                                        String[] _dm = _Dauer.Split(':');
                                        Int32 _dm_hours = 0;
                                        Int32 _dm_minutes = 0;
                                        try
                                        {
                                            _dm_hours = Convert.ToInt32(_dm[0]);
                                            _dm_minutes = Convert.ToInt32(_dm[1]);
                                        }
                                        catch
                                        {
                                        }

                                        if (_dm_hours != 0 || _dm_minutes != 0)
                                        {
                                            String _timemeasurements = "";

                                            _timemeasurements = _SetDateTime.ToShortDateString() + " " + _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00") + ":00|";

                                            _dr_new["beginn"] = _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00");

                                            _SetDateTime = _SetDateTime.AddHours(_dm_hours);
                                            _SetDateTime = _SetDateTime.AddMinutes(_dm_minutes);

                                            _dr_new["ende"] = _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00");

                                            _timemeasurements += _SetDateTime.ToShortDateString() + " " + _SetDateTime.Hour.ToString("00") + ":" + _SetDateTime.Minute.ToString("00") + ":00|";

                                            tbBeginn_5.Text = _dr_new["beginn"].ToString();
                                            tbEnde_5.Text = _dr_new["ende"].ToString();

                                            // new timemeasurements
                                            long startTick = Convert.ToDateTime(tbBeginn_5.EditValue).Ticks;
                                            long endTick = Convert.ToDateTime(tbEnde_5.EditValue).Ticks;
                                            long tick = endTick - startTick;
                                            long seconds = tick / TimeSpan.TicksPerSecond;
                                            long milliseconds = tick / TimeSpan.TicksPerMillisecond;

                                            _timemeasurements += milliseconds.ToString();

                                            _dr_new["timemeasurements"] = _timemeasurements;

                                            _dr_new.AcceptChanges();
                                        }
                                    }
                                }
                                else
                                {
                                    tbBeginn_5.Text = _dr["beginn"].ToString();
                                    tbEnde_5.Text = _dr["ende"].ToString();
                                }
                                if (_Zusatzeit != "") tbZusatzzeit_5.Text = _Zusatzeit; else tbZusatzzeit_5.Text = _dr["zusatzzeit"].ToString();
                            }
                            else
                            {
                                deDatum_5.EditValue = Convert.ToDateTime(_dr["datum"].ToString() + " 00:00:00");
                                lblDauer_5.Content = _dr["dauer"].ToString();
                                tbBeginn_5.Text = _dr["beginn"].ToString();
                                tbEnde_5.Text = _dr["ende"].ToString();
                                tbZusatzzeit_5.Text = _dr["zusatzzeit"].ToString();
                            }

                            cbTaetigkeit_5.Items.Clear();
                            String _tmp_User = _User;
                            _tmp_User = _tmp_User.Replace("\\", "_");
                            _tmp_User = _tmp_User.Replace(".", "_");

                            DataRow[] _result = _TimeEntryCache.Select("[type]='" + "FAVORITE_ACTIVITY_" + _tmp_User + "'");
                            if (_result.Length > 0)
                            {
                                foreach (DataRow _current_dr in _result)
                                {
                                    cbTaetigkeit_5.Items.Add(_current_dr["description"].ToString());
                                }
                            }
                            cbTaetigkeit_5.Text = _dr["taetigkeit"].ToString();

                            cbAusgefuehrtDurch_5.Items.Clear();
                            cbAusgefuehrtDurch_5.Text = "";
                            _result = _TimeEntryCache.Select("[type]='USER.LIST'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmuser"].ToString()) _selected_li = _li;
                                cbAusgefuehrtDurch_5.Items.Add(_li);
                            }
                            if (_selected_li != null) cbAusgefuehrtDurch_5.SelectedItem = _selected_li;

                            cbAbrechnungsweise_5.Items.Clear();
                            cbAbrechnungsweise_5.Text = "";
                            cbAbrechnungsweise_5.Tag = "";
                            _result = _TimeEntryCache.Select("[type]='PICKLIST.BILLING'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmbillingtype"].ToString()) _selected_li = _li;
                                cbAbrechnungsweise_5.Items.Add(_li);
                            }
                            if (_selected_li != null) cbAbrechnungsweise_5.SelectedItem = _selected_li;

                            // STATUSLIST.STATUSCODE
                            cbStatus_5.Items.Clear();
                            cbStatus_5.Text = "";
                            _result = _TimeEntryCache.Select("[type]='STATUSLIST.STATUSCODE'");
                            _selected_li = null;
                            foreach (DataRow _row in _result)
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                                _li.Content = _row["description"].ToString();
                                _li.Tag = _row["mscrmid"].ToString();
                                if (_row["description"].ToString() == _dr["mscrmstatus"].ToString()) _selected_li = _li;
                                cbStatus_5.Items.Add(_li);
                            }
                            if (_selected_li != null) cbStatus_5.SelectedItem = _selected_li;

                            teBeschreibung_5.Text = _dr["beschreibung"].ToString();
                            teAnmerkung_5.Text = _dr["anmerkung"].ToString();
                            ceFahrtkosten_5.IsChecked = Convert.ToBoolean(_dr["fahrtkosten"].ToString());

                            // Calc Time Values
                            if (!DoCopyTimeEntry)
                            {
                                CalcTimeValues(_dr, 5);
                            }
                            else
                            {
                                DataRow _dr_new = null;
                                if ((_dr_new = _dt.Rows.Find(_new_te_guid.ToString())) != null)
                                {
                                    CalcTimeValues(_dr_new, 5);

                                    _dr_new.AcceptChanges();
                                }
                            }

                            btnTimerCtrl_5.IsEnabled = true;
                            lblTimerCtrlInfo_5.Content = "00:00:00";
                            btnShowQuickTimeentryList_5.IsEnabled = true;

                            if (TimeEntryTabCtrl.Visibility == System.Windows.Visibility.Collapsed) TimeEntryTabCtrl.Visibility = System.Windows.Visibility.Visible;
                            TimeEntryTabCtrl.SelectedItem = tabTE_5;

                            CheckSelectedEntity(_OrganizationService, 5, false);
                            CheckSelectedEntity2(_OrganizationService, 5, false);

                            if (cbBezugEntity_5.Text.Trim() == "")
                            {
                                cbBezugEntity2_5.Items.Clear();

                                cbBezugEntity2_5.IsEnabled = false;
                                cbMSCRMEntity2_5.IsEnabled = false;
                                btnMSCRMEntityInfoImage2_5.IsEnabled = false;
                                btnExtendedSearch2_5.IsEnabled = false;
                                btnFavorites2_5.IsEnabled = false;
                                btnAddToFavorites2_5.IsEnabled = false;
                            }
                            else
                            {
                                cbBezugEntity2_5.Items.Clear();
                                cbBezugEntity2_5.Items.Add("");

                                foreach (var item in _BezugValues)
                                {
                                    if (cbBezugEntity_5.Text.Trim().ToUpper() != item) cbBezugEntity2_5.Items.Add(item);
                                }

                                cbBezugEntity2_5.IsEnabled = true;
                                cbMSCRMEntity2_5.IsEnabled = true;
                                btnMSCRMEntityInfoImage2_5.IsEnabled = true;
                                btnExtendedSearch2_5.IsEnabled = true;
                                btnFavorites2_5.IsEnabled = true;
                                btnAddToFavorites2_5.IsEnabled = true;
                            }
                        }

                        cbTaetigkeit_5.Focus();
                    }
                    else
                    {
                        MessageBox.Show(rm.GetString("max5"), rm.GetString("ImportantNote"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }

                // Check if slots are free
                if (tabTE_1.Tag.ToString() == "" || tabTE_2.Tag.ToString() == "" || tabTE_3.Tag.ToString() == "" || tabTE_4.Tag.ToString() == "" || tabTE_5.Tag.ToString() == "")
                    _rV = true;
            }
            catch (Exception ex)
            {
                FileLog("ActivateTimeEntry", "Error: " + ex.Message);
                _rV = false;
            }

            _ActivateTimeEntryActive = false;

            return _rV;
        }

        private void OnSyncMSCRMItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            if (MSCRM_ConState == MSCRM_ConectionStates.Connected)
            {
                LoadTimeEntitiesFromMSCRM(_OrganizationService, false);
            }
        }

        private void btnShowQuickTimeentryList_Click(object sender, RoutedEventArgs e)
        {
            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;
            if (_ti != null)
            {
                String _te_guid = _ti.Tag.ToString();
                DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;

                DataRow _dr = _dt.Rows.Find(_te_guid.ToString());
                if (_dr != null)
                {
                    String _te_timemeasurements = _dr["timemeasurements"].ToString().Trim();

                    EditTimeSpanItems _EditTimeSpanItems = new EditTimeSpanItems(_te_timemeasurements);
                    _EditTimeSpanItems.Topmost = false;
                    _EditTimeSpanItems.Owner = this;
                    _EditTimeSpanItems.ShowDialog();

                    if (_EditTimeSpanItems._AcceptChanges)
                    {
                        _dr["timemeasurements"] = _EditTimeSpanItems._measurements;
                        _dr.AcceptChanges();

                        UpdateTimeMeasurements();
                    }
                }
            }
        }

        private void CheckAndStopTimer()
        {
            if (TimerMode)
            {
                TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;
                if (_ti != null)
                {
                    if (_ti.Tag != null)
                    {
                        if (_ti.Tag.ToString() == _timer_info_te_guid)
                        {
                            TimerCtrl_StopTimer();
                        }
                    }
                }
            }

        }

        private void te_Active_EditValueChanged(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            try
            {
                TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;
                if (_ti != null)
                {
                    if (_ti.Tag != null)
                    {
                        String _te_guid = _ti.Tag.ToString();
                        DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
                        DataRow _dr = _dt.Rows.Find(_te_guid.ToString());
                        DateTime _tmp_time;

                        if (_dr != null)
                        {
                            if (_te_guid == tabTE_1.Tag.ToString())
                            {
                                _tmp_time = Convert.ToDateTime(tbZusatzzeit_1.EditValue);
                                _dr["zusatzzeit"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                _tmp_time = Convert.ToDateTime(tbBeginn_1.EditValue);
                                _dr["beginn"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                CalcTimeValues(_dr, 1);
                            }
                            else if (_te_guid == tabTE_2.Tag.ToString())
                            {
                                _tmp_time = Convert.ToDateTime(tbZusatzzeit_2.EditValue);
                                _dr["zusatzzeit"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                _tmp_time = Convert.ToDateTime(tbBeginn_2.EditValue);
                                _dr["beginn"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                CalcTimeValues(_dr, 2);
                            }
                            else if (_te_guid == tabTE_3.Tag.ToString())
                            {
                                _tmp_time = Convert.ToDateTime(tbZusatzzeit_3.EditValue);
                                _dr["zusatzzeit"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                _tmp_time = Convert.ToDateTime(tbBeginn_3.EditValue);
                                _dr["beginn"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                CalcTimeValues(_dr, 3);
                            }
                            else if (_te_guid == tabTE_4.Tag.ToString())
                            {
                                _tmp_time = Convert.ToDateTime(tbZusatzzeit_4.EditValue);
                                _dr["zusatzzeit"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                _tmp_time = Convert.ToDateTime(tbBeginn_4.EditValue);
                                _dr["beginn"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                CalcTimeValues(_dr, 4);
                            }
                            else if (_te_guid == tabTE_5.Tag.ToString())
                            {
                                _tmp_time = Convert.ToDateTime(tbZusatzzeit_5.EditValue);
                                _dr["zusatzzeit"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                _tmp_time = Convert.ToDateTime(tbBeginn_5.EditValue);
                                _dr["beginn"] = _tmp_time.Hour.ToString("00") + ":" + _tmp_time.Minute.ToString("00");
                                CalcTimeValues(_dr, 5);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                FileLog("te_Active_EditValueChanged", "Error: " + ex.Message);
            }
        }

        private void btnSaveAndCloseTimeentry_Click(object sender, RoutedEventArgs e)
        {
            // Speichern und schliessen
            SaveActiveTimeEntry(true, false);
        }


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

        private void cbMSCRMEntity_EditValueChanged(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            try
            {
                DevExpress.Xpf.Editors.ComboBoxEdit _cb_edit = (DevExpress.Xpf.Editors.ComboBoxEdit)sender;

                // Check Entity
                try
                {
                    if (_cb_edit.Name.Contains("_1"))
                    {
                        CheckSelectedEntity(_OrganizationService, 1, true);
                    }
                    else if (_cb_edit.Name.Contains("_2"))
                    {
                        CheckSelectedEntity(_OrganizationService, 2, true);
                    }
                    else if (_cb_edit.Name.Contains("_3"))
                    {
                        CheckSelectedEntity(_OrganizationService, 3, true);
                    }
                    else if (_cb_edit.Name.Contains("_4"))
                    {
                        CheckSelectedEntity(_OrganizationService, 4, true);
                    }
                    else if (_cb_edit.Name.Contains("_5"))
                    {
                        CheckSelectedEntity(_OrganizationService, 5, true);
                    }
                }
                catch (Exception ex)
                {
                    FileLog("cbMSCRMEntity_EditValueChanged@CheckSelectedEntity", "Error: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                FileLog("cbMSCRMEntity_EditValueChanged", "Error: " + ex.Message);
            }
        }

        private Boolean TestMSCRMEntity(IOrganizationService _Service, String MSCRMId, String EntityType)
        {
            Boolean _rV = false;

            try
            {
                if (MSCRM_ConState == MSCRM_ConectionStates.Connected)
                {
                    if (EntityType.Trim().ToUpper() == WorkPackageString)
                    {
                        try
                        {
                            if (_Service.Retrieve(a24_workpackage.EntityLogicalName, new Guid(MSCRMId), new ColumnSet(new String[] { Constants.FieldNames.TimeEntry.Name })) != null)
                                _rV = true;
                        }
                        catch
                        {
                            _rV = false;
                        }
                    }
                    else if (EntityType.Trim().ToUpper() == IncidentString)
                    {
                        try
                        {


                            if (_Service.Retrieve(Incident.EntityLogicalName, new Guid(MSCRMId), new ColumnSet(new String[] { "title" })) != null)
                                _rV = true;
                        }
                        catch
                        {
                            _rV = false;
                        }
                    }
                    else if (EntityType.Trim().ToUpper() == OpportunityString)
                    {
                        try
                        {
                            if (_Service.Retrieve(Opportunity.EntityLogicalName, new Guid(MSCRMId), new ColumnSet(new String[] { "name" })) != null)
                                _rV = true;
                        }
                        catch
                        {
                            _rV = false;
                        }
                    }
                    //else if (EntityType.Trim().ToUpper() == "REL")
                    //{
                    //    try
                    //    {
                    //        if (_Service.Retrieve(New_Version.EntityLogicalName, new Guid(MSCRMId), new ColumnSet(new String[] { "new_versionsnummer" })) != null)
                    //            _rV = true;
                    //    }
                    //    catch
                    //    {
                    //        _rV = false;
                    //    }
                    //}
                    else if (EntityType.Trim().ToUpper() == CampaignString)
                    {
                        try
                        {
                            if (_Service.Retrieve(Campaign.EntityLogicalName, new Guid(MSCRMId), new ColumnSet(new String[] { "name" })) != null)
                                _rV = true;
                        }
                        catch
                        {
                            _rV = false;
                        }
                    }
                    else if (EntityType.Trim().ToUpper() == ActionItemString)
                    {
                        try
                        {
                            if (_Service.Retrieve(a24_action_item.EntityLogicalName, new Guid(MSCRMId), new ColumnSet(new String[] { Constants.FieldNames.ActionItem.Name })) != null)
                                _rV = true;
                        }
                        catch
                        {
                            _rV = false;
                        }
                    }
                }
                else _rV = true;
            }
            catch (Exception ex)
            {
                FileLog("TestMSCRMEntity", "Error: " + ex.Message);
                _rV = false;
            }

            return _rV;
        }


        private Boolean CheckSelectedEntity(IOrganizationService _Service, Int32 ActiveTE, Boolean DoAbrechnungsweise)
        {
            try
            {
                // Check Entity
                Entity_OK = false;
                DevExpress.Xpf.Editors.ComboBoxEditItem _selected_li = null;

                if (ActiveTE == 1)
                {
                    _selected_li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_1.SelectedItem;
                    if (_selected_li != null) Entity_OK = TestMSCRMEntity(_OrganizationService, _selected_li.Tag.ToString(), cbBezugEntity_1.Text.Trim().ToUpper());
                    if (Entity_OK) { btnMSCRMEntityInfoImage_1.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/OK.png"); btnMSCRMEntityInfoImage_1.ToolTip = rm.GetString("OK.png"); }
                    else { btnMSCRMEntityInfoImage_1.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/NOK.png"); btnMSCRMEntityInfoImage_1.ToolTip = rm.GetString("NOK.png"); }

                    if (DoAbrechnungsweise)
                    {
                        if (Entity_OK && MSCRM_ConState == MSCRM_ConectionStates.Connected)
                        {
                            Guid _MSCRMBezugID = Guid.Empty;
                            DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                            _MSCRMBezugID = new Guid(_li.Tag.ToString());

                            Boolean _is_failed = true;

                            // Abrechnugsweise übernehmen (AP, ANF) 
                            String Sync_Entity_Type = cbBezugEntity_1.Text.Trim().ToUpper();
                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (wp != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_1.Items)
                                    {
                                        if (wp.a24_billing_opt != null)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == wp.a24_billing_opt.Value)
                                            {
                                                cbAbrechnungsweise_1.SelectedItem = _item; _is_failed = false; break;
                                            }

                                        }

                                    }
                                }
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (_in != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_1.Items)
                                    {
                                        if (_in.a24_billing_opt != null)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == _in.a24_billing_opt.Value)
                                            {
                                                cbAbrechnungsweise_1.SelectedItem = _item; _is_failed = false; break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (_is_failed)
                            {
                                // Abrechnugsweise offline übernehmen (AP, ANF)
                                Sync_Entity_Type = cbBezugEntity_1.Text.Trim().ToUpper();
                                String Sync_Type = "";
                                DataRow _dr = null;

                                if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                                else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                                if (Sync_Type != "")
                                {
                                    _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                    if (_li != null)
                                    {
                                        if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                        {
                                            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_1.Items)
                                            {
                                                if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                                {
                                                    cbAbrechnungsweise_1.SelectedItem = _item; break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Abrechnugsweise offline übernehmen (AP, ANF)
                            String Sync_Entity_Type = cbBezugEntity_1.Text.Trim().ToUpper();
                            String Sync_Type = "";
                            DataRow _dr = null;

                            if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                            else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                            if (Sync_Type != "")
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                if (_li != null)
                                {
                                    if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                    {
                                        foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_1.Items)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                            {
                                                cbAbrechnungsweise_1.SelectedItem = _item; break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ActiveTE == 2)
                {
                    _selected_li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_2.SelectedItem;
                    if (_selected_li != null) Entity_OK = TestMSCRMEntity(_OrganizationService, _selected_li.Tag.ToString(), cbBezugEntity_2.Text.Trim().ToUpper());
                    if (Entity_OK) { btnMSCRMEntityInfoImage_2.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/OK.png"); btnMSCRMEntityInfoImage_2.ToolTip = rm.GetString("OK.png"); }
                    else { btnMSCRMEntityInfoImage_2.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/NOK.png"); btnMSCRMEntityInfoImage_2.ToolTip = rm.GetString("NOK.png"); }

                    if (DoAbrechnungsweise)
                    {
                        if (Entity_OK && MSCRM_ConState == MSCRM_ConectionStates.Connected)
                        {
                            Guid _MSCRMBezugID = Guid.Empty;

                            DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_2.SelectedItem;
                            _MSCRMBezugID = new Guid(_li.Tag.ToString());

                            Boolean _is_failed = true;

                            // Abrechnugsweise übernehmen (AP, ANF) 
                            String Sync_Entity_Type = cbBezugEntity_2.Text.Trim().ToUpper();
                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (wp != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_2.Items)
                                    {
                                        if (wp.a24_billing_opt != null)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == wp.a24_billing_opt.Value)
                                            {
                                                cbAbrechnungsweise_2.SelectedItem = _item; _is_failed = false; break;
                                            }
                                        }
                                    }
                                }
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (_in != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_2.Items)
                                    {
                                        if (_in.a24_billing_opt != null)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == _in.a24_billing_opt.Value)
                                            {
                                                cbAbrechnungsweise_2.SelectedItem = _item; _is_failed = false; break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (_is_failed)
                            {
                                // Abrechnugsweise offline übernehmen (AP, ANF)
                                Sync_Entity_Type = cbBezugEntity_2.Text.Trim().ToUpper();
                                String Sync_Type = "";
                                DataRow _dr = null;

                                if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                                else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                                if (Sync_Type != "")
                                {
                                    _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                    if (_li != null)
                                    {
                                        if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                        {
                                            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_2.Items)
                                            {
                                                if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                                {
                                                    cbAbrechnungsweise_2.SelectedItem = _item; break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Abrechnugsweise offline übernehmen (AP, ANF)
                            String Sync_Entity_Type = cbBezugEntity_2.Text.Trim().ToUpper();
                            String Sync_Type = "";
                            DataRow _dr = null;

                            if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                            else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                            if (Sync_Type != "")
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                if (_li != null)
                                {
                                    if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                    {
                                        foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_2.Items)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                            {
                                                cbAbrechnungsweise_2.SelectedItem = _item; break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ActiveTE == 3)
                {
                    _selected_li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_3.SelectedItem;
                    if (_selected_li != null) Entity_OK = TestMSCRMEntity(_OrganizationService, _selected_li.Tag.ToString(), cbBezugEntity_3.Text.Trim().ToUpper());
                    if (Entity_OK) { btnMSCRMEntityInfoImage_3.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/OK.png"); btnMSCRMEntityInfoImage_3.ToolTip = rm.GetString("OK.png"); }
                    else { btnMSCRMEntityInfoImage_3.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/NOK.png"); btnMSCRMEntityInfoImage_3.ToolTip = rm.GetString("NOK.png"); }

                    if (DoAbrechnungsweise)
                    {
                        if (Entity_OK && MSCRM_ConState == MSCRM_ConectionStates.Connected)
                        {
                            Guid _MSCRMBezugID = Guid.Empty;

                            DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_3.SelectedItem;
                            _MSCRMBezugID = new Guid(_li.Tag.ToString());

                            Boolean _is_failed = true;

                            // Abrechnugsweise übernehmen (AP, ANF) 
                            String Sync_Entity_Type = cbBezugEntity_3.Text.Trim().ToUpper();
                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (wp != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_3.Items)
                                    {
                                        if (wp.a24_billing_opt != null)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == wp.a24_billing_opt.Value)
                                            {
                                                cbAbrechnungsweise_3.SelectedItem = _item; _is_failed = false; break;
                                            }
                                        }
                                    }
                                }
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (_in != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_3.Items)
                                    {
                                        if (_in.a24_billing_opt != null)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == _in.a24_billing_opt.Value)
                                            {
                                                cbAbrechnungsweise_3.SelectedItem = _item; _is_failed = false; break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (_is_failed)
                            {
                                // Abrechnugsweise offline übernehmen (AP, ANF)
                                Sync_Entity_Type = cbBezugEntity_3.Text.Trim().ToUpper();
                                String Sync_Type = "";
                                DataRow _dr = null;

                                if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                                else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                                if (Sync_Type != "")
                                {
                                    _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                    if (_li != null)
                                    {
                                        if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                        {
                                            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_3.Items)
                                            {
                                                if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                                {
                                                    cbAbrechnungsweise_3.SelectedItem = _item; break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Abrechnugsweise offline übernehmen (AP, ANF)
                            String Sync_Entity_Type = cbBezugEntity_3.Text.Trim().ToUpper();
                            String Sync_Type = "";
                            DataRow _dr = null;

                            if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                            else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                            if (Sync_Type != "")
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                if (_li != null)
                                {
                                    if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                    {
                                        foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_3.Items)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                            {
                                                cbAbrechnungsweise_3.SelectedItem = _item; break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ActiveTE == 4)
                {
                    _selected_li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_4.SelectedItem;
                    if (_selected_li != null) Entity_OK = TestMSCRMEntity(_OrganizationService, _selected_li.Tag.ToString(), cbBezugEntity_4.Text.Trim().ToUpper());
                    if (Entity_OK) { btnMSCRMEntityInfoImage_4.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/OK.png"); btnMSCRMEntityInfoImage_4.ToolTip = rm.GetString("OK.png"); }
                    else { btnMSCRMEntityInfoImage_4.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/NOK.png"); btnMSCRMEntityInfoImage_4.ToolTip = rm.GetString("NOK.png"); }

                    if (DoAbrechnungsweise)
                    {
                        if (Entity_OK && MSCRM_ConState == MSCRM_ConectionStates.Connected)
                        {
                            Guid _MSCRMBezugID = Guid.Empty;

                            DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_4.SelectedItem;
                            _MSCRMBezugID = new Guid(_li.Tag.ToString());

                            Boolean _is_failed = true;

                            // Abrechnugsweise übernehmen (AP, ANF) 
                            String Sync_Entity_Type = cbBezugEntity_4.Text.Trim().ToUpper();
                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (wp != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_4.Items)
                                    {
                                        if (wp.a24_billing_opt != null)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == wp.a24_billing_opt.Value)
                                            {
                                                cbAbrechnungsweise_4.SelectedItem = _item; _is_failed = false; break;
                                            }
                                        }
                                    }
                                }
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (_in != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_4.Items)
                                    {
                                        if (_in.a24_billing_opt != null)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == _in.a24_billing_opt.Value)
                                            {
                                                cbAbrechnungsweise_4.SelectedItem = _item; _is_failed = false; break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (_is_failed)
                            {
                                // Abrechnugsweise offline übernehmen (AP, ANF)
                                Sync_Entity_Type = cbBezugEntity_4.Text.Trim().ToUpper();
                                String Sync_Type = "";
                                DataRow _dr = null;

                                if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                                else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                                if (Sync_Type != "")
                                {
                                    _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                    if (_li != null)
                                    {
                                        if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                        {
                                            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_4.Items)
                                            {
                                                if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                                {
                                                    cbAbrechnungsweise_4.SelectedItem = _item; break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Abrechnugsweise offline übernehmen (AP, ANF)
                            String Sync_Entity_Type = cbBezugEntity_4.Text.Trim().ToUpper();
                            String Sync_Type = "";
                            DataRow _dr = null;

                            if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                            else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                            if (Sync_Type != "")
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                if (_li != null)
                                {
                                    if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                    {
                                        foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_4.Items)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                            {
                                                cbAbrechnungsweise_4.SelectedItem = _item; break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ActiveTE == 5)
                {
                    _selected_li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_5.SelectedItem;
                    if (_selected_li != null) Entity_OK = TestMSCRMEntity(_OrganizationService, _selected_li.Tag.ToString(), cbBezugEntity_5.Text.Trim().ToUpper());
                    if (Entity_OK) { btnMSCRMEntityInfoImage_5.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/OK.png"); btnMSCRMEntityInfoImage_5.ToolTip = rm.GetString("OK.png"); }
                    else { btnMSCRMEntityInfoImage_5.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/NOK.png"); btnMSCRMEntityInfoImage_5.ToolTip = rm.GetString("NOK.png"); }

                    if (DoAbrechnungsweise)
                    {
                        if (Entity_OK && MSCRM_ConState == MSCRM_ConectionStates.Connected)
                        {
                            Guid _MSCRMBezugID = Guid.Empty;

                            DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_5.SelectedItem;
                            _MSCRMBezugID = new Guid(_li.Tag.ToString());

                            Boolean _is_failed = true;

                            // Abrechnugsweise übernehmen (AP, ANF) 
                            String Sync_Entity_Type = cbBezugEntity_5.Text.Trim().ToUpper();
                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (wp != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_5.Items)
                                    {
                                        if (wp.a24_billing_opt != null)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == wp.a24_billing_opt.Value)
                                            {
                                                cbAbrechnungsweise_5.SelectedItem = _item; _is_failed = false; break;
                                            }
                                        }
                                    }
                                }
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (_in != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_5.Items)
                                    {
                                        if (_in.a24_billing_opt != null)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == _in.a24_billing_opt.Value)
                                            {
                                                cbAbrechnungsweise_5.SelectedItem = _item; _is_failed = false; break;
                                            }
                                        }
                                    }
                                }
                            }

                            if (_is_failed)
                            {
                                // Abrechnugsweise offline übernehmen (AP, ANF)
                                Sync_Entity_Type = cbBezugEntity_5.Text.Trim().ToUpper();
                                String Sync_Type = "";
                                DataRow _dr = null;

                                if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                                else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                                if (Sync_Type != "")
                                {
                                    _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                    if (_li != null)
                                    {
                                        if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                        {
                                            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_5.Items)
                                            {
                                                if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                                {
                                                    cbAbrechnungsweise_5.SelectedItem = _item; break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Abrechnugsweise offline übernehmen (AP, ANF)
                            String Sync_Entity_Type = cbBezugEntity_5.Text.Trim().ToUpper();
                            String Sync_Type = "";
                            DataRow _dr = null;

                            if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                            else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                            if (Sync_Type != "")
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                if (_li != null)
                                {
                                    if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                    {
                                        foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_5.Items)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                            {
                                                cbAbrechnungsweise_5.SelectedItem = _item; break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                return Entity_OK;
            }
            catch (Exception ex)
            {
                FileLog("CheckSelectedEntity", "Error: " + ex.Message);
                return false;
            }
        }

        private void cbMSCRMEntity_LostFocus(object sender, RoutedEventArgs e)
        {
            // Lost Focus ...
            DevExpress.Xpf.Editors.ComboBoxEdit _cb_edit = (DevExpress.Xpf.Editors.ComboBoxEdit)sender;

            String SelContent = "";
            Int32 _c = 0;
            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in _cb_edit.Items)
            {
                if (_item.Content.ToString().ToLower().IndexOf(_cb_edit.Text.Trim().ToLower()) != -1)
                {
                    _c++;
                    SelContent = _item.Content.ToString();
                }
            }

            if (_c == 1 && SelContent != "")
            {
                _cb_edit.Text = SelContent;
            }

            _cb_edit.ClosePopup();
        }

        private void cbMSCRMEntity_GotFocus(object sender, RoutedEventArgs e)
        {
            // Got Focus
            DevExpress.Xpf.Editors.ComboBoxEdit _cb_edit = (DevExpress.Xpf.Editors.ComboBoxEdit)sender;

            _cb_edit.SelectionLength = 0;
            _cb_edit.SelectedText = "";
        }

        private void cbBezugEntity_SelectedIndexChanged(object sender, RoutedEventArgs e)
        {
            DevExpress.Xpf.Editors.ComboBoxEdit _cb_edit = (DevExpress.Xpf.Editors.ComboBoxEdit)sender;

            String _TypeOfEntity = _cb_edit.Text.Trim().ToUpper();

            String _TOE_Name = "";
            if (_TypeOfEntity == WorkPackageString) _TOE_Name = "WORKPACKAGE";
            else if (_TypeOfEntity == OpportunityString) _TOE_Name = "OPPORTUNITY";
            else if (_TypeOfEntity == IncidentString) _TOE_Name = "INCIDENT";
            //else if (_TypeOfEntity == "REL") _TOE_Name = "NEW_VERSION";
            else if (_TypeOfEntity == CampaignString) _TOE_Name = "CAMPAIGN";
            else if (_TypeOfEntity == ActionItemString) _TOE_Name = "ACTIONITEM";

            if (_cb_edit.Name.Contains("_1"))
            {
                cbMSCRMEntity_1.Text = ""; cbMSCRMEntity_1.Items.Clear();
                if (_TypeOfEntity == OpportunityString || _TypeOfEntity == CampaignString)
                {
                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _li in cbAbrechnungsweise_1.Items)
                    {
                        if (_li.Content.ToString() == "Firmenintern") { cbAbrechnungsweise_1.SelectedItem = _li; break; }
                    }
                }

                if (cbBezugEntity_1.Text.Trim() == "")
                {
                    cbBezugEntity2_1.Text = "";
                    cbMSCRMEntity2_1.Text = ""; cbMSCRMEntity2_1.Items.Clear();

                    cbBezugEntity2_1.Items.Clear();

                    cbBezugEntity2_1.IsEnabled = false;
                    cbMSCRMEntity2_1.IsEnabled = false;
                    btnMSCRMEntityInfoImage2_1.IsEnabled = false;
                    btnExtendedSearch2_1.IsEnabled = false;
                    btnFavorites2_1.IsEnabled = false;
                    btnAddToFavorites2_1.IsEnabled = false;
                }
                else
                {
                    cbBezugEntity2_1.Items.Clear();

                    foreach (var item in _BezugValues)
                    {
                        if (cbBezugEntity_1.Text.Trim().ToUpper() != item) cbBezugEntity2_1.Items.Add(item);
                    }

                    cbBezugEntity2_1.IsEnabled = true;
                    cbMSCRMEntity2_1.IsEnabled = true;
                    btnMSCRMEntityInfoImage2_1.IsEnabled = true;
                    btnExtendedSearch2_1.IsEnabled = true;
                    btnFavorites2_1.IsEnabled = true;
                    btnAddToFavorites2_1.IsEnabled = true;

                    if (!cbBezugEntity2_1.Items.Contains(cbBezugEntity2_1.Text) || cbBezugEntity_1.Text.Trim().ToUpper() == cbBezugEntity2_1.Text.Trim().ToUpper())
                    {
                        cbBezugEntity2_1.Text = "";
                        cbMSCRMEntity2_1.Text = "";
                    }
                }

                //if (cbBezugEntity_1.Text.Trim().ToUpper() == Incident || cbBezugEntity_1.Text.Trim().ToUpper() == "REL")
                //{
                //    //btAddNewEntity_1.IsEnabled = true;
                //}
                //else
                //{
                //    //btAddNewEntity_1.IsEnabled = false;
                //}

                //if (cbBezugEntity2_1.Text.Trim().ToUpper() == Incident || cbBezugEntity2_1.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity2_1.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity2_1.IsEnabled = false;
                //}

                // Fill DropDown
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                if (_result != null)
                {
                    cbMSCRMEntity_1.Items.Clear();
                    foreach (DataRow _row in _result)
                    {
                        DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                        _li.Content = _row["description"].ToString();
                        _li.Tag = _row["mscrmid"].ToString();
                        cbMSCRMEntity_1.Items.Add(_li);
                    }
                }
            }
            else if (_cb_edit.Name.Contains("_2"))
            {
                cbMSCRMEntity_2.Text = ""; cbMSCRMEntity_2.Items.Clear();
                if (_TypeOfEntity == OpportunityString)
                {
                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _li in cbAbrechnungsweise_2.Items)
                    {
                        if (_li.Content.ToString() == "Firmenintern") { cbAbrechnungsweise_2.SelectedItem = _li; break; }
                    }
                }

                if (cbBezugEntity_2.Text.Trim() == "")
                {
                    cbBezugEntity2_2.Text = "";
                    cbMSCRMEntity2_2.Text = ""; cbMSCRMEntity2_2.Items.Clear();

                    cbBezugEntity2_2.Items.Clear();

                    cbBezugEntity2_2.IsEnabled = false;
                    cbMSCRMEntity2_2.IsEnabled = false;
                    btnMSCRMEntityInfoImage2_2.IsEnabled = false;
                    btnExtendedSearch2_2.IsEnabled = false;
                    btnFavorites2_2.IsEnabled = false;
                    btnAddToFavorites2_2.IsEnabled = false;
                }
                else
                {
                    cbBezugEntity2_2.Items.Clear();

                    foreach (var item in _BezugValues)
                    {
                        if (cbBezugEntity_2.Text.Trim().ToUpper() != item) cbBezugEntity2_2.Items.Add(item);
                    }


                    cbBezugEntity2_2.IsEnabled = true;
                    cbMSCRMEntity2_2.IsEnabled = true;
                    btnMSCRMEntityInfoImage2_2.IsEnabled = true;
                    btnExtendedSearch2_2.IsEnabled = true;
                    btnFavorites2_2.IsEnabled = true;
                    btnAddToFavorites2_2.IsEnabled = true;

                    if (cbBezugEntity_2.Text.Trim().ToUpper() == cbBezugEntity2_2.Text.Trim().ToUpper())
                    {
                        cbMSCRMEntity2_2.Text = ""; cbMSCRMEntity2_2.Items.Clear();
                    }
                }

                //if (cbBezugEntity_2.Text.Trim().ToUpper() == Incident || cbBezugEntity_2.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity_2.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity_2.IsEnabled = false;
                //}

                //if (cbBezugEntity2_2.Text.Trim().ToUpper() == Incident || cbBezugEntity2_2.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity2_2.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity2_2.IsEnabled = false;
                //}

                // Fill DropDown
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                if (_result != null)
                {
                    cbMSCRMEntity_2.Items.Clear();
                    foreach (DataRow _row in _result)
                    {
                        DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                        _li.Content = _row["description"].ToString();
                        _li.Tag = _row["mscrmid"].ToString();
                        cbMSCRMEntity_2.Items.Add(_li);
                    }
                }
            }
            else if (_cb_edit.Name.Contains("_3"))
            {
                cbMSCRMEntity_3.Text = ""; cbMSCRMEntity_3.Items.Clear();
                if (_TypeOfEntity == OpportunityString)
                {
                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _li in cbAbrechnungsweise_3.Items)
                    {
                        if (_li.Content.ToString() == "Firmenintern") { cbAbrechnungsweise_3.SelectedItem = _li; break; }
                    }
                }

                if (cbBezugEntity_3.Text.Trim() == "")
                {
                    cbBezugEntity2_3.Text = "";
                    cbMSCRMEntity2_3.Text = ""; cbMSCRMEntity2_3.Items.Clear();

                    cbBezugEntity2_3.Items.Clear();

                    cbBezugEntity2_3.IsEnabled = false;
                    cbMSCRMEntity2_3.IsEnabled = false;
                    btnMSCRMEntityInfoImage2_3.IsEnabled = false;
                    btnExtendedSearch2_3.IsEnabled = false;
                    btnFavorites2_3.IsEnabled = false;
                    btnAddToFavorites2_3.IsEnabled = false;
                }
                else
                {
                    cbBezugEntity2_3.Items.Clear();

                    foreach (var item in _BezugValues)
                    {
                        if (cbBezugEntity_3.Text.Trim().ToUpper() != item) cbBezugEntity2_3.Items.Add(item);
                    }

                    cbBezugEntity2_3.IsEnabled = true;
                    cbMSCRMEntity2_3.IsEnabled = true;
                    btnMSCRMEntityInfoImage2_3.IsEnabled = true;
                    btnExtendedSearch2_3.IsEnabled = true;
                    btnFavorites2_3.IsEnabled = true;
                    btnAddToFavorites2_3.IsEnabled = true;

                    if (cbBezugEntity_3.Text.Trim().ToUpper() == cbBezugEntity2_3.Text.Trim().ToUpper())
                    {
                        cbMSCRMEntity2_3.Text = ""; cbMSCRMEntity2_3.Items.Clear();
                    }
                }

                //if (cbBezugEntity_3.Text.Trim().ToUpper() == Incident || cbBezugEntity_3.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity_3.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity_3.IsEnabled = false;
                //}

                //if (cbBezugEntity2_3.Text.Trim().ToUpper() == Incident || cbBezugEntity2_3.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity2_3.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity2_3.IsEnabled = false;
                //}

                // Fill DropDown
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                if (_result != null)
                {
                    cbMSCRMEntity_3.Items.Clear();
                    foreach (DataRow _row in _result)
                    {
                        DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                        _li.Content = _row["description"].ToString();
                        _li.Tag = _row["mscrmid"].ToString();
                        cbMSCRMEntity_3.Items.Add(_li);
                    }
                }
            }
            else if (_cb_edit.Name.Contains("_4"))
            {
                cbMSCRMEntity_4.Text = ""; cbMSCRMEntity_4.Items.Clear();
                if (_TypeOfEntity == OpportunityString)
                {
                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _li in cbAbrechnungsweise_4.Items)
                    {
                        if (_li.Content.ToString() == "Firmenintern") { cbAbrechnungsweise_4.SelectedItem = _li; break; }
                    }
                }

                if (cbBezugEntity_4.Text.Trim() == "")
                {
                    cbBezugEntity2_4.Text = "";
                    cbMSCRMEntity2_4.Text = ""; cbMSCRMEntity2_4.Items.Clear();

                    cbBezugEntity2_4.Items.Clear();

                    cbBezugEntity2_4.IsEnabled = false;
                    cbMSCRMEntity2_4.IsEnabled = false;
                    btnMSCRMEntityInfoImage2_4.IsEnabled = false;
                    btnExtendedSearch2_4.IsEnabled = false;
                    btnFavorites2_4.IsEnabled = false;
                    btnAddToFavorites2_4.IsEnabled = false;
                }
                else
                {
                    cbBezugEntity2_4.Items.Clear();

                    foreach (var item in _BezugValues)
                    {
                        if (cbBezugEntity_4.Text.Trim().ToUpper() != item) cbBezugEntity2_4.Items.Add(item);
                    }


                    cbBezugEntity2_4.IsEnabled = true;
                    cbMSCRMEntity2_4.IsEnabled = true;
                    btnMSCRMEntityInfoImage2_4.IsEnabled = true;
                    btnExtendedSearch2_4.IsEnabled = true;
                    btnFavorites2_4.IsEnabled = true;
                    btnAddToFavorites2_4.IsEnabled = true;

                    if (cbBezugEntity_4.Text.Trim().ToUpper() == cbBezugEntity2_4.Text.Trim().ToUpper())
                    {
                        cbMSCRMEntity2_4.Text = ""; cbMSCRMEntity2_4.Items.Clear();
                    }
                }

                //if (cbBezugEntity_4.Text.Trim().ToUpper() == Incident || cbBezugEntity_4.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity_4.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity_4.IsEnabled = false;
                //}

                //if (cbBezugEntity2_4.Text.Trim().ToUpper() == Incident || cbBezugEntity2_4.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity2_4.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity2_4.IsEnabled = false;
                //}

                // Fill DropDown
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                if (_result != null)
                {
                    cbMSCRMEntity_4.Items.Clear();
                    foreach (DataRow _row in _result)
                    {
                        DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                        _li.Content = _row["description"].ToString();
                        _li.Tag = _row["mscrmid"].ToString();
                        cbMSCRMEntity_4.Items.Add(_li);
                    }
                }
            }
            else if (_cb_edit.Name.Contains("_5"))
            {
                cbMSCRMEntity_5.Text = ""; cbMSCRMEntity_5.Items.Clear();
                if (_TypeOfEntity == OpportunityString)
                {
                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _li in cbAbrechnungsweise_5.Items)
                    {
                        if (_li.Content.ToString() == "Firmenintern") { cbAbrechnungsweise_5.SelectedItem = _li; break; }
                    }
                }

                if (cbBezugEntity_5.Text.Trim() == "")
                {
                    cbBezugEntity2_5.Text = "";
                    cbMSCRMEntity2_5.Text = ""; cbMSCRMEntity2_5.Items.Clear();

                    cbBezugEntity2_5.Items.Clear();

                    cbBezugEntity2_5.IsEnabled = false;
                    cbMSCRMEntity2_5.IsEnabled = false;
                    btnMSCRMEntityInfoImage2_5.IsEnabled = false;
                    btnExtendedSearch2_5.IsEnabled = false;
                    btnFavorites2_5.IsEnabled = false;
                    btnAddToFavorites2_5.IsEnabled = false;
                }
                else
                {
                    cbBezugEntity2_5.Items.Clear();

                    foreach (var item in _BezugValues)
                    {
                        if (cbBezugEntity_5.Text.Trim().ToUpper() != item) cbBezugEntity2_5.Items.Add(item);
                    }


                    if (cbBezugEntity_5.Text.Trim().ToUpper() != OpportunityString) cbBezugEntity2_5.Items.Add(OpportunityString);
                    if (cbBezugEntity_5.Text.Trim().ToUpper() != IncidentString) cbBezugEntity2_5.Items.Add(IncidentString);
                    if (cbBezugEntity_5.Text.Trim().ToUpper() != "REL") cbBezugEntity2_5.Items.Add("REL");
                    if (cbBezugEntity_5.Text.Trim().ToUpper() != CampaignString) cbBezugEntity2_5.Items.Add(CampaignString);

                    cbBezugEntity2_5.IsEnabled = true;
                    cbMSCRMEntity2_5.IsEnabled = true;
                    btnMSCRMEntityInfoImage2_5.IsEnabled = true;
                    btnExtendedSearch2_5.IsEnabled = true;
                    btnFavorites2_5.IsEnabled = true;
                    btnAddToFavorites2_5.IsEnabled = true;

                    if (cbBezugEntity_5.Text.Trim().ToUpper() == cbBezugEntity2_5.Text.Trim().ToUpper())
                    {
                        cbMSCRMEntity2_5.Text = ""; cbMSCRMEntity2_5.Items.Clear();
                    }
                }

                //if (cbBezugEntity_5.Text.Trim().ToUpper() == Incident || cbBezugEntity_5.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity_5.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity_5.IsEnabled = false;
                //}

                //if (cbBezugEntity2_5.Text.Trim().ToUpper() == Incident || cbBezugEntity2_5.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity2_5.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity2_5.IsEnabled = false;
                //}

                // Fill DropDown
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                if (_result != null)
                {
                    cbMSCRMEntity_5.Items.Clear();
                    foreach (DataRow _row in _result)
                    {
                        DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                        _li.Content = _row["description"].ToString();
                        _li.Tag = _row["mscrmid"].ToString();
                        cbMSCRMEntity_5.Items.Add(_li);
                    }
                }
            }
        }

        private void cbAusgefuehrtDurch_LostFocus(object sender, RoutedEventArgs e)
        {
            // Lost Focus ...
            DevExpress.Xpf.Editors.ComboBoxEdit _cb_edit = (DevExpress.Xpf.Editors.ComboBoxEdit)sender;

            String _TypeOfEntity = "USER.LIST";

            String _searchText = _cb_edit.Text.Trim();
            if (_searchText.Length > 1)
            {
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TypeOfEntity + "' AND [description] LIKE '%" + EscapeLikeValue(_searchText) + "%'");
                if (_result.Length == 1)
                {
                    foreach (DataRow _row in _result)
                    {
                        _cb_edit.Text = _row["description"].ToString();
                        break;
                    }
                }
            }

            _cb_edit.ClosePopup();
        }

        private void btnExtendedSearch_Click(object sender, RoutedEventArgs e)
        {
            // Extended Entity Search
            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;

            try
            {
                EntitySearch _search_form = new EntitySearch(_TimeEntryCache, _User);
                _search_form.Topmost = false;
                _search_form.Owner = this;
                _search_form.ShowDialog();

                if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
                {
                    if (_search_form.Selected_ID != "" && _search_form.Selected_EntityType != "")
                    {
                        // Update List
                        FillEntityList(_search_form.Selected_EntityType, 1, 1);

                        // Select Entity...
                        cbBezugEntity_1.Text = _search_form.Selected_EntityType;
                        cbMSCRMEntity_1.Text = _search_form.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
                {
                    if (_search_form.Selected_ID != "" && _search_form.Selected_EntityType != "")
                    {
                        // Update List
                        FillEntityList(_search_form.Selected_EntityType, 1, 2);

                        // Select Entity...
                        cbBezugEntity_2.Text = _search_form.Selected_EntityType;
                        cbMSCRMEntity_2.Text = _search_form.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
                {
                    if (_search_form.Selected_ID != "" && _search_form.Selected_EntityType != "")
                    {
                        // Update List
                        FillEntityList(_search_form.Selected_EntityType, 1, 3);

                        // Select Entity...
                        cbBezugEntity_3.Text = _search_form.Selected_EntityType;
                        cbMSCRMEntity_3.Text = _search_form.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
                {
                    if (_search_form.Selected_ID != "" && _search_form.Selected_EntityType != "")
                    {
                        // Update List
                        FillEntityList(_search_form.Selected_EntityType, 1, 4);

                        // Select Entity...
                        cbBezugEntity_4.Text = _search_form.Selected_EntityType;
                        cbMSCRMEntity_4.Text = _search_form.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
                {
                    if (_search_form.Selected_ID != "" && _search_form.Selected_EntityType != "")
                    {
                        // Update List
                        FillEntityList(_search_form.Selected_EntityType, 1, 5);

                        // Select Entity...
                        cbBezugEntity_5.Text = _search_form.Selected_EntityType;
                        cbMSCRMEntity_5.Text = _search_form.Selected_ID;
                    }
                }
            }
            catch
            {
            }

        }

        private void btnAddToFavorites_Click(object sender, RoutedEventArgs e)
        {
            // AddToFavorites

            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;
            DevExpress.Xpf.Editors.ComboBoxEditItem _li = null;
            String _cbBezugEntity = "";

            if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
            {
                if (cbMSCRMEntity_1.SelectedItem != null && CheckSelectedEntity(_OrganizationService, 1, false)) { _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_1.SelectedItem; _cbBezugEntity = cbBezugEntity_1.Text.Trim().ToUpper(); }
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
            {
                if (cbMSCRMEntity_2.SelectedItem != null && CheckSelectedEntity(_OrganizationService, 2, false)) { _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_2.SelectedItem; _cbBezugEntity = cbBezugEntity_2.Text.Trim().ToUpper(); }
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
            {
                if (cbMSCRMEntity_3.SelectedItem != null && CheckSelectedEntity(_OrganizationService, 3, false)) { _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_3.SelectedItem; _cbBezugEntity = cbBezugEntity_3.Text.Trim().ToUpper(); }
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
            {
                if (cbMSCRMEntity_4.SelectedItem != null && CheckSelectedEntity(_OrganizationService, 4, false)) { _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_4.SelectedItem; _cbBezugEntity = cbBezugEntity_4.Text.Trim().ToUpper(); }
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
            {
                if (cbMSCRMEntity_5.SelectedItem != null && CheckSelectedEntity(_OrganizationService, 5, false)) { _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity_5.SelectedItem; _cbBezugEntity = cbBezugEntity_5.Text.Trim().ToUpper(); }
            }

            try
            {
                if (_li != null && _cbBezugEntity != "")
                {
                    String _tmp_User = _User;
                    _tmp_User = _tmp_User.Replace("\\", "_");
                    _tmp_User = _tmp_User.Replace(".", "_");

                    if (GetCacheItemByMSCRMID("FAVORITE_" + _tmp_User, _li.Tag.ToString()) == null)
                    {
                        DataRow _dr = _TimeEntryCache.NewRow();
                        _dr["id"] = Guid.NewGuid().ToString();
                        _dr["type"] = "FAVORITE_" + _tmp_User;
                        _dr["subtype"] = _cbBezugEntity;
                        _dr["mscrmid"] = _li.Tag.ToString();
                        _dr["mscrmparentid"] = "";
                        _dr["description"] = _li.Content.ToString().Trim();
                        _TimeEntryCache.Rows.Add(_dr);
                        _TimeEntryCache.AcceptChanges();

                        WriteCache();

                        staticItemInfo.Content = rm.GetString("SelectedReference");
                        // MessageBox.Show("Der ausgewählte Bezug wurde Ihrer Favoritenliste hinzugefügt!", "Hinweis", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show(rm.GetString("FavoritesList"), rm.GetString("Error"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show(rm.GetString("ValidReferences"), rm.GetString("Error"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch
            {
            }
        }

        private void Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            Hyperlink _hyperlink = (Hyperlink)sender;

            try
            {
                String _Url = _hyperlink.NavigateUri.ToString();

                _Url = _Url.Replace("[", "%7b");
                _Url = _Url.Replace("]", "%7d");
                System.Diagnostics.Process.Start("IEXPLORE.EXE", _Url);
            }
            catch
            {
            }
        }

        Double CellsTotalSum = 0.0;
        Double CellsGroupSum = 0.0;
        TimeSpan CellZusatzzeitGroupSum;
        TimeSpan CellZusatzzeitTotalSum;

        private void TimeEntryGrid_CustomSummary(object sender, DevExpress.Data.CustomSummaryEventArgs e)
        {
            if (!(
                ((GridSummaryItem)e.Item).FieldName == "dauer" ||
                ((GridSummaryItem)e.Item).FieldName == "gesamtdauer" ||
                ((GridSummaryItem)e.Item).FieldName == "zusatzzeit"
                ))
                return;

            if (e.IsTotalSummary)
            {
                if ((((GridSummaryItem)e.Item).FieldName == "dauer" || ((GridSummaryItem)e.Item).FieldName == "gesamtdauer"))
                {
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Start)
                    {
                        CellsTotalSum = 0.0;
                    }
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Calculate)
                    {
                        Double _value = 0.0;
                        try
                        {                           
                            if (e.FieldValue.ToString().Contains(","))
                            {
                                string tempField = e.FieldValue.ToString();
                                string temp = tempField.ToString().Replace(",", ".");
                                _value = Convert.ToDouble(temp, CultureInfo.GetCultureInfo("en-US").NumberFormat);
                            }
                            else
                            {
                                _value = Convert.ToDouble(e.FieldValue, CultureInfo.GetCultureInfo("en-US").NumberFormat);
                            }
                        }
                        catch
                        {
                        }
                        CellsTotalSum += _value;
                    }
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Finalize)
                    {
                        e.TotalValue = CellsTotalSum;
                    }
                }
                else
                {
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Start)
                    {
                        CellZusatzzeitTotalSum = new TimeSpan(0, 0, 0);
                    }
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Calculate)
                    {
                        String _tsp = e.FieldValue.ToString();

                        String[] _zz = _tsp.Split(':');
                        Int32 _zz_hours = 0;
                        Int32 _zz_minutes = 0;
                        try { _zz_hours = Convert.ToInt32(_zz[0]); }
                        catch { }
                        try { _zz_minutes = Convert.ToInt32(_zz[1]); }
                        catch { }

                        CellZusatzzeitTotalSum = CellZusatzzeitTotalSum.Add(new TimeSpan(_zz_hours, _zz_minutes, 0));
                    }
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Finalize)
                    {
                        e.TotalValue = CellZusatzzeitTotalSum.TotalHours.ToString("00") + ":" + CellZusatzzeitTotalSum.Minutes.ToString("00");
                    }
                }
            }

            if (e.IsGroupSummary)
            {
                if ((((GridSummaryItem)e.Item).FieldName == "dauer" || ((GridSummaryItem)e.Item).FieldName == "gesamtdauer"))
                {
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Start)
                    {
                        CellsGroupSum = 0;
                    }
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Calculate)
                    {
                        Double _value = 0.0;
                        try
                        {
                        if (e.FieldValue.ToString().Contains(","))
                            {
                                string tempField = e.FieldValue.ToString();
                                string temp = tempField.ToString().Replace(",", ".");
                                _value = Convert.ToDouble(temp, CultureInfo.GetCultureInfo("en-US").NumberFormat);
                            }
                            else
                            {
                                _value = Convert.ToDouble(e.FieldValue, CultureInfo.GetCultureInfo("en-US").NumberFormat);
                            }
                        }
                        catch
                        {
                        }
                        CellsGroupSum += _value;
                    }
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Finalize)
                    {                      
                        e.TotalValue = CellsGroupSum;
                    }
                }
                else
                {
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Start)
                    {
                        CellZusatzzeitGroupSum = new TimeSpan(0, 0, 0);
                    }
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Calculate)
                    {
                        String _tsp = e.FieldValue.ToString();

                        String[] _zz = _tsp.Split(':');
                        Int32 _zz_hours = 0;
                        Int32 _zz_minutes = 0;
                        try { _zz_hours = Convert.ToInt32(_zz[0]); }
                        catch { }
                        try { _zz_minutes = Convert.ToInt32(_zz[1]); }
                        catch { }

                        CellZusatzzeitGroupSum = CellZusatzzeitGroupSum.Add(new TimeSpan(_zz_hours, _zz_minutes, 0));
                    }
                    if (e.SummaryProcess == DevExpress.Data.CustomSummaryProcess.Finalize)
                    {
                        e.TotalValue = CellZusatzzeitGroupSum.TotalHours.ToString("00") + ":" + CellZusatzzeitGroupSum.Minutes.ToString("00");
                    }
                }

            }
        }

        private void TimeEntryGrid_CustomColumnSort(object sender, CustomColumnSortEventArgs e)
        {
            try
            {
                if (e.Column.FieldName == "datum")
                {
                    DataRow dr1 = (TimeEntryGrid.ItemsSource as DataTable).Rows[e.ListSourceRowIndex1];
                    DataRow dr2 = (TimeEntryGrid.ItemsSource as DataTable).Rows[e.ListSourceRowIndex2];

                    e.Handled = true;
                    if (dr1["datum"] != "")
                        e.Result = System.Collections.Comparer.Default.Compare(Convert.ToDateTime(dr1["datum"]),
                            Convert.ToDateTime(dr2["datum"]));
                }
            }
            catch { }
        }

        private void OnSyncMSCRMPreMonthItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            if (MSCRM_ConState == MSCRM_ConectionStates.Connected)
            {
                LoadTimeEntitiesFromMSCRM(_OrganizationService, true);
            }
        }

        private void OnHelpToolbarItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // Show Help


        }

        private void btnFavorites_Click(object sender, RoutedEventArgs e)
        {
            // Favoriten
            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;

            try
            {
                SelectFavorite _select_fav = new SelectFavorite(_TimeEntryCache, _User);
                _select_fav.Topmost = false;
                _select_fav.Owner = this;
                _select_fav.ShowDialog();

                if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
                {
                    if (_select_fav.Selected_ID != "" && _select_fav.Selected_EntityType != "")
                    {
                        // Select Entity...
                        cbBezugEntity_1.Text = _select_fav.Selected_EntityType;
                        cbMSCRMEntity_1.Text = _select_fav.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
                {
                    if (_select_fav.Selected_ID != "" && _select_fav.Selected_EntityType != "")
                    {
                        // Select Entity...
                        cbBezugEntity_2.Text = _select_fav.Selected_EntityType;
                        cbMSCRMEntity_2.Text = _select_fav.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
                {
                    if (_select_fav.Selected_ID != "" && _select_fav.Selected_EntityType != "")
                    {
                        // Select Entity...
                        cbBezugEntity_3.Text = _select_fav.Selected_EntityType;
                        cbMSCRMEntity_3.Text = _select_fav.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
                {
                    if (_select_fav.Selected_ID != "" && _select_fav.Selected_EntityType != "")
                    {
                        // Select Entity...
                        cbBezugEntity_4.Text = _select_fav.Selected_EntityType;
                        cbMSCRMEntity_4.Text = _select_fav.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
                {
                    if (_select_fav.Selected_ID != "" && _select_fav.Selected_EntityType != "")
                    {
                        // Select Entity...
                        cbBezugEntity_5.Text = _select_fav.Selected_EntityType;
                        cbMSCRMEntity_5.Text = _select_fav.Selected_ID;
                    }
                }

                WriteCache();
            }
            catch
            {
            }
        }

        private void btnAddTaetigkeitToFavorites_Click(object sender, RoutedEventArgs e)
        {
            // Tätigkeit zu den Favoriten hinzufügen

            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;

            String _cbTaetigkeit = "";

            if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
            {
                _cbTaetigkeit = cbTaetigkeit_1.Text.Trim();
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
            {
                _cbTaetigkeit = cbTaetigkeit_2.Text.Trim();
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
            {
                _cbTaetigkeit = cbTaetigkeit_3.Text.Trim();
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
            {
                _cbTaetigkeit = cbTaetigkeit_4.Text.Trim();
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
            {
                _cbTaetigkeit = cbTaetigkeit_5.Text.Trim();
            }

            try
            {
                if (_cbTaetigkeit != "")
                {
                    String _tmp_User = _User;
                    _tmp_User = _tmp_User.Replace("\\", "_");
                    _tmp_User = _tmp_User.Replace(".", "_");

                    if (GetCacheItemByDESCRIPTION("FAVORITE_ACTIVITY_" + _tmp_User, _cbTaetigkeit) == null)
                    {
                        DataRow _dr = _TimeEntryCache.NewRow();

                        _dr["id"] = Guid.NewGuid().ToString();
                        _dr["type"] = "FAVORITE_ACTIVITY_" + _tmp_User;
                        _dr["subtype"] = "";
                        _dr["mscrmid"] = "";
                        _dr["mscrmparentid"] = "";
                        _dr["description"] = _cbTaetigkeit;

                        _TimeEntryCache.Rows.Add(_dr);
                        _TimeEntryCache.AcceptChanges();

                        WriteCache();

                        if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
                        {
                            cbTaetigkeit_1.Items.Add(_cbTaetigkeit);
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
                        {
                            cbTaetigkeit_2.Items.Add(_cbTaetigkeit);
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
                        {
                            cbTaetigkeit_3.Items.Add(_cbTaetigkeit);
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
                        {
                            cbTaetigkeit_4.Items.Add(_cbTaetigkeit);
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
                        {
                            cbTaetigkeit_5.Items.Add(_cbTaetigkeit);
                        }

                        // MessageBox.Show("Die Tätigkeit '"+_cbTaetigkeit+"' wurde Ihrer Tätigkeiten-Favoritenliste hinzugefügt!", "Hinweis", MessageBoxButton.OK, MessageBoxImage.Information);
                        staticItemInfo.Content = rm.GetString("ActivityFavorite1") + " " + _cbTaetigkeit + " " + rm.GetString("ActivityFavorite2");
                    }
                    else
                    {
                        MessageBox.Show(rm.GetString("ActivityInList"), rm.GetString("Error"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show(rm.GetString("ActivityDropDown"), rm.GetString("Error"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch
            {

            }
        }

        private void btnDeleteTaetigkeitToFavorites_Click(object sender, RoutedEventArgs e)
        {
            // Remove current item from fav list

            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;

            String _cbTaetigkeit = "";

            if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
            {
                _cbTaetigkeit = cbTaetigkeit_1.Text.Trim();
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
            {
                _cbTaetigkeit = cbTaetigkeit_2.Text.Trim();
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
            {
                _cbTaetigkeit = cbTaetigkeit_3.Text.Trim();
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
            {
                _cbTaetigkeit = cbTaetigkeit_4.Text.Trim();
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
            {
                _cbTaetigkeit = cbTaetigkeit_5.Text.Trim();
            }

            try
            {
                if (_cbTaetigkeit != "")
                {
                    String _tmp_User = _User;
                    _tmp_User = _tmp_User.Replace("\\", "_");
                    _tmp_User = _tmp_User.Replace(".", "_");

                    DataRow _dr = null;
                    if ((_dr = GetCacheItemByDESCRIPTION("FAVORITE_ACTIVITY_" + _tmp_User, _cbTaetigkeit)) != null)
                    {
                        _TimeEntryCache.Rows.Remove(_dr);
                        _TimeEntryCache.AcceptChanges();

                        WriteCache();

                        if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
                        {
                            cbTaetigkeit_1.Items.Remove(_cbTaetigkeit);
                            cbTaetigkeit_1.Text = "";
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
                        {
                            cbTaetigkeit_2.Items.Remove(_cbTaetigkeit);
                            cbTaetigkeit_2.Text = "";
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
                        {
                            cbTaetigkeit_3.Items.Remove(_cbTaetigkeit);
                            cbTaetigkeit_3.Text = "";
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
                        {
                            cbTaetigkeit_4.Items.Remove(_cbTaetigkeit);
                            cbTaetigkeit_4.Text = "";
                        }
                        else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
                        {
                            cbTaetigkeit_5.Items.Remove(_cbTaetigkeit);
                            cbTaetigkeit_5.Text = "";
                        }

                        // MessageBox.Show("Die Tätigkeit '"+_cbTaetigkeit+"' wurde Ihrer Tätigkeiten-Favoritenliste hinzugefügt!", "Hinweis", MessageBoxButton.OK, MessageBoxImage.Information);
                        staticItemInfo.Content = rm.GetString("RemovedActivity1") + " " + _cbTaetigkeit + " " + rm.GetString("RemovedActivity2");
                    }
                    else
                    {
                        MessageBox.Show(rm.GetString("InternError"), rm.GetString("Error"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show(rm.GetString("ExistingActivity"), rm.GetString("Error"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch
            {
            }
        }

        /*
        private void OnSyncMSCRMDeleteItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // Delete all grid Items with status syncfrommscrm = true

            DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
            DataRow[] _result = _dt.Select("[syncfrommscrm]='true'");
            if (_result.Length > 0)
            {
                if (MessageBox.Show("Sind Sie sicher, dass Sie die von Dynamics CRM gesyncten Zeiteinträge aus dem Grid entfernen wollen?", "Dynamics CRM Zeiteinträge entfernen", MessageBoxButton.YesNoCancel, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    foreach (DataRow _dr in _result) _dt.Rows.Remove(_dr);
                }
            }
            else
            {
                MessageBox.Show("Es sind keine gesyncten Dynamics CRM Zeiteinträge im Grid vorhanden!", "Hinweis", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
        }
        */

        private void OnSyncMSCRMRefreshItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // refresh Dynamics CRM entities in cache

            if (MessageBox.Show(rm.GetString("(AP, ANF, VC)"), rm.GetString("(AP, ANF, VC)Title"), MessageBoxButton.YesNoCancel, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                // Clear PICKLIST.BILLING Cache
                DataRow[] _result = _TimeEntryCache.Select("[type]='WORKPACKAGE' OR [type]='INCIDENT' OR [type]='OPPORTUNITY' OR [type]='NEW_VERSION' OR [type]='CAMPAIGN' OR [type]='ACTIONITEM'");
                foreach (DataRow _row in _result)
                {
                    _TimeEntryCache.Rows.Remove(_row);
                }

                _LastSyncDateTime_WORKPACKAGE = Convert.ToDateTime("01/01/1900 00:00:00");
                _LastSyncDateTime_INCIDENT = Convert.ToDateTime("01/01/1900 00:00:00");
                _LastSyncDateTime_OPPORTUNITY = Convert.ToDateTime("01/01/1900 00:00:00");
                _LastSyncDateTime_CAMPAIGN = Convert.ToDateTime("01/01/1900 00:00:00");
                //_LastSyncDateTime_NEW_VERSION = Convert.ToDateTime("01/01/1900 00:00:00");
                _LastSyncDateTime_ACTIONITEM = Convert.ToDateTime("01/01/1900 00:00:00");

                WriteCache();

                _bFullSyncCompleted = false;
                MSCRM_ConState = MSCRM_ConectionStates.RequestConnection;

                // A24 JNE syncing is handled via Requesting Connection now
                //_workerSyncThread = new Thread(workerSyncThreadDoWork);
                //_workerSyncThread.Start();
            }
        }

        private void OnGoOnlineItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // Online-Mode
            _OrganizationService = null;

            // A24 JNE 24.9.2015 Connecting is now initiated via the Request Connection state.
            _bFullSyncCompleted = false;
            MSCRM_ConState = MSCRM_ConectionStates.RequestConnection;

            // A24 JNE 24.9.2015 syncing is now triggered in the  connectionsuccessful event of the connectionmanager
            //_workerSyncThread = new Thread(workerSyncThreadDoWork);
            //_workerSyncThread.Start();
        }

        private void OnGoOfflineItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // Offline-Mode

            if (MSCRM_ConState == MSCRM_ConectionStates.RequestConnection)
            {
                _workerSyncThread.Abort();
            }

            MSCRM_ConState = MSCRM_ConectionStates.NotConnected;
        }

        private void Hyperlink2_Click(object sender, RoutedEventArgs e)
        {
            Hyperlink _hyperlink = (Hyperlink)sender;

            try
            {
                String _Url = _hyperlink.NavigateUri.ToString();

                _Url = _Url.Replace("[", "%7b");
                _Url = _Url.Replace("]", "%7d");
                System.Diagnostics.Process.Start("IEXPLORE.EXE", _Url);
            }
            catch
            {
            }
        }

        private void TimeEntryGrid_CheckEdit_Checked(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;

            // Mark
            // Clear marked flag
            DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
            foreach (DataRow _dr in _dt.Rows)
            {
                _dr["marked"] = false;
            }

            for (int i = 0; i < TimeEntryGrid.VisibleRowCount; i++)
            {
                int rowHandle = TimeEntryGrid.GetRowHandleByVisibleIndex(i);

                TimeEntryGrid.SetCellValue(rowHandle, TimeEntryGrid.Columns["marked"], true);
            }

            TimeEntryGrid.RefreshData();

            this.Cursor = Cursors.Arrow;
        }

        private void TimeEntryGrid_CheckEdit_Unchecked(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;

            // Unmark
            for (int i = 0; i < TimeEntryGrid.VisibleRowCount; i++)
            {
                int rowHandle = TimeEntryGrid.GetRowHandleByVisibleIndex(i);

                TimeEntryGrid.SetCellValue(rowHandle, TimeEntryGrid.Columns["marked"], false);
            }

            TimeEntryGrid.RefreshData();

            this.Cursor = Cursors.Arrow;
        }

        private void OnFileLogOfTheDayItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // Show File Log of the Day...

            String logFilePath = System.IO.Directory.GetCurrentDirectory() + "\\logs";

            if (!System.IO.Directory.Exists(logFilePath)) System.IO.Directory.CreateDirectory(logFilePath);

            logFilePath += "\\logs" + DateTime.Now.Year.ToString() + RightStr("00" + DateTime.Now.Month.ToString(), 2) + RightStr("00" + DateTime.Now.Day.ToString(), 2) + ".txt";

            System.Diagnostics.Process.Start("notepad.exe", logFilePath);
        }

        private void OnConnectToMscrmItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            try
            {
                _fHelper.AskForConnection("ApplyConnectionToTabs");
                //{
                //    infoPanel = InformationPanel.GetInformationPanel(this, "Connecting...", 340, 120);
                //}
            }
            catch (Exception ex)
            {
                FileLog("ConnectToMscrm", "Error: " + ex.Message);
            }
        }

        private void OnManageConnectionsItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            try
            {
                _fHelper.AskForConnection("ApplyConnectionToTabs");
            }
            catch (Exception ex)
            {
                FileLog("ManageConnections", "Error: " + ex.Message);
            }
        }

        private void cbBezugEntity2_SelectedIndexChanged(object sender, RoutedEventArgs e)
        {
            DevExpress.Xpf.Editors.ComboBoxEdit _cb_edit = (DevExpress.Xpf.Editors.ComboBoxEdit)sender;

            String _TypeOfEntity = _cb_edit.Text.Trim().ToUpper();
            String _TOE_Name = "";
            if (_TypeOfEntity == WorkPackageString) _TOE_Name = "WORKPACKAGE";
            else if (_TypeOfEntity == OpportunityString) _TOE_Name = "OPPORTUNITY";
            else if (_TypeOfEntity == IncidentString) _TOE_Name = "INCIDENT";
            //else if (_TypeOfEntity == "REL") _TOE_Name = "NEW_VERSION";
            else if (_TypeOfEntity == CampaignString) _TOE_Name = "CAMPAIGN";
            else if (_TypeOfEntity == ActionItemString) _TOE_Name = "ACTIONITEM";

            if (_cb_edit.Name.Contains("_1"))
            {
                cbMSCRMEntity2_1.Text = ""; cbMSCRMEntity2_1.Items.Clear();
                if (_TypeOfEntity == OpportunityString || _TypeOfEntity == CampaignString)
                {
                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _li in cbAbrechnungsweise_1.Items)
                    {
                        if (_li.Content.ToString() == "Firmenintern") { cbAbrechnungsweise_1.SelectedItem = _li; break; }
                    }
                }


                //if (cbBezugEntity_1.Text.Trim().ToUpper() == Incident || cbBezugEntity_1.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity_1.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity_1.IsEnabled = false;
                //}

                //if (cbBezugEntity2_1.Text.Trim().ToUpper() == Incident || cbBezugEntity2_1.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity2_1.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity2_1.IsEnabled = false;
                //}

                // Fill DropDown
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                if (_result != null)
                {
                    cbMSCRMEntity2_1.Items.Clear();
                    foreach (DataRow _row in _result)
                    {
                        DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                        _li.Content = _row["description"].ToString();
                        _li.Tag = _row["mscrmid"].ToString();
                        cbMSCRMEntity2_1.Items.Add(_li);
                    }
                }
            }
            else if (_cb_edit.Name.Contains("_2"))
            {
                cbMSCRMEntity2_2.Text = ""; cbMSCRMEntity2_2.Items.Clear();
                if (_TypeOfEntity == OpportunityString)
                {
                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _li in cbAbrechnungsweise_2.Items)
                    {
                        if (_li.Content.ToString() == "Firmenintern") { cbAbrechnungsweise_2.SelectedItem = _li; break; }
                    }
                }

                //if (cbBezugEntity_2.Text.Trim().ToUpper() == Incident || cbBezugEntity_2.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity_2.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity_2.IsEnabled = false;
                //}

                //if (cbBezugEntity2_2.Text.Trim().ToUpper() == Incident || cbBezugEntity2_2.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity2_2.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity2_2.IsEnabled = false;
                //}

                // Fill DropDown
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                if (_result != null)
                {
                    cbMSCRMEntity2_2.Items.Clear();
                    foreach (DataRow _row in _result)
                    {
                        DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                        _li.Content = _row["description"].ToString();
                        _li.Tag = _row["mscrmid"].ToString();
                        cbMSCRMEntity2_2.Items.Add(_li);
                    }
                }
            }
            else if (_cb_edit.Name.Contains("_3"))
            {
                cbMSCRMEntity2_3.Text = ""; cbMSCRMEntity2_3.Items.Clear();
                if (_TypeOfEntity == OpportunityString)
                {
                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _li in cbAbrechnungsweise_3.Items)
                    {
                        if (_li.Content.ToString() == "Firmenintern") { cbAbrechnungsweise_3.SelectedItem = _li; break; }
                    }
                }

                //if (cbBezugEntity_3.Text.Trim().ToUpper() == Incident || cbBezugEntity_3.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity_3.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity_3.IsEnabled = false;
                //}

                //if (cbBezugEntity2_3.Text.Trim().ToUpper() == Incident || cbBezugEntity2_3.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity2_3.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity2_3.IsEnabled = false;
                //}

                // Fill DropDown
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                if (_result != null)
                {
                    cbMSCRMEntity2_3.Items.Clear();
                    foreach (DataRow _row in _result)
                    {
                        DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                        _li.Content = _row["description"].ToString();
                        _li.Tag = _row["mscrmid"].ToString();
                        cbMSCRMEntity2_3.Items.Add(_li);
                    }
                }
            }
            else if (_cb_edit.Name.Contains("_4"))
            {
                cbMSCRMEntity2_4.Text = ""; cbMSCRMEntity2_4.Items.Clear();
                if (_TypeOfEntity == OpportunityString)
                {
                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _li in cbAbrechnungsweise_4.Items)
                    {
                        if (_li.Content.ToString() == "Firmenintern") { cbAbrechnungsweise_4.SelectedItem = _li; break; }
                    }
                }

                //if (cbBezugEntity_4.Text.Trim().ToUpper() == Incident || cbBezugEntity_4.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity_4.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity_4.IsEnabled = false;
                //}

                //if (cbBezugEntity2_4.Text.Trim().ToUpper() == Incident || cbBezugEntity2_4.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity2_4.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity2_4.IsEnabled = false;
                //}

                // Fill DropDown
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                if (_result != null)
                {
                    cbMSCRMEntity2_4.Items.Clear();
                    foreach (DataRow _row in _result)
                    {
                        DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                        _li.Content = _row["description"].ToString();
                        _li.Tag = _row["mscrmid"].ToString();
                        cbMSCRMEntity2_4.Items.Add(_li);
                    }
                }
            }
            else if (_cb_edit.Name.Contains("_5"))
            {
                cbMSCRMEntity2_5.Text = ""; cbMSCRMEntity2_5.Items.Clear();
                if (_TypeOfEntity == OpportunityString)
                {
                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _li in cbAbrechnungsweise_5.Items)
                    {
                        if (_li.Content.ToString() == "Firmenintern") { cbAbrechnungsweise_5.SelectedItem = _li; break; }
                    }
                }

                //if (cbBezugEntity_5.Text.Trim().ToUpper() == Incident || cbBezugEntity_5.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity_5.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity_5.IsEnabled = false;
                //}

                //if (cbBezugEntity2_5.Text.Trim().ToUpper() == Incident || cbBezugEntity2_5.Text.Trim().ToUpper() == "REL")
                //{
                //    btAddNewEntity2_5.IsEnabled = true;
                //}
                //else
                //{
                //    btAddNewEntity2_5.IsEnabled = false;
                //}

                // Fill DropDown
                DataRow[] _result = _TimeEntryCache.Select("[type]='" + _TOE_Name + "'");
                if (_result != null)
                {
                    cbMSCRMEntity2_5.Items.Clear();
                    foreach (DataRow _row in _result)
                    {
                        DevExpress.Xpf.Editors.ComboBoxEditItem _li = new DevExpress.Xpf.Editors.ComboBoxEditItem();
                        _li.Content = _row["description"].ToString();
                        _li.Tag = _row["mscrmid"].ToString();
                        cbMSCRMEntity2_5.Items.Add(_li);
                    }
                }
            }

        }

        private void cbMSCRMEntity2_EditValueChanged(object sender, DevExpress.Xpf.Editors.EditValueChangedEventArgs e)
        {
            try
            {
                DevExpress.Xpf.Editors.ComboBoxEdit _cb_edit = (DevExpress.Xpf.Editors.ComboBoxEdit)sender;

                // Check Entity
                try
                {
                    if (_cb_edit.Name.Contains("_1"))
                    {
                        CheckSelectedEntity2(_OrganizationService, 1, true);
                    }
                    else if (_cb_edit.Name.Contains("_2"))
                    {
                        CheckSelectedEntity2(_OrganizationService, 2, true);
                    }
                    else if (_cb_edit.Name.Contains("_3"))
                    {
                        CheckSelectedEntity2(_OrganizationService, 3, true);
                    }
                    else if (_cb_edit.Name.Contains("_4"))
                    {
                        CheckSelectedEntity2(_OrganizationService, 4, true);
                    }
                    else if (_cb_edit.Name.Contains("_5"))
                    {
                        CheckSelectedEntity2(_OrganizationService, 5, true);
                    }
                }
                catch (Exception ex)
                {
                    FileLog("cbMSCRMEntity_EditValueChanged@CheckSelectedEntity", "Error: " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                FileLog("cbMSCRMEntity_EditValueChanged", "Error: " + ex.Message);
            }
        }

        private Boolean CheckSelectedEntity2(IOrganizationService _Service, Int32 ActiveTE, Boolean DoAbrechnungsweise)
        {
            try
            {
                // Check Entity
                Entity_OK2 = false;
                DevExpress.Xpf.Editors.ComboBoxEditItem _selected_li = null;

                if (ActiveTE == 1)
                {
                    _selected_li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_1.SelectedItem;
                    if (_selected_li != null) Entity_OK2 = TestMSCRMEntity(_Service, _selected_li.Tag.ToString(), cbBezugEntity2_1.Text.Trim().ToUpper());
                    if (Entity_OK2) { btnMSCRMEntityInfoImage2_1.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/OK.png"); btnMSCRMEntityInfoImage2_1.ToolTip = rm.GetString("OK.png"); }
                    else { btnMSCRMEntityInfoImage2_1.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/NOK.png"); btnMSCRMEntityInfoImage2_1.ToolTip = rm.GetString("NOK.png"); }

                    if (DoAbrechnungsweise)
                    {
                        if (Entity_OK2 && MSCRM_ConState == MSCRM_ConectionStates.Connected)
                        {
                            Guid _MSCRMBezugID = Guid.Empty;
                            DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                            _MSCRMBezugID = new Guid(_li.Tag.ToString());

                            Boolean _is_failed = true;

                            // Abrechnugsweise übernehmen (AP, ANF) 
                            String Sync_Entity_Type = cbBezugEntity2_1.Text.Trim().ToUpper();
                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (wp != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_1.Items)
                                    {
                                        if (Convert.ToInt32(_item.Tag) == wp.a24_billing_opt.Value)
                                        {
                                            cbAbrechnungsweise_1.SelectedItem = _item; _is_failed = false; break;
                                        }
                                    }
                                }
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (_in != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_1.Items)
                                    {
                                        if (Convert.ToInt32(_item.Tag) == _in.a24_billing_opt.Value)
                                        {
                                            cbAbrechnungsweise_1.SelectedItem = _item; _is_failed = false; break;
                                        }
                                    }
                                }
                            }

                            if (_is_failed)
                            {
                                // Abrechnugsweise offline übernehmen (AP, ANF)
                                Sync_Entity_Type = cbBezugEntity2_1.Text.Trim().ToUpper();
                                String Sync_Type = "";
                                DataRow _dr = null;

                                if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                                else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                                if (Sync_Type != "")
                                {
                                    _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                    if (_li != null)
                                    {
                                        if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                        {
                                            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_1.Items)
                                            {
                                                if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                                {
                                                    cbAbrechnungsweise_1.SelectedItem = _item; break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Abrechnugsweise offline übernehmen (AP, ANF)
                            String Sync_Entity_Type = cbBezugEntity2_1.Text.Trim().ToUpper();
                            String Sync_Type = "";
                            DataRow _dr = null;

                            if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                            else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                            if (Sync_Type != "")
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                if (_li != null)
                                {
                                    if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                    {
                                        foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_1.Items)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                            {
                                                cbAbrechnungsweise_1.SelectedItem = _item; break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ActiveTE == 2)
                {
                    _selected_li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_2.SelectedItem;
                    if (_selected_li != null) Entity_OK2 = TestMSCRMEntity(_Service, _selected_li.Tag.ToString(), cbBezugEntity2_2.Text.Trim().ToUpper());
                    if (Entity_OK2) { btnMSCRMEntityInfoImage2_2.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/OK.png"); btnMSCRMEntityInfoImage2_2.ToolTip = rm.GetString("OK.png"); }
                    else { btnMSCRMEntityInfoImage2_2.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/NOK.png"); btnMSCRMEntityInfoImage2_2.ToolTip = rm.GetString("NOK.png"); }

                    if (DoAbrechnungsweise)
                    {
                        if (Entity_OK2 && MSCRM_ConState == MSCRM_ConectionStates.Connected)
                        {
                            Guid _MSCRMBezugID = Guid.Empty;

                            DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_2.SelectedItem;
                            _MSCRMBezugID = new Guid(_li.Tag.ToString());

                            Boolean _is_failed = true;

                            // Abrechnugsweise übernehmen (AP, ANF) 
                            String Sync_Entity_Type = cbBezugEntity2_2.Text.Trim().ToUpper();
                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (wp != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_2.Items)
                                    {
                                        if (Convert.ToInt32(_item.Tag) == wp.a24_billing_opt.Value)
                                        {
                                            cbAbrechnungsweise_2.SelectedItem = _item; _is_failed = false; break;
                                        }
                                    }
                                }
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (_in != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_2.Items)
                                    {
                                        if (Convert.ToInt32(_item.Tag) == _in.a24_billing_opt.Value)
                                        {
                                            cbAbrechnungsweise_2.SelectedItem = _item; _is_failed = false; break;
                                        }
                                    }
                                }
                            }

                            if (_is_failed)
                            {
                                // Abrechnugsweise offline übernehmen (AP, ANF)
                                Sync_Entity_Type = cbBezugEntity2_2.Text.Trim().ToUpper();
                                String Sync_Type = "";
                                DataRow _dr = null;

                                if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                                else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                                if (Sync_Type != "")
                                {
                                    _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                    if (_li != null)
                                    {
                                        if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                        {
                                            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_2.Items)
                                            {
                                                if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                                {
                                                    cbAbrechnungsweise_2.SelectedItem = _item; break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Abrechnugsweise offline übernehmen (AP, ANF)
                            String Sync_Entity_Type = cbBezugEntity2_2.Text.Trim().ToUpper();
                            String Sync_Type = "";
                            DataRow _dr = null;

                            if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                            else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                            if (Sync_Type != "")
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                if (_li != null)
                                {
                                    if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                    {
                                        foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_2.Items)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                            {
                                                cbAbrechnungsweise_2.SelectedItem = _item; break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ActiveTE == 3)
                {
                    _selected_li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_3.SelectedItem;
                    if (_selected_li != null) Entity_OK2 = TestMSCRMEntity(_Service, _selected_li.Tag.ToString(), cbBezugEntity2_3.Text.Trim().ToUpper());
                    if (Entity_OK2) { btnMSCRMEntityInfoImage2_3.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/OK.png"); btnMSCRMEntityInfoImage2_3.ToolTip = rm.GetString("OK.png"); }
                    else { btnMSCRMEntityInfoImage2_3.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/NOK.png"); btnMSCRMEntityInfoImage2_3.ToolTip = rm.GetString("NOK.png"); }

                    if (DoAbrechnungsweise)
                    {
                        if (Entity_OK2 && MSCRM_ConState == MSCRM_ConectionStates.Connected)
                        {
                            Guid _MSCRMBezugID = Guid.Empty;

                            DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_3.SelectedItem;
                            _MSCRMBezugID = new Guid(_li.Tag.ToString());

                            Boolean _is_failed = true;

                            // Abrechnugsweise übernehmen (AP, ANF) 
                            String Sync_Entity_Type = cbBezugEntity2_3.Text.Trim().ToUpper();
                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (wp != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_3.Items)
                                    {
                                        if (Convert.ToInt32(_item.Tag) == wp.a24_billing_opt.Value)
                                        {
                                            cbAbrechnungsweise_3.SelectedItem = _item; _is_failed = false; break;
                                        }
                                    }
                                }
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (_in != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_3.Items)
                                    {
                                        if (Convert.ToInt32(_item.Tag) == _in.a24_billing_opt.Value)
                                        {
                                            cbAbrechnungsweise_3.SelectedItem = _item; _is_failed = false; break;
                                        }
                                    }
                                }
                            }

                            if (_is_failed)
                            {
                                // Abrechnugsweise offline übernehmen (AP, ANF)
                                Sync_Entity_Type = cbBezugEntity2_3.Text.Trim().ToUpper();
                                String Sync_Type = "";
                                DataRow _dr = null;

                                if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                                else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                                if (Sync_Type != "")
                                {
                                    _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                    if (_li != null)
                                    {
                                        if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                        {
                                            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_3.Items)
                                            {
                                                if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                                {
                                                    cbAbrechnungsweise_3.SelectedItem = _item; break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Abrechnugsweise offline übernehmen (AP, ANF)
                            String Sync_Entity_Type = cbBezugEntity2_3.Text.Trim().ToUpper();
                            String Sync_Type = "";
                            DataRow _dr = null;

                            if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                            else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                            if (Sync_Type != "")
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                if (_li != null)
                                {
                                    if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                    {
                                        foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_3.Items)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                            {
                                                cbAbrechnungsweise_3.SelectedItem = _item; break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ActiveTE == 4)
                {
                    _selected_li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_4.SelectedItem;
                    if (_selected_li != null) Entity_OK2 = TestMSCRMEntity(_Service, _selected_li.Tag.ToString(), cbBezugEntity2_4.Text.Trim().ToUpper());
                    if (Entity_OK2) { btnMSCRMEntityInfoImage2_4.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/OK.png"); btnMSCRMEntityInfoImage2_4.ToolTip = rm.GetString("OK.png"); }
                    else { btnMSCRMEntityInfoImage2_4.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/NOK.png"); btnMSCRMEntityInfoImage2_4.ToolTip = rm.GetString("NOK.png"); }

                    if (DoAbrechnungsweise)
                    {
                        if (Entity_OK2 && MSCRM_ConState == MSCRM_ConectionStates.Connected)
                        {
                            Guid _MSCRMBezugID = Guid.Empty;

                            DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_4.SelectedItem;
                            _MSCRMBezugID = new Guid(_li.Tag.ToString());

                            Boolean _is_failed = true;

                            // Abrechnugsweise übernehmen (AP, ANF) 
                            String Sync_Entity_Type = cbBezugEntity2_4.Text.Trim().ToUpper();
                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (wp != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_4.Items)
                                    {
                                        if (Convert.ToInt32(_item.Tag) == wp.a24_billing_opt.Value)
                                        {
                                            cbAbrechnungsweise_4.SelectedItem = _item; _is_failed = false; break;
                                        }
                                    }
                                }
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (_in != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_4.Items)
                                    {
                                        if (Convert.ToInt32(_item.Tag) == _in.a24_billing_opt.Value)
                                        {
                                            cbAbrechnungsweise_4.SelectedItem = _item; _is_failed = false; break;
                                        }
                                    }
                                }
                            }

                            if (_is_failed)
                            {
                                // Abrechnugsweise offline übernehmen (AP, ANF)
                                Sync_Entity_Type = cbBezugEntity2_4.Text.Trim().ToUpper();
                                String Sync_Type = "";
                                DataRow _dr = null;

                                if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                                else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                                if (Sync_Type != "")
                                {
                                    _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                    if (_li != null)
                                    {
                                        if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                        {
                                            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_4.Items)
                                            {
                                                if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                                {
                                                    cbAbrechnungsweise_4.SelectedItem = _item; break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Abrechnugsweise offline übernehmen (AP, ANF)
                            String Sync_Entity_Type = cbBezugEntity2_4.Text.Trim().ToUpper();
                            String Sync_Type = "";
                            DataRow _dr = null;

                            if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                            else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                            if (Sync_Type != "")
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                if (_li != null)
                                {
                                    if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                    {
                                        foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_4.Items)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                            {
                                                cbAbrechnungsweise_4.SelectedItem = _item; break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else if (ActiveTE == 5)
                {
                    _selected_li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_5.SelectedItem;
                    if (_selected_li != null) Entity_OK2 = TestMSCRMEntity(_Service, _selected_li.Tag.ToString(), cbBezugEntity2_5.Text.Trim().ToUpper());
                    if (Entity_OK2) { btnMSCRMEntityInfoImage2_5.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/OK.png"); btnMSCRMEntityInfoImage2_5.ToolTip = rm.GetString("OK.png"); }
                    else { btnMSCRMEntityInfoImage2_5.Source = doGetImageSourceFromResource("QuickTimeEntry", "Images/16x16neu/NOK.png"); btnMSCRMEntityInfoImage2_5.ToolTip = rm.GetString("NOK.png"); }

                    if (DoAbrechnungsweise)
                    {
                        if (Entity_OK2 && MSCRM_ConState == MSCRM_ConectionStates.Connected)
                        {
                            Guid _MSCRMBezugID = Guid.Empty;

                            DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_5.SelectedItem;
                            _MSCRMBezugID = new Guid(_li.Tag.ToString());

                            Boolean _is_failed = true;

                            // Abrechnugsweise übernehmen (AP, ANF) 
                            String Sync_Entity_Type = cbBezugEntity2_5.Text.Trim().ToUpper();
                            if (Sync_Entity_Type == WorkPackageString)
                            {
                                a24_workpackage wp = (a24_workpackage)_Service.Retrieve(a24_workpackage.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (wp != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_5.Items)
                                    {
                                        if (Convert.ToInt32(_item.Tag) == wp.a24_billing_opt.Value)
                                        {
                                            cbAbrechnungsweise_5.SelectedItem = _item; _is_failed = false; break;
                                        }
                                    }
                                }
                            }
                            else if (Sync_Entity_Type == IncidentString)
                            {
                                Incident _in = (Incident)_Service.Retrieve(Incident.EntityLogicalName, _MSCRMBezugID, new ColumnSet(true));
                                if (_in != null)
                                {
                                    foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_5.Items)
                                    {
                                        if (Convert.ToInt32(_item.Tag) == _in.a24_billing_opt.Value)
                                        {
                                            cbAbrechnungsweise_5.SelectedItem = _item; _is_failed = false; break;
                                        }
                                    }
                                }
                            }

                            if (_is_failed)
                            {
                                // Abrechnugsweise offline übernehmen (AP, ANF)
                                Sync_Entity_Type = cbBezugEntity2_5.Text.Trim().ToUpper();
                                String Sync_Type = "";
                                DataRow _dr = null;

                                if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                                else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                                if (Sync_Type != "")
                                {
                                    _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                    if (_li != null)
                                    {
                                        if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                        {
                                            foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_5.Items)
                                            {
                                                if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                                {
                                                    cbAbrechnungsweise_5.SelectedItem = _item; break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            // Abrechnugsweise offline übernehmen (AP, ANF)
                            String Sync_Entity_Type = cbBezugEntity2_5.Text.Trim().ToUpper();
                            String Sync_Type = "";
                            DataRow _dr = null;

                            if (Sync_Entity_Type == WorkPackageString) Sync_Type = "WORKPACKAGE";
                            else if (Sync_Entity_Type == IncidentString) Sync_Type = "INCIDENT";

                            if (Sync_Type != "")
                            {
                                DevExpress.Xpf.Editors.ComboBoxEditItem _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)_selected_li;
                                if (_li != null)
                                {
                                    if ((_dr = GetCacheItemByMSCRMID(Sync_Type, _li.Tag.ToString())) != null)
                                    {
                                        foreach (DevExpress.Xpf.Editors.ComboBoxEditItem _item in cbAbrechnungsweise_5.Items)
                                        {
                                            if (Convert.ToInt32(_item.Tag) == Convert.ToInt32(_dr["subtype"].ToString()))
                                            {
                                                cbAbrechnungsweise_5.SelectedItem = _item; break;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                return Entity_OK2;
            }
            catch (Exception ex)
            {
                FileLog("CheckSelectedEntity(2)", "Error: " + ex.Message);
                return false;
            }
        }

        private void btnExtendedSearch2_Click(object sender, RoutedEventArgs e)
        {
            // Extended Entity Search
            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;

            try
            {
                EntitySearch _search_form = new EntitySearch(_TimeEntryCache, _User);
                _search_form.Topmost = false;
                _search_form.Owner = this;
                _search_form.ShowDialog();

                if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
                {
                    if (_search_form.Selected_ID != "" && _search_form.Selected_EntityType != "")
                    {
                        // Update List
                        FillEntityList(_search_form.Selected_EntityType, 2, 1);

                        // Select Entity...
                        cbBezugEntity2_1.Text = _search_form.Selected_EntityType;
                        cbMSCRMEntity2_1.Text = _search_form.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
                {
                    if (_search_form.Selected_ID != "" && _search_form.Selected_EntityType != "")
                    {
                        // Update List
                        FillEntityList(_search_form.Selected_EntityType, 2, 2);

                        // Select Entity...
                        cbBezugEntity2_2.Text = _search_form.Selected_EntityType;
                        cbMSCRMEntity2_2.Text = _search_form.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
                {
                    if (_search_form.Selected_ID != "" && _search_form.Selected_EntityType != "")
                    {
                        // Update List
                        FillEntityList(_search_form.Selected_EntityType, 2, 3);

                        // Select Entity...
                        cbBezugEntity2_3.Text = _search_form.Selected_EntityType;
                        cbMSCRMEntity2_3.Text = _search_form.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
                {
                    if (_search_form.Selected_ID != "" && _search_form.Selected_EntityType != "")
                    {
                        // Update List
                        FillEntityList(_search_form.Selected_EntityType, 2, 4);

                        // Select Entity...
                        cbBezugEntity2_4.Text = _search_form.Selected_EntityType;
                        cbMSCRMEntity2_4.Text = _search_form.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
                {
                    if (_search_form.Selected_ID != "" && _search_form.Selected_EntityType != "")
                    {
                        // Update List
                        FillEntityList(_search_form.Selected_EntityType, 2, 5);

                        // Select Entity...
                        cbBezugEntity2_5.Text = _search_form.Selected_EntityType;
                        cbMSCRMEntity2_5.Text = _search_form.Selected_ID;
                    }
                }
            }
            catch
            {
            }

        }

        private void btnFavorites2_Click(object sender, RoutedEventArgs e)
        {

            // Favoriten
            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;

            try
            {
                SelectFavorite _select_fav = new SelectFavorite(_TimeEntryCache, _User);
                _select_fav.Topmost = false;
                _select_fav.Owner = this;
                _select_fav.ShowDialog();

                if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
                {
                    if (_select_fav.Selected_ID != "" && _select_fav.Selected_EntityType != "")
                    {
                        // Select Entity...
                        cbBezugEntity2_1.Text = _select_fav.Selected_EntityType;
                        cbMSCRMEntity2_1.Text = _select_fav.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
                {
                    if (_select_fav.Selected_ID != "" && _select_fav.Selected_EntityType != "")
                    {
                        // Select Entity...
                        cbBezugEntity2_2.Text = _select_fav.Selected_EntityType;
                        cbMSCRMEntity2_2.Text = _select_fav.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
                {
                    if (_select_fav.Selected_ID != "" && _select_fav.Selected_EntityType != "")
                    {
                        // Select Entity...
                        cbBezugEntity2_3.Text = _select_fav.Selected_EntityType;
                        cbMSCRMEntity2_3.Text = _select_fav.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
                {
                    if (_select_fav.Selected_ID != "" && _select_fav.Selected_EntityType != "")
                    {
                        // Select Entity...
                        cbBezugEntity2_4.Text = _select_fav.Selected_EntityType;
                        cbMSCRMEntity2_4.Text = _select_fav.Selected_ID;
                    }
                }
                else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
                {
                    if (_select_fav.Selected_ID != "" && _select_fav.Selected_EntityType != "")
                    {
                        // Select Entity...
                        cbBezugEntity2_5.Text = _select_fav.Selected_EntityType;
                        cbMSCRMEntity2_5.Text = _select_fav.Selected_ID;
                    }
                }

                WriteCache();
            }
            catch
            {
            }

        }

        private void btnAddToFavorites2_Click(object sender, RoutedEventArgs e)
        {
            // AddToFavorites (Bezug 2)

            TabItem _ti = (TabItem)TimeEntryTabCtrl.SelectedItem;
            DevExpress.Xpf.Editors.ComboBoxEditItem _li = null;
            String _cbBezugEntity = "";

            if (_ti.Header.ToString() == rm.GetString("tabTE_1"))
            {
                if (cbMSCRMEntity2_1.SelectedItem != null && CheckSelectedEntity2(_OrganizationService, 1, false)) { _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_1.SelectedItem; _cbBezugEntity = cbBezugEntity2_1.Text.Trim().ToUpper(); }
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_2"))
            {
                if (cbMSCRMEntity2_2.SelectedItem != null && CheckSelectedEntity2(_OrganizationService, 2, false)) { _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_2.SelectedItem; _cbBezugEntity = cbBezugEntity2_2.Text.Trim().ToUpper(); }
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_3"))
            {
                if (cbMSCRMEntity2_3.SelectedItem != null && CheckSelectedEntity2(_OrganizationService, 3, false)) { _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_3.SelectedItem; _cbBezugEntity = cbBezugEntity2_3.Text.Trim().ToUpper(); }
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_4"))
            {
                if (cbMSCRMEntity2_4.SelectedItem != null && CheckSelectedEntity2(_OrganizationService, 4, false)) { _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_4.SelectedItem; _cbBezugEntity = cbBezugEntity2_4.Text.Trim().ToUpper(); }
            }
            else if (_ti.Header.ToString() == rm.GetString("tabTE_5"))
            {
                if (cbMSCRMEntity2_5.SelectedItem != null && CheckSelectedEntity2(_OrganizationService, 5, false)) { _li = (DevExpress.Xpf.Editors.ComboBoxEditItem)cbMSCRMEntity2_5.SelectedItem; _cbBezugEntity = cbBezugEntity2_5.Text.Trim().ToUpper(); }
            }

            try
            {
                if (_li != null && _cbBezugEntity != "")
                {
                    String _tmp_User = _User;
                    _tmp_User = _tmp_User.Replace("\\", "_");
                    _tmp_User = _tmp_User.Replace(".", "_");

                    if (GetCacheItemByMSCRMID("FAVORITE_" + _tmp_User, _li.Tag.ToString()) == null)
                    {
                        DataRow _dr = _TimeEntryCache.NewRow();
                        _dr["id"] = Guid.NewGuid().ToString();
                        _dr["type"] = "FAVORITE_" + _tmp_User;
                        _dr["subtype"] = _cbBezugEntity;
                        _dr["mscrmid"] = _li.Tag.ToString();
                        _dr["mscrmparentid"] = "";
                        _dr["description"] = _li.Content.ToString().Trim();
                        _TimeEntryCache.Rows.Add(_dr);
                        _TimeEntryCache.AcceptChanges();

                        WriteCache();

                        staticItemInfo.Content = rm.GetString("SelectedReference");
                        // MessageBox.Show("Der ausgewählte Bezug wurde Ihrer Favoritenliste hinzugefügt!", "Hinweis", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        MessageBox.Show(rm.GetString("FavoritesList"), rm.GetString("Error"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                    }
                }
                else
                {
                    MessageBox.Show(rm.GetString("ValidReferences"), rm.GetString("Error"), MessageBoxButton.OK, MessageBoxImage.Exclamation);
                }
            }
            catch
            {
            }

        }

        // A24 JNE 17.9.2015 - Add New Entity was removed after Definition of RHE 16.9.2015

        //private void btAddNewEntity_Click(object sender, RoutedEventArgs e)
        //{
        //}

        private void btnSaveAndNewTimeentry_Click(object sender, RoutedEventArgs e)
        {
            // Speichern und neu
            SaveActiveTimeEntry(true, true);
        }

        private void OnExpandAllItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // Expand All
            TimeEntryGrid.ExpandAllGroups();
        }

        private void OnExpandNoneItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            // Expand None
            TimeEntryGrid.CollapseAllGroups();
        }

        public class DateUtils
        {
            /// <summary>
            /// Verwaltet die Daten einer Kalenderwoche
            /// </summary>
            public class CalendarWeek
            {
                /// <summary>
                /// Das Jahr
                /// </summary>
                public int Year;

                /// <summary>
                /// Die Kalenderwoche
                /// </summary>
                public int Week;

                /// <summary>
                /// Konstruktor
                /// </summary>
                /// <param name="year">Das Jahr</param>
                /// <param name="week">Die Kalenderwoche</param>
                public CalendarWeek(int year, int week)
                {
                    this.Year = year;
                    this.Week = week;
                }
            }

            /// <summary>
            /// Berechnet die Kalenderwoche eines internationalen Datums
            /// </summary>
            /// <param name="date">Das Datum</param>
            /// <returns>Gibt ein CalendarWeek-Objekt zurück</returns>
            /// <remarks>
            /// Diese Methode berechnet die Kalenderwoche eines Datums
            /// nach der GetWeekOfYear-Methode eines Calendar-Objekts
            /// und korrigiert den darin enthaltenen Fehler.
            /// </remarks>
            public static CalendarWeek GetCalendarWeek(DateTime date)
            {
                // Aktuelle Kultur ermitteln
                CultureInfo currentCulture = CultureInfo.CurrentCulture;

                // Aktuellen Kalender ermitteln
                System.Globalization.Calendar calendar = currentCulture.Calendar;

                // Kalenderwoche über das Calendar-Objekt ermitteln
                int calendarWeek = calendar.GetWeekOfYear(date,
                   currentCulture.DateTimeFormat.CalendarWeekRule,
                   currentCulture.DateTimeFormat.FirstDayOfWeek);

                if (calendarWeek > 52)
                {
                    date = date.AddDays(7);
                    int testCalendarWeek = calendar.GetWeekOfYear(date,
                       currentCulture.DateTimeFormat.CalendarWeekRule,
                       currentCulture.DateTimeFormat.FirstDayOfWeek);
                    if (testCalendarWeek == 2)
                        calendarWeek = 1;
                }

                // Das Jahr der Kalenderwoche ermitteln
                int year = date.Year;
                if (calendarWeek == 1 && date.Month == 12)
                    year++;
                if (calendarWeek >= 52 && date.Month == 1)
                    year--;

                // Die ermittelte Kalenderwoche zurückgeben
                return new CalendarWeek(year, calendarWeek);
            }


            /// <summary>
            /// Berechnet die Kalenderwoche eines deutschen Datums
            /// </summary>
            /// <param name="date">Das Datum</param>
            /// <returns>Gibt ein CalendarWeek-Objekt zurück</returns>
            /// <remarks>
            /// <para>
            /// Diese Methode gilt nur für die deutsche Kultur.
            /// Sie ist wesentlich schneller als die Methode
            /// <see cref="GetInternationalCalendarWeek"/>.
            /// </para>
            /// <para>
            /// Die Berechnung erfolgt nach dem 
            /// C++-Algorithmus von Ekkehard Hess aus einem Beitrag vom
            /// 29.7.1999 in der Newsgroup 
            /// borland.public.cppbuilder.language
            ///(freigegeben zur allgemeinen Verwendung)
            /// </para>
            /// </remarks>
            public static CalendarWeek GetGermanCalendarWeek(DateTime date)
            {
                double a = Math.Floor((14 - (date.Month)) / 12D);
                double y = date.Year + 4800 - a;
                double m = (date.Month) + (12 * a) - 3;

                double jd = date.Day + Math.Floor(((153 * m) + 2) / 5) +
                   (365 * y) + Math.Floor(y / 4) - Math.Floor(y / 100) +
                   Math.Floor(y / 400) - 32045;

                double d4 = (jd + 31741 - (jd % 7)) % 146097 % 36524 %
                   1461;
                double L = Math.Floor(d4 / 1460);
                double d1 = ((d4 - L) % 365) + L;

                // Kalenderwoche ermitteln
                int calendarWeek = (int)Math.Floor(d1 / 7) + 1;

                // Das Jahr der Kalenderwoche ermitteln
                int year = date.Year;
                if (calendarWeek == 1 && date.Month == 12)
                    year++;
                if (calendarWeek >= 52 && date.Month == 1)
                    year--;

                // Die ermittelte Kalenderwoche zurückgeben
                return new CalendarWeek(year, calendarWeek);
            }
        }

        private void CopyTimeEntry()
        {
            try
            {
                String _te_datum = TimeEntryGrid.GetFocusedRowCellValue("datum").ToString().Trim();

                String _te_dauer = "";
                String _te_zusatzzeit = TimeEntryGrid.GetFocusedRowCellValue("zusatzzeit").ToString().Trim();

                Double _dauer = Convert.ToDouble(TimeEntryGrid.GetFocusedRowCellValue("dauer").ToString().Trim());
                _dauer = _dauer * 60.0;

                TimeSpan _dauer_gemessen = new TimeSpan(0, (int)_dauer, 0);
                _te_dauer = _dauer_gemessen.Hours.ToString("00") + ":" + _dauer_gemessen.Minutes.ToString("00");

                CopyTimeEntry _copytimeentry_form = new CopyTimeEntry(Convert.ToDateTime(_te_datum, CultureInfo.CurrentCulture.DateTimeFormat), _te_dauer, _te_zusatzzeit);

                _copytimeentry_form.Topmost = false;
                _copytimeentry_form.Owner = this;
                _copytimeentry_form.ShowDialog();

                if (_copytimeentry_form.DoCopyTimeEntry)
                {
                    String _te_guid = TimeEntryGrid.GetFocusedRowCellValue("id").ToString().Trim();
                    if (_te_guid != "") ActivateTimeEntry(_te_guid, true, _copytimeentry_form.Datum, _copytimeentry_form.Dauer, _copytimeentry_form.Zusatzzeit);
                }

            }
            catch (Exception ex)
            {
                FileLog("CopyTimeEntry", "Error: " + ex.Message);
                // MessageBox.Show("TimeEntryGrid_MouseDoubleClick: " + ex.Message, "Fehlermeldung", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private void TimeEntryGrid_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            CopyTimeEntry();
        }


        /// <summary>
        /// Initialisiert die ConnectionManager Klasse und hinterlegt ihre Events.
        /// Events sind
        /// - ConnectionSucceed
        /// - ConnectionFailed
        /// - RequestPassword
        /// </summary>
        private void ManageConnectionControl()
        {
            this._cManager = new ConnectionManager();

            _cManager = new ConnectionManager();
            _cManager.RequestPassword += (sender, e) => _fHelper.RequestPassword(e.ConnectionDetail);
            _cManager.ConnectionSucceed += (sender, e) =>
            {
                _bIsConnecting = false;

                MSCRM_ConState = MSCRM_ConectionStates.Connected;

                _currentConnectionDetail = e.ConnectionDetail;
                _OrganizationService = e.OrganizationService;

                SetCurrentUser();

                // if the Full Sync has not been completed, start it.
                if (!_bFullSyncCompleted)
                {
                    _workerSyncThread = new Thread(workerSyncThreadDoWork);
                    // the ApartmentState of the worker thread needs to be set to STA so the Connection Manager works.
                    _workerSyncThread.SetApartmentState(ApartmentState.STA);
                    _workerSyncThread.Start();
                }
            };
            _cManager.ConnectionFailed += (sender, e) =>
            {
                _bIsConnecting = false;

                MSCRM_ConState = MSCRM_ConectionStates.Error;

                MessageBox.Show(this, e.FailureReason, "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.OK);

                _currentConnectionDetail = null;
                _OrganizationService = null;
            };

            _fHelper = new FormHelper(this, _cManager);
        }

        /// <summary>
        /// Sets the current user by loading data from mscrm
        /// </summary>
        private void SetCurrentUser()
        {
            try
            {
                // A24 JNE 24.9.2015 Check for orgService to avoid exception
                if (_OrganizationService != null)
                {
                    object[] UserDtls = CRMManager.GetCurrentUser(_OrganizationService, _TimeEntryCache, false, _currentConnectionDetail);
                    gDefaultUser = (Guid)UserDtls[0];
                    sDefaultUser = (string)UserDtls[1];

                    MSCRM_ConState_ErrInfo = "";
                }
            }
            catch (Exception ex)
            {
                MSCRM_ConState_ErrInfo = ex.Message;
                MSCRM_ConState = MSCRM_ConectionStates.Error;
                gDefaultUser = Guid.Empty;
                sDefaultUser = "";
            }
        }

        private void OnLanguageItemClick(object sender, DevExpress.Xpf.Bars.ItemClickEventArgs e)
        {
            Language lng = new Language();
            lng.ShowDialog();

            if (lng.DialogResult.HasValue && lng.DialogResult.Value == true)
            {
                TimeEntryGrid.TotalSummary[0].DisplayFormat = "{0} " + rm.GetString("syncedLabel");
                TimeEntryGrid.GroupSummary[0].DisplayFormat = "{0} " + rm.GetString("TimeItems");

                if (Entity_OK)
                {
                    btnMSCRMEntityInfoImage_1.ToolTip = rm.GetString("OK.png");
                    btnMSCRMEntityInfoImage_2.ToolTip = rm.GetString("OK.png");
                    btnMSCRMEntityInfoImage_3.ToolTip = rm.GetString("OK.png");
                    btnMSCRMEntityInfoImage_4.ToolTip = rm.GetString("OK.png");
                    btnMSCRMEntityInfoImage_5.ToolTip = rm.GetString("OK.png");
                }
                else
                {
                    btnMSCRMEntityInfoImage_1.ToolTip = rm.GetString("NOK.png");
                    btnMSCRMEntityInfoImage_2.ToolTip = rm.GetString("NOK.png");
                    btnMSCRMEntityInfoImage_3.ToolTip = rm.GetString("NOK.png");
                    btnMSCRMEntityInfoImage_4.ToolTip = rm.GetString("NOK.png");
                    btnMSCRMEntityInfoImage_5.ToolTip = rm.GetString("NOK.png");
                }

                if (Entity_OK2)
                {
                    btnMSCRMEntityInfoImage2_1.ToolTip = rm.GetString("OK.png");
                    btnMSCRMEntityInfoImage2_2.ToolTip = rm.GetString("OK.png");
                    btnMSCRMEntityInfoImage2_3.ToolTip = rm.GetString("OK.png");
                    btnMSCRMEntityInfoImage2_4.ToolTip = rm.GetString("OK.png");
                    btnMSCRMEntityInfoImage2_5.ToolTip = rm.GetString("OK.png");
                }
                else
                {
                    btnMSCRMEntityInfoImage2_1.ToolTip = rm.GetString("NOK.png");
                    btnMSCRMEntityInfoImage2_2.ToolTip = rm.GetString("NOK.png");
                    btnMSCRMEntityInfoImage2_3.ToolTip = rm.GetString("NOK.png");
                    btnMSCRMEntityInfoImage2_4.ToolTip = rm.GetString("NOK.png");
                    btnMSCRMEntityInfoImage2_5.ToolTip = rm.GetString("NOK.png");
                }


                //Read reference dropdown values from config file
                List<String> _BezugValuesCopy = new List<String>(_BezugValues);

                LoadRegardingEntities();

                int i1 = cbBezugEntity_1.SelectedIndex;
                int i2 = cbBezugEntity_2.SelectedIndex;
                int i3 = cbBezugEntity_3.SelectedIndex;
                int i4 = cbBezugEntity_4.SelectedIndex;
                int i5 = cbBezugEntity_5.SelectedIndex;

                int i21 = cbBezugEntity2_1.SelectedIndex;
                int i22 = cbBezugEntity2_2.SelectedIndex;
                int i23 = cbBezugEntity2_3.SelectedIndex;
                int i24 = cbBezugEntity2_4.SelectedIndex;
                int i25 = cbBezugEntity2_5.SelectedIndex;

                _BezugValues = new String[]{
                                            EmptyString              // ""
                                           , WorkPackageString        // "AP"
                                           , OpportunityString         // "VC"
                                           , IncidentString            // "ANF"               
                                           , CampaignString           // "KA"
                                           , ActionItemString         // "AIL"
                                            };
                if (i1 > -1)
                {
                    cbBezugEntity_1.Text = String.Empty;
                    cbBezugEntity_1.Text = _BezugValues[i1];
                }
                if (i2 > -1)
                {
                    cbBezugEntity_2.Text = String.Empty;
                    cbBezugEntity_2.Text = _BezugValues[i2];
                }
                if (i3 > -1)
                {
                    cbBezugEntity_3.Text = String.Empty;
                    cbBezugEntity_3.Text = _BezugValues[i3];
                }
                if (i4 > -1)
                {
                    cbBezugEntity_4.Text = String.Empty;
                    cbBezugEntity_4.Text = _BezugValues[i4];
                }
                if (i5 > -1)
                {
                    cbBezugEntity_5.Text = String.Empty;
                    cbBezugEntity_5.Text = _BezugValues[i5];
                }

                if (i21 > -1) cbBezugEntity2_1.Text = (string)cbBezugEntity2_1.Items[i21];
                if (i22 > -1) cbBezugEntity2_2.Text = (string)cbBezugEntity2_2.Items[i22];
                if (i23 > -1) cbBezugEntity2_3.Text = (string)cbBezugEntity2_3.Items[i23];
                if (i24 > -1) cbBezugEntity2_4.Text = (string)cbBezugEntity2_4.Items[i24];
                if (i25 > -1) cbBezugEntity2_5.Text = (string)cbBezugEntity2_5.Items[i25];

                cbBezugEntity_1.Items.Clear();
                cbBezugEntity_2.Items.Clear();
                cbBezugEntity_3.Items.Clear();
                cbBezugEntity_4.Items.Clear();
                cbBezugEntity_5.Items.Clear();
                foreach (var item in _BezugValues)
                {
                    cbBezugEntity_1.Items.Add(item);
                    cbBezugEntity_2.Items.Add(item);
                    cbBezugEntity_3.Items.Add(item);
                    cbBezugEntity_4.Items.Add(item);
                    cbBezugEntity_5.Items.Add(item);
                }

                //Change values in DataTable
                DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
                foreach (DataRow _dr in _dt.Rows)
                {
                    if (_BezugValuesCopy.Contains(_dr["tekz"]))
                    {
                        int i = _BezugValuesCopy.IndexOf(_dr["tekz"].ToString());
                        _dr["tekz"] = _BezugValues[i];
                    }
                }

                deDatum_1.Mask = (CultureInfo.GetCultureInfo(CultureInfo.CurrentCulture.ToString()).DateTimeFormat.ShortDatePattern);
                deDatum_2.Mask = (CultureInfo.GetCultureInfo(CultureInfo.CurrentCulture.ToString()).DateTimeFormat.ShortDatePattern);
                deDatum_3.Mask = (CultureInfo.GetCultureInfo(CultureInfo.CurrentCulture.ToString()).DateTimeFormat.ShortDatePattern);
                deDatum_4.Mask = (CultureInfo.GetCultureInfo(CultureInfo.CurrentCulture.ToString()).DateTimeFormat.ShortDatePattern);
                deDatum_5.Mask = (CultureInfo.GetCultureInfo(CultureInfo.CurrentCulture.ToString()).DateTimeFormat.ShortDatePattern);

                SetDateForAppropriateLanguage();
            }
        }

        private void SetDateForAppropriateLanguage()
        {
            DataTable _dt = (DataTable)TimeEntryGrid.ItemsSource;
            foreach (DataRow _dr in _dt.Rows)
            {
                var temp = _dr["datum"].ToString();
                DateTime? time = null;

                if (temp.Contains(".") && CultureInfo.CurrentCulture.Name.Equals("en-US"))
                {
                    time = Convert.ToDateTime(temp, CultureInfo.GetCultureInfo("de-DE").DateTimeFormat);
                }

                if (temp.Contains(@"/") && CultureInfo.CurrentCulture.Name.Equals("de-DE"))
                {
                    time = Convert.ToDateTime(temp, CultureInfo.GetCultureInfo("en-US").DateTimeFormat);
                }

                if(time!=null)
                {
                    _dr["datum"] = time.Value.ToShortDateString();
                }    
            }
        
            TimeEntryGrid.RefreshData();
        }
    }
}