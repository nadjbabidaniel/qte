using System;
using System.Collections.Generic;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Discovery;
using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;
using System.Reflection;
using System.Resources;

namespace McTools.Xrm.Connection.WpfForms
{
    public partial class ConnectionForm
    {
        #region Variables

        private ConnectionDetail _connectionDetail;

        private bool _doConnect;

        private readonly bool _isCreationMode;

        private List<OrganizationDetail> _organizations;

        #endregion Variables

        #region Constructors

        public ConnectionForm()
        {
            InitializeComponent();

            _isCreationMode = true;
            rdbIntegrated.IsChecked = true;
        }

        public ConnectionForm(bool isCreation)
        {
            InitializeComponent();

            _isCreationMode = isCreation;
            rdbIntegrated.IsChecked = true;
        }

        #endregion Constructors

        #region Properties

        public CrmConnections ConnectionList { get; set; }

        public ConnectionDetail ConnectionDetail { get { return _connectionDetail; } set { _connectionDetail = value; } }

        public bool DoConnect
        {
            get { return _doConnect; }
        }

        private static ResourceManager rm = new ResourceManager("McTools.Xrm.Connection.WpfForms.ConnectionForm", Assembly.GetExecutingAssembly());

        #endregion Properties

        #region Handlers

        #region Settings behavior

        private void ChkUseIfdChecked(object sender, RoutedEventArgs e)
        {
            chkUseSsl.IsChecked = true;
            chkUseOnline.IsEnabled = false;
            chkUserOsdp.IsEnabled = false;

            rdbCustom.IsChecked = true;
            rdbCustom.IsEnabled = false;
            rdbIntegrated.IsEnabled = false;
        }

        private void ChkUseIfdUnchecked(object sender, RoutedEventArgs e)
        {
            chkUseSsl.IsChecked = false;
            chkUseOnline.IsEnabled = true;
            chkUserOsdp.IsEnabled = true;

            rdbCustom.IsEnabled = true;
            rdbIntegrated.IsEnabled = true;
        }

        private void ChkUseOnlineChecked(object sender, RoutedEventArgs e)
        {
            chkUseIfd.IsEnabled = false;
            chkUseSsl.IsEnabled = false;
            chkUseSsl.IsChecked = true;

            rdbCustom.IsChecked = true;
            rdbCustom.IsEnabled = false;
            rdbIntegrated.IsEnabled = false;
            txtDomain.IsEnabled = false;

            txtServerPort.IsEnabled = false;
            txtHomeRealmUri.IsEnabled = false;
        }

        private void ChkUseOnlineUnchecked(object sender, RoutedEventArgs e)
        {
            chkUseIfd.IsEnabled = true;
            chkUseSsl.IsEnabled = true;
            chkUseSsl.IsChecked = false;

            rdbCustom.IsEnabled = true;
            rdbIntegrated.IsEnabled = true;

            txtDomain.IsEnabled = true;

            txtServerPort.IsEnabled = true;
            txtHomeRealmUri.IsEnabled = true;
        }

        private void ChkUserOsdpChecked(object sender, RoutedEventArgs e)
        {
            chkUseOnline.IsChecked = true;
            chkUseOnline.IsEnabled = false;
        }

        private void ChkUserOsdpUnchecked(object sender, RoutedEventArgs e)
        {
            chkUseOnline.IsEnabled = true;
        }

        private void RdbIntegratedChecked(object sender, RoutedEventArgs e)
        {
            txtDomain.IsEnabled = false;
            txtLogin.IsEnabled = false;
            txtPassword.IsEnabled = false;
            chkSavePassword.IsEnabled = false;
        }

        private void RdbCustomChecked(object sender, RoutedEventArgs e)
        {
            txtDomain.IsEnabled = true;
            txtLogin.IsEnabled = true;
            txtPassword.IsEnabled = true;
            chkSavePassword.IsEnabled = true;
        }

        #endregion Settings behavior

        private void WindowLoaded1(object sender, RoutedEventArgs e)
        {
            if (_connectionDetail != null)
            {
                FillValues();
            }
        }

        private void BtnGetOrgsClick(object sender, RoutedEventArgs e)
        {
            GetOrganizations();
        }

        private void BtnValidateClick(object sender, RoutedEventArgs e)
        {
            Validate();
            //DialogResult = true;
            Close();
        }

        private void BtnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        #endregion Handlers

        #region Methods

        private void GetOrganizations()
        {
            string warningMessage = string.Empty;
            bool goodAuthenticationData = false;
            bool goodServerData = false;

            // Check data filled by user
            if (rdbIntegrated.IsChecked != null && rdbIntegrated.IsChecked.Value ||
                (
                rdbCustom.IsChecked != null && rdbCustom.IsChecked.Value &&
                (txtDomain.Text.Length > 0
                || chkUseIfd.IsChecked != null && chkUseIfd.IsChecked.Value
                || chkUseOnline.IsChecked != null && chkUseOnline.IsChecked.Value) &&
                txtLogin.Text.Length > 0 &&
                txtPassword.Password.Length > 0
                ))
                goodAuthenticationData = true;

            if (txtServerName.GetValue(TextBox.TextProperty).ToString().Length > 0)
                goodServerData = true;

            if (!goodServerData)
            {
                warningMessage += rm.GetString("ServerName") + "\r\n";
            }
            if (!goodAuthenticationData)
            {
                warningMessage += rm.GetString("AuthenticationTextboxes") + "\r\n";
            }

            if (warningMessage.Length > 0)
            {
                MessageBox.Show(this, warningMessage, rm.GetString("Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            else
            {
                // Save connection details in structure
                if (_isCreationMode || _connectionDetail == null)
                {
                    _connectionDetail = new ConnectionDetail();
                }

                _connectionDetail.ConnectionName = txtConnectionName.Text;
                _connectionDetail.IsCustomAuth = rdbCustom.IsChecked != null && rdbCustom.IsChecked.Value;
                _connectionDetail.UseSsl = chkUseSsl.IsChecked != null && chkUseSsl.IsChecked.Value;
                _connectionDetail.UseOnline = chkUseOnline.IsChecked != null && chkUseOnline.IsChecked.Value;
                _connectionDetail.UseOsdp = chkUserOsdp.IsChecked != null && chkUserOsdp.IsChecked.Value;
                _connectionDetail.ServerName = txtServerName.Text;
                _connectionDetail.ServerPort = txtServerPort.Text;
                _connectionDetail.UserDomain = txtDomain.Text;
                _connectionDetail.UserName = txtLogin.Text;
                _connectionDetail.UserPassword = txtPassword.Password;
                _connectionDetail.SavePassword = chkSavePassword.IsChecked != null && chkSavePassword.IsChecked.Value;
                _connectionDetail.UseIfd = chkUseIfd.IsChecked != null && chkUseIfd.IsChecked.Value;
                _connectionDetail.HomeRealmUrl = (txtHomeRealmUri.Text.Length > 0 ? txtHomeRealmUri.Text : null);

                _connectionDetail.AuthType = AuthenticationProviderType.ActiveDirectory;

                if (chkUseIfd.IsChecked.Value)
                {
                    _connectionDetail.AuthType = AuthenticationProviderType.Federation;
                }
                else if (chkUserOsdp.IsChecked.Value)
                {
                    _connectionDetail.AuthType = AuthenticationProviderType.OnlineFederation;
                }
                else if (chkUseOnline.IsChecked.Value)
                {
                    _connectionDetail.AuthType = AuthenticationProviderType.LiveId;
                }

                // Launch organization retrieval
                cbbOrganizations.Items.Clear();
                //Thread retrieveOrgThread = new Thread(new ThreadStart(FillOrganizations));
                //retrieveOrgThread.Start();
                FillOrganizations();
            }
        }

        private void Validate()
        {
            if (txtConnectionName.Text.Length == 0)
            {
                MessageBox.Show(this, rm.GetString("NameSpecify"), rm.GetString("Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (cbbOrganizations.SelectedItem == null
                && !(chkUseIfd.IsChecked != null && chkUseIfd.IsChecked.Value
                || chkUserOsdp.IsChecked != null && chkUserOsdp.IsChecked.Value))
            {
                MessageBox.Show(this, rm.GetString("OrganizationSelect"), rm.GetString("Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (txtPassword.Password.Length == 0 && rdbCustom.IsChecked != null && rdbCustom.IsChecked.Value)
            {
                MessageBox.Show(this, rm.GetString("EnterPassword"), rm.GetString("Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (_connectionDetail == null)
                _connectionDetail = new ConnectionDetail();

            // Save connection details in structure
            _connectionDetail.ConnectionName = txtConnectionName.Text;
            _connectionDetail.IsCustomAuth = rdbCustom.IsChecked != null && rdbCustom.IsChecked.Value;
            _connectionDetail.UseSsl = chkUseSsl.IsChecked != null && chkUseSsl.IsChecked.Value;
            _connectionDetail.UseOnline = chkUseOnline.IsChecked != null && chkUseOnline.IsChecked.Value;
            _connectionDetail.UseOsdp = chkUserOsdp.IsChecked != null && chkUserOsdp.IsChecked.Value;
            _connectionDetail.ServerName = txtServerName.Text;
            _connectionDetail.ServerPort = txtServerPort.Text;
            _connectionDetail.UserDomain = txtDomain.Text;
            _connectionDetail.UserName = txtLogin.Text;
            _connectionDetail.UserPassword = txtPassword.Password;
            _connectionDetail.SavePassword = chkSavePassword.IsChecked != null && chkSavePassword.IsChecked.Value;
            _connectionDetail.UseIfd = chkUseIfd.IsChecked != null && chkUseIfd.IsChecked.Value;
            _connectionDetail.HomeRealmUrl = (txtHomeRealmUri.Text.Length > 0 ? txtHomeRealmUri.Text : null);

            try
            {
                FillDetails();

                if (MessageBox.Show(this, rm.GetString("ServerConn"), rm.GetString("Question"), MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    _doConnect = true;
                }

                DialogResult = true;
                Close();
            }
            catch (Exception error)
            {
                MessageBox.Show(this, error.Message, rm.GetString("Error"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FillDetails()
        {
            bool hasFoundOrg = false;

            OrganizationDetail selectedOrganization = cbbOrganizations.SelectedItem != null ? ((OrgDetail)cbbOrganizations.SelectedItem).Detail: null;

            if (_organizations == null || _organizations.Count == 0)
            {
                FillOrganizations();
            }

            if (_organizations != null && selectedOrganization != null)
                foreach (var organization in _organizations)
                {
                    if (organization.UniqueName == selectedOrganization.UniqueName)
                    {
                        _connectionDetail.OrganizationServiceUrl =
                            organization.Endpoints[EndpointType.OrganizationService];
                        _connectionDetail.Organization = organization.UniqueName;
                        _connectionDetail.OrganizationFriendlyName = organization.FriendlyName;

                        _connectionDetail.ConnectionName = txtConnectionName.Text;

                        if (_isCreationMode)
                        {
                            _connectionDetail.ConnectionId = Guid.NewGuid();
                        }

                        hasFoundOrg = true;

                        break;
                    }
                }

            if (!hasFoundOrg)
            {
                throw new Exception(rm.GetString("UnableToMatch"));
            }
        }

        public void FillOrganizations()
        {
            try
            {
                Guid applicationId = new Guid("6f4cad4a-e3d4-41d7-8ac7-cd17a69c3997");// Guid.NewGuid();//

                //// SetCursor(this, Cursors.WaitCursor);


                WebRequest.GetSystemWebProxy();

                var connection = CrmConnection.Parse(_connectionDetail.GetDiscoveryCrmConnectionString());
                var service = new DiscoveryService(connection);

                var request = new RetrieveOrganizationsRequest();
                var response = (RetrieveOrganizationsResponse)service.Execute(request);

                _organizations = new List<OrganizationDetail>();

                foreach (OrganizationDetail orgDetail in response.Details)
                {
                    _organizations.Add(orgDetail);

                    cbbOrganizations.Items.Add(new OrgDetail{Detail = orgDetail});
                }

                cbbOrganizations.SelectedIndex = 0;

                //var sc = new ServerConnection();

                //sc.config = sc.GetServerConfiguration(_connectionDetail, applicationId);

                //WebRequest.DefaultWebProxy = WebRequest.GetSystemWebProxy();

                //using (var serviceProxy = new DiscoveryServiceProxy(sc.config.DiscoveryUri, sc.config.HomeRealmUri, sc.config.Credentials, sc.config.DeviceCredentials))
                //{
                    
                //    _organizations = new List<OrganizationDetail>();

                //    // Obtain information about the organizations that the system user belongs to.
                //    var orgsCollection = sc.DiscoverOrganizations(serviceProxy);

                //    foreach (var orgDetail in orgsCollection)
                //    {
                //        _organizations.Add(orgDetail);

                //        cbbOrganizations.Items.Add(new OrgDetail{Detail = orgDetail});

                //        //AddItem(comboBoxOrganizations, new Organization() { OrganizationDetail = orgDetail });

                //        //SelectUniqueValue(comboBoxOrganizations);
                //    }

                //    cbbOrganizations.SelectedIndex = 0;
                //}

                //// SetEnableState(comboBoxOrganizations, true);



            }
            catch (Exception error)
            {
                string errorMessage = CrmExceptionHelper.GetErrorMessage(error, false);

                MessageBox.Show(errorMessage);

                //if (error.InnerException != null)
                //DisplayMessageBox(this, "Error while retrieving organizations : " + errorMessage + "\r\n" + error.InnerException.Message + "\r\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //else
                //DisplayMessageBox(this, "Error while retrieving organizations : " + errorMessage + "\r\n" + error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //SetCursor(this, Cursors.Default);
            }
        }

        private void FillValues()
        {
            rdbCustom.IsChecked = _connectionDetail.IsCustomAuth;

            txtConnectionName.Text = _connectionDetail.ConnectionName;
            txtServerName.Text = _connectionDetail.ServerName;
            txtServerPort.Text = _connectionDetail.ServerPort;
            txtDomain.Text = _connectionDetail.UserDomain;
            txtLogin.Text = _connectionDetail.UserName;
            txtPassword.Password= _connectionDetail.UserPassword;
            chkSavePassword.IsChecked = _connectionDetail.SavePassword;
            cbbOrganizations.Text = _connectionDetail.OrganizationFriendlyName;
            chkUseIfd.IsChecked = _connectionDetail.UseIfd;
            chkUserOsdp.IsChecked = _connectionDetail.UseOsdp;
            chkUseOnline.IsChecked = _connectionDetail.UseOnline;
            chkUseSsl.IsChecked = _connectionDetail.UseSsl;
            txtHomeRealmUri.Text = _connectionDetail.HomeRealmUrl;

            txtHomeRealmUri.IsEnabled = _connectionDetail.UseIfd;
        }

        #endregion Methods
    }
}
