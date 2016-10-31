using System;
using System.Net;
using System.Windows;

namespace McTools.Xrm.Connection.WpfForms
{
    public class FormHelper
    {
        public FormHelper(Window innerAppForm, ConnectionManager connectionManager)
        {
            _innerAppForm = innerAppForm;
            _connectionManager = connectionManager;
        }

        private readonly Window _innerAppForm;
        private readonly ConnectionManager _connectionManager;

        /// <summary>
        /// Checks the existence of a user password and returns it
        /// </summary>
        /// <param name="detail">Details of the Crm connection</param>
        /// <returns>True if password defined</returns>
        public bool RequestPassword(ConnectionDetail detail)
        {
            if (!string.IsNullOrEmpty(detail.UserPassword))
                return true;

            bool returnValue = false;

            var pForm = new PasswordForm
                {
                    UserLogin = detail.UserName,
                    UserDomain = detail.UserDomain,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner
                };

            var result = pForm.ShowDialog();

            if (result != null && result.Value)
            {
                detail.UserPassword = pForm.UserPassword;
                detail.SavePassword = pForm.SavePassword;
                returnValue = true;
            }

            return returnValue;
        }

       

        /// <summary>
        /// Asks this manager to select a Crm connection to use
        /// </summary>
        [STAThreadAttribute]
        public void AskForConnection(object connectionParameter)
        {
            var cs = new ConnectionSelector(_connectionManager.ConnectionsList)
            {
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ConnectionList = _connectionManager.LoadConnectionsList()
            };

            var result = cs.ShowDialog();

            if (result != null && result.Value)
            {
                _connectionManager.ConnectToServer(cs.SelectedConnection, connectionParameter);

                //if (cs.HadCreatedNewConnection)
                //{
                //    _connectionManager.ConnectionsList.Connections.Add(cs.SelectedConnection);
                //    _connectionManager.SaveConnectionsFile(_connectionManager.ConnectionsList);
                //}
            }
        }


        /// <summary>
        /// Creates or updates a Crm connection
        /// </summary>
        /// <param name="isCreation">Indicates if it is a connection creation</param>
        /// <param name="connectionToUpdate">Details of the connection to update</param>
        /// <returns>Created or updated connection</returns>
        public ConnectionDetail EditConnection(bool isCreation, ConnectionDetail connectionToUpdate)
        {
            var cForm = new ConnectionForm(isCreation) { WindowStartupLocation = WindowStartupLocation.CenterOwner };

            if (!isCreation)
            {
                cForm.ConnectionDetail = connectionToUpdate;
            }

            var result = cForm.ShowDialog();

            if (result != null && result.Value)
            {
                if (isCreation)
                {
                    _connectionManager.ConnectionsList.Connections.Add(cForm.ConnectionDetail);
                }
                else
                {
                    foreach (ConnectionDetail detail in _connectionManager.ConnectionsList.Connections)
                    {
                        if (detail.ConnectionId == cForm.ConnectionDetail.ConnectionId)
                        {
                            #region Update connection details

                            detail.ConnectionName = cForm.ConnectionDetail.ConnectionName;
                            detail.OrganizationServiceUrl = cForm.ConnectionDetail.OrganizationServiceUrl;
                            detail.CrmTicket = cForm.ConnectionDetail.CrmTicket;
                            detail.IsCustomAuth = cForm.ConnectionDetail.IsCustomAuth;
                            detail.Organization = cForm.ConnectionDetail.Organization;
                            detail.OrganizationFriendlyName = cForm.ConnectionDetail.OrganizationFriendlyName;
                            detail.ServerName = cForm.ConnectionDetail.ServerName;
                            detail.ServerPort = cForm.ConnectionDetail.ServerPort;
                            detail.UseIfd = cForm.ConnectionDetail.UseIfd;
                            detail.UseOnline = cForm.ConnectionDetail.UseOnline;
                            detail.UseOsdp = cForm.ConnectionDetail.UseOsdp;
                            detail.UserDomain = cForm.ConnectionDetail.UserDomain;
                            detail.UserName = cForm.ConnectionDetail.UserName;
                            detail.UserPassword = cForm.ConnectionDetail.UserPassword;
                            detail.UseSsl = cForm.ConnectionDetail.UseSsl;
                            detail.HomeRealmUrl = cForm.ConnectionDetail.HomeRealmUrl;

                            #endregion
                        }
                    }
                }

                if (cForm.DoConnect)
                {
                    _connectionManager.ConnectToServer(cForm.ConnectionDetail);
                }

                _connectionManager.SaveConnectionsFile(_connectionManager.ConnectionsList);

                return cForm.ConnectionDetail;
            }

            return null;
        }

        /// <summary>
        /// Deletes a Crm connection from the connections list
        /// </summary>
        /// <param name="connectionToDelete">Details of the connection to delete</param>
        public void DeleteConnection(ConnectionDetail connectionToDelete)
        {
            _connectionManager.ConnectionsList.Connections.Remove(connectionToDelete);
            _connectionManager.SaveConnectionsFile(_connectionManager.ConnectionsList);
        }
    }
}
