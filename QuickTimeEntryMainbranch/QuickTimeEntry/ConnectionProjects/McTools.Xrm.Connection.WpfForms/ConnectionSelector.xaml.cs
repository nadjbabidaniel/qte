using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace McTools.Xrm.Connection.WpfForms
{
    /// <summary>
    /// Logique d'interaction pour ConnectionSelector.xaml
    /// </summary>
    public partial class ConnectionSelector : Window
    {
        #region Variables

        /// <summary>
        /// Obtient la connexion sélectionnée
        /// </summary>
        public ConnectionDetail SelectedConnection { get; private set; }

        //public bool HadCreatedNewConnection { get; private set; }

        private readonly ObservableCollection<ConnectionItem> _connectionCollection;

        public CrmConnections ConnectionList { get; set; }

        #endregion

        #region Constuctor

        /// <summary>
        /// Initializes a new instance of class ConnectionSelector
        /// </summary>
        public ConnectionSelector()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Initializes a new instance of class ConnectionSelector
        /// </summary>
        /// <param name="connections">Available connections list</param>
        public ConnectionSelector(CrmConnections connections)
        {
            connections.Connections.Sort();
            _connectionCollection = new ObservableCollection<ConnectionItem>();

            foreach (var detail in connections.Connections)
            {
                _connectionCollection.Add(new ConnectionItem
                   {
                       ConnectionInfo = detail,
                       Name = detail.ConnectionName,
                       Organization = detail.OrganizationFriendlyName,
                       Server = detail.ServerName
                   });
            }

            InitializeComponent();
        }

        #endregion Constuctor

        #region Properties

        public ObservableCollection<ConnectionItem> ConnectionCollection { get { return _connectionCollection; } }

        #endregion  Properties

        #region Handlers

        private void CreateButtonClick(object sender, RoutedEventArgs e)
        {
            ConnectionForm cForm = new ConnectionForm(true);
            cForm.ConnectionList = ConnectionList;
            //cForm.StartPosition = FormStartPosition.CenterParent;

            //if(true)
            if ((bool)cForm.ShowDialog())
            {
                SelectedConnection = cForm.ConnectionDetail;
                //HadCreatedNewConnection = true;

                _connectionCollection.Add(new ConnectionItem
                    {
                        ConnectionInfo = SelectedConnection,
                        Name = SelectedConnection.ConnectionName,
                        Organization = SelectedConnection.OrganizationFriendlyName,
                        Server = SelectedConnection.ServerName
                    });

                
                
                lvConnections.SelectedItems.Clear();

                for (int i = 0; i < lvConnections.Items.Count; i++)
                {
                    if (((ConnectionItem)lvConnections.Items[i]).Name == SelectedConnection.ConnectionName)
                    {
                        lvConnections.SelectedIndex = i;
                        break;
                    }
                }

                ValidateButtonClick(null, null);
            }
            else
            {
                CancelButtonClick(null, null);
            }
        }

        private void EditButtonClick(object sender, RoutedEventArgs e)
        {
            if (lvConnections.SelectedItems.Count < 1)
                return;

            

            ConnectionForm cForm = new ConnectionForm(false);
            cForm.ConnectionDetail = ((ConnectionItem)lvConnections.SelectedItems[0]).ConnectionInfo;
            cForm.ConnectionList = ConnectionList;
            //cForm.StartPosition = FormStartPosition.CenterParent;

            //if(true)
            if ((bool)cForm.ShowDialog())
            {
                SelectedConnection = cForm.ConnectionDetail;
                //HadCreatedNewConnection = false;

                ValidateButtonClick(null, null);
            }
            else
            {
                CancelButtonClick(null, null);
            }
        }

        private void LvConnectionsMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ValidateButtonClick(null, null);
        }

        private void ValidateButtonClick(object sender, RoutedEventArgs e)
        {
            if (lvConnections.SelectedItems.Count > 0)
            {
                SelectedConnection = ((ConnectionItem)lvConnections.SelectedItems[0]).ConnectionInfo;
                // MessageBox.Show(SelectedConnection.ConnectionName);
                DialogResult = true;
                Close();
            }
        }

        private void CancelButtonClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        #endregion Handlers

       
    }
}
