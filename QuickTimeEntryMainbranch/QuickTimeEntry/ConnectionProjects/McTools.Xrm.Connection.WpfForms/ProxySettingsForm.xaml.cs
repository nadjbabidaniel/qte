using System.Reflection;
using System.Resources;
using System.Windows;
using System.Windows.Controls;

namespace McTools.Xrm.Connection.WpfForms
{
    /// <summary>
    /// Logique d'interaction pour ProxySettingsForm.xaml
    /// </summary>
    public partial class ProxySettingsForm
    {
        public ProxySettingsForm()
        {
            InitializeComponent();
        }

        #region Properties

        public string ProxyAddress { get; set; }
        public string ProxyPort { get; set; }
        public string UserName { get; set; }
        public string Password { get; set; }
        public bool UseCustomProxy { get; set; }

        private static ResourceManager rm = new ResourceManager("McTools.Xrm.Connection.WpfForms.ProxySettingsForm", Assembly.GetExecutingAssembly());

        #endregion Properties

        private void WindowLoaded(object sender, RoutedEventArgs e)
        {
            chkUseCustomProxy.IsChecked = UseCustomProxy;
            txtProxyAddress.Text = ProxyAddress ?? string.Empty;
            txtProxyPort.Text = ProxyPort ?? string.Empty;
            txtLogin.Text = UserName ?? string.Empty;
            txtPassword.Password = Password ?? string.Empty;
        }

        private void ChkUseCustomProxyChecked(object sender, RoutedEventArgs e)
        {
            txtProxyAddress.IsEnabled = true;
            txtProxyPort.IsEnabled = true;
            txtLogin.IsEnabled = true;
            txtPassword.IsEnabled = true;
        }

        private void ChkUseCustomProxyUnChecked(object sender, RoutedEventArgs e)
        {
            txtProxyAddress.IsEnabled = false;
            txtProxyPort.IsEnabled = false;
            txtLogin.IsEnabled = false;
            txtPassword.IsEnabled = false;
        }

        private void BtnValidateClick(object sender, RoutedEventArgs e)
        {
            ProxyAddress = txtProxyAddress.Text;
            ProxyPort = txtProxyPort.Text;
            UserName = txtLogin.Text;
            Password = txtPassword.Password;
            UseCustomProxy = chkUseCustomProxy.IsChecked != null && chkUseCustomProxy.IsChecked.Value;

            DialogResult = true;
            Close();
        }

        private void BtnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void TxtProxyPortTextChanged(object sender, TextChangedEventArgs e)
        {
            int port;

            if (!int.TryParse(txtProxyPort.Text, out port))
            {
                MessageBox.Show(this, rm.GetString("NumericCharacters"), rm.GetString("Warning"), MessageBoxButton.OK, MessageBoxImage.Warning);
                txtProxyPort.Text = txtProxyPort.Text.Substring(0, txtProxyPort.Text.Length - 1);
                txtProxyPort.Select(txtProxyPort.Text.Length, 0);
            }
        }
    }
}
