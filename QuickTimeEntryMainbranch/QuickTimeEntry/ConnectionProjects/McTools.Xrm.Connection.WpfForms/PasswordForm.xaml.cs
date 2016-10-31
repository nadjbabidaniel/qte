using System.Reflection;
using System.Resources;
using System.Windows;
using System.Windows.Input;

namespace McTools.Xrm.Connection.WpfForms
{
    /// <summary>
    /// Logique d'interaction pour PasswordForm.xaml
    /// </summary>
    public partial class PasswordForm
    {
        #region Variables

        /// <summary>
        /// Mot de passe de l'utilisateur
        /// </summary>
        string _userPassword;

        #endregion

        #region Constructeur

        /// <summary>
        /// Créé une nouvelle instance de la classe PasswordForm
        /// </summary>
        public PasswordForm(string userPassword)
        {
            _userPassword = userPassword;
            InitializeComponent();
        }

        /// <summary>
        /// Créé une nouvelle instance de la classe PasswordForm
        /// </summary>
        public PasswordForm()
        {
            InitializeComponent();
        }

        #endregion

        #region Propriétés

        /// <summary>
        /// Obtient ou définit le login de l'utilisateur
        /// </summary>
        public string UserLogin { get; set; }

        /// <summary>
        /// Obtient ou définit le nom de domaine pour l'utilisateur
        /// </summary>
        public string UserDomain { get; set; }

        /// <summary>
        /// Obtient le mot de passe de l'utilisateur
        /// </summary>
        public string UserPassword
        {
            get { return _userPassword; }
        }

        public bool SavePassword { get; set; }


        private static ResourceManager rm = new ResourceManager("McTools.Xrm.Connection.WpfForms.PasswordForm", Assembly.GetExecutingAssembly());


        #endregion

        private void ShowPasswordChecked(object sender, RoutedEventArgs e)
        {
            txtPassword.PasswordChar = (char)0;
        }

        private void ShowPasswordUnchecked(object sender, RoutedEventArgs e)
        {
            txtPassword.PasswordChar = '•';
        }

        private void BtnCancelClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void BtnValidateClick(object sender, RoutedEventArgs e)
        {
            bool go = true;

            if (txtPassword.Password.Length == 0)
            {
                if (MessageBox.Show(this, rm.GetString("PasswordEmpty"), rm.GetString("Question"), MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.No)
                {
                    go = false;
                }
            }

            if (go)
            {
                _userPassword = txtPassword.Password;
                SavePassword = chkSavePassword.IsChecked != null && chkSavePassword.IsChecked.Value;
                DialogResult = true;
                Close();
            }
        }

        private void TxtPasswordKeyUp(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                BtnValidateClick(null, null);
            }
        }
    }
}
