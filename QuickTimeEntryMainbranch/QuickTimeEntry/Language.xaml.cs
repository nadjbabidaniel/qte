using Infralution.Localization.Wpf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace QuickTimeEntry
{
    /// <summary>
    /// Interaction logic for Languages.xaml
    /// </summary>
    public partial class Language : Window
    {
        public Language()
        {
            InitializeComponent();
            string lang = CultureManager.UICulture.TwoLetterISOLanguageName.ToLower();
            if (lang == "de")
            {
                radioButton1.IsChecked = true;
                radioButton2.IsChecked = false;                
            }
            if (lang == "en")
            {
                radioButton1.IsChecked = false;
                radioButton2.IsChecked = true;
            }

            Ok_button.IsDefault = true;
        }

        private void Ok_button_Click(object sender, RoutedEventArgs e)
        {
            if (radioButton1.IsChecked == true)
            {
                CultureManager.UICulture = new CultureInfo("de");               
            }
            if (radioButton2.IsChecked == true)
            {
                CultureManager.UICulture = new CultureInfo("en");              
            }

            DialogResult = true;

            Close();
        }

        private void Cancel_button_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
