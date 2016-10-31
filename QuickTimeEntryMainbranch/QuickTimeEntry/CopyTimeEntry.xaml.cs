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
using System.Globalization;

namespace QuickTimeEntry
{
    /// <summary>
    /// Interaktionslogik für CopyTimeEntry.xaml
    /// </summary>
    public partial class CopyTimeEntry : Window
    {
        public Boolean DoCopyTimeEntry = false;

        public DateTime Datum;
        public String Dauer;
        public String Zusatzzeit;

        public CopyTimeEntry(DateTime _Datum, String _Dauer, String _Zusatzzeit)
        {
            InitializeComponent();

            deDatum.MaskType = DevExpress.Xpf.Editors.MaskType.DateTime;
            deDatum.MaskCulture = CultureInfo.CurrentCulture;
            deDatum.MaskUseAsDisplayFormat = true;
            deDatum.Mask = "dd.MM.yyyy";

            //var tempDN = _Datum.Day.ToString("00") + "." + _Datum.Month.ToString("00") + "." + _Datum.Year.ToString() + " 00:00:00";
            deDatum.EditValue = Convert.ToDateTime(_Datum.ToShortDateString(), CultureInfo.CurrentCulture.DateTimeFormat);

            teDauer.Text = _Dauer;
            teZusatz.Text = _Zusatzzeit;
        }

        private void btnCopyTimeEntry_Click(object sender, RoutedEventArgs e)
        {
            // OK
            DoCopyTimeEntry = true;

            DateTime _dauer = Convert.ToDateTime(teDauer.EditValue);
            DateTime _zusatz = Convert.ToDateTime(teZusatz.EditValue);
            DateTime _datum = deDatum.DateTime;

            Datum = Convert.ToDateTime(_datum.ToShortDateString(), CultureInfo.CurrentCulture.DateTimeFormat);
            Dauer = _dauer.Hour.ToString("00")+":"+_dauer.Minute.ToString("00");
            Zusatzzeit = _zusatz.Hour.ToString("00") + ":" + _zusatz.Minute.ToString("00");

            this.Close();
        }

        private void btnAbort_Click(object sender, RoutedEventArgs e)
        {
            // Abort
            DoCopyTimeEntry = false;
            this.Close();
        }
    }
}
