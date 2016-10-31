using System.Windows.Data;
using System.Windows.Media;
using System;
using DevExpress.Xpf.Grid;
using System.Windows;

namespace QuickTimeEntry
{
    public class GroupColumnSummaryForegroundConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            string fieldName = (string)values[0];

            if (fieldName == "gesamtdauer" && values[4] != DependencyProperty.UnsetValue && values[4] != null)
            {
                if ((double)values[4] < 20) return new SolidColorBrush(Colors.Red);
                else if ((double)values[4] < 30) return new SolidColorBrush(Colors.Yellow);
                else if ((double)values[4] >= 30) return new SolidColorBrush(Colors.Green);
            }

            return new SolidColorBrush(Colors.Black);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
