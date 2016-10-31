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
using System.Reflection;
using System.Resources;

namespace QuickTimeEntry
{
    /// <summary>
    /// Interaktionslogik für EditTimeSpanItems.xaml
    /// </summary>
    public partial class EditTimeSpanItems : Window
    {
        public String _measurements;
        public Boolean _AcceptChanges = false;

        private DataTable _dt = new DataTable("time_measurements");

        private static ResourceManager rm = new ResourceManager("QuickTimeEntry.EditTimeSpanItems", Assembly.GetExecutingAssembly());

        private void FillDataGrid()
        {
            GridTimeMeasurements.ItemsSource = _dt;
            GridTimeMeasurements.IsFilterEnabled = true;
            GridTimeMeasurements.RefreshData();

            if (_dt.Rows.Count == 0)
            {
                GridTimeMeasurementsInfo.Content = rm.GetString("NoEntries");
                btnTimeMeasurementDelete.IsEnabled = false;
                btnTimeMeasurementAccept.IsEnabled = false;
            }
            else
            {
                if (_dt.Rows.Count > 1)
                    GridTimeMeasurementsInfo.Content = _dt.Rows.Count.ToString() + " "+ rm.GetString("Entry"); 
                else GridTimeMeasurementsInfo.Content = rm.GetString("AnEntry");
                btnTimeMeasurementDelete.IsEnabled = true;
                btnTimeMeasurementAccept.IsEnabled = true;
            }
        }

        public EditTimeSpanItems(String time_measurements)
        {
            InitializeComponent();

            this.Title = rm.GetString("TimeMeasurements");

            _measurements = time_measurements;

            GridTimeMeasurementsInfo.Content = "";
            
            _dt.Columns.Add("marked", Type.GetType("System.Boolean"));
            _dt.Columns.Add("TimeFrom", Type.GetType("System.DateTime"));
            _dt.Columns.Add("TimeTo", Type.GetType("System.DateTime"));
            _dt.Columns.Add("TimeDelay", Type.GetType("System.String"));

            if (_measurements != "")
            {
                String[] _te_measurements = _measurements.Split(';');

                foreach (String _item in _te_measurements)
                {
                    String[] _te_parts = _item.Split('|');

                    DataRow _r = _dt.NewRow();

                    _r["marked"] = false;
                    try
                    {
                        _r["TimeFrom"] = Convert.ToDateTime(_te_parts[0]);
                        
                    }
                    catch {
                        _r["TimeFrom"] = DBNull.Value;
                    }

                    try
                    {
                        _r["TimeTo"] = Convert.ToDateTime(_te_parts[1]);
                    }
                    catch {
                        _r["TimeTo"] = DBNull.Value;
                    }

                    String _time_info = "";

                    _time_info += TimeSpan.FromMilliseconds(Convert.ToDouble(_te_parts[2].ToString())).Hours.ToString("00");
                    _time_info += ":";
                    _time_info += TimeSpan.FromMilliseconds(Convert.ToDouble(_te_parts[2].ToString())).Minutes.ToString("00");
                    _time_info += ":";
                    _time_info += TimeSpan.FromMilliseconds(Convert.ToDouble(_te_parts[2].ToString())).Seconds.ToString("00");

                    _r["TimeDelay"] = _time_info;

                    _dt.Rows.Add(_r);
                }

                FillDataGrid();
            }
        }

        private void btnTimeMeasurementDelete_Click(object sender, RoutedEventArgs e)
        {
            // Delete
            Int32 _Count_Marked = 0;

            DataTable _dt = (DataTable)GridTimeMeasurements.ItemsSource;
            foreach (DataRow _dr in _dt.Rows)
            {
                if (_dr["marked"] != DBNull.Value)
                {
                    if (Convert.ToBoolean(_dr["marked"].ToString()))
                    {
                        _Count_Marked++;
                    }
                }
            }

            if (_Count_Marked > 0)
            {
                try
                {
                    _dt = (DataTable)GridTimeMeasurements.ItemsSource;
                    while (_Count_Marked > 0)
                    {
                        foreach (DataRow _dr in _dt.Rows)
                        {
                            if (_dr["marked"] != DBNull.Value)
                            {
                                if (Convert.ToBoolean(_dr["marked"].ToString()))
                                {
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

                GridTimeMeasurements.RefreshData();
            }
        }

        private void btnTimeMeasurementAccept_Click(object sender, RoutedEventArgs e)
        {
            // Accept Changes
            try
            {
                _AcceptChanges = true;
                _measurements = "";

                foreach (DataRow _r in _dt.Rows)
                {
                    if (_measurements != "") _measurements += ";";

                    _measurements += _r["TimeFrom"].ToString();
                    _measurements += "|";
                    _measurements += _r["TimeTo"].ToString();
                    _measurements += "|";

                    long startTick = Convert.ToDateTime(_r["TimeFrom"].ToString()).Ticks;
                    long endTick = Convert.ToDateTime(_r["TimeTo"].ToString()).Ticks;
                    long tick = endTick - startTick;
                    long seconds = tick / TimeSpan.TicksPerSecond;
                    long milliseconds = tick / TimeSpan.TicksPerMillisecond;

                    _measurements += milliseconds.ToString();
                }

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(rm.GetString("ErrorConversion"));
            }
        }

        private void btnTimeMeasurementClose_Click(object sender, RoutedEventArgs e)
        {
            // Close Window

            _AcceptChanges = false;

            this.Close();
        }

        private void GridTimeMeasurements_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Mouse Double Klick
        }

        private void GridTimeMeasurements_CheckEdit_Checked(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;

            // Clear marked flag
            DataTable _dt = (DataTable)GridTimeMeasurements.DataSource;
            foreach (DataRow _dr in _dt.Rows)
            {
                _dr["marked"] = false;
            }

            for (int i = 0; i < GridTimeMeasurements.VisibleRowCount; i++)
            {
                int rowHandle = GridTimeMeasurements.GetRowHandleByVisibleIndex(i);

                GridTimeMeasurements.SetCellValue(rowHandle, GridTimeMeasurements.Columns["marked"], true);
            }

            GridTimeMeasurements.RefreshData();

            this.Cursor = Cursors.Arrow;

        }

        private void GridTimeMeasurements_CheckEdit_Unchecked(object sender, RoutedEventArgs e)
        {
            this.Cursor = Cursors.Wait;

            // Unmark
            for (int i = 0; i < GridTimeMeasurements.VisibleRowCount; i++)
            {
                int rowHandle = GridTimeMeasurements.GetRowHandleByVisibleIndex(i);

                GridTimeMeasurements.SetCellValue(rowHandle, GridTimeMeasurements.Columns["marked"], false);
            }

            GridTimeMeasurements.RefreshData();

            this.Cursor = Cursors.Arrow;

        }
    }
}
