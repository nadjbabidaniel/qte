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
using System.Resources;
using System.Reflection;

namespace QuickTimeEntry
{
    /// <summary>
    /// Interaktionslogik für SelectFavorite.xaml
    /// </summary>
    public partial class SelectFavorite : Window
    {
        DataTable TimeEntryCache;

        DataTable Favorites;

        public String Selected_ID = "";
        public String Selected_EntityType = "";

        private String UserName;
        private static ResourceManager rm = new ResourceManager("QuickTimeEntry.SelectFavorite", Assembly.GetExecutingAssembly());


        private void FillDataGrid()
        {
            Favorites = new DataTable("favorites");
            Favorites.Columns.Add("gctype", typeof(String));
            Favorites.Columns.Add("gcdescription", typeof(String));
            Favorites.Columns.Add("gcmscrmid", typeof(String));

            DataRow[] _result = TimeEntryCache.Select("[type]='FAVORITE_" + UserName + "'");
            foreach (DataRow _dr in _result)
            {
                DataRow _nr = Favorites.NewRow();

                _nr["gctype"] = _dr["subtype"];
                _nr["gcdescription"] = _dr["description"];
                _nr["gcmscrmid"] = _dr["mscrmid"];

                Favorites.Rows.Add(_nr);
            }

            FavGrid.DataSource = Favorites;
            FavGrid.IsFilterEnabled = true;
            FavGrid.RefreshData();

            FavGrid.GroupBy(FavGrid.Columns["gctype"], DevExpress.Data.ColumnSortOrder.Ascending);
            FavGrid.ExpandAllGroups();

            if (Favorites.Rows.Count == 0)
            {
                FAVSearchInfo.Content = rm.GetString("NoEntries");
                btnFAVDelete.IsEnabled = false;
                btnFAVAccept.IsEnabled = false;
            }
            else
            {
                if (Favorites.Rows.Count > 1)
                    FAVSearchInfo.Content = Favorites.Rows.Count.ToString() + rm.GetString("EntriesList"); 
                else FAVSearchInfo.Content = rm.GetString("AnEntry");
                btnFAVDelete.IsEnabled = true;
                btnFAVAccept.IsEnabled = true;
            }
        }

        public SelectFavorite(DataTable _TimeEntryCache, String _User)
        {
            InitializeComponent();
            
            TimeEntryCache = _TimeEntryCache;

            this.Title = rm.GetString("SelectionFavorites");

            String _tmp_User = _User;
            _tmp_User = _tmp_User.Replace("\\", "_");
            _tmp_User = _tmp_User.Replace(".", "_");
            UserName = _tmp_User;

            FillDataGrid();

            Selected_ID = "";
        }

        private void btnFAVAccept_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                String _desc = FavGrid.GetFocusedRowCellValue("gcdescription").ToString().Trim();
                String _entitytype = FavGrid.GetFocusedRowCellValue("gctype").ToString().Trim();
                if (_desc != "" && _entitytype != "")
                {
                    Selected_EntityType = _entitytype;
                    Selected_ID = _desc;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("btnFAVAccept_Click: " + ex.Message, rm.GetString("ErrorNotification"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void btnFAVClose_Click(object sender, RoutedEventArgs e)
        {
            // Close
            Selected_ID = "";
            Selected_EntityType = "";
            this.Close();
        }

        private void FavGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            // Klick
            try
            {
                String _desc = FavGrid.GetFocusedRowCellValue("gcdescription").ToString().Trim();
                String _entitytype = FavGrid.GetFocusedRowCellValue("gctype").ToString().Trim();
                if (_desc != "" && _entitytype != "")
                {
                    Selected_EntityType = _entitytype;
                    Selected_ID = _desc;
                    this.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("FavGrid_MouseDoubleClick: " + ex.Message, rm.GetString("ErrorNotification"), MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private DataRow GetCacheItemByMSCRMID(String itemTYPE, String itemGUID)
        {
            DataRow _rV = null;

            DataRow[] _result = TimeEntryCache.Select("[type]='" + itemTYPE + "' AND [mscrmid]='" + itemGUID + "'");
            if (_result.Length == 1)
            {
                _rV = _result[0];
            }

            return _rV;
        }

        private void btnFAVDelete_Click(object sender, RoutedEventArgs e)
        {
            // Aktuellen Eintrag entfernen
            if (FavGrid.GetFocusedRowCellValue("gcmscrmid") != null)
            {
                String _mscrmid = FavGrid.GetFocusedRowCellValue("gcmscrmid").ToString().Trim();
                String _entitytype = FavGrid.GetFocusedRowCellValue("gctype").ToString().Trim();
                DataRow _dr = null;

                if ((_dr = GetCacheItemByMSCRMID("FAVORITE_" + UserName, _mscrmid)) != null)
                {
                    TimeEntryCache.Rows.Remove(_dr);
                    TimeEntryCache.AcceptChanges();

                    FillDataGrid();                    
                }
            }
        }
    }
}
