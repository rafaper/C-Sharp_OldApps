using System;
using System.Collections.Generic;
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
using System.Windows.Navigation;
using System.Windows.Shapes;


namespace XL_Auto_UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        XL_Control _xlControl;
        UST_WeeklyClosed _ustWeeklyClosed;

        public MainWindow()
        { 

        }

        private void btnGenerate_Click(object sender, RoutedEventArgs e)
        {
            GenerateUSTWeekly();
        }

        private void GenerateUSTWeekly()
        {
            _xlControl = new XL_Control();

            _xlControl.Original_File = tbFilePath_USTClosedWkly.Text.Replace("\"", "");

            _xlControl.Open(_xlControl.Original_File);
            _ustWeeklyClosed.Delete_Irrelevant_Columns();
            _ustWeeklyClosed.Delete_Irrelevant_Rows();
            CheckIfDatesAreSelected();
            _xlControl.Show_App();
        }

        private void CheckIfDatesAreSelected()
        {
            var start = dpStart_USTClosedWkly.Text;
            var end = dpEnd_USTClosedWkly.Text;

            if (string.IsNullOrWhiteSpace(start) || string.IsNullOrWhiteSpace(end))
            {
                MessageBox.Show("Please select a start and end date", "Date Selection", MessageBoxButton.OK);
            }
            else
            {
                _ustWeeklyClosed.DeleteIfOutsideDateRange(dpStart_USTClosedWkly.SelectedDate.Value, dpEnd_USTClosedWkly.SelectedDate.Value);
            }
        }

        //Clean Up Later
        private void btnBrowseFile_Click(object sender, RoutedEventArgs e)
        {
            var ofd = new Microsoft.Win32.OpenFileDialog()
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            };

            ofd.ShowDialog();

            tbFilePath_USTClosedWkly.Text = ofd.FileName;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            DisplayStackPanel(sp_MainMenu.Name);
        }

        private void btn_GoTo_USTWklyClosed_Click(object sender, RoutedEventArgs e)
        {
            DisplayStackPanel(sp_USTClosedWkly.Name);

        }

        private void btn_GoTo_USTWklyActive_Click(object sender, RoutedEventArgs e)
        {
            DisplayStackPanel(sp_USTActiveWkly.Name);
        }

        private void btn_BackMaster_Click(object sender, RoutedEventArgs e)
        {
            DisplayStackPanel(sp_MainMenu.Name);
        }

        private void btnGenerate_USTActiveWkly_Click(object sender, RoutedEventArgs e)
        {
            UST_WeeklyActive _ustWklyActive = new UST_WeeklyActive();
            _ustWklyActive.FileName = tbFilePath_USTActiveWkly.Text.Replace("\"", "");

            _xlControl.Open(_ustWklyActive.FileName);
            _ustWklyActive.GrabData();
            _ustWklyActive.WriteDataToTxt();
        }

        private void DisplayStackPanel(string sp)
        {
            sp_MainMenu.Visibility = Visibility.Collapsed;
            sp_USTClosedWkly.Visibility = Visibility.Collapsed;
            sp_USTActiveWkly.Visibility = Visibility.Collapsed;

            switch (sp)
            {
                case "sp_MainMenu":
                    sp_MainMenu.Visibility = Visibility.Visible;
                    btn_BackMaster.Visibility = Visibility.Hidden;
                    break;
                case "sp_USTClosedWkly":
                    sp_USTClosedWkly.Visibility = Visibility.Visible;
                    btn_BackMaster.Visibility = Visibility.Visible;
                    break;
                case "sp_USTActiveWkly":
                    sp_USTActiveWkly.Visibility = Visibility.Visible;
                    btn_BackMaster.Visibility = Visibility.Visible;
                    break;
            }
        }
    }
}
