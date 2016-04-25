using RateRouteParser;
using System;
using System.IO;
using System.Windows;
using System.Windows.Forms;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace RRparser
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        public MainWindow()
        {
            InitializeComponent();
        }



        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                InitialDirectory = "c:\\",
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == true)
                FileNameTextBox.Text = openFileDialog.FileName;

        }

        private void btnDirectory_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            dialog.ShowDialog();
            DirectoryName.Text = dialog.SelectedPath;
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {

     
            var inputPath = FileNameTextBox.Text;
            var outputPath = DirectoryName.Text;
            try
            {
                if (RadioRates.IsChecked == true)
                {
                    var exportRates = new UpsRates(inputPath);
                    exportRates.ExportToCsv(outputPath);

                }
                if (RadioRoutes.IsChecked == true)
                {
                    if (string.IsNullOrWhiteSpace(TextFromZip.Text))
                    {
                        var exportZones = new UpsZones(inputPath);
                        exportZones.ExportToCsv(outputPath);
                    }
                    else
                    {
                        var exportZones = new UpsZones(inputPath, TextFromZip.Text);
                        exportZones.ExportToCsv(outputPath);
                    }
                }
                ShowSuccess("Data sucessfully exported");

            }
          
            catch (Exception exception)
            {
                ShowError(exception.Message);
            }
         

        }

        private void RadioRates_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                TextFromZip.Visibility = Visibility.Hidden;
                LabelZip.Visibility = Visibility.Hidden;
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private void RadioRoutes_Checked(object sender, RoutedEventArgs e)
        {
            try
            {
                TextFromZip.Visibility = Visibility.Visible;
                LabelZip.Visibility = Visibility.Visible;
            }
            catch (Exception)
            {
                // ignored
            }
        }

        public static void ShowError(string error)
        {
            System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(new Action(() =>
            {
                System.Windows.Forms.MessageBox.Show(error, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }));
        }


        public static void ShowSuccess(string success)
        {
            System.Windows.Threading.Dispatcher.CurrentDispatcher.BeginInvoke(new Action(() =>
            {
                System.Windows.Forms.MessageBox.Show(success, "Success", MessageBoxButtons.OK, MessageBoxIcon.Question);
            }));
        }



    }



}
