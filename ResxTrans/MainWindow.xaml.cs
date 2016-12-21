using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using MessageBox = System.Windows.MessageBox;

namespace ResxTrans
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public string ResxFileName { get; set; }
        public string XlsFileName { get; set; }
        public MainWindow()
        {
            InitializeComponent();
        }


        private void OpenResxFile_OnClick(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".resx";
            dlg.Filter = "Resource Files (*.resx)|*.resx";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string fileName = dlg.FileName;
                FileNameBox.Text = fileName;
                ResxFileName = fileName;
            }
        }

        private void ExportToXls_OnClick(object sender, RoutedEventArgs e)
        {
            string fileName = ResxFileName;

            // Create OpenFileDialog 
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();

            // Set filter for file extension and default file extension 
            FileInfo fi = new FileInfo(fileName);
            dlg.FileName = fi.Name.Split('.')[0];
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();
            string xlsFileName;

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                xlsFileName = dlg.FileName;
            }
            else
            {
                return;
            }

            ResxXlsHelper.ExportToXls(fileName, xlsFileName);

            MessageBox.Show(this, "Export to xls has finished", "Export", MessageBoxButton.OK);

        }

        private void OpenXlsFile_OnClick(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx；.xls";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Files (*.xls)|*.xls";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string fileName = dlg.FileName;
                XlsFileNameBox.Text = fileName;
                XlsFileName = fileName;
            }
        }

        private void ExportToResx_OnClick(object sender, RoutedEventArgs e)
        {
            string fileName = XlsFileName;

            // Create OpenFileDialog 
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();

            // Set filter for file extension and default file extension 
            FileInfo fi = new FileInfo(fileName);
            dlg.FileName = fi.Name.Split('.')[0];
            dlg.DefaultExt = ".resx";
            dlg.Filter = "Resource Files (*.resx)|*.resx";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();
            string resxFileName;

            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                resxFileName = dlg.FileName;
            }
            else
            {
                return;
            }

            ResxXlsHelper.ExportToResx(fileName, resxFileName);

            MessageBox.Show(this, "Export to resx has finished", "Export", MessageBoxButton.OK);
        }
    }


}
