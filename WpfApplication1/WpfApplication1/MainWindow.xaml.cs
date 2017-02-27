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

namespace WpfGenerateExcel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string ExcelFileName { get; set; }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog()
            {
                FileName = "AnyExcelFile",
                DefaultExt = ".xlsx",
                Filter = "Excel Document (.xlsx)|*.xlsx"
            };

            // Show save file dialog box
            var result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                ExcelFileName = dlg.FileName;

                var generateExcel = new GeneratedClass();
                generateExcel.CreatePackage(ExcelFileName);

                StatusBar.Text = $"Success!  Generated: {ExcelFileName}";

            }
        }
    }
}
