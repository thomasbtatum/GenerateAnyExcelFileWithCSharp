using System.Windows;

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
                FileName = "RightToLeftExcelFile",
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
