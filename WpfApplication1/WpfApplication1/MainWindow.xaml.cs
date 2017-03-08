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
                generateExcel.CreatePackage(ExcelFileName, GeneratedExcelData());

                StatusBar.Text = $"Success!  Generated: {ExcelFileName}";

            }
        }

        private List<ExcelData> GeneratedExcelData()
        {
            return new[]
            {
                new ExcelData() {Barcode="Barcode",Class="Class",Client="Client",Client_Barcode="C_Barcode",Content_Owner="CO",IsClientVisible="T",Language="Lang",MaterialId="MID",Material_Type="MT",Standard="North",Title_Name="TITLE" },
                new ExcelData() {Barcode="Barcode1",Class="Class1",Client="Client1",Client_Barcode="C_Barcode1",Content_Owner="CO1",IsClientVisible="T1",Language="Lang1",MaterialId="MID1",Material_Type="MT1",Standard="North1",Title_Name="TITLE1" },
            }.ToList();
        }
    }
}
