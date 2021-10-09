using System.Windows;

namespace Wpf_Excel_DataGrid
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            ExcelDados excelDados = new ExcelDados();
            this.grdExcel.DataContext = excelDados.DadosExcel;
        }
    }
}
