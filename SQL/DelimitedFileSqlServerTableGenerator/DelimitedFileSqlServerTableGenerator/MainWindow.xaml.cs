using DelimitedFileSqlServerTableGenerator.ViewModels;
using System.Windows;

namespace DelimitedFileSqlServerTableGenerator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        internal MainWindowViewModel ViewModel { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            ViewModel = new MainWindowViewModel();
            this.DataContext = ViewModel;
        }

        private void OpenFile(object sender, RoutedEventArgs e)
        {
            ViewModel.SelectFile();
        }

        private void ParseFile(object sender, RoutedEventArgs e)
        {
            ViewModel.ParseFile();
        }

        private void RefreshSql(object sender, System.EventArgs e)
        {
            ViewModel.RefreshSqlServerCreateStatement();
        }

        private void RefreshSql(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            ViewModel.RefreshSqlServerCreateStatement();
        }

        private void CopyInserts(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(ViewModel.SqlServerInsertStatement);
        }

        private void CopyCreateTable(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(ViewModel.SqlServerCreateStatement);
        }
    }
}
