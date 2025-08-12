// ==================== Using Statements ====================
using System.Windows;
using System.Globalization;
using System.Windows.Media.Imaging;
using System;
using System.Threading.Tasks;
using System.Windows.Controls;

// ==================== Namespace Declaration ====================
namespace ITSupportToolGUI
{
    // ==================== MainWindow Class Definition ====================
    public partial class MainWindow : Window
    {
        // ==================== Member Variables ====================
        private readonly MainViewModel _viewModel;
        private int secretButtonClickCount = 0;

        // ==================== Constructor and Initialization ====================
        public MainWindow()
        {
            InitializeComponent();
            _viewModel = new MainViewModel();
            this.DataContext = _viewModel;
            Loaded += async (s, e) => await _viewModel.InitializeAsync();
        }

        // ==================== Button Click Handlers ====================
        private void btnOpenTicketSystem_Click(object sender, RoutedEventArgs e) => _viewModel.OpenTicketSystem();
        private void btnOpenITMessages_Click(object sender, RoutedEventArgs e) => _viewModel.OpenITMessages();
        private void btnOpenPasswordReset_Click(object sender, RoutedEventArgs e) => _viewModel.OpenPasswordReset();

        private void btnRunScript_Click(object sender, RoutedEventArgs e)
        {
            _viewModel.RunPowerShellScriptCommand.Execute(null);
        }

        private async void btnSwitchToEnglish_Click(object sender, RoutedEventArgs e)
        {
            await _viewModel.SetLanguage("en-US");
        }

        private async void btnSwitchToNative_Click(object sender, RoutedEventArgs e)
        {
            string nativeCulture = _viewModel.GetCountry() switch
            {
                "SE" => "sv-SE",
                "NO" => "nb-NO",
                _ => "da-DK",
            };
            await _viewModel.SetLanguage(nativeCulture);
        }

        private void SecretButton_Click(object sender, RoutedEventArgs e)
        {
            secretButtonClickCount++;
            if (secretButtonClickCount == 3)
            {
                MessageBox.Show("This application was made by Mathias West Robinson");
                secretButtonClickCount = 0;
            }
        }
    }
}
