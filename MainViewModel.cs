// ==================== Using Statements ====================
using ITSupportToolGUI.Models;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Management;
using System.Net.NetworkInformation;
using System.Resources;
using System.Runtime.CompilerServices;
using System.Runtime.Versioning;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using Windows.Devices.WiFi;

// ==================== Namespace Declaration ====================
namespace ITSupportToolGUI
{
    // ==================== MainViewModel Class Definition ====================
    public class MainViewModel : INotifyPropertyChanged
    {
        // ==================== Member Variables ====================
        private readonly ResourceManager _resourceManager = new ResourceManager("ITSupportToolGUI.Resources.Strings", typeof(MainViewModel).Assembly);
        private bool _isServiceDeskViewVisible = true;
        private bool _isComputerInfoViewVisible = false;
        private bool _isInstructionsViewVisible = false;
        private bool _isPrinterViewVisible = false;
        private int _windowWidth = 520; // Initial width
        private string _siteId = string.Empty;
        private string _printerStatusMessage = string.Empty;
        private string? _nativeFlagIconPath;
        private bool _isSearchingPrinters;


        // ==================== INotifyPropertyChanged Implementation ====================
        public event PropertyChangedEventHandler? PropertyChanged;

        protected void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        // ==================== Commands ====================
        public ICommand CopyToClipboardCommand { get; }
        public ICommand RunPowerShellScriptCommand { get; }
        public ICommand ExecuteWifiFixScriptCommand { get; }
        public ICommand BackToInfoViewCommand { get; }
        public ICommand CopyAllInfoCommand { get; }
        public ICommand ShowPrinterViewCommand { get; }
        public ICommand ShowComputerInfoViewCommand { get; }
        public ICommand SearchPrintersCommand { get; }
        public ICommand AddSelectedPrintersCommand { get; }
        public ICommand OpenRemoteHelpCommand { get; }


        // ==================== Visibility & Layout Properties ====================
        public bool IsServiceDeskViewVisible { get => _isServiceDeskViewVisible; set { _isServiceDeskViewVisible = value; OnPropertyChanged(); } }
        public bool IsComputerInfoViewVisible { get => _isComputerInfoViewVisible; set { _isComputerInfoViewVisible = value; OnPropertyChanged(); } }
        public bool IsInstructionsViewVisible { get => _isInstructionsViewVisible; set { _isInstructionsViewVisible = value; OnPropertyChanged(); } }
        public bool IsPrinterViewVisible { get => _isPrinterViewVisible; set { _isPrinterViewVisible = value; OnPropertyChanged(); } }
        public int WindowWidth { get => _windowWidth; set { _windowWidth = value; OnPropertyChanged(); } }


        // ==================== System Information Properties ====================
        public string? ComputerName { get; private set; }
        public string? UserName { get; private set; }
        public string? OsVersion { get; private set; }
        public string? Ssid { get; private set; }
        public string? IpAddress { get; private set; }
        public string? VpnStatus { get; private set; }
        public string? VpnIpAddress { get; private set; }
        public string? LastRestart { get; private set; }
        public bool IsVpnConnected { get; private set; }
        public bool IsPowerShellButtonVisible { get; private set; }
        public bool IsPrinterButtonVisible { get; private set; }
        public bool IsDenmarkServiceDeskVisible { get; private set; }
        public bool IsNorwayServiceDeskVisible { get; private set; }
        public bool IsSwedenServiceDeskVisible { get; private set; }


        // ==================== Printer Properties ====================
        public ObservableCollection<Printer> AvailablePrinters { get; } = new ObservableCollection<Printer>();
        public string SiteId { get => _siteId; set { _siteId = value; OnPropertyChanged(); } }
        public string PrinterStatusMessage { get => _printerStatusMessage; set { _printerStatusMessage = value; OnPropertyChanged(); } }
        public string PrinterNotice => "OPS: Kan kun tilføje Ricoh printere";


        // ==================== UI Properties ====================
        public string? NativeFlagIconPath { get => _nativeFlagIconPath; private set { _nativeFlagIconPath = value; OnPropertyChanged(); } }
        public string OpenTicketSystemTip => _resourceManager.GetString("OpenTicketSystemTip") ?? string.Empty;
        public string OpenITMessagesTip => _resourceManager.GetString("OpenITMessagesTip") ?? string.Empty;
        public string OpenRemoteHelpTip => _resourceManager.GetString("OpenRemoteHelpTip") ?? string.Empty;
        public string OpenPasswordResetTip => _resourceManager.GetString("OpenPasswordResetTip") ?? string.Empty;
        public string SwitchToDanishTip => _resourceManager.GetString("SwitchToDanishTip") ?? string.Empty;
        public string SwitchToEnglishTip => _resourceManager.GetString("SwitchToEnglishTip") ?? string.Empty;
        public string RunScriptTip => _resourceManager.GetString("RunScriptTip") ?? string.Empty;
        public string ComputerNameLabel => _resourceManager.GetString("ComputerNameLabel") ?? string.Empty;
        public string UserNameLabel => _resourceManager.GetString("UsernameLabel") ?? string.Empty;
        public string OsVersionLabel => _resourceManager.GetString("OSVersionLabel") ?? string.Empty;
        public string SsidLabel => _resourceManager.GetString("SSIDLabel") ?? string.Empty;
        public string VpnStatusLabel => _resourceManager.GetString("VPNStatusLabel") ?? string.Empty;
        public string IpAddressLabel => _resourceManager.GetString("IPAddressLabel") ?? string.Empty;
        public string VpnIpAddressLabel => _resourceManager.GetString("VPNIPAddressLabel") ?? string.Empty;
        public string LastRestartLabel => _resourceManager.GetString("LastRestartLabel") ?? string.Empty;
        public string OpenTicketSystemLabel => _resourceManager.GetString("OpenTicketSystemLabel") ?? string.Empty;
        public string OpenITMessagesLabel => _resourceManager.GetString("OpenITMessagesLabel") ?? string.Empty;
        public string OpenRemoteHelpLabel => _resourceManager.GetString("OpenRemoteHelpLabel") ?? string.Empty;
        public string OpenPasswordResetLabel => _resourceManager.GetString("OpenPasswordResetLabel") ?? string.Empty;
        public string RunScriptLabel => _resourceManager.GetString("RunScriptLabel") ?? string.Empty;
        public string CopyAllInfoTip => _resourceManager.GetString("CopyInfoTip") ?? string.Empty;
        public string ComputerInformationTitle => _resourceManager.GetString("ComputerInformationTitle") ?? string.Empty;

        public string InstructionsTitle => _resourceManager.GetString("InstructionsTitle") ?? string.Empty;
        public string InstructionsStep1 => _resourceManager.GetString("InstructionsStep1") ?? string.Empty;
        public string InstructionsStep1Body => _resourceManager.GetString("InstructionsStep1Body") ?? string.Empty;
        public string InstructionsStep2 => _resourceManager.GetString("InstructionsStep2") ?? string.Empty;
        public string InstructionsStep2Body => _resourceManager.GetString("InstructionsStep2Body") ?? string.Empty;
        public string InstructionsBullet2 => _resourceManager.GetString("InstructionsBullet2") ?? string.Empty;
        public string InstructionsBullet3 => _resourceManager.GetString("InstructionsBullet3") ?? string.Empty;
        public string ExecuteScriptButtonText => _resourceManager.GetString("ExecuteScriptButtonText") ?? string.Empty;
        public string BackButtonText => _resourceManager.GetString("BackButtonText") ?? string.Empty;
        public string InstructionsExampleUsername => _resourceManager.GetString("InstructionsExampleUsername") ?? string.Empty;
        public string InstructionsExamplePassword => _resourceManager.GetString("InstructionsExamplePassword") ?? string.Empty;

        public string AddPrinterLabel => _resourceManager.GetString("AddPrinterLabel") ?? string.Empty;
        public string AddPrinterToolTip => _resourceManager.GetString("AddPrinterToolTip") ?? string.Empty;
        public string PrinterViewTitle => _resourceManager.GetString("PrinterViewTitle") ?? string.Empty;
        public string SiteIdLabel => _resourceManager.GetString("SiteIdLabel") ?? string.Empty;
        public string SearchButtonText => _resourceManager.GetString("SearchButtonText") ?? string.Empty;
        public string AddSelectedButtonText => _resourceManager.GetString("AddSelectedButtonText") ?? string.Empty;

        public string ServiceDeskTitle => _resourceManager.GetString("ServiceDeskTitle") ?? string.Empty;
        public string ServiceDeskHours => _resourceManager.GetString("ServiceDeskHours") ?? string.Empty;
        public string ServiceDeskDenmarkPhone => _resourceManager.GetString("ServiceDeskDenmarkPhone") ?? string.Empty;
        public string ServiceDeskDenmarkEmail => _resourceManager.GetString("ServiceDeskDenmarkEmail") ?? string.Empty;
        public string ServiceDeskNorwayPhone => _resourceManager.GetString("ServiceDeskNorwayPhone") ?? string.Empty;
        public string ServiceDeskNorwayEmail => _resourceManager.GetString("ServiceDeskNorwayEmail") ?? string.Empty;
        public string ServiceDeskSwedenPhone => _resourceManager.GetString("ServiceDeskSwedenPhone") ?? string.Empty;
        public string ServiceDeskSwedenEmail => _resourceManager.GetString("ServiceDeskSwedenEmail") ?? string.Empty;
        public string ShowComputerInfoLabel => _resourceManager.GetString("ShowComputerInfoLabel") ?? string.Empty;
        public string ShowComputerInfoToolTip => _resourceManager.GetString("ShowComputerInfoToolTip") ?? string.Empty;
        public string ServiceDeskDenmarkTitle => _resourceManager.GetString("ServiceDeskDenmarkTitle") ?? string.Empty;
        public string ServiceDeskNorwayTitle => _resourceManager.GetString("ServiceDeskNorwayTitle") ?? string.Empty;
        public string ServiceDeskSwedenTitle => _resourceManager.GetString("ServiceDeskSwedenTitle") ?? string.Empty;
        public string ShowServiceDeskLabel => _resourceManager.GetString("ShowServiceDeskLabel") ?? string.Empty;
        public string ShowServiceDeskToolTip => _resourceManager.GetString("ShowServiceDeskToolTip") ?? string.Empty;
        public string ServiceDeskHoursLabel => _resourceManager.GetString("ServiceDeskHoursLabel") ?? string.Empty;
        public string ServiceDeskPhoneLabel => _resourceManager.GetString("ServiceDeskPhoneLabel") ?? string.Empty;
        public string ServiceDeskEmailLabel => _resourceManager.GetString("ServiceDeskEmailLabel") ?? string.Empty;


        // ==================== Initialization ====================
        public MainViewModel()
        {
            CopyToClipboardCommand = new RelayCommand(CopyToClipboard);
            RunPowerShellScriptCommand = new RelayCommand(HandleWifiFixClick);
            ExecuteWifiFixScriptCommand = new RelayCommand(RunWifiFixScript);
            BackToInfoViewCommand = new RelayCommand(ShowInfoView);
            CopyAllInfoCommand = new RelayCommand(_ => CopyInfo());
            ShowPrinterViewCommand = new RelayCommand(ShowPrinterView);
            ShowComputerInfoViewCommand = new RelayCommand(ShowComputerInfoView);
            SearchPrintersCommand = new RelayCommand(async _ => await SearchPrintersAsync(), _ => !_isSearchingPrinters);
            AddSelectedPrintersCommand = new RelayCommand(async _ => await AddSelectedPrintersAsync(), _ => AvailablePrinters.Any(p => p.IsSelected));
            OpenRemoteHelpCommand = new RelayCommand(_ => OpenRemoteHelp());
        }

        public async Task InitializeAsync()
        {
            await FetchSystemInfo();
            await SetLanguage(GetInitialLanguage());
        }

        // ==================== View Switching Logic ====================
        private void HandleWifiFixClick(object? parameter)
        {
            if (IsInstructionsViewVisible)
            {
                RunWifiFixScript(null);
            }
            else
            {
                ShowScriptInstructionsView(null);
            }
        }

        private void ShowScriptInstructionsView(object? parameter)
        {
            WindowWidth = 570; // Expand width
            IsServiceDeskViewVisible = false;
            IsComputerInfoViewVisible = false;
            IsPrinterViewVisible = false;
            IsInstructionsViewVisible = true;
        }
        private void ShowInfoView(object? parameter)
        {
            WindowWidth = 520; // Shrink width back to normal
            IsServiceDeskViewVisible = true;
            IsComputerInfoViewVisible = false;
            IsInstructionsViewVisible = false;
            IsPrinterViewVisible = false;
        }

        private void ShowComputerInfoView(object? parameter)
        {
            WindowWidth = 520; // Shrink width back to normal
            IsServiceDeskViewVisible = false;
            IsInstructionsViewVisible = false;
            IsPrinterViewVisible = false;
            IsComputerInfoViewVisible = true;
        }

        private void ShowPrinterView(object? parameter)
        {
            WindowWidth = 800; // Expand width for printer view
            IsServiceDeskViewVisible = false;
            IsComputerInfoViewVisible = false;
            IsInstructionsViewVisible = false;
            IsPrinterViewVisible = true;
        }

        // ==================== Logic for Button Actions ====================
        private void RunWifiFixScript(object? parameter)
        {
            try
            {
                string scriptCommands = "netsh wlan delete profile name='StarkGroup'; explorer.exe ms-availablenetworks:";
                var processStartInfo = new ProcessStartInfo("powershell.exe", $"-NoProfile -ExecutionPolicy Bypass -Command \"{scriptCommands}\"") { CreateNoWindow = true, UseShellExecute = false, WindowStyle = ProcessWindowStyle.Hidden };
                Process.Start(processStartInfo);
            }
            catch (Exception ex) { MessageBox.Show($"Error running Wi-Fi fix script: {ex.Message}"); }
        }

        private void CopyToClipboard(object? parameter)
        {
            if (parameter is string text && !string.IsNullOrEmpty(text))
            {
                try { Clipboard.SetText(text); }
                catch (Exception ex) { MessageBox.Show($"Could not copy to clipboard: {ex.Message}"); }
            }
        }

        public void OpenTicketSystem() { try { Process.Start(new ProcessStartInfo("https://starkprod.service-now.com/ssp") { UseShellExecute = true }); } catch (Exception ex) { MessageBox.Show($"Error opening ticket system URL: {ex.Message}"); } }
        public void OpenPasswordReset() { try { Process.Start(new ProcessStartInfo("https://mysignins.microsoft.com/security-info/password/change") { UseShellExecute = true }); } catch (Exception ex) { MessageBox.Show($"Error opening password reset URL: {ex.Message}"); } }
        public void OpenRemoteHelp()
        {
            try
            {
                string callingCardPath = FindCallingCard();
                if (!string.IsNullOrEmpty(callingCardPath))
                {
                    Process.Start(callingCardPath);
                }
                else
                {
                    MessageBox.Show("Could not find the 'LogMeIn Rescue Calling Card' application. Please ensure it is installed.", "Application Not Found", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening Stark Remote Help: {ex.Message}");
            }
        }

        private string FindCallingCard()
        {
            string baseDir = @"C:\Program Files (x86)\LogMeIn Rescue Calling Card";
            if (Directory.Exists(baseDir))
            {
                // Search recursively for CallingCard.exe
                var files = Directory.GetFiles(baseDir, "CallingCard.exe", SearchOption.AllDirectories);
                if (files.Length > 0)
                {
                    return files[0]; // Return the first one found
                }
            }
            return string.Empty;
        }
        public void OpenITMessages()
        {
            try
            {
                string url;
                switch (GetCountry())
                {
                    case "SE": url = "https://starkworkspace.sharepoint.com/sites/toolboxbeijer/_layouts/15/news.aspx?title=IT-Info%20%26%20Driftst%C3%B6rningar&newsSource=3&instanceId=8c765d0e-5f7c-4792-b5d3-ab91d5410598&webPartId=8c88f208-6c77-4bdb-86a0-0c47b4316588&serverRelativeUrl=%2Fsites%2Ftoolboxbeijer%2FSitePages%2FHome.aspx&pagesListId=2b8827a3-0d95-495e-8d57-ae0cb9587734"; break;
                    case "NO": url = "https://starkworkspace.sharepoint.com/sites/toolboxneumann/_layouts/15/news.aspx?title=IT%20driftsmeldinger&newsSource=3&instanceId=474050fe-7cf7-4d0b-b4f3-6eafe78aae78&webPartId=8c88f208-6c77-4bdb-86a0-0c47b4316588&serverRelativeUrl=%2Fsites%2Ftoolboxneumann%2FSitePages%2FHome.aspx&pagesListId=847af74d-84e4-4a37-89ce-cbb1a044f0fc"; break;
                    case "GROUP": url = "https://starkworkspace.sharepoint.com/sites/toolboxstarkgroup/_layouts/15/news.aspx?title=Latest%20news&newsSource=3&instanceId=3c3acc5f-432c-4b7b-a6b3-0426a6dad5cf&webPartId=8c88f208-6c77-4bdb-86a0-0c47b4316588&serverRelativeUrl=%2Fsites%2Ftoolboxstarkgroup%2FSitePages%2FHome.aspx&pagesListId=b5016441-c0e2-46b8-8e37-38673915ad42&audienceTargetingEnabled=true"; break;
                    default: url = "https://starkworkspace.sharepoint.com/sites/toolboxstarkdk/_layouts/15/news.aspx?title=IT%20Servicemeddelelser&newsSource=3&instanceId=f0432ad5-c752-4696-bb24-3006c93c8e22&webPartId=8c88f208-6c77-4bdb-86a0-0c47b4316588&serverRelativeUrl=%2Fsites%2Ftoolboxstarkdk%2FSitePages%2FHome.aspx&pagesListId=1485bbde-3d57-4a10-8d68-ff781af0d722"; break;
                }
                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
            }
            catch (Exception ex) { MessageBox.Show($"Error opening IT messages URL: {ex.Message}"); }
        }

        public void CopyInfo()
        {
            try
            {
                var infoText = new StringBuilder();
                infoText.AppendLine($"{ComputerNameLabel} {ComputerName}");
                infoText.AppendLine($"{UserNameLabel} {UserName}");
                infoText.AppendLine($"{OsVersionLabel} {OsVersion}");
                infoText.AppendLine($"{IpAddressLabel} {IpAddress}");
                infoText.AppendLine($"{SsidLabel} {Ssid}");
                infoText.AppendLine($"{LastRestartLabel} {LastRestart}");

                if (IsVpnConnected)
                {
                    infoText.AppendLine($"{VpnStatusLabel} {VpnStatus}");
                    infoText.AppendLine($"{VpnIpAddressLabel} {VpnIpAddress}");
                }
                else
                {
                    infoText.AppendLine($"{VpnStatusLabel} {VpnStatus}");
                }
                Clipboard.SetText(infoText.ToString());
            }
            catch (Exception ex) { MessageBox.Show($"Error copying information to clipboard: {ex.Message}"); }
        }

        // ==================== Printer Logic ====================
        [SupportedOSPlatform("windows")]
        private async Task SearchPrintersAsync()
        {
            if (_isSearchingPrinters) return;

            try
            {
                _isSearchingPrinters = true;
                PrinterStatusMessage = _resourceManager.GetString("PrinterStatusSearching") ?? "Searching for printers...";
                Application.Current.Dispatcher.Invoke(() => AvailablePrinters.Clear());

                var foundPrinters = new ObservableCollection<Printer>();

                await Task.Run(() =>
                {
                    var scriptBuilder = new StringBuilder();
                    scriptBuilder.AppendLine("$printers = @()");
                    // Always search for the FollowYou printer
                    scriptBuilder.AppendLine("$printers += Get-Printer -ComputerName stgprfy01 -Name 'FollowYou' -ErrorAction SilentlyContinue");

                    // Also search for Ricoh and Zebra printers if a Site ID is provided
                    if (!string.IsNullOrWhiteSpace(SiteId))
                    {
                        // Ricoh printers
                        scriptBuilder.AppendLine($"$printers += Get-Printer -ComputerName STGPRPRD01.dt2kmeta.dansketraelast.dk -Name '*{SiteId}*' -ErrorAction SilentlyContinue");
                        // Zebra printers
                        scriptBuilder.AppendLine($"$printers += Get-Printer -ComputerName dt2rpr50 -Name '*{SiteId}PRZ*' -ErrorAction SilentlyContinue");
                    }

                    scriptBuilder.AppendLine("$printers | Select-Object Name, @{Name='ServerName';Expression={$_.ComputerName}}, DriverName, Comment | ConvertTo-Csv -NoTypeInformation");

                    var processStartInfo = new ProcessStartInfo("powershell.exe")
                    {
                        ArgumentList = { "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", scriptBuilder.ToString() },
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        CreateNoWindow = true
                    };

                    try
                    {
                        using (var process = Process.Start(processStartInfo))
                        {
                            if (process == null)
                            {
                                PrinterStatusMessage = "Failed to start PowerShell process.";
                                return;
                            }

                            string output = process.StandardOutput.ReadToEnd();
                            process.WaitForExit();

                            var lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                            if (lines.Length > 1)
                            {
                                foreach (var line in lines.Skip(1))
                                {
                                    var values = line.Split(',').Select(v => v.Trim('"')).ToArray();
                                    if (values.Length >= 4)
                                    {
                                        var printer = new Printer
                                        {
                                            Name = values[0],
                                            ServerName = values[1],
                                            DriverName = values[2],
                                            Comment = values[3]
                                        };
                                        foundPrinters.Add(printer);
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Application.Current.Dispatcher.Invoke(() => {
                            PrinterStatusMessage = $"An error occurred while running PowerShell: {ex.Message}";
                        });
                    }
                });

                Application.Current.Dispatcher.Invoke(() =>
                {
                    foreach (var p in foundPrinters) AvailablePrinters.Add(p);

                    if (!AvailablePrinters.Any())
                    {
                        PrinterStatusMessage = _resourceManager.GetString("PrinterStatusNotFound") ?? "No printers found.";
                    }
                    else
                    {
                        PrinterStatusMessage = $"Found {AvailablePrinters.Count} printers.";
                    }
                });
            }
            finally
            {
                _isSearchingPrinters = false;
            }
        }


        [SupportedOSPlatform("windows")]
        private async Task AddSelectedPrintersAsync()
        {
            var selectedPrinters = AvailablePrinters.Where(p => p.IsSelected).ToList();
            if (!selectedPrinters.Any())
            {
                MessageBox.Show(
                    _resourceManager.GetString("PrinterErrorNoPrinterSelected") ?? "Please select at least one printer to add.",
                    _resourceManager.GetString("PrinterErrorNoPrinterSelectedTitle") ?? "No Printer Selected",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            PrinterStatusMessage = _resourceManager.GetString("PrinterStatusInstalling") ?? "Installing selected printers, please wait...";
            var successfullyAdded = new StringBuilder();
            var failedToAdd = new StringBuilder();

            await Task.Run(() =>
            {
                foreach (var printer in selectedPrinters)
                {
                    if (string.IsNullOrEmpty(printer.Name) || string.IsNullOrEmpty(printer.ServerName)) continue;

                    try
                    {
                        var script = $"Add-Printer -ConnectionName \\\\{printer.ServerName}\\{printer.Name}";
                        var processStartInfo = new ProcessStartInfo("powershell.exe")
                        {
                            ArgumentList = { "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script },
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            RedirectStandardError = true
                        };

                        using (var process = Process.Start(processStartInfo))
                        {
                            if (process == null) continue;
                            string error = process.StandardError.ReadToEnd();
                            process.WaitForExit();

                            if (process.ExitCode == 0)
                            {
                                successfullyAdded.AppendLine($"- {printer.Name}");
                            }
                            else
                            {
                                failedToAdd.AppendLine($"- {printer.Name} (Error: {error.Trim()})");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        failedToAdd.AppendLine($"- {printer.Name} (Exception: {ex.Message})");
                    }
                }
            });

            var report = new StringBuilder();
            report.AppendLine((_resourceManager.GetString("PrinterReportComplete") ?? "Printer installation complete.") + "\n");
            if (successfullyAdded.Length > 0)
            {
                report.AppendLine(_resourceManager.GetString("PrinterReportSuccess") ?? "Successfully Added:");
                report.Append(successfullyAdded);
            }
            if (failedToAdd.Length > 0)
            {
                report.AppendLine("\n" + (_resourceManager.GetString("PrinterReportFailed") ?? "Failed to Add:"));
                report.Append(failedToAdd);
            }

            MessageBox.Show(report.ToString(), _resourceManager.GetString("PrinterReportTitle") ?? "Installation Report", MessageBoxButton.OK, MessageBoxImage.Information);
            PrinterStatusMessage = _resourceManager.GetString("PrinterStatusFinished") ?? "Installation process finished.";
        }


        // ==================== Language and Culture Logic ====================
        public Task SetLanguage(string culture)
        {
            // Set the culture first
            CultureInfo.CurrentUICulture = new CultureInfo(culture);

            // Update all state that depends on the new culture
            UpdateVpnStatus();
            string country = GetCountry();
            var countriesWithWifiFix = new[] { "DK", "NO", "SE", "GROUP" };
            var countriesWithPrinterButton = new[] { "DK", "GROUP" };

            IsPowerShellButtonVisible = countriesWithWifiFix.Contains(country);
            IsPrinterButtonVisible = countriesWithPrinterButton.Contains(country);

            IsDenmarkServiceDeskVisible = country == "DK" || country == "GROUP";
            IsNorwayServiceDeskVisible = country == "NO";
            IsSwedenServiceDeskVisible = country == "SE";

            switch (GetCountry())
            {
                case "SE":
                    NativeFlagIconPath = "Resources/swe-svg.ico";
                    break;
                case "NO":
                    NativeFlagIconPath = "Resources/nor-svg.ico";
                    break;
                case "DK":
                    NativeFlagIconPath = "Resources/dnk-svg.ico";
                    break;
                default:
                    // For any other country, including GROUP and US, we show the Danish flag as the 'native' option.
                    NativeFlagIconPath = "Resources/dnk-svg.ico";
                    break;
            }

            // Notify the UI to update everything at once.
            OnPropertyChanged(string.Empty);

            // Start the longer running background tasks without waiting for them.
            // This makes the UI feel instant.
            _ = UpdateSSID();

            return Task.CompletedTask;
        }

        private string GetInitialLanguage()
        {
            return GetCountry() switch
            {
                "SE" => "sv-SE",
                "NO" => "nb-NO",
                "DK" => "da-DK",
                _ => "en-US", // Default for GROUP and any other region
            };
        }

        public string GetCountry()
        {
            // For testing purposes, you can temporarily override the country code here.
            // For example: return "SE"; or return "NO";
            // Set to an empty string to use the actual system region.
            string _testCountryCode = "";

            if (!string.IsNullOrEmpty(_testCountryCode))
            {
                return _testCountryCode;
            }

            // Check for GROUP PC name first.
            if (Environment.MachineName.StartsWith("GROUP-", StringComparison.OrdinalIgnoreCase))
            {
                return "GROUP";
            }

            try { return RegionInfo.CurrentRegion.TwoLetterISORegionName; }
            catch (Exception) { return "US"; } // Fallback in case of an error
        }

        // ==================== System Information Fetching ====================
        private async Task FetchSystemInfo()
        {
            try
            {
                var computerNameTask = Task.Run(() => Environment.MachineName);
                var usernameTask = Task.Run(() => Environment.UserName);
                var ssidTask = GetSSIDAsync();
                var ipAddressTask = GetIPAddressAsync();
                var osVersionTask = GetOSVersionAsync();
                var lastRestartTask = GetLastRestartAsync();

                await Task.WhenAll(computerNameTask, usernameTask, ssidTask, ipAddressTask, osVersionTask, lastRestartTask);

                ComputerName = computerNameTask.Result;
                UserName = usernameTask.Result;
                Ssid = ssidTask.Result;
                IpAddress = ipAddressTask.Result;
                OsVersion = osVersionTask.Result;
                LastRestart = lastRestartTask.Result;

                // Also update VPN status after initial info is fetched
                UpdateVpnStatus();

                OnPropertyChanged(string.Empty);
            }
            catch (Exception ex) { MessageBox.Show($"An error occurred while fetching system information: {ex.Message}"); }
        }

        private async Task UpdateSSID()
        {
            try { Ssid = await GetSSIDAsync(); OnPropertyChanged(nameof(Ssid)); }
            catch (Exception ex) { MessageBox.Show($"An error occurred while updating SSID: {ex.Message}"); }
        }

        [SupportedOSPlatform("windows10.0.17763.0")]
        private async Task<string> GetSSIDAsync()
        {
            try
            {
                var access = await WiFiAdapter.RequestAccessAsync();
                if (access == WiFiAccessStatus.Allowed)
                {
                    var wifiAdapters = await WiFiAdapter.FindAllAdaptersAsync();
                    if (wifiAdapters != null)
                    {
                        foreach (var adapter in wifiAdapters)
                        {
                            var profile = await adapter.NetworkAdapter.GetConnectedProfileAsync();
                            if (profile != null)
                            {
                                return profile.ProfileName;
                            }
                        }
                    }
                }
            }
            catch
            {
                // This can fail if Wi-Fi is disabled or hardware is missing. We'll proceed to check for Ethernet.
            }

            var ethernetInterface = NetworkInterface.GetAllNetworkInterfaces()
                .FirstOrDefault(ni =>
                    ni.NetworkInterfaceType == NetworkInterfaceType.Ethernet &&
                    ni.OperationalStatus == OperationalStatus.Up &&
                    ni.GetIPProperties().GatewayAddresses.Any());

            if (ethernetInterface != null)
            {
                return _resourceManager.GetString("ConnectedViaEthernet") ?? string.Empty;
            }

            return _resourceManager.GetString("NotConnectedToWiFi") ?? string.Empty;
        }


        private async Task<string> GetIPAddressAsync()
        {
            return await Task.Run(() =>
            {
                var activeInterface = NetworkInterface.GetAllNetworkInterfaces().FirstOrDefault(ni => ni.OperationalStatus == OperationalStatus.Up && (ni.NetworkInterfaceType == NetworkInterfaceType.Ethernet || ni.NetworkInterfaceType == NetworkInterfaceType.Wireless80211) && ni.GetIPProperties().GatewayAddresses.Any());
                return activeInterface?.GetIPProperties().UnicastAddresses.FirstOrDefault(ip => ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)?.Address.ToString() ?? _resourceManager.GetString("NoActiveConnection") ?? string.Empty;
            });
        }

        [SupportedOSPlatform("windows")]
        private async Task<string> GetOSVersionAsync()
        {
            return await Task.Run(() =>
            {
                try
                {
                    using var searcher = new ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem");
                    string? fullOsName = searcher.Get().Cast<ManagementObject>().FirstOrDefault()?["Caption"]?.ToString();
                    if (fullOsName != null)
                    {
                        if (fullOsName.Contains("Windows 11")) return "Windows 11";
                        if (fullOsName.Contains("Windows 10")) return "Windows 10";
                        return fullOsName;
                    }
                    return "Unknown OS";
                }
                catch { return Environment.OSVersion.VersionString; }
            });
        }

        [SupportedOSPlatform("windows")]
        private async Task<string> GetLastRestartAsync()
        {
            return await Task.Run(() =>
            {
                try
                {
                    var searcher = new ManagementObjectSearcher("SELECT LastBootUpTime FROM Win32_OperatingSystem");
                    var lastBootUpTimeStr = searcher.Get().Cast<ManagementObject>().FirstOrDefault()?["LastBootUpTime"]?.ToString();
                    if (lastBootUpTimeStr != null) { return ManagementDateTimeConverter.ToDateTime(lastBootUpTimeStr).ToString("dd-MM-yyyy HH:mm:ss"); }
                    return "N/A";
                }
                catch { return "N/A"; }
            });
        }

        private void UpdateVpnStatus()
        {
            IsVpnConnected = IsFortiClientConnected();
            if (IsVpnConnected)
            {
                VpnStatus = _resourceManager.GetString("FortiClientConnected");
                VpnIpAddress = GetVpnIpAddress();
            }
            else
            {
                VpnStatus = _resourceManager.GetString("FortiClientDisconnected");
                VpnIpAddress = string.Empty;
            }
        }

        private bool IsFortiClientConnected()
        {
            return NetworkInterface.GetAllNetworkInterfaces().Any(ni => ni.Description.Contains("Fortinet SSL VPN Virtual Ethernet Adapter") && ni.OperationalStatus == OperationalStatus.Up);
        }

        private string? GetVpnIpAddress()
        {
            return NetworkInterface.GetAllNetworkInterfaces().Where(ni => ni.Description.Contains("Fortinet SSL VPN Virtual Ethernet Adapter") && ni.OperationalStatus == OperationalStatus.Up).SelectMany(ni => ni.GetIPProperties().UnicastAddresses).FirstOrDefault(ip => ip.Address.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)?.Address.ToString();
        }
    }

    // ==================== RelayCommand Class Definition ====================
    public class RelayCommand : ICommand
    {
        private readonly Action<object?> _execute;
        private readonly Predicate<object?>? _canExecute;
        public RelayCommand(Action<object?> execute, Predicate<object?>? canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object? parameter)
        {
            return _canExecute == null || _canExecute(parameter);
        }

        public void Execute(object? parameter) => _execute(parameter);

        public event EventHandler? CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}
