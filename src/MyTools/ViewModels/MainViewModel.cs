using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;
using MyTools.Services;

namespace MyTools.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private System.Collections.ObjectModel.ObservableCollection<NetworkData> _networkList;
        public System.Collections.ObjectModel.ObservableCollection<NetworkData> NetworkList
        {
            get => _networkList;
            set { _networkList = value; OnPropertyChanged(); }
        }

        private System.Collections.ObjectModel.ObservableCollection<StartupItem> _startupList;
        public System.Collections.ObjectModel.ObservableCollection<StartupItem> StartupList
        {
            get => _startupList;
            set { _startupList = value; OnPropertyChanged(); }
        }

        private string _wgInterfaceName = "wg0";
        public string WgInterfaceName
        {
            get => _wgInterfaceName;
            set { _wgInterfaceName = value; OnPropertyChanged(); }
        }

        private string _wgConfig;
        public string WgConfig
        {
            get => _wgConfig;
            set { _wgConfig = value; OnPropertyChanged(); }
        }

        private bool _isWgConnected;
        public bool IsWgConnected
        {
            get => _isWgConnected;
            set { _isWgConnected = value; OnPropertyChanged(); }
        }

        private string _wgStatusText = "未连接";
        public string WgStatusText
        {
            get => _wgStatusText;
            set { _wgStatusText = value; OnPropertyChanged(); }
        }

        private string _wgEndpoint;
        public string WgEndpoint
        {
            get => _wgEndpoint;
            set { _wgEndpoint = value; OnPropertyChanged(); }
        }

        private string _wgAddress;
        public string WgAddress
        {
            get => _wgAddress;
            set { _wgAddress = value; OnPropertyChanged(); }
        }

        private string _wgServerPublicKey;
        public string WgServerPublicKey
        {
            get => _wgServerPublicKey;
            set { _wgServerPublicKey = value; OnPropertyChanged(); }
        }

        private bool _isWgSettingsOpen;
        public bool IsWgSettingsOpen
        {
            get => _isWgSettingsOpen;
            set { _isWgSettingsOpen = value; OnPropertyChanged(); }
        }

        private string _currentModule;
        public string CurrentModule
        {
            get => _currentModule;
            set { _currentModule = value; OnPropertyChanged(); }
        }

        public ICommand RefreshCommand { get; }
        public ICommand ShowNetworkCommand { get; }
        public ICommand ShowStartupCommand { get; }
        public ICommand ShowWireGuardCommand { get; }
        public ICommand ShowSystemCommand { get; }
        public ICommand ToggleStartupCommand { get; }
        public ICommand DeleteStartupCommand { get; }
        public ICommand ToggleWireGuardCommand { get; }
        public ICommand ToggleWgSettingsCommand { get; }
        public ICommand GenerateConfigCommand { get; }
        public ICommand LockWin10Command { get; }
        public ICommand ExitCommand { get; }
        public ICommand RestoreCommand { get; }

        public MainViewModel()
        {
            NetworkList = new System.Collections.ObjectModel.ObservableCollection<NetworkData>();
            StartupList = new System.Collections.ObjectModel.ObservableCollection<StartupItem>();
            
            RefreshCommand = new RelayCommand(Refresh);
            ShowNetworkCommand = new RelayCommand(() => { CurrentModule = "Network"; Refresh(); });
            ShowStartupCommand = new RelayCommand(() => { CurrentModule = "Startup"; Refresh(); });
            ShowWireGuardCommand = new RelayCommand(() => { CurrentModule = "WireGuard"; Refresh(); });
            ShowSystemCommand = new RelayCommand(() => { CurrentModule = "System"; Refresh(); });
            ToggleStartupCommand = new RelayParameterCommand(obj => {
                if (obj is StartupItem item) {
                    StartupService.ToggleStartupItem(item);
                    Refresh();
                }
            });
            DeleteStartupCommand = new RelayParameterCommand(obj => {
                if (obj is StartupItem item) {
                    StartupService.DeleteStartupItem(item);
                    Refresh();
                }
            });
            ToggleWireGuardCommand = new RelayCommand(async () => await ToggleWireGuard());
            ToggleWgSettingsCommand = new RelayCommand(() => IsWgSettingsOpen = !IsWgSettingsOpen);
            GenerateConfigCommand = new RelayCommand(GenerateConfigFromSettings);
            LockWin10Command = new RelayCommand(LockWin10Version);
            ExitCommand = new RelayCommand(ExitApplication);
            RestoreCommand = new RelayCommand(RestoreWindow);
            
            CurrentModule = "Home"; 
        }

        private void Refresh()
        {
            if (CurrentModule == "Network")
            {
                var data = NetworkService.GetAllNetworkDetails();
                NetworkList.Clear();
                foreach (var item in data) NetworkList.Add(item);
            }
            else if (CurrentModule == "Startup")
            {
                var data = StartupService.GetStartupItems();
                StartupList.Clear();
                foreach (var item in data) StartupList.Add(item);
            }
            else if (CurrentModule == "WireGuard")
            {
                UpdateWgStatus();
            }
        }

        private void GenerateConfigFromSettings()
        {
            // Simple generation logic
            string config = "[Interface]\n";
            config += "PrivateKey = <请输入您的私钥>\n";
            config += $"Address = {WgAddress}\n";
            config += "DNS = 8.8.8.8\n\n";
            config += "[Peer]\n";
            config += $"PublicKey = {WgServerPublicKey}\n";
            config += $"Endpoint = {WgEndpoint}\n";
            config += "AllowedIPs = 0.0.0.0/0\n";
            config += "PersistentKeepalive = 25";
            
            WgConfig = config;
            IsWgSettingsOpen = false;
        }

        private void UpdateWgStatus()
        {
            var status = WireGuardService.GetCurrentStatus(WgInterfaceName);
            IsWgConnected = status.IsConnected;
            WgStatusText = IsWgConnected ? $"已连接: {status.IpAddress}" : "未连接";
        }

        private async System.Threading.Tasks.Task ToggleWireGuard()
        {
            if (IsWgConnected)
            {
                await WireGuardService.DisconnectAsync(WgInterfaceName);
            }
            else
            {
                if (string.IsNullOrWhiteSpace(WgConfig)) return;
                var status = await WireGuardService.ConnectAsync(WgInterfaceName, WgConfig);
                if (!string.IsNullOrEmpty(status.ErrorMessage))
                {
                    System.Windows.MessageBox.Show(status.ErrorMessage, "WireGuard Error");
                }
            }
            UpdateWgStatus();
        }

        private void RestoreWindow()
        {
            var window = System.Windows.Application.Current.MainWindow;
            if (window != null)
            {
                window.Show();
                if (window.WindowState == System.Windows.WindowState.Minimized)
                    window.WindowState = System.Windows.WindowState.Normal;
                window.Activate();
            }
        }

        private void ExitApplication()
        {
            App.IsExiting = true;
            System.Windows.Application.Current.Shutdown();
        }

        private void LockWin10Version()
        {
            try
            {
                // Only for Windows 10
                var os = System.Environment.OSVersion;
                if (os.Version.Major == 10 && os.Version.Build < 22000) // Simple Win10 check
                {
                    string scriptPath = System.IO.Path.Combine(System.AppDomain.CurrentDomain.BaseDirectory, "LockWin10_22H2.ps1");
                    
                    var startInfo = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{scriptPath}\"",
                        UseShellExecute = true,
                        Verb = "runas"
                    };
                    
                    System.Diagnostics.Process.Start(startInfo);
                    System.Windows.MessageBox.Show("命令已提交，请在弹出的 UAC 窗口中确认。", "操作提示");
                }
                else
                {
                    System.Windows.MessageBox.Show("当前系统不是 Windows 10，无需执行此操作。", "版本不匹配");
                }
            }
            catch (System.Exception ex)
            {
                System.Windows.MessageBox.Show("执行失败: " + ex.Message, "错误");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly System.Action _execute;
        public RelayCommand(System.Action execute) => _execute = execute;
        public bool CanExecute(object parameter) => true;
        public void Execute(object parameter) => _execute();
        public event System.EventHandler CanExecuteChanged;
    }

    public class RelayParameterCommand : ICommand
    {
        private readonly System.Action<object> _execute;
        public RelayParameterCommand(System.Action<object> execute) => _execute = execute;
        public bool CanExecute(object parameter) => true;
        public void Execute(object parameter) => _execute(parameter);
        public event System.EventHandler CanExecuteChanged;
    }
}
