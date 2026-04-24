using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32;
using MyTools.Services;

namespace MyTools.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<NetworkData> _networkList;
        private ObservableCollection<StartupItem> _startupList;
        private ObservableCollection<DatabaseItem> _sqlDatabaseList;
        private ObservableCollection<TableItem> _sqlTableList;
        private ObservableCollection<TableItem> _allSqlTableList;
        private ObservableCollection<string> _sqlServerAddressHistory;
        private ObservableCollection<string> _sqlUsernameHistory;
        private ObservableCollection<string> _sqlPasswordHistory;
        private string _wgInterfaceName = "wg0";
        private string _wgConfig;
        private bool _isWgConnected;
        private string _wgStatusText = "未连接";
        private string _wgEndpoint;
        private string _wgAddress;
        private string _wgServerPublicKey;
        private bool _isWgSettingsOpen;
        private string _currentModule;
        private string _sqlServerAddress;
        private string _sqlPort = "1433";
        private string _sqlUsername;
        private string _sqlPassword;
        private string _sqlTableSearchText;
        private DatabaseItem _selectedSqlDatabase;
        private TableItem _selectedSqlTable;
        private ICollectionView _filteredSqlTableView;
        private string _sqlStatusMessage = "请输入 SQL Server 连接信息后测试连接。";
        private bool _isSqlBusy;
        private CancellationTokenSource _loadTablesCancellationTokenSource;

        private readonly AsyncRelayCommand _testSqlConnectionCommand;
        private readonly AsyncRelayCommand _exportSqlTableCommand;

        public MainViewModel()
        {
            NetworkList = new ObservableCollection<NetworkData>();
            StartupList = new ObservableCollection<StartupItem>();
            SqlDatabaseList = new ObservableCollection<DatabaseItem>();
            SqlTableList = new ObservableCollection<TableItem>();
            AllSqlTableList = new ObservableCollection<TableItem>();
            SqlServerAddressHistory = new ObservableCollection<string>();
            SqlUsernameHistory = new ObservableCollection<string>();
            SqlPasswordHistory = new ObservableCollection<string>();
            FilteredSqlTableView = CollectionViewSource.GetDefaultView(SqlTableList);
            FilteredSqlTableView.Filter = FilterSqlTable;

            RefreshCommand = new RelayCommand(Refresh);
            ShowNetworkCommand = new RelayCommand(() => { CurrentModule = "Network"; Refresh(); });
            ShowStartupCommand = new RelayCommand(() => { CurrentModule = "Startup"; Refresh(); });
            ShowWireGuardCommand = new RelayCommand(() => { CurrentModule = "WireGuard"; Refresh(); });
            ShowSystemCommand = new RelayCommand(() => { CurrentModule = "System"; Refresh(); });
            ShowSqlExportCommand = new RelayCommand(() => { CurrentModule = "SqlExport"; Refresh(); });
            ToggleStartupCommand = new RelayParameterCommand(obj =>
            {
                if (obj is StartupItem item)
                {
                    StartupService.ToggleStartupItem(item);
                    Refresh();
                }
            });
            DeleteStartupCommand = new RelayParameterCommand(obj =>
            {
                if (obj is StartupItem item)
                {
                    StartupService.DeleteStartupItem(item);
                    Refresh();
                }
            });
            ToggleWireGuardCommand = new AsyncRelayCommand(ToggleWireGuardAsync);
            ToggleWgSettingsCommand = new RelayCommand(() => IsWgSettingsOpen = !IsWgSettingsOpen);
            GenerateConfigCommand = new RelayCommand(GenerateConfigFromSettings);
            LockWin10Command = new RelayCommand(LockWin10Version);
            ExitCommand = new RelayCommand(ExitApplication);
            RestoreCommand = new RelayCommand(RestoreWindow);

            _testSqlConnectionCommand = new AsyncRelayCommand(TestSqlConnectionAsync, () => !IsSqlBusy);
            _exportSqlTableCommand = new AsyncRelayCommand(ExportSelectedTableAsync, CanExportSqlTable);
            TestSqlConnectionCommand = _testSqlConnectionCommand;
            ExportSqlTableCommand = _exportSqlTableCommand;

            CurrentModule = "Home";
            _ = LoadSqlConnectionHistoryAsync();
        }

        public ObservableCollection<NetworkData> NetworkList
        {
            get => _networkList;
            set { _networkList = value; OnPropertyChanged(); }
        }

        public ObservableCollection<StartupItem> StartupList
        {
            get => _startupList;
            set { _startupList = value; OnPropertyChanged(); }
        }

        public ObservableCollection<DatabaseItem> SqlDatabaseList
        {
            get => _sqlDatabaseList;
            set { _sqlDatabaseList = value; OnPropertyChanged(); }
        }

        public ObservableCollection<TableItem> SqlTableList
        {
            get => _sqlTableList;
            set
            {
                _sqlTableList = value;
                FilteredSqlTableView = CollectionViewSource.GetDefaultView(_sqlTableList);
                FilteredSqlTableView.Filter = FilterSqlTable;
                OnPropertyChanged();
            }
        }

        public ObservableCollection<TableItem> AllSqlTableList
        {
            get => _allSqlTableList;
            set { _allSqlTableList = value; OnPropertyChanged(); }
        }

        public ICollectionView FilteredSqlTableView
        {
            get => _filteredSqlTableView;
            set { _filteredSqlTableView = value; OnPropertyChanged(); }
        }

        public ObservableCollection<string> SqlServerAddressHistory
        {
            get => _sqlServerAddressHistory;
            set { _sqlServerAddressHistory = value; OnPropertyChanged(); }
        }

        public ObservableCollection<string> SqlUsernameHistory
        {
            get => _sqlUsernameHistory;
            set { _sqlUsernameHistory = value; OnPropertyChanged(); }
        }

        public ObservableCollection<string> SqlPasswordHistory
        {
            get => _sqlPasswordHistory;
            set { _sqlPasswordHistory = value; OnPropertyChanged(); }
        }

        public string WgInterfaceName
        {
            get => _wgInterfaceName;
            set { _wgInterfaceName = value; OnPropertyChanged(); }
        }

        public string WgConfig
        {
            get => _wgConfig;
            set { _wgConfig = value; OnPropertyChanged(); }
        }

        public bool IsWgConnected
        {
            get => _isWgConnected;
            set { _isWgConnected = value; OnPropertyChanged(); }
        }

        public string WgStatusText
        {
            get => _wgStatusText;
            set { _wgStatusText = value; OnPropertyChanged(); }
        }

        public string WgEndpoint
        {
            get => _wgEndpoint;
            set { _wgEndpoint = value; OnPropertyChanged(); }
        }

        public string WgAddress
        {
            get => _wgAddress;
            set { _wgAddress = value; OnPropertyChanged(); }
        }

        public string WgServerPublicKey
        {
            get => _wgServerPublicKey;
            set { _wgServerPublicKey = value; OnPropertyChanged(); }
        }

        public bool IsWgSettingsOpen
        {
            get => _isWgSettingsOpen;
            set { _isWgSettingsOpen = value; OnPropertyChanged(); }
        }

        public string CurrentModule
        {
            get => _currentModule;
            set { _currentModule = value; OnPropertyChanged(); }
        }

        public string SqlServerAddress
        {
            get => _sqlServerAddress;
            set { _sqlServerAddress = value; OnPropertyChanged(); }
        }

        public string SqlPort
        {
            get => _sqlPort;
            set { _sqlPort = value; OnPropertyChanged(); }
        }

        public string SqlUsername
        {
            get => _sqlUsername;
            set { _sqlUsername = value; OnPropertyChanged(); }
        }

        public string SqlPassword
        {
            get => _sqlPassword;
            set { _sqlPassword = value; OnPropertyChanged(); }
        }

        public string SqlTableSearchText
        {
            get => _sqlTableSearchText;
            set
            {
                if (_sqlTableSearchText == value)
                {
                    return;
                }

                _sqlTableSearchText = value;
                OnPropertyChanged();
                if (SelectedSqlTable != null && !string.Equals(SelectedSqlTable.DisplayName, value, StringComparison.OrdinalIgnoreCase))
                {
                    SelectedSqlTable = null;
                }

                FilteredSqlTableView?.Refresh();
            }
        }

        public DatabaseItem SelectedSqlDatabase
        {
            get => _selectedSqlDatabase;
            set
            {
                if (_selectedSqlDatabase == value)
                {
                    return;
                }

                _selectedSqlDatabase = value;
                OnPropertyChanged();
                SelectedSqlTable = null;
                TriggerCommandRequery();
                _ = LoadTablesForSelectedDatabaseAsync();
            }
        }

        public TableItem SelectedSqlTable
        {
            get => _selectedSqlTable;
            set
            {
                _selectedSqlTable = value;
                OnPropertyChanged();
                TriggerCommandRequery();
            }
        }

        public string SqlStatusMessage
        {
            get => _sqlStatusMessage;
            set { _sqlStatusMessage = value; OnPropertyChanged(); }
        }

        public bool IsSqlBusy
        {
            get => _isSqlBusy;
            set
            {
                _isSqlBusy = value;
                OnPropertyChanged();
                TriggerCommandRequery();
            }
        }

        public ICommand RefreshCommand { get; }
        public ICommand ShowNetworkCommand { get; }
        public ICommand ShowStartupCommand { get; }
        public ICommand ShowWireGuardCommand { get; }
        public ICommand ShowSystemCommand { get; }
        public ICommand ShowSqlExportCommand { get; }
        public ICommand ToggleStartupCommand { get; }
        public ICommand DeleteStartupCommand { get; }
        public ICommand ToggleWireGuardCommand { get; }
        public ICommand ToggleWgSettingsCommand { get; }
        public ICommand GenerateConfigCommand { get; }
        public ICommand LockWin10Command { get; }
        public ICommand ExitCommand { get; }
        public ICommand RestoreCommand { get; }
        public ICommand TestSqlConnectionCommand { get; }
        public ICommand ExportSqlTableCommand { get; }

        private void Refresh()
        {
            if (CurrentModule == "Network")
            {
                var data = NetworkService.GetAllNetworkDetails();
                NetworkList.Clear();
                foreach (var item in data)
                {
                    NetworkList.Add(item);
                }
            }
            else if (CurrentModule == "Startup")
            {
                var data = StartupService.GetStartupItems();
                StartupList.Clear();
                foreach (var item in data)
                {
                    StartupList.Add(item);
                }
            }
            else if (CurrentModule == "WireGuard")
            {
                UpdateWgStatus();
            }
            else if (CurrentModule == "SqlExport")
            {
                if (SelectedSqlDatabase != null && SqlTableList.Count == 0)
                {
                    _ = LoadTablesForSelectedDatabaseAsync();
                }
            }
        }

        private void GenerateConfigFromSettings()
        {
            string config = "[Interface]\n";
            config += "PrivateKey = <请手动填写您的私钥>\n";
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

        private async Task ToggleWireGuardAsync()
        {
            if (IsWgConnected)
            {
                await WireGuardService.DisconnectAsync(WgInterfaceName);
            }
            else
            {
                if (string.IsNullOrWhiteSpace(WgConfig))
                {
                    return;
                }

                var status = await WireGuardService.ConnectAsync(WgInterfaceName, WgConfig);
                if (!string.IsNullOrEmpty(status.ErrorMessage))
                {
                    MessageBox.Show(status.ErrorMessage, "WireGuard Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            UpdateWgStatus();
        }

        private async Task TestSqlConnectionAsync()
        {
            try
            {
                IsSqlBusy = true;
                SqlStatusMessage = "正在连接 SQL Server...";
                SqlDatabaseList.Clear();
                SqlTableList.Clear();
                AllSqlTableList.Clear();
                SelectedSqlDatabase = null;
                SelectedSqlTable = null;
                SqlTableSearchText = string.Empty;

                var options = BuildSqlConnectionOptions();
                await SqlExportService.TestConnectionAsync(options, CancellationToken.None);
                await SaveSqlConnectionHistoryAsync(options);

                SqlStatusMessage = "连接成功，正在读取数据库列表...";
                var databases = await SqlExportService.GetDatabasesAsync(options, CancellationToken.None);
                SqlDatabaseList.Clear();
                foreach (var database in databases)
                {
                    SqlDatabaseList.Add(database);
                }

                SqlStatusMessage = databases.Count > 0
                    ? $"连接成功，已加载 {databases.Count} 个数据库，请继续选择数据库和表。"
                    : "连接成功，但当前账号没有可访问的数据库。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "SQL connection test failed for {ServerAddress}", SqlServerAddress ?? string.Empty);
                SqlStatusMessage = "连接失败，请检查服务器地址、端口、用户名和密码。";
                MessageBox.Show(ex.Message, "SQL Server 连接失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsSqlBusy = false;
            }
        }

        private async Task LoadTablesForSelectedDatabaseAsync()
        {
            CancelPendingTableLoad();
            SqlTableList.Clear();
            AllSqlTableList.Clear();
            SqlTableSearchText = string.Empty;

            if (SelectedSqlDatabase == null)
            {
                SqlStatusMessage = SqlDatabaseList.Count == 0
                    ? "请输入 SQL Server 连接信息后测试连接。"
                    : "请选择数据库以加载数据表。";
                return;
            }

            var cancellationTokenSource = new CancellationTokenSource();
            _loadTablesCancellationTokenSource = cancellationTokenSource;

            try
            {
                IsSqlBusy = true;
                SqlStatusMessage = $"正在读取数据库 {SelectedSqlDatabase.Name} 的表列表...";

                var tables = await SqlExportService.GetTablesAsync(
                    BuildSqlConnectionOptions(),
                    SelectedSqlDatabase.Name,
                    cancellationTokenSource.Token);

                if (cancellationTokenSource.IsCancellationRequested)
                {
                    return;
                }

                SqlTableList.Clear();
                AllSqlTableList.Clear();
                foreach (var table in tables)
                {
                    AllSqlTableList.Add(table);
                    SqlTableList.Add(table);
                }
                FilteredSqlTableView?.Refresh();

                SqlStatusMessage = tables.Count > 0
                    ? $"已加载 {tables.Count} 张表，请选择需要导出的表。"
                    : "当前数据库下没有可导出的用户表。";
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading SQL tables failed for {DatabaseName}", SelectedSqlDatabase?.Name ?? string.Empty);
                SqlStatusMessage = "读取表列表失败。";
                MessageBox.Show(ex.Message, "读取表列表失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (ReferenceEquals(_loadTablesCancellationTokenSource, cancellationTokenSource))
                {
                    _loadTablesCancellationTokenSource = null;
                }

                IsSqlBusy = false;
                cancellationTokenSource.Dispose();
            }
        }

        private async Task ExportSelectedTableAsync()
        {
            try
            {
                var options = BuildSqlConnectionOptions();
                if (SelectedSqlDatabase == null)
                {
                    throw new InvalidOperationException("请先选择数据库。");
                }

                if (SelectedSqlTable == null)
                {
                    throw new InvalidOperationException("请先选择数据表。");
                }

                var dialog = new SaveFileDialog
                {
                    Filter = "Excel 工作簿 (*.xlsx)|*.xlsx",
                    FileName = SqlExportService.BuildDefaultFileName(SqlServerAddress, SelectedSqlDatabase.Name, SelectedSqlTable),
                    DefaultExt = ".xlsx",
                    AddExtension = true,
                    OverwritePrompt = true
                };

                if (dialog.ShowDialog() != true)
                {
                    return;
                }

                IsSqlBusy = true;
                SqlStatusMessage = "正在检查数据量并导出 Excel...";

                var exportResult = await SqlExportService.ExportTableAsync(
                    options,
                    SelectedSqlDatabase.Name,
                    SelectedSqlTable,
                    dialog.FileName,
                    CancellationToken.None);

                SqlStatusMessage = $"导出完成，共 {exportResult.RowCount} 行。";
                MessageBox.Show(
                    $"导出成功。\n文件路径：{exportResult.FilePath}",
                    "导出完成",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(
                    ex,
                    "SQL export failed for {DatabaseName}.{TableName}",
                    SelectedSqlDatabase?.Name ?? string.Empty,
                    SelectedSqlTable?.DisplayName ?? string.Empty);
                SqlStatusMessage = "导出失败。";
                MessageBox.Show(ex.Message, "导出失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsSqlBusy = false;
            }
        }

        private SqlServerConnectionOptions BuildSqlConnectionOptions()
        {
            return new SqlServerConnectionOptions
            {
                ServerAddress = SqlServerAddress,
                Port = SqlPort,
                Username = SqlUsername,
                Password = SqlPassword
            };
        }

        private async Task LoadSqlConnectionHistoryAsync()
        {
            var history = await SqlConnectionHistoryService.LoadAsync();
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                ReplaceItems(SqlServerAddressHistory, history.ServerAddresses);
                ReplaceItems(SqlUsernameHistory, history.Usernames);
                ReplaceItems(SqlPasswordHistory, history.Passwords);

                SqlServerAddress = history.LastServerAddress;
                SqlPort = string.IsNullOrWhiteSpace(history.LastPort) ? "1433" : history.LastPort;
                SqlUsername = history.LastUsername;
                SqlPassword = history.LastPassword;
            });
        }

        private async Task SaveSqlConnectionHistoryAsync(SqlServerConnectionOptions options)
        {
            await SqlConnectionHistoryService.SaveAsync(options);
            var history = await SqlConnectionHistoryService.LoadAsync();
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                ReplaceItems(SqlServerAddressHistory, history.ServerAddresses);
                ReplaceItems(SqlUsernameHistory, history.Usernames);
                ReplaceItems(SqlPasswordHistory, history.Passwords);
            });
        }

        private static void ReplaceItems<T>(ObservableCollection<T> target, IEnumerable<T> values)
        {
            target.Clear();
            foreach (var value in values ?? Enumerable.Empty<T>())
            {
                target.Add(value);
            }
        }

        private bool FilterSqlTable(object item)
        {
            if (!(item is TableItem table))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(SqlTableSearchText))
            {
                return true;
            }

            return table.DisplayName.IndexOf(SqlTableSearchText.Trim(), StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private bool CanExportSqlTable()
        {
            return !IsSqlBusy && SelectedSqlDatabase != null && SelectedSqlTable != null;
        }

        private void CancelPendingTableLoad()
        {
            if (_loadTablesCancellationTokenSource == null)
            {
                return;
            }

            _loadTablesCancellationTokenSource.Cancel();
            _loadTablesCancellationTokenSource.Dispose();
            _loadTablesCancellationTokenSource = null;
        }

        private void RestoreWindow()
        {
            var window = Application.Current.MainWindow;
            if (window != null)
            {
                window.Show();
                if (window.WindowState == WindowState.Minimized)
                {
                    window.WindowState = WindowState.Normal;
                }

                window.Activate();
            }
        }

        private void ExitApplication()
        {
            App.IsExiting = true;
            CancelPendingTableLoad();
            Application.Current.Shutdown();
        }

        private void LockWin10Version()
        {
            try
            {
                var os = Environment.OSVersion;
                if (os.Version.Major == 10 && os.Version.Build < 22000)
                {
                    string scriptPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LockWin10_22H2.ps1");

                    var startInfo = new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "powershell.exe",
                        Arguments = $"-NoProfile -ExecutionPolicy Bypass -File \"{scriptPath}\"",
                        UseShellExecute = true,
                        Verb = "runas"
                    };

                    System.Diagnostics.Process.Start(startInfo);
                    MessageBox.Show("命令已提交，请在弹出的 UAC 窗口中确认。", "操作提示", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    MessageBox.Show("当前系统不是 Windows 10，无需执行此操作。", "版本不匹配", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("执行失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void TriggerCommandRequery()
        {
            _testSqlConnectionCommand?.RaiseCanExecuteChanged();
            _exportSqlTableCommand?.RaiseCanExecuteChanged();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class RelayCommand : ICommand
    {
        private readonly Action _execute;
        private readonly Func<bool> _canExecute;

        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute();
        }

        public void Execute(object parameter)
        {
            _execute();
        }

        public event EventHandler CanExecuteChanged;

        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public class RelayParameterCommand : ICommand
    {
        private readonly Action<object> _execute;
        private readonly Func<object, bool> _canExecute;

        public RelayParameterCommand(Action<object> execute, Func<object, bool> canExecute = null)
        {
            _execute = execute;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute(parameter);
        }

        public void Execute(object parameter)
        {
            _execute(parameter);
        }

        public event EventHandler CanExecuteChanged;

        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public class AsyncRelayCommand : ICommand
    {
        private readonly Func<Task> _executeAsync;
        private readonly Func<bool> _canExecute;
        private bool _isExecuting;

        public AsyncRelayCommand(Func<Task> executeAsync, Func<bool> canExecute = null)
        {
            _executeAsync = executeAsync;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return !_isExecuting && (_canExecute == null || _canExecute());
        }

        public async void Execute(object parameter)
        {
            if (!CanExecute(parameter))
            {
                return;
            }

            try
            {
                _isExecuting = true;
                RaiseCanExecuteChanged();
                await _executeAsync();
            }
            finally
            {
                _isExecuting = false;
                RaiseCanExecuteChanged();
            }
        }

        public event EventHandler CanExecuteChanged;

        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
