using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
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
        private ObservableCollection<CodexProfileItem> _codexProfiles;
        private string _wgInterfaceName = "wg0";
        private string _wgConfig;
        private bool _isWgConnected;
        private string _wgStatusText = "鏈繛鎺?";
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
        private string _sqlStatusMessage = "璇疯緭鍏?SQL Server 杩炴帴淇℃伅鍚庢祴璇曡繛鎺ャ€?";
        private bool _isSqlBusy;
        private CancellationTokenSource _loadTablesCancellationTokenSource;
        private bool _suppressSqlTableAutoLoad;
        private bool _isApplyingSqlHistory;
        private bool _hasUserModifiedSqlConnectionInputs;
        private SqlServerConnectionOptions _activeSqlConnectionOptions;
        private bool _isDefenderEnabled = true;
        private bool _isAutoUpdateEnabled = true;
        private string _systemStatusMessage = string.Empty;
        private bool _isWgBusy;
        private string _sqlQueryText = string.Empty;
        private DataView _sqlQueryResult;
        private bool _isQueryBusy;
        private string _queryStatusMessage = string.Empty;

        private bool _showEditorAfterCapture = true;
        private string _screenshotHotkeyText = "Ctrl+Shift+Z";
        private bool _isCapturingHotkey;
        private bool _isScreenshotBusy;
        private ScreenshotEditorWindow _screenshotEditorWindow;
        private uint _pendingModifiers = 0x0006;
        private uint _pendingKey = 0x5A;
        private string _codexProfilesStatusMessage = "鎷栧叆鍖呭惈 config.toml 鍜?auth.json 鐨勬枃浠跺す锛岀敓鎴愬彲搴旂敤鐨?Codex 閰嶇疆璁板綍銆?";

        private readonly AsyncRelayCommand _executeQueryCommand;
        private readonly AsyncRelayCommand _exportQueryResultCommand;
        private readonly AsyncRelayParameterCommand _applyCodexProfileCommand;

        private readonly AsyncRelayCommand _testSqlConnectionCommand;
        private readonly AsyncRelayCommand _exportSqlTableCommand;
        private readonly AsyncRelayCommand _toggleDefenderCommand;
        private readonly AsyncRelayCommand _toggleAutoUpdateCommand;
        private readonly AsyncRelayCommand _triggerUpdateNowCommand;

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
            CodexProfiles = new ObservableCollection<CodexProfileItem>();
            FilteredSqlTableView = CollectionViewSource.GetDefaultView(SqlTableList);
            FilteredSqlTableView.Filter = FilterSqlTable;

            RefreshCommand = new RelayCommand(Refresh);
            ShowNetworkCommand = new RelayCommand(() => { CurrentModule = "Network"; Refresh(); });
            ShowStartupCommand = new RelayCommand(() => { CurrentModule = "Startup"; Refresh(); });
            ShowWireGuardCommand = new RelayCommand(() => { CurrentModule = "WireGuard"; Refresh(); });
            ShowSystemCommand = new RelayCommand(() => { CurrentModule = "System"; Refresh(); });
            ShowSqlExportCommand = new RelayCommand(() => { CurrentModule = "SqlExport"; Refresh(); });
            ShowCodexProfilesCommand = new RelayCommand(() => { CurrentModule = "CodexProfiles"; });
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
                if (!(obj is StartupItem item)) return;
                var confirm = MessageBox.Show(
                    $"纭畾瑕佹案涔呭垹闄ゅ惎鍔ㄩ」 \"{item.Name}\" 鍚楋紵\n姝ゆ搷浣滀笉鍙仮澶嶃€?",
                    "纭鍒犻櫎",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);
                if (confirm == MessageBoxResult.Yes)
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

            _toggleDefenderCommand = new AsyncRelayCommand(ToggleDefenderAsync);
            _toggleAutoUpdateCommand = new AsyncRelayCommand(ToggleAutoUpdateAsync);
            _triggerUpdateNowCommand = new AsyncRelayCommand(TriggerUpdateNowAsync);
            ToggleDefenderCommand = _toggleDefenderCommand;
            ToggleAutoUpdateCommand = _toggleAutoUpdateCommand;
            TriggerUpdateNowCommand = _triggerUpdateNowCommand;
            RefreshSystemStatusCommand = new RelayCommand(RefreshSystemStatus);

            _executeQueryCommand = new AsyncRelayCommand(ExecuteSqlQueryAsync, CanExecuteSqlQuery);
            _exportQueryResultCommand = new AsyncRelayCommand(ExportQueryResultAsync, CanExportQueryResult);
            ExecuteSqlQueryCommand = _executeQueryCommand;
            ExportQueryResultCommand = _exportQueryResultCommand;

            _testSqlConnectionCommand = new AsyncRelayCommand(TestSqlConnectionAsync, () => !IsSqlBusy);
            _exportSqlTableCommand = new AsyncRelayCommand(ExportSelectedTableAsync, CanExportSqlTable);
            TestSqlConnectionCommand = _testSqlConnectionCommand;
            ExportSqlTableCommand = _exportSqlTableCommand;

            ShowScreenshotCommand = new RelayCommand(() => { CurrentModule = "Screenshot"; });
            TakeScreenshotNowCommand = new AsyncRelayCommand(TriggerScreenshotAsync);
            StartCaptureHotkeyCommand = new RelayCommand(() => IsCapturingHotkey = true);
            CancelCaptureHotkeyCommand = new RelayCommand(() => IsCapturingHotkey = false);
            SaveScreenshotSettingsCommand = new AsyncRelayCommand(SaveScreenshotSettingsAsync);
            EditClipboardImageCommand = new RelayCommand(EditClipboardImage);

            _applyCodexProfileCommand = new AsyncRelayParameterCommand(ApplyCodexProfileAsync);
            ApplyCodexProfileCommand = _applyCodexProfileCommand;
            DeleteCodexProfileCommand = new RelayParameterCommand(DeleteCodexProfile);

            CurrentModule = "Home";
            _ = LoadSqlConnectionHistoryAsync();
            _ = LoadScreenshotSettingsAsync();
            _ = LoadCodexProfilesAsync();
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

        public ObservableCollection<CodexProfileItem> CodexProfiles
        {
            get => _codexProfiles;
            set { _codexProfiles = value; OnPropertyChanged(); }
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
            set
            {
                if (string.Equals(_sqlServerAddress, value, StringComparison.Ordinal))
                {
                    return;
                }

                _sqlServerAddress = value;
                MarkSqlConnectionInputAsUserModified();
                OnPropertyChanged();
            }
        }

        public string SqlPort
        {
            get => _sqlPort;
            set
            {
                if (string.Equals(_sqlPort, value, StringComparison.Ordinal))
                {
                    return;
                }

                _sqlPort = value;
                MarkSqlConnectionInputAsUserModified();
                OnPropertyChanged();
            }
        }

        public string SqlUsername
        {
            get => _sqlUsername;
            set
            {
                if (string.Equals(_sqlUsername, value, StringComparison.Ordinal))
                {
                    return;
                }

                _sqlUsername = value;
                MarkSqlConnectionInputAsUserModified();
                OnPropertyChanged();
            }
        }

        public string SqlPassword
        {
            get => _sqlPassword;
            set
            {
                if (string.Equals(_sqlPassword, value, StringComparison.Ordinal))
                {
                    return;
                }

                _sqlPassword = value;
                MarkSqlConnectionInputAsUserModified();
                OnPropertyChanged();
            }
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
                if (!_suppressSqlTableAutoLoad)
                {
                    _ = LoadTablesForSelectedDatabaseAsync();
                }
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
        public ICommand ShowCodexProfilesCommand { get; }
        public ICommand ToggleStartupCommand { get; }
        public ICommand DeleteStartupCommand { get; }
        public ICommand ToggleWireGuardCommand { get; }
        public ICommand ToggleWgSettingsCommand { get; }
        public ICommand GenerateConfigCommand { get; }
        public bool IsWgBusy
        {
            get => _isWgBusy;
            set { _isWgBusy = value; OnPropertyChanged(); }
        }

        public bool HasNoStartupItems => StartupList.Count == 0;

        public bool WgExeFound => WireGuardService.IsExeAvailable;
        public bool WgExeNotFound => !WireGuardService.IsExeAvailable;

        public string SqlQueryText
        {
            get => _sqlQueryText;
            set { _sqlQueryText = value; OnPropertyChanged(); _executeQueryCommand?.RaiseCanExecuteChanged(); }
        }

        public DataView SqlQueryResult
        {
            get => _sqlQueryResult;
            set { _sqlQueryResult = value; OnPropertyChanged(); _exportQueryResultCommand?.RaiseCanExecuteChanged(); }
        }

        public bool IsQueryBusy
        {
            get => _isQueryBusy;
            set { _isQueryBusy = value; OnPropertyChanged(); _executeQueryCommand?.RaiseCanExecuteChanged(); _exportQueryResultCommand?.RaiseCanExecuteChanged(); }
        }

        public string QueryStatusMessage
        {
            get => _queryStatusMessage;
            set { _queryStatusMessage = value; OnPropertyChanged(); }
        }

        public ICommand ExecuteSqlQueryCommand { get; }
        public ICommand ExportQueryResultCommand { get; }

        public bool IsDefenderEnabled
        {
            get => _isDefenderEnabled;
            set { _isDefenderEnabled = value; OnPropertyChanged(); }
        }

        public bool IsAutoUpdateEnabled
        {
            get => _isAutoUpdateEnabled;
            set { _isAutoUpdateEnabled = value; OnPropertyChanged(); }
        }

        public string SystemStatusMessage
        {
            get => _systemStatusMessage;
            set { _systemStatusMessage = value; OnPropertyChanged(); }
        }

        public bool ShowEditorAfterCapture
        {
            get => _showEditorAfterCapture;
            set { _showEditorAfterCapture = value; OnPropertyChanged(); }
        }

        public string ScreenshotHotkeyText
        {
            get => _screenshotHotkeyText;
            set { _screenshotHotkeyText = value; OnPropertyChanged(); }
        }

        public bool IsCapturingHotkey
        {
            get => _isCapturingHotkey;
            set { _isCapturingHotkey = value; OnPropertyChanged(); }
        }

        public ICommand ShowScreenshotCommand { get; }
        public ICommand TakeScreenshotNowCommand { get; }
        public ICommand StartCaptureHotkeyCommand { get; }
        public ICommand CancelCaptureHotkeyCommand { get; }
        public ICommand SaveScreenshotSettingsCommand { get; }
        public ICommand EditClipboardImageCommand { get; }
        public ICommand ApplyCodexProfileCommand { get; }
        public ICommand DeleteCodexProfileCommand { get; }

        public string CodexProfilesStatusMessage
        {
            get => _codexProfilesStatusMessage;
            set { _codexProfilesStatusMessage = value; OnPropertyChanged(); }
        }

        public ICommand LockWin10Command { get; }
        public ICommand ExitCommand { get; }
        public ICommand RestoreCommand { get; }
        public ICommand TestSqlConnectionCommand { get; }
        public ICommand ExportSqlTableCommand { get; }
        public ICommand ToggleDefenderCommand { get; }
        public ICommand ToggleAutoUpdateCommand { get; }
        public ICommand TriggerUpdateNowCommand { get; }
        public ICommand RefreshSystemStatusCommand { get; }

        public void AddCodexProfileFolders(IEnumerable<string> folderPaths)
        {
            _ = AddCodexProfileFoldersAsync(folderPaths);
        }

        public async Task AddCodexProfileFoldersAsync(IEnumerable<string> folderPaths)
        {
            var addedCount = 0;
            var updatedCount = 0;
            var failedCount = 0;

            foreach (var folderPath in folderPaths ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath))
                {
                    continue;
                }

                var fullPath = NormalizeFolderPath(folderPath);
                try
                {
                    var sourceFiles = await CodexConfigProfileService.ReadProfileFromFolderAsync(fullPath, CancellationToken.None);
                    var protectedConfig = CodexConfigProfileService.ProtectBytesToBase64(sourceFiles.ConfigTomlBytes);
                    var protectedAuth = CodexConfigProfileService.ProtectBytesToBase64(sourceFiles.AuthJsonBytes);

                    var existing = CodexProfiles.FirstOrDefault(item =>
                        string.Equals(item.FolderPath, fullPath, StringComparison.OrdinalIgnoreCase));
                    if (existing != null)
                    {
                        existing.ConfigTomlContentProtected = protectedConfig;
                        existing.AuthJsonContentProtected = protectedAuth;
                        existing.StatusMessage = "配置内容已更新。";
                        updatedCount++;
                        continue;
                    }

                    var profileItem = CreateCodexProfileItem(fullPath, null, null, protectedConfig, protectedAuth);
                    profileItem.StatusMessage = "已保存配置内容。";
                    AddCodexProfileItem(profileItem);
                    addedCount++;
                }
                catch (Exception ex)
                {
                    failedCount++;
                    AppLogService.Error(ex, "Adding Codex profile folder failed for {FolderPath}", fullPath);
                }
            }

            if (addedCount > 0 || updatedCount > 0)
            {
                CurrentModule = "CodexProfiles";
                CodexProfilesStatusMessage = $"已保存 {addedCount + updatedCount} 条记录（新增 {addedCount}，更新 {updatedCount}）。";
                if (failedCount > 0)
                {
                    CodexProfilesStatusMessage += $" 失败 {failedCount} 条（请检查是否同时包含 {CodexConfigProfileService.ConfigFileName} 和 {CodexConfigProfileService.AuthFileName}）。";
                }

                await SaveCodexProfilesAsync();
                return;
            }

            CodexProfilesStatusMessage = failedCount > 0
                ? $"未添加任何记录。失败 {failedCount} 条（文件夹需同时包含 {CodexConfigProfileService.ConfigFileName} 和 {CodexConfigProfileService.AuthFileName}）。"
                : "未检测到可添加的文件夹。";
        }

        private async Task LoadCodexProfilesAsync()
        {
            try
            {
                var settings = await AppSettingsService.LoadAsync();
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    foreach (var item in CodexProfiles)
                    {
                        item.PropertyChanged -= CodexProfileItem_OnPropertyChanged;
                    }

                    CodexProfiles.Clear();
                    var profiles = settings.CodexProfiles ?? new List<CodexProfileSettings>();
                    foreach (var profile in profiles)
                    {
                        if (profile == null)
                        {
                            continue;
                        }

                        AddCodexProfileItem(CreateCodexProfileItem(
                            profile.FolderPath,
                            profile.Name,
                            profile.Remark,
                            profile.ConfigTomlContentProtected,
                            profile.AuthJsonContentProtected));
                    }
                });
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading Codex config profiles failed.");
                CodexProfilesStatusMessage = "璇诲彇 Codex 閰嶇疆璁板綍澶辫触銆?";
            }
        }

        private CodexProfileItem CreateCodexProfileItem(
            string folderPath,
            string name,
            string remark,
            string configTomlContentProtected,
            string authJsonContentProtected)
        {
            var normalizedPath = NormalizeFolderPath(folderPath);
            var defaultName = ResolveCodexProfileName(name, normalizedPath);
            return new CodexProfileItem
            {
                Name = defaultName,
                Remark = string.IsNullOrWhiteSpace(remark) ? defaultName : remark,
                FolderPath = normalizedPath,
                ConfigTomlContentProtected = configTomlContentProtected,
                AuthJsonContentProtected = authJsonContentProtected,
                StatusMessage = string.Empty
            };
        }

        private void AddCodexProfileItem(CodexProfileItem item)
        {
            item.PropertyChanged += CodexProfileItem_OnPropertyChanged;
            CodexProfiles.Add(item);
        }

        private void CodexProfileItem_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(CodexProfileItem.Remark))
            {
                _ = SaveCodexProfilesAsync();
            }
        }

        private async Task SaveCodexProfilesAsync()
        {
            var profiles = CodexProfiles
                .Where(item =>
                    !string.IsNullOrWhiteSpace(item.Name)
                    && (!string.IsNullOrWhiteSpace(item.FolderPath)
                        || !string.IsNullOrWhiteSpace(item.ConfigTomlContentProtected)
                        || !string.IsNullOrWhiteSpace(item.AuthJsonContentProtected)))
                .Select(item => new CodexProfileSettings
                {
                    Name = item.Name,
                    Remark = item.Remark,
                    FolderPath = NormalizeFolderPath(item.FolderPath),
                    ConfigTomlContentProtected = item.ConfigTomlContentProtected,
                    AuthJsonContentProtected = item.AuthJsonContentProtected
                })
                .ToList();

            await AppSettingsService.UpdateAsync(settings => settings.CodexProfiles = profiles);
        }

        private async Task ApplyCodexProfileAsync(object parameter)
        {
            if (!(parameter is CodexProfileItem item))
            {
                return;
            }

            try
            {
                item.IsApplying = true;
                item.StatusMessage = "正在应用...";

                var configTomlBytes = CodexConfigProfileService.UnprotectBytesFromBase64(item.ConfigTomlContentProtected);
                var authJsonBytes = CodexConfigProfileService.UnprotectBytesFromBase64(item.AuthJsonContentProtected);

                if (configTomlBytes == null || authJsonBytes == null)
                {
                    var fallbackFolderPath = NormalizeFolderPath(item.FolderPath);
                    if (string.IsNullOrWhiteSpace(fallbackFolderPath) || !Directory.Exists(fallbackFolderPath))
                    {
                        throw new InvalidOperationException("该记录未保存配置内容，且来源文件夹已不存在，请重新拖入配置文件夹。");
                    }

                    var sourceFiles = await CodexConfigProfileService.ReadProfileFromFolderAsync(fallbackFolderPath, CancellationToken.None);
                    configTomlBytes = sourceFiles.ConfigTomlBytes;
                    authJsonBytes = sourceFiles.AuthJsonBytes;
                    item.ConfigTomlContentProtected = CodexConfigProfileService.ProtectBytesToBase64(configTomlBytes);
                    item.AuthJsonContentProtected = CodexConfigProfileService.ProtectBytesToBase64(authJsonBytes);
                }

                var result = await CodexConfigProfileService.ApplyAsync(configTomlBytes, authJsonBytes, CancellationToken.None);
                item.StatusMessage = $"已应用到 {result.TargetFolderPath}：{DateTime.Now:yyyy-MM-dd HH:mm:ss}";
                CodexProfilesStatusMessage = $"已应用「{item.Name}」。";
                await SaveCodexProfilesAsync();
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Applying Codex config profile failed for {ProfileName}", item.Name ?? string.Empty);
                item.StatusMessage = "应用失败：" + ex.Message;
                MessageBox.Show(ex.Message, "Codex 配置应用失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                item.IsApplying = false;
            }
        }

        private void DeleteCodexProfile(object parameter)
        {
            if (!(parameter is CodexProfileItem item))
            {
                return;
            }

            var result = MessageBox.Show(
                $"确定删除记录 \"{item.Name}\" 吗？",
                "确认删除",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            item.PropertyChanged -= CodexProfileItem_OnPropertyChanged;
            CodexProfiles.Remove(item);
            CodexProfilesStatusMessage = $"已删除「{item.Name}」。";
            _ = SaveCodexProfilesAsync();
        }

        private static string ResolveCodexProfileName(string name, string normalizedFolderPath)
        {
            if (!string.IsNullOrWhiteSpace(name))
            {
                return name;
            }

            if (!string.IsNullOrWhiteSpace(normalizedFolderPath))
            {
                try
                {
                    return new DirectoryInfo(normalizedFolderPath).Name;
                }
                catch
                {
                }
            }

            return "閰嶇疆璁板綍";
        }

        private static string NormalizeFolderPath(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
            {
                return string.Empty;
            }

            try
            {
                return Path.GetFullPath(folderPath);
            }
            catch
            {
                return folderPath;
            }
        }

        private void MarkSqlConnectionInputAsUserModified()
        {
            if (_isApplyingSqlHistory)
            {
                return;
            }

            _hasUserModifiedSqlConnectionInputs = true;
            _activeSqlConnectionOptions = null;
            if (SqlDatabaseList.Count == 0 && SqlTableList.Count == 0 && SelectedSqlDatabase == null && SelectedSqlTable == null)
            {
                return;
            }

            CancelPendingTableLoad();
            _suppressSqlTableAutoLoad = true;
            try
            {
                SelectedSqlDatabase = null;
                SelectedSqlTable = null;
            }
            finally
            {
                _suppressSqlTableAutoLoad = false;
            }

            SqlDatabaseList.Clear();
            SqlTableList.Clear();
            AllSqlTableList.Clear();
            SqlTableSearchText = string.Empty;
            SqlStatusMessage = "杩炴帴淇℃伅宸插彉鍖栵紝璇烽噸鏂版祴璇曡繛鎺ャ€?";
        }

        private void EditClipboardImage()
        {
            var image = System.Windows.Clipboard.GetImage();
            if (image == null)
            {
                MessageBox.Show("鍓创鏉夸腑娌℃湁鍥剧墖", "鎻愮ず", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            ShowScreenshotEditorWindow(image);
        }

        public void ReRegisterHotkey()
        {
            if (_pendingKey == 0)
            {
                return;
            }

            _ = TryRegisterHotkeyAsync(_pendingModifiers, _pendingKey, false);
        }

        public void ApplyPendingHotkey(uint modifiers, uint key)
        {
            var previousModifiers = _pendingModifiers;
            var previousKey = _pendingKey;
            var previousDisplayText = ScreenshotHotkeyText;

            _pendingModifiers = modifiers;
            _pendingKey = key;
            ScreenshotHotkeyText = HotkeyService.BuildDisplayText(modifiers, key);

            _ = TryRegisterHotkeyAsync(modifiers, key, true, () =>
            {
                _pendingModifiers = previousModifiers;
                _pendingKey = previousKey;
                ScreenshotHotkeyText = previousDisplayText;
            });

            IsCapturingHotkey = false;
        }

        public async Task TriggerScreenshotAsync()
        {
            if (_isScreenshotBusy)
            {
                return;
            }

            _isScreenshotBusy = true;
            Window mainWin = null;
            var shouldRestoreMainWindow = false;

            try
            {
                mainWin = Application.Current?.MainWindow;
                var wasVisible = mainWin?.IsVisible == true;
                if (wasVisible)
                {
                    mainWin.Hide();
                    shouldRestoreMainWindow = true;
                }

                await Task.Delay(150);

                var screenshot = await Task.Run(() => ScreenshotService.CaptureFullScreen());

                System.Windows.Clipboard.SetImage(screenshot);

                if (ShowEditorAfterCapture)
                {
                    ShowScreenshotEditorWindow(screenshot, () =>
                    {
                        if (!wasVisible)
                        {
                            return;
                        }

                        mainWin?.Show();
                        mainWin?.Activate();
                    });
                    shouldRestoreMainWindow = false;
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Screenshot failed");
                MessageBox.Show(ex.Message, "鎴浘澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (shouldRestoreMainWindow)
                {
                    mainWin?.Show();
                    mainWin?.Activate();
                }

                _isScreenshotBusy = false;
            }
        }

        private void ShowScreenshotEditorWindow(System.Windows.Media.Imaging.BitmapSource screenshot, Action onClosed = null)
        {
            if (_screenshotEditorWindow != null)
            {
                _screenshotEditorWindow.Closed -= HandleScreenshotEditorWindowClosed;
                _screenshotEditorWindow.Close();
                _screenshotEditorWindow = null;
            }

            var editor = new ScreenshotEditorWindow();
            editor.Closed += HandleScreenshotEditorWindowClosed;
            if (onClosed != null)
            {
                editor.Closed += (sender, args) => onClosed();
            }

            editor.LoadScreenshot(screenshot);
            editor.Show();
            _screenshotEditorWindow = editor;
        }

        private void HandleScreenshotEditorWindowClosed(object sender, EventArgs e)
        {
            if (ReferenceEquals(_screenshotEditorWindow, sender))
            {
                _screenshotEditorWindow = null;
            }
        }

        private async Task LoadScreenshotSettingsAsync()
        {
            try
            {
                var settings = await AppSettingsService.LoadAsync();
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    _pendingModifiers = settings.ScreenshotHotkey.Modifiers;
                    _pendingKey = settings.ScreenshotHotkey.Key;
                    ShowEditorAfterCapture = settings.ShowEditorAfterCapture;
                    ScreenshotHotkeyText = string.IsNullOrWhiteSpace(settings.ScreenshotHotkey.DisplayText)
                        ? HotkeyService.BuildDisplayText(_pendingModifiers, _pendingKey)
                        : settings.ScreenshotHotkey.DisplayText;
                });

                await TryRegisterHotkeyAsync(_pendingModifiers, _pendingKey, false);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "LoadScreenshotSettings failed");
            }
        }

        private async Task SaveScreenshotSettingsAsync()
        {
            try
            {
                if (_pendingKey == 0) return;
                await AppSettingsService.UpdateAsync(settings =>
                {
                    settings.ScreenshotHotkey = new HotkeySettings
                    {
                        Modifiers = _pendingModifiers,
                        Key = _pendingKey,
                        DisplayText = ScreenshotHotkeyText
                    };
                    settings.ShowEditorAfterCapture = ShowEditorAfterCapture;
                });

                var registered = await TryRegisterHotkeyAsync(_pendingModifiers, _pendingKey, true);
                if (!registered)
                {
                    return;
                }
                Application.Current?.Dispatcher.Invoke(() =>
                    MessageBox.Show("设置已保存", "完成", MessageBoxButton.OK, MessageBoxImage.Information));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "SaveScreenshotSettings failed");
            }
        }

        private async Task<bool> TryRegisterHotkeyAsync(uint modifiers, uint key, bool showMessageOnFailure, Action onRegisterFailed = null)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null)
            {
                return false;
            }

            return await dispatcher.InvokeAsync(() =>
            {
                var registered = HotkeyService.Register(modifiers, key);
                if (registered || !showMessageOnFailure)
                {
                    return registered;
                }

                onRegisterFailed?.Invoke();
                MessageBox.Show(
                    BuildHotkeyRegistrationErrorMessage(modifiers, key, HotkeyService.LastWin32ErrorCode),
                    "蹇嵎閿笉鍙敤",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return false;
            });
        }

        private static string BuildHotkeyRegistrationErrorMessage(uint modifiers, uint key, int errorCode)
        {
            var hotkeyText = HotkeyService.BuildDisplayText(modifiers, key);
            if (errorCode == 1409)
            {
                return $"蹇嵎閿?{hotkeyText} 宸茶鍏朵粬绋嬪簭鍗犵敤锛岃鎹竴涓粍鍚堛€?";
            }

            if (errorCode > 0)
            {
                return $"鏃犳硶娉ㄥ唽蹇嵎閿?{hotkeyText}锛學in32 閿欒鐮侊細{errorCode}銆?";
            }

            return $"鏃犳硶娉ㄥ唽蹇嵎閿?{hotkeyText}锛岃鎹竴涓粍鍚堛€?";
        }

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
                OnPropertyChanged(nameof(HasNoStartupItems));
            }
            else if (CurrentModule == "WireGuard")
            {
                UpdateWgStatus();
                if (string.IsNullOrWhiteSpace(WgConfig))
                {
                    var saved = WireGuardService.GetSavedConfig(WgInterfaceName);
                    if (saved != null) WgConfig = saved;
                }
            }
            else if (CurrentModule == "SqlExport")
            {
                if (SelectedSqlDatabase != null && SqlTableList.Count == 0)
                {
                    _ = LoadTablesForSelectedDatabaseAsync();
                }
            }
            else if (CurrentModule == "System")
            {
                RefreshSystemStatus();
            }
        }

        private void GenerateConfigFromSettings()
        {
            string config = "[Interface]\n";
            config += "PrivateKey = <璇锋墜鍔ㄥ～鍐欐偍鐨勭閽?\n";
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
            WgStatusText = IsWgConnected ? $"宸茶繛鎺? {status.IpAddress}" : "鏈繛鎺?";
        }

        private async Task ToggleWireGuardAsync()
        {
            IsWgBusy = true;
            try
            {
                if (IsWgConnected)
                {
                    WgStatusText = "姝ｅ湪鏂紑...";
                    await WireGuardService.DisconnectAsync(WgInterfaceName);
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(WgConfig)) return;
                    WgStatusText = "姝ｅ湪杩炴帴...";
                    var status = await WireGuardService.ConnectAsync(WgInterfaceName, WgConfig);
                    if (!string.IsNullOrEmpty(status.ErrorMessage))
                    {
                        MessageBox.Show(status.ErrorMessage, "WireGuard 杩炴帴澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
            }
            finally
            {
                IsWgBusy = false;
                UpdateWgStatus();
            }
        }

        private async Task TestSqlConnectionAsync()
        {
            try
            {
                IsSqlBusy = true;
                SqlStatusMessage = "姝ｅ湪杩炴帴 SQL Server...";
                CancelPendingTableLoad(clearBusy: false);
                SqlDatabaseList.Clear();
                SqlTableList.Clear();
                AllSqlTableList.Clear();
                _suppressSqlTableAutoLoad = true;
                try
                {
                    SelectedSqlDatabase = null;
                    SelectedSqlTable = null;
                }
                finally
                {
                    _suppressSqlTableAutoLoad = false;
                }

                SqlTableSearchText = string.Empty;

                var options = BuildSqlConnectionOptions();
                await SqlExportService.TestConnectionAsync(options, CancellationToken.None);
                _activeSqlConnectionOptions = CloneSqlConnectionOptions(options);
                _hasUserModifiedSqlConnectionInputs = false;
                await SaveSqlConnectionHistoryAsync(options);

                SqlStatusMessage = "杩炴帴鎴愬姛锛屾鍦ㄨ鍙栨暟鎹簱鍒楄〃...";
                var databases = await SqlExportService.GetDatabasesAsync(options, CancellationToken.None);
                SqlDatabaseList.Clear();
                foreach (var database in databases)
                {
                    SqlDatabaseList.Add(database);
                }

                SqlStatusMessage = databases.Count > 0
                    ? $"杩炴帴鎴愬姛锛屽凡鍔犺浇 {databases.Count} 涓暟鎹簱锛岃缁х画閫夋嫨鏁版嵁搴撳拰琛ㄣ€?"
                    : "杩炴帴鎴愬姛锛屼絾褰撳墠璐﹀彿娌℃湁鍙闂殑鏁版嵁搴撱€?";
            }
            catch (Exception ex)
            {
                _activeSqlConnectionOptions = null;
                AppLogService.Error(ex, "SQL connection test failed for {ServerAddress}", SqlServerAddress ?? string.Empty);
                SqlStatusMessage = "杩炴帴澶辫触锛岃妫€鏌ユ湇鍔″櫒鍦板潃銆佺鍙ｃ€佺敤鎴峰悕鍜屽瘑鐮併€?";
                MessageBox.Show(ex.Message, "SQL Server 杩炴帴澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsSqlBusy = false;
            }
        }

        private async Task LoadTablesForSelectedDatabaseAsync()
        {
            CancelPendingTableLoad(clearBusy: false);
            SqlTableList.Clear();
            AllSqlTableList.Clear();
            SqlTableSearchText = string.Empty;

            if (SelectedSqlDatabase == null)
            {
                IsSqlBusy = false;
                SqlStatusMessage = SqlDatabaseList.Count == 0
                    ? "璇疯緭鍏?SQL Server 杩炴帴淇℃伅鍚庢祴璇曡繛鎺ャ€?"
                    : "璇烽€夋嫨鏁版嵁搴撲互鍔犺浇鏁版嵁琛ㄣ€?";
                return;
            }

            var cancellationTokenSource = new CancellationTokenSource();
            _loadTablesCancellationTokenSource = cancellationTokenSource;

            try
            {
                IsSqlBusy = true;
                SqlStatusMessage = $"姝ｅ湪璇诲彇鏁版嵁搴?{SelectedSqlDatabase.Name} 鐨勮〃鍒楄〃...";

                var tables = await SqlExportService.GetTablesAsync(
                    GetEffectiveSqlConnectionOptions(),
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
                    ? $"宸插姞杞?{tables.Count} 寮犺〃锛岃閫夋嫨闇€瑕佸鍑虹殑琛ㄣ€?"
                    : "褰撳墠鏁版嵁搴撲笅娌℃湁鍙鍑虹殑鐢ㄦ埛琛ㄣ€?";
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading SQL tables failed for {DatabaseName}", SelectedSqlDatabase?.Name ?? string.Empty);
                SqlStatusMessage = "璇诲彇琛ㄥ垪琛ㄥけ璐ャ€?";
                MessageBox.Show(ex.Message, "读取表列表失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                if (ReferenceEquals(_loadTablesCancellationTokenSource, cancellationTokenSource))
                {
                    _loadTablesCancellationTokenSource = null;
                    IsSqlBusy = false;
                }

                cancellationTokenSource.Dispose();
            }
        }

        private async Task ExportSelectedTableAsync()
        {
            try
            {
                var options = GetEffectiveSqlConnectionOptions();
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
                    Filter = "Excel 宸ヤ綔绨?(*.xlsx)|*.xlsx",
                    FileName = SqlExportService.BuildDefaultFileName(options.ServerAddress, SelectedSqlDatabase.Name, SelectedSqlTable),
                    DefaultExt = ".xlsx",
                    AddExtension = true,
                    OverwritePrompt = true
                };

                if (dialog.ShowDialog() != true)
                {
                    return;
                }

                IsSqlBusy = true;
                SqlStatusMessage = "姝ｅ湪妫€鏌ユ暟鎹噺骞跺鍑?Excel...";

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
                SqlStatusMessage = "瀵煎嚭澶辫触銆?";
                MessageBox.Show(ex.Message, "瀵煎嚭澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
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
                ServerAddress = SqlServerAddress?.Trim(),
                Port = SqlPort?.Trim(),
                Username = SqlUsername?.Trim(),
                Password = SqlPassword
            };
        }

        private SqlServerConnectionOptions GetEffectiveSqlConnectionOptions()
        {
            return _activeSqlConnectionOptions != null
                ? CloneSqlConnectionOptions(_activeSqlConnectionOptions)
                : BuildSqlConnectionOptions();
        }

        private static SqlServerConnectionOptions CloneSqlConnectionOptions(SqlServerConnectionOptions options)
        {
            if (options == null)
            {
                return null;
            }

            return new SqlServerConnectionOptions
            {
                ServerAddress = options.ServerAddress,
                Port = options.Port,
                Username = options.Username,
                Password = options.Password
            };
        }

        private async Task LoadSqlConnectionHistoryAsync()
        {
            var history = await SqlConnectionHistoryService.LoadAsync();
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                _isApplyingSqlHistory = true;
                try
                {
                    ReplaceItems(SqlServerAddressHistory, history.ServerAddresses);
                    ReplaceItems(SqlUsernameHistory, history.Usernames);
                    ReplaceItems(SqlPasswordHistory, history.Passwords);

                    if (!_hasUserModifiedSqlConnectionInputs)
                    {
                        SqlServerAddress = history.LastServerAddress;
                        SqlPort = string.IsNullOrWhiteSpace(history.LastPort) ? "1433" : history.LastPort;
                        SqlUsername = history.LastUsername;
                        SqlPassword = history.LastPassword;
                    }
                }
                finally
                {
                    _isApplyingSqlHistory = false;
                }
            });
        }

        private async Task SaveSqlConnectionHistoryAsync(SqlServerConnectionOptions options)
        {
            await SqlConnectionHistoryService.SaveAsync(options);
            var history = await SqlConnectionHistoryService.LoadAsync();
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                _isApplyingSqlHistory = true;
                try
                {
                    ReplaceItems(SqlServerAddressHistory, history.ServerAddresses);
                    ReplaceItems(SqlUsernameHistory, history.Usernames);
                    ReplaceItems(SqlPasswordHistory, history.Passwords);
                }
                finally
                {
                    _isApplyingSqlHistory = false;
                }
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

        private void CancelPendingTableLoad(bool clearBusy = true)
        {
            if (_loadTablesCancellationTokenSource == null)
            {
                return;
            }

            _loadTablesCancellationTokenSource.Cancel();
            _loadTablesCancellationTokenSource = null;
            if (clearBusy)
            {
                IsSqlBusy = false;
            }
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
            _screenshotEditorWindow?.Close();
            _screenshotEditorWindow = null;
            Application.Current.Shutdown();
        }

        private void RefreshSystemStatus()
        {
            IsDefenderEnabled = WindowsSecurityService.GetDefenderRealtimeStatus();
            IsAutoUpdateEnabled = WindowsSecurityService.GetAutoUpdateStatus();
        }

        private async Task ToggleDefenderAsync()
        {
            bool target = !IsDefenderEnabled;
            SystemStatusMessage = target ? "姝ｅ湪鎭㈠瀹炴椂闃叉姢锛岃鍦?UAC 寮圭獥涓‘璁?.." : "姝ｅ湪鍏抽棴瀹炴椂闃叉姢锛岃鍦?UAC 寮圭獥涓‘璁?..";
            try
            {
                await WindowsSecurityService.SetDefenderRealtimeAsync(target);
                await Task.Delay(1500);
                RefreshSystemStatus();
                SystemStatusMessage = target ? "实时防护已恢复。" : "实时防护已关闭。";
            }
            catch (OperationCanceledException)
            {
                SystemStatusMessage = "鎿嶄綔宸插彇娑堬紙UAC 鏈巿鏉冿級銆?";
            }
            catch (Exception ex)
            {
                SystemStatusMessage = "操作失败：" + ex.Message;
                MessageBox.Show(ex.Message, "鎿嶄綔澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ToggleAutoUpdateAsync()
        {
            bool target = !IsAutoUpdateEnabled;
            SystemStatusMessage = target ? "姝ｅ湪鎭㈠鑷姩鏇存柊锛岃鍦?UAC 寮圭獥涓‘璁?.." : "姝ｅ湪鍋滄鑷姩鏇存柊锛岃鍦?UAC 寮圭獥涓‘璁?..";
            try
            {
                await WindowsSecurityService.SetAutoUpdateAsync(target);
                await Task.Delay(1500);
                RefreshSystemStatus();
                SystemStatusMessage = target ? "自动更新已恢复。" : "自动更新已停止。";
            }
            catch (OperationCanceledException)
            {
                SystemStatusMessage = "鎿嶄綔宸插彇娑堬紙UAC 鏈巿鏉冿級銆?";
            }
            catch (Exception ex)
            {
                SystemStatusMessage = "操作失败：" + ex.Message;
                MessageBox.Show(ex.Message, "鎿嶄綔澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task TriggerUpdateNowAsync()
        {
            SystemStatusMessage = "姝ｅ湪瑙﹀彂绔嬪嵆鏇存柊锛岃鍦?UAC 寮圭獥涓‘璁?..";
            try
            {
                await WindowsSecurityService.TriggerImmediateUpdateAsync();
                RefreshSystemStatus();
                SystemStatusMessage = "鏇存柊浠诲姟宸蹭笅鍙戯紝Windows Update 璁剧疆椤靛凡鎵撳紑锛屽彲鍦ㄥ叾涓煡鐪嬭繘搴︺€?";
            }
            catch (OperationCanceledException)
            {
                SystemStatusMessage = "鎿嶄綔宸插彇娑堬紙UAC 鏈巿鏉冿級銆?";
            }
            catch (Exception ex)
            {
                SystemStatusMessage = "鎿嶄綔澶辫触锛? + ex.Message";
                MessageBox.Show(ex.Message, "鎿嶄綔澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
            }
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
                MessageBox.Show("鎵ц澶辫触: " + ex.Message, "閿欒", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool CanExecuteSqlQuery()
            => !IsQueryBusy && SelectedSqlDatabase != null && !string.IsNullOrWhiteSpace(SqlQueryText);

        private bool CanExportQueryResult()
            => !IsQueryBusy && SqlQueryResult != null && SqlQueryResult.Count > 0;

        private async Task ExecuteSqlQueryAsync()
        {
            IsQueryBusy = true;
            QueryStatusMessage = "姝ｅ湪鎵ц鏌ヨ...";
            SqlQueryResult = null;
            try
            {
                var options = GetEffectiveSqlConnectionOptions();
                var table = await SqlExportService.ExecuteQueryAsync(
                    options,
                    SelectedSqlDatabase.Name,
                    SqlQueryText,
                    CancellationToken.None);

                SqlQueryResult = table.DefaultView;
                QueryStatusMessage = $"共 {table.Rows.Count} 行，{table.Columns.Count} 列。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "SQL query execution failed");
                QueryStatusMessage = "执行失败：" + ex.Message;
                MessageBox.Show(ex.Message, "SQL 鎵ц澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsQueryBusy = false;
            }
        }

        private async Task ExportQueryResultAsync()
        {
            if (SqlQueryResult == null) return;
            var dialog = new SaveFileDialog
            {
                Filter = "Excel 宸ヤ綔绨?(*.xlsx)|*.xlsx",
                FileName = $"QueryResult_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                DefaultExt = ".xlsx",
                AddExtension = true,
                OverwritePrompt = true
            };
            if (dialog.ShowDialog() != true) return;

            IsQueryBusy = true;
            QueryStatusMessage = "姝ｅ湪瀵煎嚭 Excel...";
            try
            {
                var table = SqlQueryResult.Table;
                var result = await SqlExportService.ExportDataTableAsync(
                    table,
                    "QueryResult",
                    dialog.FileName,
                    CancellationToken.None);

                QueryStatusMessage = $"瀵煎嚭瀹屾垚锛屽叡 {result.RowCount} 琛屻€?";
                MessageBox.Show(
                    $"导出成功。\n文件路径：{result.FilePath}",
                    "导出完成",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Query result export failed");
                QueryStatusMessage = "瀵煎嚭澶辫触锛? + ex.Message";
                MessageBox.Show(ex.Message, "瀵煎嚭澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsQueryBusy = false;
            }
        }

        private void TriggerCommandRequery()
        {
            _testSqlConnectionCommand?.RaiseCanExecuteChanged();
            _exportSqlTableCommand?.RaiseCanExecuteChanged();
            _executeQueryCommand?.RaiseCanExecuteChanged();
            _exportQueryResultCommand?.RaiseCanExecuteChanged();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class CodexProfileItem : INotifyPropertyChanged
    {
        private string _remark;
        private bool _isApplying;
        private string _statusMessage;
        private string _configTomlContentProtected;
        private string _authJsonContentProtected;

        public string Name { get; set; }
        public string FolderPath { get; set; }

        public string Remark
        {
            get => _remark;
            set
            {
                if (string.Equals(_remark, value, StringComparison.Ordinal))
                {
                    return;
                }

                _remark = value;
                OnPropertyChanged();
            }
        }

        public bool IsApplying
        {
            get => _isApplying;
            set
            {
                if (_isApplying == value)
                {
                    return;
                }

                _isApplying = value;
                OnPropertyChanged();
            }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (string.Equals(_statusMessage, value, StringComparison.Ordinal))
                {
                    return;
                }

                _statusMessage = value;
                OnPropertyChanged();
            }
        }

        public string ConfigTomlContentProtected
        {
            get => _configTomlContentProtected;
            set
            {
                if (string.Equals(_configTomlContentProtected, value, StringComparison.Ordinal))
                {
                    return;
                }

                _configTomlContentProtected = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasEmbeddedContent));
                OnPropertyChanged(nameof(ContentStorageSummary));
            }
        }

        public string AuthJsonContentProtected
        {
            get => _authJsonContentProtected;
            set
            {
                if (string.Equals(_authJsonContentProtected, value, StringComparison.Ordinal))
                {
                    return;
                }

                _authJsonContentProtected = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasEmbeddedContent));
                OnPropertyChanged(nameof(ContentStorageSummary));
            }
        }

        public bool HasEmbeddedContent =>
            !string.IsNullOrWhiteSpace(ConfigTomlContentProtected)
            && !string.IsNullOrWhiteSpace(AuthJsonContentProtected);

        public string ContentStorageSummary =>
            HasEmbeddedContent
                ? "宸插唴缃繚瀛?config.toml 鍜?auth.json 鍐呭"
                : "鏈唴缃厤缃唴瀹癸紙寤鸿閲嶆柊鎷栧叆鏂囦欢澶癸級";

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
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Command execution failed.");
                MessageBox.Show(ex.Message, "鎿嶄綔澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
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

    public class AsyncRelayParameterCommand : ICommand
    {
        private readonly Func<object, Task> _executeAsync;
        private readonly Func<object, bool> _canExecute;
        private bool _isExecuting;

        public AsyncRelayParameterCommand(Func<object, Task> executeAsync, Func<object, bool> canExecute = null)
        {
            _executeAsync = executeAsync;
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return !_isExecuting && (_canExecute == null || _canExecute(parameter));
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
                await _executeAsync(parameter);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Parameterized command execution failed.");
                MessageBox.Show(ex.Message, "鎿嶄綔澶辫触", MessageBoxButton.OK, MessageBoxImage.Error);
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




