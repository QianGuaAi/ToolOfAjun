using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Threading;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using MyTools.Services;
using WinForms = System.Windows.Forms;

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
        private string _wgStatusText = "未连接";
        private string _wgEndpoint;
        private string _wgAddress;
        private string _wgServerPublicKey;
        private bool _isWgSettingsOpen;
        private string _currentModule;
        private string _sqlServerAddress;
        private string _sqlPort = "1433";
        private SqlProviderKind _selectedSqlProvider = SqlProviderKind.SqlServer;
        private string _sqlUsername;
        private string _sqlPassword;
        private string _sqlTableSearchText;
        private DatabaseItem _selectedSqlDatabase;
        private TableItem _selectedSqlTable;
        private ICollectionView _filteredSqlTableView;
        private string _sqlStatusMessage = "请输入 SQL Server 连接信息后测试连接。";
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
        private ObservableCollection<OptimizationReportItem> _optimizationReports;
        private ObservableCollection<JunkCandidate> _junkCandidates;
        private ObservableCollection<WeChatCleanupCandidate> _weChatCleanupCandidates;
        private ObservableCollection<WeChatRoot> _weChatRoots;
        private WeChatRoot _selectedWeChatRoot;
        private bool _isAutoOptimizeBusy;
        private string _autoOptimizeStatusMessage = "点击“开始自动优化”执行白名单优化流程。";
        private bool _isJunkBusy;
        private string _junkStatusMessage = "点击“开始扫描”分析可清理项目。";
        private bool _isWeChatCleanupBusy;
        private string _weChatCleanupStatusMessage = "请选择日期和类别后先扫描，再执行清理。";
        private bool _isWeChatBackupBusy;
        private string _weChatBackupStatusMessage = "请选择日期和类别，设置输出目录后开始备份。";
        private bool _isWeChatRestoreBusy;
        private string _weChatRestoreStatusMessage = "请选择备份文件并确认恢复路径。";
        private DateTime? _weChatCleanupStartDate = DateTime.Today.AddDays(-30);
        private DateTime? _weChatCleanupEndDate = DateTime.Today;
        private DateTime? _weChatBackupStartDate = DateTime.Today.AddDays(-30);
        private DateTime? _weChatBackupEndDate = DateTime.Today;
        private bool _weChatCleanupIncludeText;
        private bool _weChatCleanupIncludeImage = true;
        private bool _weChatCleanupIncludeVideo = true;
        private bool _weChatCleanupIncludeVoice = true;
        private bool _weChatCleanupIncludeFile = true;
        private bool _weChatCleanupIncludeCache = true;
        private bool _weChatBackupIncludeText;
        private bool _weChatBackupIncludeImage = true;
        private bool _weChatBackupIncludeVideo = true;
        private bool _weChatBackupIncludeVoice = true;
        private bool _weChatBackupIncludeFile = true;
        private bool _weChatBackupIncludeCache;
        private string _weChatBackupOutputFolder = string.Empty;
        private string _weChatRestoreZipPath = string.Empty;
        private string _weChatRestoreManifestSummary = "未选择备份文件。";
        private bool _restoreToOriginal = true;
        private string _weChatRestoreTargetRoot = string.Empty;

        private bool _showEditorAfterCapture = true;
        private string _screenshotHotkeyText = "Ctrl+Shift+Z";
        private bool _isCapturingHotkey;
        private bool _isScreenshotBusy;
        private ScreenshotEditorWindow _screenshotEditorWindow;
        private RecordRegionWindow _recordRegionWindow;
        private readonly RecordingService _recordingService = new RecordingService();
        private bool _isVideoRecording;
        private bool _isAudioRecording;
        private DateTime _audioRecordingStartedAt;
        private string _audioRecordingIndicator = string.Empty;
        private string _activeVideoOutputPath = string.Empty;
        private string _activeAudioOutputPath = string.Empty;
        private readonly DispatcherTimer _audioRecordingTimer;
        private uint _pendingModifiers = 0x0006;
        private uint _pendingKey = 0x5A;
        private string _codexProfilesStatusMessage = "拖入包含 config.toml 和 auth.json 的文件夹，生成可应用的 Codex 配置记录。";

        private readonly AsyncRelayCommand _executeQueryCommand;
        private readonly AsyncRelayCommand _exportQueryResultCommand;
        private readonly AsyncRelayParameterCommand _applyCodexProfileCommand;
        private readonly AsyncRelayParameterCommand _exportCodexProfileCommand;
        private readonly AsyncRelayCommand _importCodexProfileCommand;
        private readonly AsyncRelayCommand _openRecordRegionCommand;
        private readonly AsyncRelayCommand _toggleAudioRecordingCommand;

        private readonly AsyncRelayCommand _testSqlConnectionCommand;
        private readonly AsyncRelayCommand _exportSqlTableCommand;
        private readonly AsyncRelayCommand _toggleDefenderCommand;
        private readonly AsyncRelayCommand _toggleAutoUpdateCommand;
        private readonly AsyncRelayCommand _triggerUpdateNowCommand;
        private readonly AsyncRelayCommand _startAutoOptimizeCommand;
        private readonly AsyncRelayCommand _startJunkScanCommand;
        private readonly AsyncRelayCommand _runJunkCleanupCommand;
        private readonly AsyncRelayCommand _scanWeChatCleanupCommand;
        private readonly AsyncRelayCommand _startWeChatCleanupCommand;
        private readonly AsyncRelayCommand _startWeChatBackupCommand;
        private readonly AsyncRelayCommand _startWeChatRestoreCommand;
        private readonly RelayCommand _deleteSelectedReportsCommand;
        private readonly RelayCommand _selectWeChatBackupOutputFolderCommand;
        private readonly RelayCommand _selectWeChatRestoreZipCommand;
        private readonly RelayCommand _selectWeChatRestoreTargetRootCommand;
        private readonly RelayParameterCommand _showReportDetailsCommand;

        private readonly OptimizationReportService _optimizationReportService = new OptimizationReportService();
        private readonly SystemOptimizationService _systemOptimizationService = new SystemOptimizationService();
        private readonly JunkCleanupService _junkCleanupService = new JunkCleanupService();
        private readonly WeChatDataLocator _weChatDataLocator = new WeChatDataLocator();
        private readonly WeChatCleanupService _weChatCleanupService = new WeChatCleanupService();
        private readonly WeChatBackupService _weChatBackupService = new WeChatBackupService();
        private static readonly IReadOnlyList<SqlProviderOption> SqlProviderOptionItems = new List<SqlProviderOption>
        {
            new SqlProviderOption(SqlProviderKind.SqlServer, "SQL Server"),
            new SqlProviderOption(SqlProviderKind.PostgreSql, "PostgreSQL"),
            new SqlProviderOption(SqlProviderKind.MySql, "MySQL")
        };

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
            OptimizationReports = new ObservableCollection<OptimizationReportItem>();
            JunkCandidates = new ObservableCollection<JunkCandidate>();
            WeChatCleanupCandidates = new ObservableCollection<WeChatCleanupCandidate>();
            WeChatRoots = new ObservableCollection<WeChatRoot>();
            FilteredSqlTableView = CollectionViewSource.GetDefaultView(SqlTableList);
            FilteredSqlTableView.Filter = FilterSqlTable;
            _audioRecordingTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _audioRecordingTimer.Tick += (sender, args) => UpdateAudioRecordingIndicator();

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
                    $"确定要永久删除启动项 \"{item.Name}\" 吗？\n此操作不可恢复。",
                    "确认删除",
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
            _startAutoOptimizeCommand = new AsyncRelayCommand(StartAutoOptimizeAsync, () => !IsAutoOptimizeBusy);
            _startJunkScanCommand = new AsyncRelayCommand(StartJunkScanAsync, () => !IsJunkBusy);
            _runJunkCleanupCommand = new AsyncRelayCommand(RunJunkCleanupAsync, CanRunJunkCleanup);
            _scanWeChatCleanupCommand = new AsyncRelayCommand(ScanWeChatCleanupAsync, () => !IsWeChatCleanupBusy);
            _startWeChatCleanupCommand = new AsyncRelayCommand(StartWeChatCleanupAsync, CanStartWeChatCleanup);
            _startWeChatBackupCommand = new AsyncRelayCommand(StartWeChatBackupAsync, CanStartWeChatBackup);
            _startWeChatRestoreCommand = new AsyncRelayCommand(StartWeChatRestoreAsync, CanStartWeChatRestore);
            _deleteSelectedReportsCommand = new RelayCommand(DeleteSelectedReports, CanDeleteSelectedReports);
            _selectWeChatBackupOutputFolderCommand = new RelayCommand(SelectWeChatBackupOutputFolder, () => !IsWeChatBackupBusy);
            _selectWeChatRestoreZipCommand = new RelayCommand(SelectWeChatRestoreZip, () => !IsWeChatRestoreBusy);
            _selectWeChatRestoreTargetRootCommand = new RelayCommand(SelectWeChatRestoreTargetRoot, () => !IsWeChatRestoreBusy);
            _showReportDetailsCommand = new RelayParameterCommand(ShowReportDetails);
            StartAutoOptimizeCommand = _startAutoOptimizeCommand;
            StartJunkScanCommand = _startJunkScanCommand;
            RunJunkCleanupCommand = _runJunkCleanupCommand;
            ScanWeChatCleanupCommand = _scanWeChatCleanupCommand;
            StartWeChatCleanupCommand = _startWeChatCleanupCommand;
            StartWeChatBackupCommand = _startWeChatBackupCommand;
            StartWeChatRestoreCommand = _startWeChatRestoreCommand;
            DeleteSelectedReportsCommand = _deleteSelectedReportsCommand;
            SelectWeChatBackupOutputFolderCommand = _selectWeChatBackupOutputFolderCommand;
            SelectWeChatRestoreZipCommand = _selectWeChatRestoreZipCommand;
            SelectWeChatRestoreTargetRootCommand = _selectWeChatRestoreTargetRootCommand;
            ShowReportDetailsCommand = _showReportDetailsCommand;

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
            _openRecordRegionCommand = new AsyncRelayCommand(OpenRecordRegionAsync, () => !IsVideoRecording && !IsAudioRecording);
            StartVideoRecordingCommand = _openRecordRegionCommand;
            _toggleAudioRecordingCommand = new AsyncRelayCommand(ToggleAudioRecordingAsync, () => !IsVideoRecording);
            ToggleAudioRecordingCommand = _toggleAudioRecordingCommand;
            StartCaptureHotkeyCommand = new RelayCommand(() => IsCapturingHotkey = true);
            CancelCaptureHotkeyCommand = new RelayCommand(() => IsCapturingHotkey = false);
            SaveScreenshotSettingsCommand = new AsyncRelayCommand(SaveScreenshotSettingsAsync);
            EditClipboardImageCommand = new RelayCommand(EditClipboardImage);

            _applyCodexProfileCommand = new AsyncRelayParameterCommand(ApplyCodexProfileAsync);
            ApplyCodexProfileCommand = _applyCodexProfileCommand;
            _exportCodexProfileCommand = new AsyncRelayParameterCommand(ExportCodexProfileAsync);
            ExportCodexProfileCommand = _exportCodexProfileCommand;
            _importCodexProfileCommand = new AsyncRelayCommand(ImportCodexProfileAsync);
            ImportCodexProfileCommand = _importCodexProfileCommand;
            DeleteCodexProfileCommand = new RelayParameterCommand(DeleteCodexProfile);

            CurrentModule = "Home";
            _ = LoadSqlConnectionHistoryAsync();
            _ = LoadScreenshotSettingsAsync();
            _ = LoadCodexProfilesAsync();
            _ = LoadOptimizationReportsAsync();
            _ = LoadWeChatRootsAsync();
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

        public ObservableCollection<OptimizationReportItem> OptimizationReports
        {
            get => _optimizationReports;
            set { _optimizationReports = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasSelectedOptimizationReports)); }
        }

        public ObservableCollection<JunkCandidate> JunkCandidates
        {
            get => _junkCandidates;
            set { _junkCandidates = value; OnPropertyChanged(); }
        }

        public ObservableCollection<WeChatCleanupCandidate> WeChatCleanupCandidates
        {
            get => _weChatCleanupCandidates;
            set { _weChatCleanupCandidates = value; OnPropertyChanged(); }
        }

        public ObservableCollection<WeChatRoot> WeChatRoots
        {
            get => _weChatRoots;
            set { _weChatRoots = value; OnPropertyChanged(); }
        }

        public WeChatRoot SelectedWeChatRoot
        {
            get => _selectedWeChatRoot;
            set
            {
                if (ReferenceEquals(_selectedWeChatRoot, value))
                {
                    return;
                }

                _selectedWeChatRoot = value;
                OnPropertyChanged();
                TriggerCommandRequery();
            }
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

        public string AppVersionText => BuildAppVersionText();

        public IReadOnlyList<SqlProviderOption> SqlProviderOptions => SqlProviderOptionItems;

        public SqlProviderKind SelectedSqlProvider
        {
            get => _selectedSqlProvider;
            set
            {
                if (_selectedSqlProvider == value)
                {
                    return;
                }

                _selectedSqlProvider = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(SqlPortHint));
                ApplyDefaultSqlPortIfNeeded();
                MarkSqlConnectionInputAsUserModified();
                _ = LoadSqlConnectionHistoryAsync();
            }
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

        public string SqlPortHint
        {
            get
            {
                switch (SelectedSqlProvider)
                {
                    case SqlProviderKind.PostgreSql:
                        return "端口（默认 5432）";
                    case SqlProviderKind.MySql:
                        return "端口（默认 3306）";
                    default:
                        return "端口（默认 1433）";
                }
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

        public bool IsAutoOptimizeBusy
        {
            get => _isAutoOptimizeBusy;
            set
            {
                if (_isAutoOptimizeBusy == value)
                {
                    return;
                }

                _isAutoOptimizeBusy = value;
                OnPropertyChanged();
                TriggerCommandRequery();
            }
        }

        public string AutoOptimizeStatusMessage
        {
            get => _autoOptimizeStatusMessage;
            set { _autoOptimizeStatusMessage = value; OnPropertyChanged(); }
        }

        public bool IsJunkBusy
        {
            get => _isJunkBusy;
            set
            {
                if (_isJunkBusy == value)
                {
                    return;
                }

                _isJunkBusy = value;
                OnPropertyChanged();
                TriggerCommandRequery();
            }
        }

        public string JunkStatusMessage
        {
            get => _junkStatusMessage;
            set { _junkStatusMessage = value; OnPropertyChanged(); }
        }

        public bool IsWeChatCleanupBusy
        {
            get => _isWeChatCleanupBusy;
            set
            {
                if (_isWeChatCleanupBusy == value)
                {
                    return;
                }

                _isWeChatCleanupBusy = value;
                OnPropertyChanged();
                TriggerCommandRequery();
            }
        }

        public string WeChatCleanupStatusMessage
        {
            get => _weChatCleanupStatusMessage;
            set { _weChatCleanupStatusMessage = value; OnPropertyChanged(); }
        }

        public bool IsWeChatBackupBusy
        {
            get => _isWeChatBackupBusy;
            set
            {
                if (_isWeChatBackupBusy == value)
                {
                    return;
                }

                _isWeChatBackupBusy = value;
                OnPropertyChanged();
                TriggerCommandRequery();
            }
        }

        public string WeChatBackupStatusMessage
        {
            get => _weChatBackupStatusMessage;
            set { _weChatBackupStatusMessage = value; OnPropertyChanged(); }
        }

        public bool IsWeChatRestoreBusy
        {
            get => _isWeChatRestoreBusy;
            set
            {
                if (_isWeChatRestoreBusy == value)
                {
                    return;
                }

                _isWeChatRestoreBusy = value;
                OnPropertyChanged();
                TriggerCommandRequery();
            }
        }

        public string WeChatRestoreStatusMessage
        {
            get => _weChatRestoreStatusMessage;
            set { _weChatRestoreStatusMessage = value; OnPropertyChanged(); }
        }

        public DateTime? WeChatCleanupStartDate
        {
            get => _weChatCleanupStartDate;
            set { _weChatCleanupStartDate = value; OnPropertyChanged(); }
        }

        public DateTime? WeChatCleanupEndDate
        {
            get => _weChatCleanupEndDate;
            set { _weChatCleanupEndDate = value; OnPropertyChanged(); }
        }

        public DateTime? WeChatBackupStartDate
        {
            get => _weChatBackupStartDate;
            set { _weChatBackupStartDate = value; OnPropertyChanged(); }
        }

        public DateTime? WeChatBackupEndDate
        {
            get => _weChatBackupEndDate;
            set { _weChatBackupEndDate = value; OnPropertyChanged(); }
        }

        public bool WeChatCleanupIncludeText
        {
            get => _weChatCleanupIncludeText;
            set { _weChatCleanupIncludeText = value; OnPropertyChanged(); }
        }

        public bool WeChatCleanupIncludeImage
        {
            get => _weChatCleanupIncludeImage;
            set { _weChatCleanupIncludeImage = value; OnPropertyChanged(); }
        }

        public bool WeChatCleanupIncludeVideo
        {
            get => _weChatCleanupIncludeVideo;
            set { _weChatCleanupIncludeVideo = value; OnPropertyChanged(); }
        }

        public bool WeChatCleanupIncludeVoice
        {
            get => _weChatCleanupIncludeVoice;
            set { _weChatCleanupIncludeVoice = value; OnPropertyChanged(); }
        }

        public bool WeChatCleanupIncludeFile
        {
            get => _weChatCleanupIncludeFile;
            set { _weChatCleanupIncludeFile = value; OnPropertyChanged(); }
        }

        public bool WeChatCleanupIncludeCache
        {
            get => _weChatCleanupIncludeCache;
            set { _weChatCleanupIncludeCache = value; OnPropertyChanged(); }
        }

        public bool WeChatBackupIncludeText
        {
            get => _weChatBackupIncludeText;
            set { _weChatBackupIncludeText = value; OnPropertyChanged(); }
        }

        public bool WeChatBackupIncludeImage
        {
            get => _weChatBackupIncludeImage;
            set { _weChatBackupIncludeImage = value; OnPropertyChanged(); }
        }

        public bool WeChatBackupIncludeVideo
        {
            get => _weChatBackupIncludeVideo;
            set { _weChatBackupIncludeVideo = value; OnPropertyChanged(); }
        }

        public bool WeChatBackupIncludeVoice
        {
            get => _weChatBackupIncludeVoice;
            set { _weChatBackupIncludeVoice = value; OnPropertyChanged(); }
        }

        public bool WeChatBackupIncludeFile
        {
            get => _weChatBackupIncludeFile;
            set { _weChatBackupIncludeFile = value; OnPropertyChanged(); }
        }

        public bool WeChatBackupIncludeCache
        {
            get => _weChatBackupIncludeCache;
            set { _weChatBackupIncludeCache = value; OnPropertyChanged(); }
        }

        public string WeChatBackupOutputFolder
        {
            get => _weChatBackupOutputFolder;
            set { _weChatBackupOutputFolder = value; OnPropertyChanged(); TriggerCommandRequery(); }
        }

        public string WeChatRestoreZipPath
        {
            get => _weChatRestoreZipPath;
            set { _weChatRestoreZipPath = value; OnPropertyChanged(); TriggerCommandRequery(); }
        }

        public string WeChatRestoreManifestSummary
        {
            get => _weChatRestoreManifestSummary;
            set { _weChatRestoreManifestSummary = value; OnPropertyChanged(); }
        }

        public bool RestoreToOriginal
        {
            get => _restoreToOriginal;
            set
            {
                if (_restoreToOriginal == value)
                {
                    return;
                }

                _restoreToOriginal = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(RestoreToCustom));
            }
        }

        public bool RestoreToCustom
        {
            get => !_restoreToOriginal;
            set
            {
                if (value == !_restoreToOriginal)
                {
                    return;
                }

                _restoreToOriginal = !value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(RestoreToOriginal));
            }
        }

        public string WeChatRestoreTargetRoot
        {
            get => _weChatRestoreTargetRoot;
            set { _weChatRestoreTargetRoot = value; OnPropertyChanged(); TriggerCommandRequery(); }
        }

        public bool HasSelectedOptimizationReports => OptimizationReports.Any(x => x.IsSelected);

        public bool HasJunkCandidates => JunkCandidates.Any();

        public bool HasWeChatCleanupCandidates => WeChatCleanupCandidates.Any();

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
        public ICommand StartVideoRecordingCommand { get; }
        public ICommand ToggleAudioRecordingCommand { get; }
        public ICommand StartCaptureHotkeyCommand { get; }
        public ICommand CancelCaptureHotkeyCommand { get; }
        public ICommand SaveScreenshotSettingsCommand { get; }
        public ICommand EditClipboardImageCommand { get; }
        public ICommand ApplyCodexProfileCommand { get; }
        public ICommand ExportCodexProfileCommand { get; }
        public ICommand ImportCodexProfileCommand { get; }
        public ICommand DeleteCodexProfileCommand { get; }

        public bool IsVideoRecording
        {
            get => _isVideoRecording;
            private set
            {
                if (_isVideoRecording == value)
                {
                    return;
                }

                _isVideoRecording = value;
                OnPropertyChanged();
                TriggerCommandRequery();
            }
        }

        public bool IsAudioRecording
        {
            get => _isAudioRecording;
            private set
            {
                if (_isAudioRecording == value)
                {
                    return;
                }

                _isAudioRecording = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasAudioRecordingIndicator));
                TriggerCommandRequery();
            }
        }

        public string AudioRecordingIndicator
        {
            get => _audioRecordingIndicator;
            private set
            {
                if (string.Equals(_audioRecordingIndicator, value, StringComparison.Ordinal))
                {
                    return;
                }

                _audioRecordingIndicator = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasAudioRecordingIndicator));
            }
        }

        public bool HasAudioRecordingIndicator => !string.IsNullOrWhiteSpace(AudioRecordingIndicator);

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
        public ICommand StartAutoOptimizeCommand { get; }
        public ICommand StartJunkScanCommand { get; }
        public ICommand RunJunkCleanupCommand { get; }
        public ICommand ScanWeChatCleanupCommand { get; }
        public ICommand StartWeChatCleanupCommand { get; }
        public ICommand StartWeChatBackupCommand { get; }
        public ICommand StartWeChatRestoreCommand { get; }
        public ICommand DeleteSelectedReportsCommand { get; }
        public ICommand SelectWeChatBackupOutputFolderCommand { get; }
        public ICommand SelectWeChatRestoreZipCommand { get; }
        public ICommand SelectWeChatRestoreTargetRootCommand { get; }
        public ICommand ShowReportDetailsCommand { get; }

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
                CodexProfilesStatusMessage = "读取 Codex 配置记录失败。";
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

        private async Task ExportCodexProfileAsync(object parameter)
        {
            if (!(parameter is CodexProfileItem item))
            {
                return;
            }

            try
            {
                if (string.IsNullOrWhiteSpace(item.ConfigTomlContentProtected) || string.IsNullOrWhiteSpace(item.AuthJsonContentProtected))
                {
                    MessageBox.Show(
                        "该记录暂无可导出的内容，请先确保已成功添加配置。",
                        "导出 Codex 配置",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                using (var dialog = new WinForms.FolderBrowserDialog())
                {
                    dialog.Description = "请选择导出的目标文件夹";
                    dialog.ShowNewFolderButton = true;
                    if (dialog.ShowDialog() != WinForms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedPath))
                    {
                        return;
                    }

                    var parentFolder = dialog.SelectedPath;
                    var safeName = BuildSafeCodexFolderName(item.Name);
                    var targetFolder = BuildUniqueChildFolder(parentFolder, safeName);
                    Directory.CreateDirectory(targetFolder);

                    var configTomlBytes = CodexConfigProfileService.UnprotectBytesFromBase64(item.ConfigTomlContentProtected);
                    var authJsonBytes = CodexConfigProfileService.UnprotectBytesFromBase64(item.AuthJsonContentProtected);
                    if (configTomlBytes == null || authJsonBytes == null)
                    {
                        throw new InvalidOperationException("该记录暂无可导出的内容，请先确保已成功添加配置。");
                    }

                    var configPath = Path.Combine(targetFolder, CodexConfigProfileService.ConfigFileName);
                    var authPath = Path.Combine(targetFolder, CodexConfigProfileService.AuthFileName);
                    await WriteAllBytesAsync(configPath, configTomlBytes, CancellationToken.None);
                    await WriteAllBytesAsync(authPath, authJsonBytes, CancellationToken.None);

                    CodexProfilesStatusMessage = $"已导出「{item.Name}」到 {targetFolder}。";
                    AppLogService.Information("Exported Codex profile {Name} to {Folder}", item.Name ?? string.Empty, targetFolder);
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Exporting Codex profile failed for {Name}", item.Name ?? string.Empty);
                MessageBox.Show(ex.Message, "导出失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ImportCodexProfileAsync()
        {
            try
            {
                var dialog = new OpenFileDialog
                {
                    Multiselect = true,
                    Filter = "Codex 配置文件|config.toml;auth.json|所有文件 (*.*)|*.*",
                    Title = "请选择 config.toml 和 auth.json"
                };

                if (dialog.ShowDialog() != true || dialog.FileNames == null || dialog.FileNames.Length == 0)
                {
                    return;
                }

                var configFilePath = dialog.FileNames.FirstOrDefault(path =>
                    string.Equals(Path.GetFileName(path), CodexConfigProfileService.ConfigFileName, StringComparison.OrdinalIgnoreCase));
                var authFilePath = dialog.FileNames.FirstOrDefault(path =>
                    string.Equals(Path.GetFileName(path), CodexConfigProfileService.AuthFileName, StringComparison.OrdinalIgnoreCase));

                if (string.IsNullOrWhiteSpace(configFilePath) || string.IsNullOrWhiteSpace(authFilePath))
                {
                    MessageBox.Show("请同时选择 config.toml 和 auth.json 两个文件。", "导入 Codex 配置", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var configFolderPath = NormalizeFolderPath(Path.GetDirectoryName(configFilePath));
                var authFolderPath = NormalizeFolderPath(Path.GetDirectoryName(authFilePath));
                if (!string.Equals(configFolderPath, authFolderPath, StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("config.toml 和 auth.json 必须来自同一个文件夹。", "导入 Codex 配置", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var inputName = Interaction.InputBox("请输入新记录名称", "导入 Codex 配置", string.Empty);
                var fallbackName = Path.GetFileName(configFolderPath);
                var name = string.IsNullOrWhiteSpace(inputName)
                    ? fallbackName
                    : inputName.Trim();
                if (string.IsNullOrWhiteSpace(name))
                {
                    name = $"codex-profile-{DateTime.Now:yyyyMMddHHmmss}";
                }

                var configBytes = await ReadAllBytesAsync(configFilePath, CancellationToken.None);
                var authBytes = await ReadAllBytesAsync(authFilePath, CancellationToken.None);
                var item = CreateCodexProfileItem(
                    configFolderPath,
                    name,
                    name,
                    CodexConfigProfileService.ProtectBytesToBase64(configBytes),
                    CodexConfigProfileService.ProtectBytesToBase64(authBytes));
                item.StatusMessage = "已导入配置内容。";
                AddCodexProfileItem(item);

                await SaveCodexProfilesAsync();
                CodexProfilesStatusMessage = $"已导入「{item.Name}」。";
                AppLogService.Information("Imported Codex profile {Name} from {Folder}", item.Name ?? string.Empty, configFolderPath);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Importing Codex profile failed.");
                MessageBox.Show(ex.Message, "导入失败", MessageBoxButton.OK, MessageBoxImage.Error);
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

            return "配置记录";
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

        private static string BuildSafeCodexFolderName(string name)
        {
            var safeName = string.IsNullOrWhiteSpace(name) ? "codex-profile" : name.Trim();
            foreach (var invalidFileNameChar in Path.GetInvalidFileNameChars())
            {
                safeName = safeName.Replace(invalidFileNameChar, '_');
            }

            return string.IsNullOrWhiteSpace(safeName) ? "codex-profile" : safeName;
        }

        private static string BuildUniqueChildFolder(string parentFolder, string safeName)
        {
            var baseName = string.IsNullOrWhiteSpace(safeName) ? "codex-profile" : safeName;
            var index = 0;
            while (true)
            {
                var folderName = index == 0 ? baseName : $"{baseName}_{index}";
                var fullPath = Path.Combine(parentFolder, folderName);
                if (!Directory.Exists(fullPath))
                {
                    return fullPath;
                }

                index++;
            }
        }

        private static async Task<byte[]> ReadAllBytesAsync(string filePath, CancellationToken cancellationToken)
        {
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, useAsync: true))
            using (var memoryStream = new MemoryStream())
            {
                await stream.CopyToAsync(memoryStream, 81920, cancellationToken);
                return memoryStream.ToArray();
            }
        }

        private static async Task WriteAllBytesAsync(string filePath, byte[] bytes, CancellationToken cancellationToken)
        {
            using (var stream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true))
            {
                await stream.WriteAsync(bytes, 0, bytes.Length, cancellationToken);
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
            SqlStatusMessage = "连接信息已变化，请重新测试连接。";
        }

        private void EditClipboardImage()
        {
            var image = System.Windows.Clipboard.GetImage();
            if (image == null)
            {
                MessageBox.Show("剪贴板中没有图片", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
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
                AppLogService.Information("Screenshot capture started.");
                mainWin = Application.Current?.MainWindow;
                var wasVisible = mainWin?.IsVisible == true;
                if (wasVisible)
                {
                    mainWin.Hide();
                    shouldRestoreMainWindow = true;
                }

                await Task.Delay(150);

                var screenshot = await Task.Run(() => ScreenshotService.CaptureFullScreen());

                var dispatcher = Application.Current?.Dispatcher;
                if (dispatcher == null || dispatcher.CheckAccess())
                {
                    Clipboard.SetImage(screenshot);
                }
                else
                {
                    await dispatcher.InvokeAsync(() => Clipboard.SetImage(screenshot));
                }

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
                AppLogService.Error(ex, "Screenshot failed: {Detail}", ex.ToString());
                SystemStatusMessage = "截图失败：" + ex.Message;
                MessageBox.Show(ex.Message, "截图失败", MessageBoxButton.OK, MessageBoxImage.Error);
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

            if (Application.Current?.MainWindow != null)
            {
                editor.Owner = Application.Current.MainWindow;
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

        private async Task OpenRecordRegionAsync()
        {
            if (IsVideoRecording)
            {
                _recordRegionWindow?.Activate();
                return;
            }

            if (IsAudioRecording)
            {
                MessageBox.Show("当前已有录像/录音任务在进行，请先停止。", "录像", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!EnsureFfmpegAvailable())
            {
                return;
            }

            var outputFolder = await EnsureRecordingOutputFolderAsync();
            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                return;
            }

            if (_recordRegionWindow != null)
            {
                _recordRegionWindow.Activate();
                return;
            }

            var mainWindow = Application.Current?.MainWindow;
            if (mainWindow?.IsVisible == true)
            {
                mainWindow.Hide();
            }

            var recordWindow = new RecordRegionWindow();
            recordWindow.ToggleRecordingRequested += RecordRegionWindow_OnToggleRecordingRequested;
            recordWindow.Closed += async (sender, args) =>
            {
                _recordRegionWindow = null;
                if (IsVideoRecording)
                {
                    await StopVideoRecordingInternalAsync(showMessage: true);
                }

                RestoreWindow();
            };

            _recordRegionWindow = recordWindow;
            recordWindow.Show();
            recordWindow.Activate();
        }

        private async void RecordRegionWindow_OnToggleRecordingRequested(object sender, EventArgs e)
        {
            if (IsVideoRecording)
            {
                await StopVideoRecordingInternalAsync(showMessage: true);
                _recordRegionWindow?.SetRecordingState(false);
                _recordRegionWindow?.Close();
                return;
            }

            if (!EnsureFfmpegAvailable())
            {
                return;
            }

            var outputFolder = await EnsureRecordingOutputFolderAsync();
            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                return;
            }

            try
            {
                var window = sender as RecordRegionWindow ?? _recordRegionWindow;
                if (window == null)
                {
                    return;
                }

                var region = window.GetCaptureRegion();
                var audioDevice = await _recordingService.ResolvePreferredAudioDeviceAsync(CancellationToken.None);
                if (string.IsNullOrWhiteSpace(audioDevice))
                {
                    var continueWithoutAudio = MessageBox.Show(
                        "未检测到可用音频设备，将仅录制画面，是否继续？",
                        "录像",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);
                    if (continueWithoutAudio != MessageBoxResult.Yes)
                    {
                        return;
                    }
                }

                _activeVideoOutputPath = BuildRecordingOutputPath(outputFolder, "录像", ".mp4");
                await _recordingService.StartVideoRecordingAsync(region, _activeVideoOutputPath, audioDevice, CancellationToken.None);
                IsVideoRecording = true;
                window.SetRecordingState(true);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Starting video recording failed.");
                MessageBox.Show(ex.Message, "录像失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ToggleAudioRecordingAsync()
        {
            if (IsAudioRecording)
            {
                await StopAudioRecordingInternalAsync(showMessage: true);
                return;
            }

            if (IsVideoRecording)
            {
                MessageBox.Show("当前已有录像/录音任务在进行，请先停止。", "录音", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!EnsureFfmpegAvailable())
            {
                return;
            }

            var outputFolder = await EnsureAudioOutputFolderAsync();
            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                return;
            }

            try
            {
                _activeAudioOutputPath = BuildRecordingOutputPath(outputFolder, "录音", ".m4a");
                await _recordingService.StartAudioOnlyAsync(_activeAudioOutputPath, CancellationToken.None);
                _audioRecordingStartedAt = DateTime.Now;
                IsAudioRecording = true;
                UpdateAudioRecordingIndicator();
                _audioRecordingTimer.Start();
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Starting audio recording failed.");
                MessageBox.Show(ex.Message, "录音失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task StopVideoRecordingInternalAsync(bool showMessage)
        {
            try
            {
                var result = await _recordingService.StopVideoRecordingAsync();
                IsVideoRecording = false;
                if (!showMessage)
                {
                    return;
                }

                if (result.TimedOut)
                {
                    MessageBox.Show(
                        $"录像文件可能未正确写入，请检查 {_activeVideoOutputPath}。",
                        "录像",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
                else
                {
                    MessageBox.Show($"录像已保存到 {_activeVideoOutputPath}", "录像", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Stopping video recording failed.");
                MessageBox.Show(ex.Message, "录像失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsVideoRecording = false;
            }
        }

        private async Task StopAudioRecordingInternalAsync(bool showMessage)
        {
            try
            {
                var result = await _recordingService.StopAudioOnlyAsync();
                _audioRecordingTimer.Stop();
                IsAudioRecording = false;
                AudioRecordingIndicator = string.Empty;
                if (!showMessage)
                {
                    return;
                }

                if (result.TimedOut)
                {
                    MessageBox.Show(
                        $"录音文件可能未正确写入，请检查 {_activeAudioOutputPath}。",
                        "录音",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
                else
                {
                    MessageBox.Show($"录音已保存到 {_activeAudioOutputPath}", "录音", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Stopping audio recording failed.");
                MessageBox.Show(ex.Message, "录音失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _audioRecordingTimer.Stop();
                IsAudioRecording = false;
                AudioRecordingIndicator = string.Empty;
            }
        }

        private void UpdateAudioRecordingIndicator()
        {
            if (!IsAudioRecording)
            {
                AudioRecordingIndicator = string.Empty;
                return;
            }

            var elapsed = DateTime.Now - _audioRecordingStartedAt;
            AudioRecordingIndicator = $"录音中 {elapsed:mm\\:ss}";
        }

        private bool EnsureFfmpegAvailable()
        {
            if (_recordingService.TryGetFfmpegPath(out _))
            {
                return true;
            }

            MessageBox.Show(
                $"请将 ffmpeg.exe 放到 {_recordingService.ExpectedFfmpegPath} 后再试。",
                "缺少 ffmpeg",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return false;
        }

        private async Task<string> EnsureRecordingOutputFolderAsync()
        {
            var settings = await AppSettingsService.LoadAsync();
            if (!string.IsNullOrWhiteSpace(settings.RecordingOutputFolder) && Directory.Exists(settings.RecordingOutputFolder))
            {
                return settings.RecordingOutputFolder;
            }

            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "请选择录像保存的文件夹（仅本次首设）";
                dialog.ShowNewFolderButton = true;
                if (dialog.ShowDialog() != WinForms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    return string.Empty;
                }

                await AppSettingsService.UpdateAsync(current => current.RecordingOutputFolder = dialog.SelectedPath);
                return dialog.SelectedPath;
            }
        }

        private async Task<string> EnsureAudioOutputFolderAsync()
        {
            var settings = await AppSettingsService.LoadAsync();
            if (!string.IsNullOrWhiteSpace(settings.AudioOutputFolder) && Directory.Exists(settings.AudioOutputFolder))
            {
                return settings.AudioOutputFolder;
            }

            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "请选择录音保存的文件夹（仅本次首设）";
                dialog.ShowNewFolderButton = true;
                if (dialog.ShowDialog() != WinForms.DialogResult.OK || string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    return string.Empty;
                }

                await AppSettingsService.UpdateAsync(current => current.AudioOutputFolder = dialog.SelectedPath);
                return dialog.SelectedPath;
            }
        }

        private static string BuildRecordingOutputPath(string outputFolder, string prefix, string extension)
        {
            Directory.CreateDirectory(outputFolder);
            return Path.Combine(outputFolder, $"{prefix}_{DateTime.Now:yyyyMMdd_HHmmss}{extension}");
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
                    "快捷键不可用",
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
                return $"快捷键 {hotkeyText} 已被其他程序占用，请换一个组合。";
            }

            if (errorCode > 0)
            {
                return $"无法注册快捷键 {hotkeyText}，Win32 错误码：{errorCode}。";
            }

            return $"无法注册快捷键 {hotkeyText}，请换一个组合。";
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
            WgStatusText = IsWgConnected ? $"已连接 {status.IpAddress}" : "未连接";
        }

        private async Task ToggleWireGuardAsync()
        {
            IsWgBusy = true;
            try
            {
                if (IsWgConnected)
                {
                    WgStatusText = "正在断开...";
                    await WireGuardService.DisconnectAsync(WgInterfaceName);
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(WgConfig)) return;
                    WgStatusText = "正在连接...";
                    var status = await WireGuardService.ConnectAsync(WgInterfaceName, WgConfig);
                    if (!string.IsNullOrEmpty(status.ErrorMessage))
                    {
                        MessageBox.Show(status.ErrorMessage, "WireGuard 连接失败", MessageBoxButton.OK, MessageBoxImage.Error);
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
                SqlStatusMessage = $"正在连接 {GetSqlProviderDisplayName(SelectedSqlProvider)}...";
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
                var provider = SqlExportProviderFactory.GetProvider(options.ProviderKind);
                await provider.TestConnectionAsync(options, CancellationToken.None);
                _activeSqlConnectionOptions = CloneSqlConnectionOptions(options);
                _hasUserModifiedSqlConnectionInputs = false;
                await SaveSqlConnectionHistoryAsync(options);

                SqlStatusMessage = "连接成功，正在读取数据库列表...";
                var databases = await provider.GetDatabasesAsync(options, CancellationToken.None);
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
                _activeSqlConnectionOptions = null;
                AppLogService.Error(ex, "SQL connection test failed for {ServerAddress}", SqlServerAddress ?? string.Empty);
                SqlStatusMessage = "连接失败，请检查服务器地址、端口、用户名和密码。";
                MessageBox.Show(ex.Message, $"{GetSqlProviderDisplayName(SelectedSqlProvider)} 连接失败", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    ? $"请输入 {GetSqlProviderDisplayName(SelectedSqlProvider)} 连接信息后测试连接。"
                    : "请选择数据库以加载数据表。";
                return;
            }

            var cancellationTokenSource = new CancellationTokenSource();
            _loadTablesCancellationTokenSource = cancellationTokenSource;

            try
            {
                IsSqlBusy = true;
                SqlStatusMessage = $"正在读取数据库 {SelectedSqlDatabase.Name} 的表列表...";

                var options = GetEffectiveSqlConnectionOptions();
                var provider = SqlExportProviderFactory.GetProvider(options.ProviderKind);
                var tables = await provider.GetTablesAsync(
                    options,
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
                    Filter = "Excel 工作簿 (*.xlsx)|*.xlsx",
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
                SqlStatusMessage = "正在检查数据量并导出 Excel...";

                var provider = SqlExportProviderFactory.GetProvider(options.ProviderKind);
                var exportResult = await provider.ExportTableAsync(
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
                ProviderKind = SelectedSqlProvider,
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
                ProviderKind = options.ProviderKind,
                ServerAddress = options.ServerAddress,
                Port = options.Port,
                Username = options.Username,
                Password = options.Password
            };
        }

        private async Task LoadSqlConnectionHistoryAsync()
        {
            var history = await SqlConnectionHistoryService.LoadAsync(SelectedSqlProvider);
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
                        var defaultPort = GetDefaultSqlPort(SelectedSqlProvider);
                        SqlPort = string.IsNullOrWhiteSpace(history.LastPort) ? defaultPort : history.LastPort;
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
            var history = await SqlConnectionHistoryService.LoadAsync(options?.ProviderKind ?? SelectedSqlProvider);
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

        private void ApplyDefaultSqlPortIfNeeded()
        {
            if (_isApplyingSqlHistory)
            {
                return;
            }

            var current = SqlPort?.Trim();
            if (string.IsNullOrWhiteSpace(current)
                || string.Equals(current, "1433", StringComparison.Ordinal)
                || string.Equals(current, "5432", StringComparison.Ordinal)
                || string.Equals(current, "3306", StringComparison.Ordinal))
            {
                _sqlPort = GetDefaultSqlPort(SelectedSqlProvider);
                OnPropertyChanged(nameof(SqlPort));
            }
        }

        private static string GetDefaultSqlPort(SqlProviderKind providerKind)
        {
            switch (providerKind)
            {
                case SqlProviderKind.PostgreSql:
                    return "5432";
                case SqlProviderKind.MySql:
                    return "3306";
                default:
                    return "1433";
            }
        }

        private static string GetSqlProviderDisplayName(SqlProviderKind providerKind)
        {
            switch (providerKind)
            {
                case SqlProviderKind.PostgreSql:
                    return "PostgreSQL";
                case SqlProviderKind.MySql:
                    return "MySQL";
                default:
                    return "SQL Server";
            }
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
            _recordRegionWindow?.Close();
            _recordRegionWindow = null;
            if (IsVideoRecording)
            {
                _recordingService.StopVideoRecordingAsync().GetAwaiter().GetResult();
            }

            if (IsAudioRecording)
            {
                _recordingService.StopAudioOnlyAsync().GetAwaiter().GetResult();
            }
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
            SystemStatusMessage = target ? "正在恢复实时防护，请在 UAC 弹窗中确认..." : "正在关闭实时防护，请在 UAC 弹窗中确认...";
            try
            {
                await WindowsSecurityService.SetDefenderRealtimeAsync(target);
                await Task.Delay(1500);
                RefreshSystemStatus();
                SystemStatusMessage = target ? "实时防护已恢复。" : "实时防护已关闭。";
            }
            catch (OperationCanceledException)
            {
                SystemStatusMessage = "操作已取消（UAC 未授权）。";
            }
            catch (Exception ex)
            {
                SystemStatusMessage = "操作失败：" + ex.Message;
                MessageBox.Show(ex.Message, "操作失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ToggleAutoUpdateAsync()
        {
            bool target = !IsAutoUpdateEnabled;
            SystemStatusMessage = target ? "正在恢复自动更新，请在 UAC 弹窗中确认..." : "正在停止自动更新，请在 UAC 弹窗中确认...";
            try
            {
                await WindowsSecurityService.SetAutoUpdateAsync(target);
                await Task.Delay(1500);
                RefreshSystemStatus();
                SystemStatusMessage = target ? "自动更新已恢复。" : "自动更新已停止。";
            }
            catch (OperationCanceledException)
            {
                SystemStatusMessage = "操作已取消（UAC 未授权）。";
            }
            catch (Exception ex)
            {
                SystemStatusMessage = "操作失败：" + ex.Message;
                MessageBox.Show(ex.Message, "操作失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task TriggerUpdateNowAsync()
        {
            SystemStatusMessage = "正在触发立即更新，请在 UAC 弹窗中确认...";
            try
            {
                await WindowsSecurityService.TriggerImmediateUpdateAsync();
                RefreshSystemStatus();
                SystemStatusMessage = "更新任务已下发，Windows Update 设置页已打开，可在其中查看进度。";
            }
            catch (OperationCanceledException)
            {
                SystemStatusMessage = "操作已取消（UAC 未授权）。";
            }
            catch (Exception ex)
            {
                SystemStatusMessage = "操作失败：" + ex.Message;
                MessageBox.Show(ex.Message, "操作失败", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show("执行失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool CanExecuteSqlQuery()
            => !IsQueryBusy && SelectedSqlDatabase != null && !string.IsNullOrWhiteSpace(SqlQueryText);

        private bool CanExportQueryResult()
            => !IsQueryBusy && SqlQueryResult != null && SqlQueryResult.Count > 0;

        private async Task ExecuteSqlQueryAsync()
        {
            IsQueryBusy = true;
            QueryStatusMessage = "正在执行查询...";
            SqlQueryResult = null;
            try
            {
                var options = GetEffectiveSqlConnectionOptions();
                var provider = SqlExportProviderFactory.GetProvider(options.ProviderKind);
                var table = await provider.ExecuteQueryAsync(
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
                MessageBox.Show(ex.Message, "SQL 执行失败", MessageBoxButton.OK, MessageBoxImage.Error);
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
                Filter = "Excel 工作簿 (*.xlsx)|*.xlsx",
                FileName = $"QueryResult_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx",
                DefaultExt = ".xlsx",
                AddExtension = true,
                OverwritePrompt = true
            };
            if (dialog.ShowDialog() != true) return;

            IsQueryBusy = true;
            QueryStatusMessage = "正在导出 Excel...";
            try
            {
                var table = SqlQueryResult.Table;
                var result = await SqlExportService.ExportDataTableAsync(
                    table,
                    "QueryResult",
                    dialog.FileName,
                    CancellationToken.None);

                QueryStatusMessage = $"导出完成，共 {result.RowCount} 行。";
                MessageBox.Show(
                    $"导出成功。\n文件路径：{result.FilePath}",
                    "导出完成",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Query result export failed");
                QueryStatusMessage = "导出失败：" + ex.Message;
                MessageBox.Show(ex.Message, "导出失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsQueryBusy = false;
            }
        }

        private async Task LoadOptimizationReportsAsync()
        {
            try
            {
                var reports = await _optimizationReportService.LoadAllAsync();
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    foreach (var existing in OptimizationReports)
                    {
                        existing.PropertyChanged -= OptimizationReport_OnPropertyChanged;
                    }

                    OptimizationReports.Clear();
                    foreach (var report in reports)
                    {
                        report.PropertyChanged += OptimizationReport_OnPropertyChanged;
                        OptimizationReports.Add(report);
                    }

                    OnPropertyChanged(nameof(HasSelectedOptimizationReports));
                });
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading optimization reports failed.");
            }
        }

        private async Task LoadWeChatRootsAsync()
        {
            try
            {
                var roots = await Task.Run(() => _weChatDataLocator.LocateRoots());
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    WeChatRoots.Clear();
                    foreach (var root in roots)
                    {
                        WeChatRoots.Add(root);
                    }

                    SelectedWeChatRoot = WeChatRoots.FirstOrDefault();
                    if (SelectedWeChatRoot == null)
                    {
                        WeChatCleanupStatusMessage = "未检测到本机微信数据。";
                        WeChatBackupStatusMessage = "未检测到本机微信数据。";
                    }
                });
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading WeChat roots failed.");
            }
        }

        private async Task StartAutoOptimizeAsync()
        {
            if (IsAutoOptimizeBusy)
            {
                return;
            }

            var restartExplorerResult = MessageBox.Show(
                "清理缩略图缓存需要临时重启资源管理器（explorer.exe），是否允许该步骤执行？",
                "自动优化确认",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            IsAutoOptimizeBusy = true;
            try
            {
                _systemOptimizationService.AllowExplorerRestartForThumbnailCleanup = restartExplorerResult == MessageBoxResult.Yes;
                var progress = new Progress<string>(message => AutoOptimizeStatusMessage = message);
                var report = await Task.Run(() => _systemOptimizationService.RunAsync(progress, CancellationToken.None));
                await _optimizationReportService.SaveAsync(report);

                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    report.PropertyChanged += OptimizationReport_OnPropertyChanged;
                    OptimizationReports.Insert(0, report);
                    OnPropertyChanged(nameof(HasSelectedOptimizationReports));
                });

                AutoOptimizeStatusMessage = $"自动优化完成：{report.Summary}，释放空间 {report.TotalBytesFreedDisplay}。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Auto optimize execution failed.");
                AutoOptimizeStatusMessage = "自动优化失败：" + ex.Message;
                MessageBox.Show(ex.Message, "自动优化失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsAutoOptimizeBusy = false;
            }
        }

        private async Task StartJunkScanAsync()
        {
            IsJunkBusy = true;
            try
            {
                var progress = new Progress<string>(message => JunkStatusMessage = message);
                var scanned = await Task.Run(() => _junkCleanupService.ScanAsync(progress, CancellationToken.None));
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    foreach (var candidate in JunkCandidates)
                    {
                        candidate.PropertyChanged -= JunkCandidate_OnPropertyChanged;
                    }

                    JunkCandidates.Clear();
                    foreach (var candidate in scanned)
                    {
                        candidate.PropertyChanged += JunkCandidate_OnPropertyChanged;
                        JunkCandidates.Add(candidate);
                    }

                    OnPropertyChanged(nameof(HasJunkCandidates));
                    TriggerCommandRequery();
                });

                var totalBytes = scanned.Sum(x => x.Bytes);
                JunkStatusMessage = scanned.Count == 0
                    ? "未发现可清理项目。"
                    : $"扫描完成：{scanned.Count} 项，预计可释放 {FileSizeFormatter.Format(totalBytes)}。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Junk scan failed.");
                JunkStatusMessage = "扫描失败：" + ex.Message;
                MessageBox.Show(ex.Message, "垃圾扫描失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsJunkBusy = false;
            }
        }

        private bool CanRunJunkCleanup()
        {
            return !IsJunkBusy && JunkCandidates.Any(x => x.IsSelected);
        }

        private async Task RunJunkCleanupAsync()
        {
            var selected = JunkCandidates.Where(x => x.IsSelected).ToList();
            if (selected.Count == 0)
            {
                return;
            }

            var totalBytes = selected.Sum(x => x.Bytes);
            var confirm = MessageBox.Show(
                $"将删除 {selected.Count} 项，约 {FileSizeFormatter.Format(totalBytes)}，操作不可撤销。是否继续？",
                "确认清理",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            IsJunkBusy = true;
            try
            {
                var progress = new Progress<string>(message => JunkStatusMessage = message);
                var execution = await Task.Run(() => _junkCleanupService.CleanupAsync(selected, progress, CancellationToken.None));
                JunkStatusMessage = $"已清理 {execution.DeletedCount} 项，释放空间 {FileSizeFormatter.Format(execution.FreedBytes)}。";

                var report = new OptimizationReportItem
                {
                    Id = Guid.NewGuid().ToString("N").Substring(0, 8),
                    StartedAt = DateTime.Now,
                    FinishedAt = DateTime.Now,
                    ReportType = "JunkCleanup",
                    Steps = execution.Steps,
                    TotalBytesFreed = execution.FreedBytes,
                    Summary = $"垃圾清理：成功 {execution.DeletedCount} 项。"
                };

                await _optimizationReportService.SaveAsync(report);
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    report.PropertyChanged += OptimizationReport_OnPropertyChanged;
                    OptimizationReports.Insert(0, report);
                    OnPropertyChanged(nameof(HasSelectedOptimizationReports));
                });

                MessageBox.Show(
                    $"已清理 {execution.DeletedCount} 项，释放空间 {FileSizeFormatter.Format(execution.FreedBytes)}。",
                    "清理完成",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Junk cleanup failed.");
                JunkStatusMessage = "清理失败：" + ex.Message;
                MessageBox.Show(ex.Message, "垃圾清理失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsJunkBusy = false;
            }
        }

        private async Task ScanWeChatCleanupAsync()
        {
            if (SelectedWeChatRoot == null)
            {
                WeChatCleanupStatusMessage = "未检测到本机微信数据。";
                return;
            }

            IsWeChatCleanupBusy = true;
            try
            {
                var options = BuildWeChatCleanupOptions(
                    WeChatCleanupStartDate,
                    WeChatCleanupEndDate,
                    WeChatCleanupIncludeText,
                    WeChatCleanupIncludeImage,
                    WeChatCleanupIncludeVideo,
                    WeChatCleanupIncludeVoice,
                    WeChatCleanupIncludeFile,
                    WeChatCleanupIncludeCache);

                var progress = new Progress<string>(message => WeChatCleanupStatusMessage = message);
                var scanResult = await Task.Run(() => _weChatCleanupService.ScanAsync(
                    new[] { SelectedWeChatRoot },
                    options,
                    progress,
                    CancellationToken.None));

                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    foreach (var item in WeChatCleanupCandidates)
                    {
                        item.PropertyChanged -= WeChatCleanupCandidate_OnPropertyChanged;
                    }

                    WeChatCleanupCandidates.Clear();
                    foreach (var candidate in scanResult.Candidates)
                    {
                        candidate.PropertyChanged += WeChatCleanupCandidate_OnPropertyChanged;
                        WeChatCleanupCandidates.Add(candidate);
                    }

                    OnPropertyChanged(nameof(HasWeChatCleanupCandidates));
                    TriggerCommandRequery();
                });

                var totalBytes = scanResult.Candidates.Sum(x => x.Bytes);
                var note = scanResult.PendingNotes.Count > 0
                    ? " " + string.Join("；", scanResult.PendingNotes)
                    : string.Empty;
                WeChatCleanupStatusMessage = scanResult.Candidates.Count == 0
                    ? "未发现符合条件的微信数据。" + note
                    : $"扫描完成：{scanResult.Candidates.Count} 项，约 {FileSizeFormatter.Format(totalBytes)}。{note}";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "WeChat cleanup scan failed.");
                WeChatCleanupStatusMessage = "扫描失败：" + ex.Message;
                MessageBox.Show(ex.Message, "微信清理扫描失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsWeChatCleanupBusy = false;
            }
        }

        private bool CanStartWeChatCleanup()
        {
            return !IsWeChatCleanupBusy && WeChatCleanupCandidates.Any(x => x.IsSelected) && SelectedWeChatRoot != null;
        }

        private async Task StartWeChatCleanupAsync()
        {
            var selected = WeChatCleanupCandidates.Where(x => x.IsSelected).ToList();
            if (selected.Count == 0)
            {
                WeChatCleanupStatusMessage = "请先扫描并选择要清理的项目。";
                return;
            }

            var totalBytes = selected.Sum(x => x.Bytes);
            var confirm = MessageBox.Show(
                $"将删除 {selected.Count} 项微信数据（{FileSizeFormatter.Format(totalBytes)}），操作不可撤销，建议先备份。是否继续？",
                "确认微信清理",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            IsWeChatCleanupBusy = true;
            try
            {
                var progress = new Progress<string>(message => WeChatCleanupStatusMessage = message);
                var execution = await Task.Run(() => _weChatCleanupService.CleanupAsync(
                    selected,
                    new[] { SelectedWeChatRoot },
                    progress,
                    CancellationToken.None));

                var report = new OptimizationReportItem
                {
                    Id = Guid.NewGuid().ToString("N").Substring(0, 8),
                    StartedAt = DateTime.Now,
                    FinishedAt = DateTime.Now,
                    ReportType = "WeChatCleanup",
                    Steps = execution.Steps,
                    TotalBytesFreed = execution.FreedBytes,
                    Summary = $"微信清理：成功 {execution.DeletedCount} 项。"
                };

                await _optimizationReportService.SaveAsync(report);
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    report.PropertyChanged += OptimizationReport_OnPropertyChanged;
                    OptimizationReports.Insert(0, report);
                    OnPropertyChanged(nameof(HasSelectedOptimizationReports));
                });

                WeChatCleanupStatusMessage = $"微信清理完成：成功 {execution.DeletedCount} 项，释放空间 {FileSizeFormatter.Format(execution.FreedBytes)}。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "WeChat cleanup execution failed.");
                WeChatCleanupStatusMessage = "清理失败：" + ex.Message;
                MessageBox.Show(ex.Message, "微信清理失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsWeChatCleanupBusy = false;
            }
        }

        private bool CanStartWeChatBackup()
        {
            return !IsWeChatBackupBusy
                && SelectedWeChatRoot != null
                && !string.IsNullOrWhiteSpace(WeChatBackupOutputFolder);
        }

        private async Task StartWeChatBackupAsync()
        {
            if (SelectedWeChatRoot == null)
            {
                WeChatBackupStatusMessage = "未检测到本机微信数据。";
                return;
            }

            if (string.IsNullOrWhiteSpace(WeChatBackupOutputFolder))
            {
                SelectWeChatBackupOutputFolder();
                if (string.IsNullOrWhiteSpace(WeChatBackupOutputFolder))
                {
                    return;
                }
            }

            IsWeChatBackupBusy = true;
            try
            {
                var options = new WeChatBackupOptions
                {
                    Root = SelectedWeChatRoot,
                    StartDate = (WeChatBackupStartDate ?? DateTime.Today.AddDays(-30)).Date,
                    EndDate = (WeChatBackupEndDate ?? DateTime.Today).Date,
                    Categories = WeChatCleanupService.BuildCategories(
                        WeChatBackupIncludeText,
                        WeChatBackupIncludeImage,
                        WeChatBackupIncludeVideo,
                        WeChatBackupIncludeVoice,
                        WeChatBackupIncludeFile,
                        WeChatBackupIncludeCache),
                    OutputDirectory = WeChatBackupOutputFolder
                };

                var progress = new Progress<WeChatBackupProgress>(p =>
                {
                    WeChatBackupStatusMessage = $"备份中：{p.RelativePath}（{p.Current}/{p.Total}）";
                });

                var backupResult = await Task.Run(() => _weChatBackupService.BackupAsync(options, progress, CancellationToken.None));
                WeChatBackupStatusMessage = $"备份完成：{backupResult.FileCount} 个文件，{FileSizeFormatter.Format(backupResult.TotalBytes)}。";
                MessageBox.Show(
                    $"已备份 {backupResult.FileCount} 个文件（{FileSizeFormatter.Format(backupResult.TotalBytes)}）到 {backupResult.ZipPath}。",
                    "备份完成",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "WeChat backup failed.");
                WeChatBackupStatusMessage = "备份失败：" + ex.Message;
                MessageBox.Show(ex.Message, "微信备份失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsWeChatBackupBusy = false;
            }
        }

        private bool CanStartWeChatRestore()
        {
            if (IsWeChatRestoreBusy || string.IsNullOrWhiteSpace(WeChatRestoreZipPath))
            {
                return false;
            }

            if (RestoreToOriginal)
            {
                return true;
            }

            return !string.IsNullOrWhiteSpace(WeChatRestoreTargetRoot);
        }

        private async Task StartWeChatRestoreAsync()
        {
            if (string.IsNullOrWhiteSpace(WeChatRestoreZipPath))
            {
                WeChatRestoreStatusMessage = "请先选择备份文件。";
                return;
            }

            if (!RestoreToOriginal && string.IsNullOrWhiteSpace(WeChatRestoreTargetRoot))
            {
                SelectWeChatRestoreTargetRoot();
                if (string.IsNullOrWhiteSpace(WeChatRestoreTargetRoot))
                {
                    return;
                }
            }

            IsWeChatRestoreBusy = true;
            try
            {
                var restoreResult = await Task.Run(() => _weChatBackupService.RestoreAsync(
                    new WeChatRestoreOptions
                    {
                        ZipPath = WeChatRestoreZipPath,
                        RestoreToOriginal = RestoreToOriginal,
                        CustomTargetRoot = RestoreToOriginal ? null : WeChatRestoreTargetRoot
                    },
                    new Progress<string>(msg => WeChatRestoreStatusMessage = msg),
                    CancellationToken.None));

                WeChatRestoreStatusMessage = $"恢复完成：成功 {restoreResult.Success}，失败 {restoreResult.Failed}。";
                if (restoreResult.Failed > 0)
                {
                    MessageBox.Show(
                        $"恢复完成：成功 {restoreResult.Success}，失败 {restoreResult.Failed}（详见日志）。",
                        "微信恢复",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
                else
                {
                    MessageBox.Show(
                        $"恢复完成：成功 {restoreResult.Success}，失败 {restoreResult.Failed}（详见日志）。",
                        "微信恢复",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "WeChat restore failed.");
                WeChatRestoreStatusMessage = "恢复失败：" + ex.Message;
                MessageBox.Show(ex.Message, "微信恢复失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsWeChatRestoreBusy = false;
            }
        }

        private void SelectWeChatBackupOutputFolder()
        {
            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "请选择微信备份输出目录";
                dialog.ShowNewFolderButton = true;
                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    WeChatBackupOutputFolder = dialog.SelectedPath;
                }
            }
        }

        private void SelectWeChatRestoreZip()
        {
            var dialog = new OpenFileDialog
            {
                Filter = "微信备份 (*.zip)|*.zip",
                Title = "选择微信备份文件"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            WeChatRestoreZipPath = dialog.FileName;
            _ = LoadRestoreManifestSummaryAsync(dialog.FileName);
        }

        private async Task LoadRestoreManifestSummaryAsync(string zipPath)
        {
            try
            {
                var manifest = await _weChatBackupService.ReadManifestAsync(zipPath, CancellationToken.None);
                var count = manifest.Entries?.Count ?? 0;
                var total = manifest.Entries?.Sum(x => x.Size) ?? 0L;
                var categories = manifest.Categories == null || manifest.Categories.Count == 0
                    ? "无"
                    : string.Join(", ", manifest.Categories);

                var builder = new StringBuilder();
                builder.AppendLine($"微信版本：{manifest.WechatVariant}");
                builder.AppendLine($"wxId：{manifest.WxId}");
                builder.AppendLine($"时间范围：{manifest.DateRange?.Start} ~ {manifest.DateRange?.End}");
                builder.AppendLine($"类别：{categories}");
                builder.AppendLine($"条目数：{count}");
                builder.Append($"总大小：{FileSizeFormatter.Format(total)}");
                WeChatRestoreManifestSummary = builder.ToString();
                WeChatRestoreStatusMessage = "已解析备份清单。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Reading restore manifest failed for {Path}", zipPath ?? string.Empty);
                WeChatRestoreManifestSummary = "清单读取失败。";
                WeChatRestoreStatusMessage = "无法解析备份清单：" + ex.Message;
            }
        }

        private void SelectWeChatRestoreTargetRoot()
        {
            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "请选择恢复目标目录";
                dialog.ShowNewFolderButton = true;
                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    WeChatRestoreTargetRoot = dialog.SelectedPath;
                }
            }
        }

        private async void DeleteSelectedReports()
        {
            var selected = OptimizationReports.Where(x => x.IsSelected).ToList();
            if (selected.Count == 0)
            {
                return;
            }

            var confirm = MessageBox.Show(
                $"确定删除选中的 {selected.Count} 条优化报告吗？",
                "确认删除",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            try
            {
                await _optimizationReportService.DeleteAsync(selected.Select(x => x.Id));
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    foreach (var item in selected)
                    {
                        item.PropertyChanged -= OptimizationReport_OnPropertyChanged;
                        OptimizationReports.Remove(item);
                    }

                    OnPropertyChanged(nameof(HasSelectedOptimizationReports));
                    TriggerCommandRequery();
                });
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Deleting selected optimization reports failed.");
                MessageBox.Show(ex.Message, "删除失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool CanDeleteSelectedReports()
        {
            return OptimizationReports.Any(x => x.IsSelected)
                && !IsAutoOptimizeBusy
                && !IsJunkBusy
                && !IsWeChatCleanupBusy;
        }

        private void ShowReportDetails(object parameter)
        {
            if (!(parameter is OptimizationReportItem report))
            {
                return;
            }

            var lines = new List<string>
            {
                $"报告类型：{report.ReportTypeDisplay}",
                $"开始时间：{report.StartedAt:yyyy-MM-dd HH:mm:ss}",
                $"结束时间：{report.FinishedAt:yyyy-MM-dd HH:mm:ss}",
                $"释放空间：{report.TotalBytesFreedDisplay}",
                $"摘要：{report.Summary}",
                "",
                "步骤明细："
            };

            if (report.Steps == null || report.Steps.Count == 0)
            {
                lines.Add("（无步骤明细）");
            }
            else
            {
                foreach (var step in report.Steps)
                {
                    lines.Add($"- {step.Name} [{step.Status}] 释放 {FileSizeFormatter.Format(step.BytesFreed)}，耗时 {step.Duration:g}");
                    if (!string.IsNullOrWhiteSpace(step.Detail))
                    {
                        lines.Add($"  {step.Detail}");
                    }
                }
            }

            MessageBox.Show(string.Join(Environment.NewLine, lines), "优化报告详情", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private static WeChatCleanupScanOptions BuildWeChatCleanupOptions(
            DateTime? startDate,
            DateTime? endDate,
            bool includeText,
            bool includeImage,
            bool includeVideo,
            bool includeVoice,
            bool includeFile,
            bool includeCache)
        {
            var start = (startDate ?? DateTime.Today.AddDays(-30)).Date;
            var end = (endDate ?? DateTime.Today).Date;
            if (start > end)
            {
                var swap = start;
                start = end;
                end = swap;
            }

            return new WeChatCleanupScanOptions
            {
                StartDate = start,
                EndDate = end,
                Categories = WeChatCleanupService.BuildCategories(
                    includeText,
                    includeImage,
                    includeVideo,
                    includeVoice,
                    includeFile,
                    includeCache)
            };
        }

        private void OptimizationReport_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(OptimizationReportItem.IsSelected))
            {
                OnPropertyChanged(nameof(HasSelectedOptimizationReports));
                TriggerCommandRequery();
            }
        }

        private void JunkCandidate_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(JunkCandidate.IsSelected))
            {
                TriggerCommandRequery();
            }
        }

        private void WeChatCleanupCandidate_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(WeChatCleanupCandidate.IsSelected))
            {
                TriggerCommandRequery();
            }
        }

        private void TriggerCommandRequery()
        {
            _testSqlConnectionCommand?.RaiseCanExecuteChanged();
            _exportSqlTableCommand?.RaiseCanExecuteChanged();
            _executeQueryCommand?.RaiseCanExecuteChanged();
            _exportQueryResultCommand?.RaiseCanExecuteChanged();
            _startAutoOptimizeCommand?.RaiseCanExecuteChanged();
            _startJunkScanCommand?.RaiseCanExecuteChanged();
            _runJunkCleanupCommand?.RaiseCanExecuteChanged();
            _scanWeChatCleanupCommand?.RaiseCanExecuteChanged();
            _startWeChatCleanupCommand?.RaiseCanExecuteChanged();
            _startWeChatBackupCommand?.RaiseCanExecuteChanged();
            _startWeChatRestoreCommand?.RaiseCanExecuteChanged();
            _deleteSelectedReportsCommand?.RaiseCanExecuteChanged();
            _selectWeChatBackupOutputFolderCommand?.RaiseCanExecuteChanged();
            _selectWeChatRestoreZipCommand?.RaiseCanExecuteChanged();
            _selectWeChatRestoreTargetRootCommand?.RaiseCanExecuteChanged();
            _openRecordRegionCommand?.RaiseCanExecuteChanged();
            _toggleAudioRecordingCommand?.RaiseCanExecuteChanged();
            _importCodexProfileCommand?.RaiseCanExecuteChanged();
            _exportCodexProfileCommand?.RaiseCanExecuteChanged();
        }

        private static string BuildAppVersionText()
        {
            var version = typeof(MainViewModel).Assembly.GetName().Version;
            if (version == null)
            {
                return "v1.0.0";
            }

            var build = version.Build < 0 ? 0 : version.Build;
            return $"v{version.Major}.{version.Minor}.{build}";
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
                ? "已内置保存 config.toml 和 auth.json 内容"
                : "未内置配置内容（建议重新拖入文件夹）";

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class SqlProviderOption
    {
        public SqlProviderOption(SqlProviderKind kind, string displayName)
        {
            Kind = kind;
            DisplayName = displayName ?? string.Empty;
        }

        public SqlProviderKind Kind { get; }
        public string DisplayName { get; }
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
                MessageBox.Show(ex.Message, "操作失败", MessageBoxButton.OK, MessageBoxImage.Error);
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
                MessageBox.Show(ex.Message, "操作失败", MessageBoxButton.OK, MessageBoxImage.Error);
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




