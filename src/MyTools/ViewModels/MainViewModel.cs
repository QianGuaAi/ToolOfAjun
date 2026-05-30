using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using MyTools.Services;
using MyTools.Shared;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;
using DrawingImaging = System.Drawing.Imaging;
using Media = System.Windows.Media;

namespace MyTools.ViewModels
{
    public class MainViewModel : INotifyPropertyChanged, IDisposable
    {
        private ObservableCollection<NetworkData> _networkList;
        private ObservableCollection<StartupItem> _startupList;
        private ICollectionView _filteredStartupView;
        private ObservableCollection<DatabaseItem> _sqlDatabaseList;
        private ObservableCollection<TableItem> _sqlTableList;
        private ObservableCollection<TableItem> _allSqlTableList;
        private ObservableCollection<InstalledProgram> _installedPrograms;
        private ICollectionView _filteredInstalledProgramsView;
        private ObservableCollection<string> _sqlServerAddressHistory;
        private ObservableCollection<string> _sqlUsernameHistory;
        private ObservableCollection<string> _sqlPasswordHistory;
        private ObservableCollection<SqlConnectionHistoryItem> _sqlRecentConnections;
        private ObservableCollection<CodexProfileItem> _codexProfiles;
        private ObservableCollection<HomeCommandItem> _homeCommandItems;
        private ICollectionView _filteredHomeCommandView;
        private ICollectionView _filteredVideoViewerPlaylistView;
        private ObservableCollection<HomeRecentItem> _homeRecentItems;
        private string _homeCommandSearchText = string.Empty;
        private ObservableCollection<ScreenshotHistoryItem> _screenshotHistoryItems;
        private string _currentModule;
        private string _currentSystemSection = "Startup";
        private int _selectedSystemSectionIndex;
        private ScheduleViewModel _schedule;
        private SystemSettingsViewModel _systemSettings;
        private MultimediaViewModel _multimedia;
        private FrpViewModel _frp;
        private string _sqlServerAddress;
        private string _sqlPort = "1433";
        private SqlProviderKind _selectedSqlProvider = SqlProviderKind.SqlServer;
        private string _sqlUsername;
        private string _sqlPassword;
        private string _sqlTableSearchText;
        private string _startupSearchText = string.Empty;
        private string _installedProgramSearchText = string.Empty;
        private string _installedProgramSortMode = "名称 A-Z";
        private string _installedProgramSizeFilter = "全部大小";
        private string _installedProgramDateFilter = "全部日期";
        private DatabaseItem _selectedSqlDatabase;
        private TableItem _selectedSqlTable;
        private ICollectionView _filteredSqlTableView;
        private string _sqlStatusMessage = "请输入 SQL Server 连接信息后测试连接。";
        private bool _isSqlBusy;
        private CancellationTokenSource _sqlExportCancellationTokenSource;
        private CancellationTokenSource _loadTablesCancellationTokenSource;
        private bool _suppressSqlTableAutoLoad;
        private bool _isApplyingSqlHistory;
        private bool _hasUserModifiedSqlConnectionInputs;
        private SqlServerConnectionOptions _activeSqlConnectionOptions;
        private bool _isDefenderEnabled = true;
        private bool _isAutoUpdateEnabled = true;
        private string _systemStatusMessage = string.Empty;
        private string _sqlQueryText = string.Empty;
        private DataView _sqlQueryResult;
        private bool _isQueryBusy;
        private string _queryStatusMessage = string.Empty;
        private bool _isInstalledProgramsBusy;
        private string _installedProgramsStatusMessage = "点击刷新加载可卸载程序列表。";
        private ObservableCollection<OptimizationReportItem> _optimizationReports;
        private ObservableCollection<JunkCandidate> _junkCandidates;
        private ObservableCollection<WeChatCleanupCandidate> _weChatCleanupCandidates;
        private ObservableCollection<WeChatRoot> _weChatRoots;
        private ObservableCollection<RecentWeChatBackupItem> _recentWeChatBackups;
        private WeChatRoot _selectedWeChatRoot;
        private bool _isAutoOptimizeBusy;
        private string _autoOptimizeStatusMessage = "点击“开始自动优化”执行白名单优化流程。";
        private bool _isPowerPolicyBusy;
        private string _powerPolicyStatusMessage = "点击应用后，当前电源计划将设置为不关闭屏幕、不关闭硬盘、不进入睡眠/休眠。";
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
        private bool _weChatRestoreIncludeText = true;
        private bool _weChatRestoreIncludeImage = true;
        private bool _weChatRestoreIncludeVideo = true;
        private bool _weChatRestoreIncludeVoice = true;
        private bool _weChatRestoreIncludeFile = true;
        private bool _weChatRestoreIncludeCache = true;
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
        private RecordingService _recordingService;
        private bool _isVideoRecording;
        private bool _isAudioRecording;
        private bool _isDisposed;
        private DateTime _audioRecordingStartedAt;
        private string _audioRecordingIndicator = string.Empty;
        private string _recordingOutputFolderText = "未设置";
        private string _audioOutputFolderText = "未设置";
        private string _activeVideoOutputPath = string.Empty;
        private string _activeAudioOutputPath = string.Empty;
        private bool _isGifRecordingMode;
        private string _recordingFrameRateOption = "30 FPS";
        private string _recordingQualityOption = "均衡";
        private readonly DispatcherTimer _audioRecordingTimer;
        private uint _pendingModifiers = 0x0006;
        private uint _pendingKey = 0x5A;
        private uint _pendingVideoRecordModifiers;
        private uint _pendingVideoRecordKey;
        private string _videoRecordHotkeyText = "\u672a\u8bbe\u7f6e";
        private bool _isCapturingVideoRecordHotkey;
        private uint _pendingAudioRecordModifiers;
        private uint _pendingAudioRecordKey;
        private string _audioRecordHotkeyText = "\u672a\u8bbe\u7f6e";
        private bool _isCapturingAudioRecordHotkey;
        private string _codexProfilesStatusMessage = "拖入包含 config.toml 和 auth.json 的文件夹，生成可应用的 Codex 配置记录。";
        private string _codexNextSwitchPreview = "无可用轮换目标";
        private CodexRotationSettings _codexRotationSettings = new CodexRotationSettings();
        private bool _isAutoStartEnabled;
        private FileHashResult _currentFileHashResult;
        private string _fileHashResult = string.Empty;
        private string _expectedFileHash = string.Empty;
        private string _fileHashCompareResult = string.Empty;
        private ObservableCollection<FileHashResult> _batchFileHashResults = new ObservableCollection<FileHashResult>();
        private Dictionary<string, ImportedHashEntry> _importedFileHashEntries = new Dictionary<string, ImportedHashEntry>(StringComparer.OrdinalIgnoreCase);
        private bool _isFileHashBusy;
        private string _fileHashStatusMessage = "选择文件后计算 MD5 / SHA-1 / SHA-256 / CRC32。";
        private BitmapSource _imageViewerOriginalImage;
        private BitmapSource _imageViewerPreviewImage;
        private string _imageViewerFilePath = string.Empty;
        private string _imageViewerStatusMessage = "选择图片后可预览、缩放、旋转、翻转、灰度、复制或另存。";
        private double _imageViewerZoom = 1.0;
        private int _imageViewerRotationDegrees;
        private bool _imageViewerFlipHorizontal;
        private bool _imageViewerFlipVertical;
        private bool _imageViewerIsGrayscale;
        private string _imageViewerCropMode = "原图";
        private int _imageViewerBrightness;
        private int _imageViewerContrast;
        private int _imageViewerSharpenAmount;
        private List<string> _imageViewerDirectoryFiles = new List<string>();
        private int _imageViewerDirectoryIndex = -1;
        private bool _imageViewerFitToFrame;
        private Uri _videoViewerSource;
        private string _videoViewerFilePath = string.Empty;
        private string _videoViewerStatusMessage = "选择音频或视频后可播放、暂停、拖动进度或调节音量。";
        private ObservableCollection<VideoPlaylistItem> _videoViewerPlaylist = new ObservableCollection<VideoPlaylistItem>();
        private int _videoViewerPlaylistIndex = -1;
        private bool _isVideoViewerCapturingFrame;
        private bool _isVideoViewerGeneratingWaveform;
        private BitmapImage _videoViewerWaveformImage;
        private string _videoViewerWaveformPath = string.Empty;
        private bool _isVideoViewerPlaying;
        private double _videoViewerPositionSeconds;
        private double _videoViewerDurationSeconds;
        private double _videoViewerVolume = 0.75;
        private bool _videoViewerIsMuted;
        private double _videoViewerSpeedRatio = 1.0;
        private string _videoViewerPlayMode = "顺序";
        private double _videoViewerLoopStartSeconds = -1;
        private double _videoViewerLoopEndSeconds = -1;
        private bool _videoViewerIsLoopEnabled;
        private ObservableCollection<VideoLoopRangeItem> _videoViewerLoopRanges = new ObservableCollection<VideoLoopRangeItem>();
        private ObservableCollection<VideoBookmarkItem> _videoViewerBookmarks = new ObservableCollection<VideoBookmarkItem>();
        private string _videoViewerSubtitleFilePath = string.Empty;
        private string _videoViewerSubtitleStatus = "未载入字幕";
        private string _videoViewerSubtitleText = string.Empty;
        private IReadOnlyList<SubtitleCue> _videoViewerSubtitles = new List<SubtitleCue>();
        private int _videoViewerSubtitleIndex = -1;
        private double _videoViewerSubtitleOffsetSeconds;
        private double _videoViewerSubtitleFontSize = 16;
        private ObservableCollection<RecentPlaylistItem> _videoViewerRecentPlaylists;
        private ObservableCollection<FavoritePlaylistItem> _videoViewerFavoritePlaylists;
        private string _videoViewerPlaylistSearchText = string.Empty;
        private readonly Random _videoViewerRandom = new Random();

        private readonly AsyncRelayCommand _executeQueryCommand;
        private readonly AsyncRelayCommand _exportQueryResultCommand;
        private readonly AsyncRelayCommand _exportQueryResultCsvCommand;
        private readonly RelayCommand _cancelQueryCommand;
        private readonly AsyncRelayParameterCommand _applyCodexProfileCommand;
        private readonly AsyncRelayParameterCommand _exportCodexProfileCommand;
        private readonly AsyncRelayParameterCommand _previewCodexProfileDiffCommand;
        private readonly AsyncRelayCommand _importCodexProfileCommand;
        private readonly AsyncRelayCommand _importCodexCpaTokenCommand;
        private readonly AsyncRelayParameterCommand _refreshCodexProfileCommand;
        private readonly AsyncRelayParameterCommand _renameCodexProfileCommand;
        private readonly AsyncRelayParameterCommand _editCodexProfileNoteCommand;
        private readonly AsyncRelayCommand _restoreLastCodexBackupCommand;
        private readonly AsyncRelayCommand _exportCodexProfilesEncBoxCommand;
        private readonly AsyncRelayCommand _importCodexProfilesEncBoxCommand;
        private readonly AsyncRelayCommand _rotateToNextCodexProfileCommand;
        private readonly AsyncRelayCommand _restartCodexDesktopCommand;
        private readonly AsyncRelayParameterCommand _toggleCodexProfileRotationCommand;
        private readonly AsyncRelayCommand _openRecordRegionCommand;
        private readonly AsyncRelayCommand _toggleAudioRecordingCommand;
        private readonly AsyncRelayCommand _refreshInstalledProgramsCommand;
        private readonly AsyncRelayParameterCommand _uninstallProgramCommand;

        private readonly AsyncRelayCommand _testSqlConnectionCommand;
        private readonly AsyncRelayCommand _exportSqlTableCommand;
        private readonly AsyncRelayCommand _toggleDefenderCommand;
        private readonly AsyncRelayCommand _toggleAutoUpdateCommand;
        private readonly AsyncRelayCommand _triggerUpdateNowCommand;
        private readonly AsyncRelayCommand _startAutoOptimizeCommand;
        private readonly AsyncRelayCommand _startJunkScanCommand;
        private readonly AsyncRelayCommand _runJunkCleanupCommand;
        private readonly RelayCommand _exportJunkCleanupPlanCommand;
        private readonly AsyncRelayCommand _scanWeChatCleanupCommand;
        private readonly AsyncRelayCommand _startWeChatCleanupCommand;
        private readonly AsyncRelayCommand _startWeChatBackupCommand;
        private readonly AsyncRelayCommand _startWeChatRestoreCommand;
        private readonly RelayCommand _deleteSelectedReportsCommand;
        private readonly RelayCommand _selectWeChatBackupOutputFolderCommand;
        private readonly RelayCommand _selectWeChatRestoreZipCommand;
        private readonly RelayCommand _selectWeChatRestoreTargetRootCommand;
        private readonly RelayParameterCommand _showReportDetailsCommand;
        private readonly RelayCommand _copyBenchmarkResultsCommand;
        private readonly RelayCommand _exportBenchmarkResultsCommand;
        private readonly AsyncRelayCommand _applyAlwaysOnPowerPolicyCommand;

        private OptimizationReportService _optimizationReportService;
        private SystemOptimizationService _systemOptimizationService;
        private JunkCleanupService _junkCleanupService;
        private WeChatDataLocator _weChatDataLocator;
        private WeChatCleanupService _weChatCleanupService;
        private WeChatBackupService _weChatBackupService;
        private bool _sqlHistoryLoadRequested;
        private bool _screenshotHotkeysLoadRequested;
        private bool _screenshotStartupLoadRequested;
        private bool _videoViewerStartupLoadRequested;
        private bool _codexProfilesLoadRequested;
        private bool _systemOptimizationLoadRequested;
        private bool _weChatStartupLoadRequested;
        private bool _frpConfigLoadRequested;
        private static readonly IReadOnlyList<SqlProviderOption> SqlProviderOptionItems = new List<SqlProviderOption>
        {
            new SqlProviderOption(SqlProviderKind.SqlServer, "SQL Server"),
            new SqlProviderOption(SqlProviderKind.PostgreSql, "PostgreSQL"),
            new SqlProviderOption(SqlProviderKind.MySql, "MySQL")
        };

        private RecordingService Recording => _recordingService ?? (_recordingService = new RecordingService());
        private OptimizationReportService OptimizationReportsStore => _optimizationReportService ?? (_optimizationReportService = new OptimizationReportService());
        private SystemOptimizationService SystemOptimizer => _systemOptimizationService ?? (_systemOptimizationService = new SystemOptimizationService());
        private JunkCleanupService JunkCleaner => _junkCleanupService ?? (_junkCleanupService = new JunkCleanupService());
        private WeChatDataLocator WeChatLocator => _weChatDataLocator ?? (_weChatDataLocator = new WeChatDataLocator());
        private WeChatCleanupService WeChatCleaner => _weChatCleanupService ?? (_weChatCleanupService = new WeChatCleanupService());
        private WeChatBackupService WeChatBackupStore => _weChatBackupService ?? (_weChatBackupService = new WeChatBackupService());

        public MainViewModel()
        {
            var constructorStopwatch = Stopwatch.StartNew();
            NetworkList = new ObservableCollection<NetworkData>();
            StartupList = new ObservableCollection<StartupItem>();
            FilteredStartupView = CollectionViewSource.GetDefaultView(StartupList);
            FilteredStartupView.Filter = FilterStartupItem;
            SqlDatabaseList = new ObservableCollection<DatabaseItem>();
            SqlTableList = new ObservableCollection<TableItem>();
            AllSqlTableList = new ObservableCollection<TableItem>();
            InstalledPrograms = new ObservableCollection<InstalledProgram>();
            FilteredInstalledProgramsView = CollectionViewSource.GetDefaultView(InstalledPrograms);
            FilteredInstalledProgramsView.Filter = FilterInstalledProgram;
            ApplyInstalledProgramSort();
            SqlServerAddressHistory = new ObservableCollection<string>();
            SqlUsernameHistory = new ObservableCollection<string>();
            SqlPasswordHistory = new ObservableCollection<string>();
            SqlRecentConnections = new ObservableCollection<SqlConnectionHistoryItem>();
            CodexProfiles = new ObservableCollection<CodexProfileItem>();
            HomeCommandItems = new ObservableCollection<HomeCommandItem>();
            FilteredHomeCommandView = CollectionViewSource.GetDefaultView(HomeCommandItems);
            FilteredHomeCommandView.Filter = FilterHomeCommandItem;
            HomeRecentItems = new ObservableCollection<HomeRecentItem>();
            ScreenshotHistoryItems = new ObservableCollection<ScreenshotHistoryItem>();
            FilteredVideoViewerPlaylistView = CollectionViewSource.GetDefaultView(VideoViewerPlaylist);
            FilteredVideoViewerPlaylistView.Filter = FilterVideoViewerPlaylistItem;
            VideoViewerRecentPlaylists = new ObservableCollection<RecentPlaylistItem>();
            VideoViewerFavoritePlaylists = new ObservableCollection<FavoritePlaylistItem>();
            OptimizationReports = new ObservableCollection<OptimizationReportItem>();
            JunkCandidates = new ObservableCollection<JunkCandidate>();
            WeChatCleanupCandidates = new ObservableCollection<WeChatCleanupCandidate>();
            WeChatRoots = new ObservableCollection<WeChatRoot>();
            RecentWeChatBackups = new ObservableCollection<RecentWeChatBackupItem>();
            FilteredSqlTableView = CollectionViewSource.GetDefaultView(SqlTableList);
            FilteredSqlTableView.Filter = FilterSqlTable;
            _audioRecordingTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _audioRecordingTimer.Tick += AudioRecordingTimer_OnTick;

            RefreshCommand = new RelayCommand(Refresh);
            ShowHomeCommand = new RelayCommand(() => SwitchModule("Home"));
            ShowNetworkCommand = new RelayCommand(() => SwitchModule("Network"));
            ShowStartupCommand = new RelayCommand(() => SwitchModule("Startup"));
            ShowSystemCommand = new RelayCommand(() => SwitchModule("Optimization"));
            ShowUninstallCommand = new RelayCommand(() => SwitchModule("Uninstall"));
            ShowSqlExportCommand = new RelayCommand(() => { SwitchModule("SqlExport"); Refresh(); });
            ShowFrpCommand = new RelayCommand(() => SwitchModule("Frp"));
            ShowCodexProfilesCommand = new RelayCommand(() => SwitchModule("CodexProfiles"));
            ShowSystemInfoCommand = new RelayCommand(() => SwitchModule("SystemInfo"));
            ShowFileVerifyCommand = new RelayCommand(() => SwitchModule("FileVerify"));
            ShowMultimediaCommand = new RelayCommand(() => ShowMultimedia(MultimediaPreferredFilter.All));
            ShowConvertCommand = new RelayCommand(() => ShowMultimedia(MultimediaPreferredFilter.All));
            ShowImageViewerCommand = new RelayCommand(() => ShowMultimedia(MultimediaPreferredFilter.Image));
            ShowVideoViewerCommand = new RelayCommand(() => ShowMultimedia(MultimediaPreferredFilter.AudioVideo));
            ShowBenchmarkCommand = new RelayCommand(() => SwitchModule("Benchmark"));
            ShowScheduleCommand = new RelayCommand(() => SwitchModule("Schedule"));
            ShowSystemSettingsCommand = new RelayCommand(() => SwitchModule("SystemSettings"));
            LoadSystemInfoCommand = new AsyncRelayCommand(LoadSystemInfoAsync, () => !_isSystemInfoBusy);
            ToggleHardwareSensorsCommand = new RelayCommand(ToggleHardwareSensors);
            RefreshSensorsOnceCommand = new RelayCommand(() => SensorTimer_OnTick(this, EventArgs.Empty), () => _sensorService != null);
            RestartAsAdminCommand = new RelayCommand(RestartAsAdmin);
            VerifyFileCommand = new AsyncRelayCommand(VerifyFileAsync, () => !_isFileHashBusy);
            ConvertImageCommand = new AsyncRelayCommand(ConvertImageAsync, () => !_isConvertBusy);
            ConvertMediaCommand = new AsyncRelayCommand(ConvertMediaAsync, () => !_isConvertBusy);
            CancelConvertCommand = new RelayCommand(CancelConvert, () => IsConvertBusy);
            OpenConvertOutputFolderCommand = new RelayCommand(OpenConvertOutputFolder, () => HasConvertOutputTarget);
            SelectConvertOutputFolderCommand = new RelayCommand(SelectConvertOutputFolder);
            ClearConvertQueueCommand = new RelayCommand(ClearConvertQueue, () => HasConvertQueueItems && !IsConvertBusy);
            OpenConvertQueueOutputCommand = new RelayParameterCommand(OpenConvertQueueOutput, CanOpenConvertQueueOutput);
            CopyConvertQueueMessageCommand = new RelayParameterCommand(CopyConvertQueueMessage, CanCopyConvertQueueMessage);
            RetryConvertQueueItemCommand = new AsyncRelayParameterCommand(RetryConvertQueueItemAsync, CanRetryConvertQueueItem);
            RemoveConvertQueueItemCommand = new RelayParameterCommand(RemoveConvertQueueItem, parameter => !IsConvertBusy && parameter is ConvertQueueItem);
            MoveConvertQueueItemUpCommand = new RelayParameterCommand(parameter => MoveConvertQueueItem(parameter, -1), CanMoveConvertQueueItemUp);
            MoveConvertQueueItemDownCommand = new RelayParameterCommand(parameter => MoveConvertQueueItem(parameter, 1), CanMoveConvertQueueItemDown);
            ToggleConvertPauseCommand = new RelayCommand(ToggleConvertPause, () => IsConvertBusy);
            ApplyImageCompressPresetCommand = new RelayCommand(ApplyImageCompressPreset);
            ApplyImageAvatarPresetCommand = new RelayCommand(ApplyImageAvatarPreset);
            ApplyMediaMp4PresetCommand = new RelayCommand(ApplyMediaMp4Preset);
            ApplyMediaMp3PresetCommand = new RelayCommand(ApplyMediaMp3Preset);
            RunAllBenchmarksCommand = new AsyncRelayCommand(RunAllBenchmarksAsync, () => !_isBenchmarkBusy);
            RunSingleBenchmarkCommand = new AsyncRelayParameterCommand(RunSingleBenchmarkAsync);
            _copyBenchmarkResultsCommand = new RelayCommand(CopyBenchmarkResults, () => HasBenchmarkResults);
            _exportBenchmarkResultsCommand = new RelayCommand(ExportBenchmarkResults, () => HasBenchmarkResults);
            CopyBenchmarkResultsCommand = _copyBenchmarkResultsCommand;
            ExportBenchmarkResultsCommand = _exportBenchmarkResultsCommand;
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
            LockWin10Command = new RelayCommand(LockCurrentWindowsVersion);
            ExitCommand = new RelayCommand(ExitApplication);
            RestoreCommand = new RelayCommand(RestoreWindow);
            OpenLogFolderCommand = new RelayCommand(OpenLogFolder);
            ToggleAutoStartCommand = new RelayCommand(ToggleAutoStart);
            _isAutoStartEnabled = false;

            _toggleDefenderCommand = new AsyncRelayCommand(ToggleDefenderAsync);
            _toggleAutoUpdateCommand = new AsyncRelayCommand(ToggleAutoUpdateAsync);
            _triggerUpdateNowCommand = new AsyncRelayCommand(TriggerUpdateNowAsync);
            ToggleDefenderCommand = _toggleDefenderCommand;
            ToggleAutoUpdateCommand = _toggleAutoUpdateCommand;
            TriggerUpdateNowCommand = _triggerUpdateNowCommand;
            RefreshSystemStatusCommand = new RelayCommand(RefreshSystemStatus);
            _applyAlwaysOnPowerPolicyCommand = new AsyncRelayCommand(ApplyAlwaysOnPowerPolicyAsync, () => !IsPowerPolicyBusy);
            ApplyAlwaysOnPowerPolicyCommand = _applyAlwaysOnPowerPolicyCommand;
            _startAutoOptimizeCommand = new AsyncRelayCommand(StartAutoOptimizeAsync, () => !IsAutoOptimizeBusy);
            _startJunkScanCommand = new AsyncRelayCommand(StartJunkScanAsync, () => !IsJunkBusy);
            _runJunkCleanupCommand = new AsyncRelayCommand(RunJunkCleanupAsync, CanRunJunkCleanup);
            _exportJunkCleanupPlanCommand = new RelayCommand(ExportJunkCleanupPlan, CanExportJunkCleanupPlan);
            _scanWeChatCleanupCommand = new AsyncRelayCommand(ScanWeChatCleanupAsync, () => !IsWeChatCleanupBusy);
            _startWeChatCleanupCommand = new AsyncRelayCommand(StartWeChatCleanupAsync, CanStartWeChatCleanup);
            _startWeChatBackupCommand = new AsyncRelayCommand(StartWeChatBackupAsync, CanStartWeChatBackup);
            _startWeChatRestoreCommand = new AsyncRelayCommand(StartWeChatRestoreAsync, CanStartWeChatRestore);
            _deleteSelectedReportsCommand = new RelayCommand(DeleteSelectedReports, CanDeleteSelectedReports);
            _selectWeChatBackupOutputFolderCommand = new RelayCommand(SelectWeChatBackupOutputFolder, () => !IsWeChatBackupBusy);
            _selectWeChatRestoreZipCommand = new RelayCommand(SelectWeChatRestoreZip, () => !IsWeChatRestoreBusy);
            _selectWeChatRestoreTargetRootCommand = new RelayCommand(SelectWeChatRestoreTargetRoot, () => !IsWeChatRestoreBusy);
            OpenRecentWeChatBackupCommand = new RelayParameterCommand(OpenRecentWeChatBackup, parameter => !IsWeChatRestoreBusy && parameter is RecentWeChatBackupItem);
            _showReportDetailsCommand = new RelayParameterCommand(ShowReportDetails);
            StartAutoOptimizeCommand = _startAutoOptimizeCommand;
            StartJunkScanCommand = _startJunkScanCommand;
            RunJunkCleanupCommand = _runJunkCleanupCommand;
            ExportJunkCleanupPlanCommand = _exportJunkCleanupPlanCommand;
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
            _exportQueryResultCsvCommand = new AsyncRelayCommand(ExportQueryResultCsvAsync, CanExportQueryResult);
            _cancelQueryCommand = new RelayCommand(CancelSqlQuery, () => IsQueryBusy);
            ExecuteSqlQueryCommand = _executeQueryCommand;
            ExportQueryResultCommand = _exportQueryResultCommand;
            ExportQueryResultCsvCommand = _exportQueryResultCsvCommand;
            CancelSqlQueryCommand = _cancelQueryCommand;

            _testSqlConnectionCommand = new AsyncRelayCommand(TestSqlConnectionAsync, () => !IsSqlBusy);
            _exportSqlTableCommand = new AsyncRelayCommand(ExportSelectedTableAsync, CanExportSqlTable);
            CancelSqlExportCommand = new RelayCommand(CancelSqlExport, () => IsSqlBusy && _sqlExportCancellationTokenSource != null);
            TestSqlConnectionCommand = _testSqlConnectionCommand;
            ExportSqlTableCommand = _exportSqlTableCommand;
            ApplySqlRecentConnectionCommand = new RelayParameterCommand(ApplySqlRecentConnection, parameter => !IsSqlBusy && parameter is SqlConnectionHistoryItem);
            _refreshInstalledProgramsCommand = new AsyncRelayCommand(LoadInstalledProgramsAsync, () => !IsInstalledProgramsBusy);
            _uninstallProgramCommand = new AsyncRelayParameterCommand(UninstallProgramAsync, parameter => parameter is InstalledProgram);
            RefreshInstalledProgramsCommand = _refreshInstalledProgramsCommand;
            UninstallProgramCommand = _uninstallProgramCommand;
            SelectFilteredInstalledProgramsCommand = new RelayCommand(SelectFilteredInstalledPrograms, () => FilteredInstalledProgramsView != null && !FilteredInstalledProgramsView.IsEmpty);
            ClearInstalledProgramSelectionCommand = new RelayCommand(ClearInstalledProgramSelection, () => SelectedInstalledProgramsCount > 0);
            ExportInstalledProgramListCommand = new RelayCommand(ExportInstalledProgramList, () => GetInstalledProgramsForExport().Count > 0);

            ShowScreenshotCommand = new RelayCommand(() => SwitchModule("Screenshot"));
            TakeScreenshotNowCommand = new AsyncRelayCommand(TriggerScreenshotAsync);
            _openRecordRegionCommand = new AsyncRelayCommand(OpenRecordRegionAsync, () => !IsVideoRecording && !IsAudioRecording);
            StartVideoRecordingCommand = _openRecordRegionCommand;
            _toggleAudioRecordingCommand = new AsyncRelayCommand(ToggleAudioRecordingAsync, () => !IsVideoRecording);
            ToggleAudioRecordingCommand = _toggleAudioRecordingCommand;
            SelectRecordingOutputFolderCommand = new AsyncRelayCommand(SelectRecordingOutputFolderAsync);
            SelectAudioOutputFolderCommand = new AsyncRelayCommand(SelectAudioOutputFolderAsync);
            StartCaptureHotkeyCommand = new RelayCommand(() => IsCapturingHotkey = true);
            CancelCaptureHotkeyCommand = new RelayCommand(() => IsCapturingHotkey = false);
            SaveScreenshotSettingsCommand = new AsyncRelayCommand(SaveScreenshotSettingsAsync);
            StartCaptureVideoRecordHotkeyCommand = new RelayCommand(() => IsCapturingVideoRecordHotkey = true);
            CancelCaptureVideoRecordHotkeyCommand = new RelayCommand(() => IsCapturingVideoRecordHotkey = false);
            StartCaptureAudioRecordHotkeyCommand = new RelayCommand(() => IsCapturingAudioRecordHotkey = true);
            CancelCaptureAudioRecordHotkeyCommand = new RelayCommand(() => IsCapturingAudioRecordHotkey = false);
            SaveRecordingHotkeySettingsCommand = new AsyncRelayCommand(SaveRecordingHotkeySettingsAsync);
            EditClipboardImageCommand = new RelayCommand(EditClipboardImage);
            OcrClipboardImageCommand = new AsyncRelayCommand(OcrClipboardImageAsync, () => OcrService.IsSupported);
            ComputeFileHashCommand = new AsyncRelayCommand(ComputeFileHashAsync, () => !_isFileHashBusy);
            ExportFileHashListCommand = new RelayCommand(ExportFileHashList, () => HasBatchFileHashResults);
            ImportFileHashListCommand = new RelayCommand(ImportFileHashList, () => !_isFileHashBusy);
            CopyBatchFileHashColumnCommand = new RelayParameterCommand(CopyBatchFileHashColumn, parameter => HasBatchFileHashResults && GetHashKindFromParameter(parameter) != null);
            OpenImageViewerFileCommand = new RelayCommand(OpenImageViewerFile);
            CopyImageViewerImageCommand = new RelayCommand(CopyImageViewerImage, () => HasImageViewerImage);
            SaveImageViewerImageAsCommand = new RelayCommand(SaveImageViewerImageAs, () => HasImageViewerImage);
            EditImageViewerImageCommand = new RelayCommand(EditImageViewerImage, () => HasImageViewerImage);
            OpenPreviousImageViewerFileCommand = new RelayCommand(() => OpenImageViewerSibling(-1), () => CanOpenImageViewerSibling(-1));
            OpenNextImageViewerFileCommand = new RelayCommand(() => OpenImageViewerSibling(1), () => CanOpenImageViewerSibling(1));
            ImageViewerZoomInCommand = new RelayCommand(() => ImageViewerZoom = ImageViewerZoom + 0.1, () => HasImageViewerImage);
            ImageViewerZoomOutCommand = new RelayCommand(() => ImageViewerZoom = ImageViewerZoom - 0.1, () => HasImageViewerImage);
            ImageViewerZoomActualCommand = new RelayCommand(() => ImageViewerZoom = 1.0, () => HasImageViewerImage);
            RotateImageViewerLeftCommand = new RelayCommand(() => RotateImageViewer(-90), () => HasImageViewerImage);
            RotateImageViewerRightCommand = new RelayCommand(() => RotateImageViewer(90), () => HasImageViewerImage);
            FlipImageViewerHorizontalCommand = new RelayCommand(() => { _imageViewerFlipHorizontal = !_imageViewerFlipHorizontal; UpdateImageViewerPreview("已水平翻转。"); }, () => HasImageViewerImage);
            FlipImageViewerVerticalCommand = new RelayCommand(() => { _imageViewerFlipVertical = !_imageViewerFlipVertical; UpdateImageViewerPreview("已垂直翻转。"); }, () => HasImageViewerImage);
            ToggleImageViewerGrayscaleCommand = new RelayCommand(() => { ImageViewerIsGrayscale = !ImageViewerIsGrayscale; UpdateImageViewerPreview(ImageViewerIsGrayscale ? "已应用灰度效果。" : "已取消灰度效果。"); }, () => HasImageViewerImage);
            ResetImageViewerAdjustmentsCommand = new RelayCommand(ResetImageViewerAdjustments, () => HasImageViewerImage);
            ImageViewerFitToFrameCommand = new RelayCommand(ToggleImageViewerFitToFrame, () => HasImageViewerImage);
            OpenVideoViewerFileCommand = new RelayCommand(OpenVideoViewerFile);
            OpenVideoViewerInExternalPlayerCommand = new RelayCommand(OpenVideoViewerInExternalPlayer, () => HasVideoViewerVideo);
            CaptureVideoViewerFrameCommand = new AsyncRelayCommand(CaptureVideoViewerFrameAsync, () => HasVideoViewerVideo && !_isVideoViewerCapturingFrame);
            GenerateVideoViewerWaveformCommand = new AsyncRelayCommand(GenerateVideoViewerWaveformAsync, () => HasVideoViewerVideo && !_isVideoViewerGeneratingWaveform);
            ToggleVideoViewerMuteCommand = new RelayCommand(() => VideoViewerIsMuted = !VideoViewerIsMuted, () => HasVideoViewerVideo);
            OpenPreviousVideoViewerPlaylistItemCommand = new RelayCommand(() => OpenVideoViewerPlaylistSibling(-1), () => CanOpenVideoViewerPlaylistSibling(-1));
            OpenNextVideoViewerPlaylistItemCommand = new RelayCommand(() => OpenVideoViewerPlaylistSibling(1), () => CanOpenVideoViewerPlaylistSibling(1));
            OpenVideoViewerPlaylistItemCommand = new RelayParameterCommand(OpenVideoViewerPlaylistItem, item => item is VideoPlaylistItem);
            RemoveVideoViewerPlaylistItemCommand = new RelayParameterCommand(RemoveVideoViewerPlaylistItem, item => item is VideoPlaylistItem && _videoViewerPlaylist != null && _videoViewerPlaylist.Count > 0);
            CopyVideoViewerPlaylistCommand = new RelayCommand(CopyVideoViewerPlaylist, () => _videoViewerPlaylist != null && _videoViewerPlaylist.Count > 0);
            CleanInvalidVideoViewerPlaylistItemsCommand = new RelayCommand(CleanInvalidVideoViewerPlaylistItems, () => _videoViewerPlaylist != null && _videoViewerPlaylist.Count > 0);
            ClearVideoViewerPlaylistCommand = new RelayCommand(ClearVideoViewerPlaylist, () => _videoViewerPlaylist != null && _videoViewerPlaylist.Count > 0);
            SaveVideoViewerPlaylistCommand = new RelayCommand(SaveVideoViewerPlaylist, () => _videoViewerPlaylist != null && _videoViewerPlaylist.Count > 0);
            LoadVideoViewerPlaylistCommand = new RelayCommand(LoadVideoViewerPlaylistFromFile);
            OpenRecentVideoViewerPlaylistCommand = new RelayParameterCommand(OpenRecentVideoViewerPlaylist, item => item is RecentPlaylistItem);
            FavoriteVideoViewerPlaylistCommand = new RelayCommand(FavoriteVideoViewerPlaylist, () => _videoViewerPlaylist != null && _videoViewerPlaylist.Count > 0);
            OpenFavoriteVideoViewerPlaylistCommand = new RelayParameterCommand(OpenFavoriteVideoViewerPlaylist, item => item is FavoritePlaylistItem);
            RemoveFavoriteVideoViewerPlaylistCommand = new RelayParameterCommand(RemoveFavoriteVideoViewerPlaylist, item => item is FavoritePlaylistItem);
            SetVideoViewerLoopStartCommand = new RelayCommand(() => SetVideoViewerLoopPoint(true), () => HasVideoViewerVideo);
            SetVideoViewerLoopEndCommand = new RelayCommand(() => SetVideoViewerLoopPoint(false), () => HasVideoViewerVideo);
            ClearVideoViewerLoopCommand = new RelayCommand(ClearVideoViewerLoop, () => HasVideoViewerLoopRange || VideoViewerIsLoopEnabled);
            ToggleVideoViewerLoopCommand = new RelayCommand(() => VideoViewerIsLoopEnabled = !VideoViewerIsLoopEnabled, () => HasVideoViewerLoopRange);
            SaveVideoViewerLoopRangeCommand = new RelayCommand(SaveVideoViewerLoopRange, () => HasVideoViewerLoopRange);
            OpenVideoViewerLoopRangeCommand = new RelayParameterCommand(OpenVideoViewerLoopRange, item => item is VideoLoopRangeItem);
            RemoveVideoViewerLoopRangeCommand = new RelayParameterCommand(RemoveVideoViewerLoopRange, item => item is VideoLoopRangeItem);
            AddVideoViewerBookmarkCommand = new RelayCommand(AddVideoViewerBookmark, () => HasVideoViewerVideo);
            OpenVideoViewerBookmarkCommand = new RelayParameterCommand(OpenVideoViewerBookmark, item => item is VideoBookmarkItem);
            RemoveVideoViewerBookmarkCommand = new RelayParameterCommand(RemoveVideoViewerBookmark, item => item is VideoBookmarkItem);
            LoadVideoViewerSubtitleCommand = new AsyncRelayCommand(LoadVideoViewerSubtitleAsync, () => HasVideoViewerVideo);
            ClearVideoViewerSubtitleCommand = new RelayCommand(() => ClearVideoViewerSubtitle("已清除字幕。"), () => HasVideoViewerSubtitle);
            CopyScreenshotHistoryItemCommand = new RelayParameterCommand(CopyScreenshotHistoryItem, CanUseScreenshotHistoryItem);
            EditScreenshotHistoryItemCommand = new RelayParameterCommand(EditScreenshotHistoryItem, CanUseScreenshotHistoryItem);
            OpenScreenshotHistoryItemCommand = new RelayParameterCommand(OpenScreenshotHistoryItem, CanUseScreenshotHistoryItem);
            DeleteScreenshotHistoryItemCommand = new RelayParameterCommand(DeleteScreenshotHistoryItem, parameter => parameter is ScreenshotHistoryItem);
            ExecuteHomeCommandItemCommand = new RelayParameterCommand(ExecuteHomeCommandItem, parameter => parameter is HomeCommandItem);
            OpenHomeRecentItemCommand = new RelayParameterCommand(OpenHomeRecentItem, parameter => parameter is HomeRecentItem);
            InitializeHomeCommandItems();

            _applyCodexProfileCommand = new AsyncRelayParameterCommand(ApplyCodexProfileAsync, parameter => parameter is CodexProfileItem);
            ApplyCodexProfileCommand = _applyCodexProfileCommand;
            SwitchCodexProfileCommand = _applyCodexProfileCommand;
            _exportCodexProfileCommand = new AsyncRelayParameterCommand(ExportCodexProfileAsync, parameter => parameter is CodexProfileItem);
            ExportCodexProfileCommand = _exportCodexProfileCommand;
            _previewCodexProfileDiffCommand = new AsyncRelayParameterCommand(PreviewCodexProfileDiffAsync, parameter => parameter is CodexProfileItem);
            PreviewCodexProfileDiffCommand = _previewCodexProfileDiffCommand;
            _importCodexProfileCommand = new AsyncRelayCommand(ImportCodexProfileAsync);
            ImportCodexProfileCommand = _importCodexProfileCommand;
            _importCodexCpaTokenCommand = new AsyncRelayCommand(ImportCodexCpaTokenAsync);
            ImportCodexCpaTokenCommand = _importCodexCpaTokenCommand;
            _refreshCodexProfileCommand = new AsyncRelayParameterCommand(RefreshCodexProfileAsync, parameter => parameter is CodexProfileItem);
            RefreshCodexProfileCommand = _refreshCodexProfileCommand;
            _renameCodexProfileCommand = new AsyncRelayParameterCommand(RenameCodexProfileAsync, parameter => parameter is CodexProfileItem);
            RenameCodexProfileCommand = _renameCodexProfileCommand;
            _editCodexProfileNoteCommand = new AsyncRelayParameterCommand(EditCodexProfileNoteAsync, parameter => parameter is CodexProfileItem);
            EditCodexProfileNoteCommand = _editCodexProfileNoteCommand;
            _restoreLastCodexBackupCommand = new AsyncRelayCommand(RestoreLastCodexBackupAsync);
            RestoreLastCodexBackupCommand = _restoreLastCodexBackupCommand;
            _exportCodexProfilesEncBoxCommand = new AsyncRelayCommand(ExportCodexProfilesEncBoxAsync);
            ExportCodexProfilesEncBoxCommand = _exportCodexProfilesEncBoxCommand;
            _importCodexProfilesEncBoxCommand = new AsyncRelayCommand(ImportCodexProfilesEncBoxAsync);
            ImportCodexProfilesEncBoxCommand = _importCodexProfilesEncBoxCommand;
            DeleteCodexProfileCommand = new AsyncRelayParameterCommand(DeleteCodexProfileAsync, parameter => parameter is CodexProfileItem);
            EditCodexConfigTomlCommand = new AsyncRelayParameterCommand(p => EditCodexFileAsync(p, CodexConfigProfileService.ConfigFileName), parameter => parameter is CodexProfileItem);
            EditCodexAuthJsonCommand = new AsyncRelayParameterCommand(p => EditCodexFileAsync(p, CodexConfigProfileService.AuthFileName), parameter => parameter is CodexProfileItem);
            _rotateToNextCodexProfileCommand = new AsyncRelayCommand(RotateToNextCodexProfileAsync, () => IsCodexRotationAvailable);
            RotateToNextCodexProfileCommand = _rotateToNextCodexProfileCommand;
            _restartCodexDesktopCommand = new AsyncRelayCommand(RestartCodexDesktopAsync);
            RestartCodexDesktopCommand = _restartCodexDesktopCommand;
            _toggleCodexProfileRotationCommand = new AsyncRelayParameterCommand(ToggleCodexProfileRotationAsync, parameter => parameter is CodexProfileItem);
            ToggleCodexProfileRotationCommand = _toggleCodexProfileRotationCommand;

            CurrentModule = "Home";
            ScheduleStartupBackgroundLoads();
            constructorStopwatch.Stop();
            AppLogService.InformationIfInitialized("MainViewModel constructed in {ElapsedMs} ms", constructorStopwatch.ElapsedMilliseconds);
        }

        private void ScheduleStartupBackgroundLoads()
        {
            var dispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;
            dispatcher.BeginInvoke(
                DispatcherPriority.ContextIdle,
                new Action(() => SafeFireAndForget(LoadStartupShellStateAsync())));
            dispatcher.BeginInvoke(
                DispatcherPriority.ContextIdle,
                new Action(() => StartSensorsBackground()));
        }

        private void StartSensorsBackground()
        {
            _sensorService = new HardwareSensorService();
            if (!_sensorService.TryStart())
            {
                SensorStatusMessage = "传感器启动失败：" + (_sensorService.LastError ?? "未知错误。请以管理员身份重新运行 MyTools。");
                HomeSensorRiskText = "传感器启动失败，无法读取温度风险。";
                _sensorService.Dispose();
                _sensorService = null;
                return;
            }

            SensorTimer.Tick += SensorTimer_OnTick;
            ApplySensorRefreshMode();
            SensorTimer_OnTick(this, EventArgs.Empty);
            _isSensorsRunning = true;
            OnPropertyChanged(nameof(IsSensorsRunning));

            if (SensorReadings.Count == 0)
            {
                SensorStatusMessage = HardwareSensorService.IsRunningAsAdmin
                    ? "未读取到传感器数据，硬件可能不被支持。"
                    : "当前为非管理员模式：温度/风扇/电压通常无法读取。请以管理员身份重启 MyTools。";
                HomeSensorRiskText = HardwareSensorService.IsRunningAsAdmin
                    ? "未读取到传感器数据，硬件可能不支持。"
                    : "未读到传感器：建议以管理员身份重启后查看温度风险。";
            }
            else
            {
                SensorStatusMessage = HardwareSensorService.IsRunningAsAdmin
                    ? $"已启用 · {SensorReadings.Count} 项 · {SensorRefreshMode}刷新"
                    : $"已启用（非管理员，仅 {SensorReadings.Count} 项可读）";
                HomeSensorRiskText = $"传感器正常：已读取 {SensorReadings.Count} 项，未发现高温或高负载。";
            }
        }

        private async Task LoadStartupShellStateAsync()
        {
            try
            {
                IsAutoStartEnabled = await Task.Run(() => ReadAutoStartStatus()).ConfigureAwait(true);
            }
            catch
            {
                IsAutoStartEnabled = false;
            }
        }

        public ObservableCollection<NetworkData> NetworkList
        {
            get => _networkList;
            set { _networkList = value; OnPropertyChanged(); }
        }

        public ObservableCollection<StartupItem> StartupList
        {
            get => _startupList;
            set
            {
                _startupList = value;
                FilteredStartupView = CollectionViewSource.GetDefaultView(_startupList);
                FilteredStartupView.Filter = FilterStartupItem;
                OnPropertyChanged();
            }
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

        public ObservableCollection<InstalledProgram> InstalledPrograms
        {
            get => _installedPrograms;
            set
            {
                _installedPrograms = value;
                FilteredInstalledProgramsView = CollectionViewSource.GetDefaultView(_installedPrograms);
                FilteredInstalledProgramsView.Filter = FilterInstalledProgram;
                ApplyInstalledProgramSort();
                AttachInstalledProgramSelectionHandlers();
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasNoInstalledPrograms));
                OnPropertyChanged(nameof(InstalledProgramsCountText));
                OnPropertyChanged(nameof(SelectedInstalledProgramsCountText));
            }
        }

        public ICollectionView FilteredStartupView
        {
            get => _filteredStartupView;
            set { _filteredStartupView = value; OnPropertyChanged(); }
        }

        public ICollectionView FilteredInstalledProgramsView
        {
            get => _filteredInstalledProgramsView;
            set { _filteredInstalledProgramsView = value; OnPropertyChanged(); }
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

        public ObservableCollection<SqlConnectionHistoryItem> SqlRecentConnections
        {
            get => _sqlRecentConnections;
            set
            {
                _sqlRecentConnections = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasSqlRecentConnections));
            }
        }

        public bool HasSqlRecentConnections => SqlRecentConnections.Count > 0;

        public ObservableCollection<CodexProfileItem> CodexProfiles
        {
            get => _codexProfiles;
            set { _codexProfiles = value; OnPropertyChanged(); }
        }

        public ObservableCollection<HomeCommandItem> HomeCommandItems
        {
            get => _homeCommandItems;
            set
            {
                _homeCommandItems = value;
                OnPropertyChanged();
            }
        }

        public ICollectionView FilteredHomeCommandView
        {
            get => _filteredHomeCommandView;
            private set { _filteredHomeCommandView = value; OnPropertyChanged(); }
        }

        public ICollectionView FilteredVideoViewerPlaylistView
        {
            get => _filteredVideoViewerPlaylistView;
            private set { _filteredVideoViewerPlaylistView = value; OnPropertyChanged(); }
        }

        public ObservableCollection<HomeRecentItem> HomeRecentItems
        {
            get => _homeRecentItems;
            set
            {
                _homeRecentItems = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasHomeRecentItems));
            }
        }

        public bool HasHomeRecentItems => HomeRecentItems != null && HomeRecentItems.Count > 0;

        public ObservableCollection<ScreenshotHistoryItem> ScreenshotHistoryItems
        {
            get => _screenshotHistoryItems;
            set
            {
                _screenshotHistoryItems = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasScreenshotHistoryItems));
                OnPropertyChanged(nameof(ScreenshotHistorySummary));
            }
        }

        public bool HasScreenshotHistoryItems => ScreenshotHistoryItems != null && ScreenshotHistoryItems.Count > 0;

        public string ScreenshotHistorySummary => HasScreenshotHistoryItems
            ? $"最近 {ScreenshotHistoryItems.Count} 张截图"
            : "暂无截图历史";

        public string HomeCommandSearchText
        {
            get => _homeCommandSearchText;
            set
            {
                if (string.Equals(_homeCommandSearchText, value, StringComparison.Ordinal))
                {
                    return;
                }

                _homeCommandSearchText = value ?? string.Empty;
                OnPropertyChanged();
                FilteredHomeCommandView?.Refresh();
            }
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

        public ObservableCollection<RecentWeChatBackupItem> RecentWeChatBackups
        {
            get => _recentWeChatBackups;
            set
            {
                _recentWeChatBackups = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasRecentWeChatBackups));
            }
        }

        public bool HasRecentWeChatBackups => RecentWeChatBackups != null && RecentWeChatBackups.Count > 0;

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

        public string CurrentModule
        {
            get => _currentModule;
            set
            {
                if (_currentModule == value) return;
                var prev = _currentModule;
                _currentModule = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CurrentNavigationText));
                NotifySystemSectionVisibilityChanged();

                // Auto-pause sensor polling when leaving SystemInfo page to save CPU
                if (prev == "SystemInfo" && _isSensorsRunning && SensorTimer.IsEnabled)
                {
                    SensorTimer.Stop();
                }
                else if (value == "SystemInfo" && _isSensorsRunning && !SensorTimer.IsEnabled)
                {
                    SensorTimer.Start();
                }
                else if (prev == "System" && value != "System" && string.Equals(CurrentSystemSection, "SystemInfo", StringComparison.Ordinal) && _isSensorsRunning && SensorTimer.IsEnabled)
                {
                    SensorTimer.Stop();
                }
            }
        }

        public string CurrentSystemSection
        {
            get => _currentSystemSection;
            private set
            {
                if (string.Equals(_currentSystemSection, value, StringComparison.Ordinal)) return;
                var previous = _currentSystemSection;
                _currentSystemSection = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(CurrentNavigationText));
                NotifySystemSectionVisibilityChanged();
                if (string.Equals(previous, "SystemInfo", StringComparison.Ordinal) && _isSensorsRunning && SensorTimer.IsEnabled)
                {
                    SensorTimer.Stop();
                }
                else if (string.Equals(value, "SystemInfo", StringComparison.Ordinal) && _isSensorsRunning && !SensorTimer.IsEnabled)
                {
                    SensorTimer.Start();
                }
            }
        }

        public int SelectedSystemSectionIndex
        {
            get => _selectedSystemSectionIndex;
            set
            {
                if (value < 0 || value >= SystemSectionKeys.Length) return;
                if (_selectedSystemSectionIndex == value && string.Equals(CurrentModule, "System", StringComparison.Ordinal)) return;
                ShowSystemSection(SystemSectionKeys[value]);
            }
        }

        public bool IsSystemOptimizationVisible => IsSystemSectionVisible("Optimization");
        public bool IsSystemNetworkVisible => IsSystemSectionVisible("Network");
        public bool IsSystemStartupVisible => IsSystemSectionVisible("Startup");
        public bool IsSystemUninstallVisible => IsSystemSectionVisible("Uninstall");
        public bool IsSystemInfoVisible => IsSystemSectionVisible("SystemInfo");
        public bool IsSystemBenchmarkVisible => IsSystemSectionVisible("Benchmark");
        public bool IsSystemSettingsVisible => IsSystemSectionVisible("SystemSettings");
        public string AppVersionText => BuildAppVersionText();
        public string CurrentNavigationText => BuildCurrentNavigationText();

        public bool IsWindows10OrGreater => OsVersionService.IsWindows10OrGreater;
        public bool IsWindows11OrGreater => OsVersionService.IsWindows11OrGreater;
        public bool IsOcrSupported => OcrService.IsSupported;
        public string OsDisplayName => OsVersionService.DisplayName;
        private SystemInfoSnapshot _systemInfoSnapshot;
        public SystemInfoSnapshot SystemInfo
        {
            get => _systemInfoSnapshot;
            private set { _systemInfoSnapshot = value; OnPropertyChanged(); }
        }

        private async Task LoadSystemInfoSnapshotAsync()
        {
            try { SystemInfo = await SystemInfoService.GetSnapshotAsync().ConfigureAwait(false); }
            catch (Exception ex) { AppLogService.Warning("SystemInfoSnapshot load failed: {Msg}", ex.Message); }
        }

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
                SafeFireAndForget(LoadSqlConnectionHistoryAsync());
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
                    SafeFireAndForget(LoadTablesForSelectedDatabaseAsync());
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
        public ICommand ExecuteHomeCommandItemCommand { get; }
        public ICommand OpenHomeRecentItemCommand { get; }
        public ICommand ShowHomeCommand { get; }
        public ICommand ShowNetworkCommand { get; }
        public ICommand ShowStartupCommand { get; }
        public ICommand ShowSystemCommand { get; }
        public ICommand ShowUninstallCommand { get; }
        public ICommand ShowSqlExportCommand { get; }
        public ICommand ShowFrpCommand { get; }
        public ICommand ShowCodexProfilesCommand { get; }
        public ICommand ShowMultimediaCommand { get; }
        public ICommand ShowImageViewerCommand { get; }
        public ICommand ShowVideoViewerCommand { get; }
        public ICommand ToggleStartupCommand { get; }
        public ICommand DeleteStartupCommand { get; }
        public string StartupSearchText
        {
            get => _startupSearchText;
            set
            {
                _startupSearchText = value ?? string.Empty;
                OnPropertyChanged();
                FilteredStartupView?.Refresh();
                OnPropertyChanged(nameof(HasNoStartupItems));
            }
        }

        public string InstalledProgramSearchText
        {
            get => _installedProgramSearchText;
            set
            {
                _installedProgramSearchText = value ?? string.Empty;
                OnPropertyChanged();
                RefreshInstalledProgramView();
            }
        }

        public string[] InstalledProgramSortOptions { get; } =
        {
            "名称 A-Z",
            "安装日期新到旧",
            "安装日期旧到新",
            "大小从大到小",
            "大小从小到大",
            "发布者 A-Z"
        };

        public string[] InstalledProgramSizeFilterOptions { get; } =
        {
            "全部大小",
            "未知大小",
            "小于 100 MB",
            "100 MB - 1 GB",
            "大于 1 GB"
        };

        public string[] InstalledProgramDateFilterOptions { get; } =
        {
            "全部日期",
            "未知日期",
            "最近 30 天",
            "最近 90 天",
            "最近 1 年"
        };

        public string InstalledProgramSortMode
        {
            get => _installedProgramSortMode;
            set
            {
                _installedProgramSortMode = string.IsNullOrWhiteSpace(value) ? "名称 A-Z" : value;
                OnPropertyChanged();
                ApplyInstalledProgramSort();
            }
        }

        public string InstalledProgramSizeFilter
        {
            get => _installedProgramSizeFilter;
            set
            {
                _installedProgramSizeFilter = string.IsNullOrWhiteSpace(value) ? "全部大小" : value;
                OnPropertyChanged();
                RefreshInstalledProgramView();
            }
        }

        public string InstalledProgramDateFilter
        {
            get => _installedProgramDateFilter;
            set
            {
                _installedProgramDateFilter = string.IsNullOrWhiteSpace(value) ? "全部日期" : value;
                OnPropertyChanged();
                RefreshInstalledProgramView();
            }
        }

        public bool HasNoStartupItems => FilteredStartupView == null || FilteredStartupView.IsEmpty;

        public bool HasNoInstalledPrograms => (FilteredInstalledProgramsView == null || FilteredInstalledProgramsView.IsEmpty) && !IsInstalledProgramsBusy;

        public string InstalledProgramsCountText => InstalledPrograms.Count == 0
            ? "未加载"
            : !HasInstalledProgramViewFilter()
                ? $"共 {InstalledPrograms.Count} 个"
                : $"筛选 {FilteredInstalledProgramsView?.Cast<object>().Count() ?? 0} / {InstalledPrograms.Count} 个";

        public int SelectedInstalledProgramsCount => InstalledPrograms.Count(program => program.IsSelected);

        public string SelectedInstalledProgramsCountText => SelectedInstalledProgramsCount == 0
            ? "未选择"
            : $"已选 {SelectedInstalledProgramsCount} 个";


        public string SqlQueryText
        {
            get => _sqlQueryText;
            set { _sqlQueryText = value; OnPropertyChanged(); _executeQueryCommand?.RaiseCanExecuteChanged(); }
        }

        public DataView SqlQueryResult
        {
            get => _sqlQueryResult;
            set
            {
                _sqlQueryResult = value;
                OnPropertyChanged();
                _exportQueryResultCommand?.RaiseCanExecuteChanged();
                _exportQueryResultCsvCommand?.RaiseCanExecuteChanged();
            }
        }

        public bool IsQueryBusy
        {
            get => _isQueryBusy;
            set
            {
                _isQueryBusy = value;
                OnPropertyChanged();
                _executeQueryCommand?.RaiseCanExecuteChanged();
                _exportQueryResultCommand?.RaiseCanExecuteChanged();
                _exportQueryResultCsvCommand?.RaiseCanExecuteChanged();
            }
        }

        public string QueryStatusMessage
        {
            get => _queryStatusMessage;
            set { _queryStatusMessage = value; OnPropertyChanged(); }
        }

        public bool IsInstalledProgramsBusy
        {
            get => _isInstalledProgramsBusy;
            set
            {
                _isInstalledProgramsBusy = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasNoInstalledPrograms));
                TriggerCommandRequery();
            }
        }

        public string InstalledProgramsStatusMessage
        {
            get => _installedProgramsStatusMessage;
            set { _installedProgramsStatusMessage = value; OnPropertyChanged(); }
        }

        public ICommand ExecuteSqlQueryCommand { get; }
        public ICommand ExportQueryResultCommand { get; }
        public ICommand ExportQueryResultCsvCommand { get; }
        public ICommand CancelSqlQueryCommand { get; }

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

        public bool IsPowerPolicyBusy
        {
            get => _isPowerPolicyBusy;
            set
            {
                if (_isPowerPolicyBusy == value)
                {
                    return;
                }

                _isPowerPolicyBusy = value;
                OnPropertyChanged();
                TriggerCommandRequery();
            }
        }

        public string PowerPolicyStatusMessage
        {
            get => _powerPolicyStatusMessage;
            set { _powerPolicyStatusMessage = value; OnPropertyChanged(); }
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

        public bool WeChatRestoreIncludeText
        {
            get => _weChatRestoreIncludeText;
            set { _weChatRestoreIncludeText = value; OnPropertyChanged(); TriggerCommandRequery(); }
        }

        public bool WeChatRestoreIncludeImage
        {
            get => _weChatRestoreIncludeImage;
            set { _weChatRestoreIncludeImage = value; OnPropertyChanged(); TriggerCommandRequery(); }
        }

        public bool WeChatRestoreIncludeVideo
        {
            get => _weChatRestoreIncludeVideo;
            set { _weChatRestoreIncludeVideo = value; OnPropertyChanged(); TriggerCommandRequery(); }
        }

        public bool WeChatRestoreIncludeVoice
        {
            get => _weChatRestoreIncludeVoice;
            set { _weChatRestoreIncludeVoice = value; OnPropertyChanged(); TriggerCommandRequery(); }
        }

        public bool WeChatRestoreIncludeFile
        {
            get => _weChatRestoreIncludeFile;
            set { _weChatRestoreIncludeFile = value; OnPropertyChanged(); TriggerCommandRequery(); }
        }

        public bool WeChatRestoreIncludeCache
        {
            get => _weChatRestoreIncludeCache;
            set { _weChatRestoreIncludeCache = value; OnPropertyChanged(); TriggerCommandRequery(); }
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
            set
            {
                if (_showEditorAfterCapture == value) return;
                _showEditorAfterCapture = value;
                OnPropertyChanged();
                SafeFireAndForget(PersistScreenshotBehaviorAsync());
            }
        }

        // ===== 截图模式：FullScreen / Region / Window =====
        private string _screenshotMode = "FullScreen";
        public string ScreenshotMode
        {
            get => _screenshotMode;
            set
            {
                var v = value ?? "FullScreen";
                if (v != "FullScreen" && v != "Region" && v != "Window") v = "FullScreen";
                if (_screenshotMode == v) return;
                _screenshotMode = v;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsScreenshotModeFullScreen));
                OnPropertyChanged(nameof(IsScreenshotModeRegion));
                OnPropertyChanged(nameof(IsScreenshotModeWindow));
                SafeFireAndForget(PersistScreenshotBehaviorAsync());
            }
        }
        public bool IsScreenshotModeFullScreen
        {
            get => _screenshotMode == "FullScreen";
            set { if (value) ScreenshotMode = "FullScreen"; }
        }
        public bool IsScreenshotModeRegion
        {
            get => _screenshotMode == "Region";
            set { if (value) ScreenshotMode = "Region"; }
        }
        public bool IsScreenshotModeWindow
        {
            get => _screenshotMode == "Window";
            set { if (value) ScreenshotMode = "Window"; }
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

        public string VideoRecordHotkeyText
        {
            get => _videoRecordHotkeyText;
            set { _videoRecordHotkeyText = value; OnPropertyChanged(); }
        }

        public bool IsCapturingVideoRecordHotkey
        {
            get => _isCapturingVideoRecordHotkey;
            set { _isCapturingVideoRecordHotkey = value; OnPropertyChanged(); }
        }

        public string AudioRecordHotkeyText
        {
            get => _audioRecordHotkeyText;
            set { _audioRecordHotkeyText = value; OnPropertyChanged(); }
        }

        public bool IsCapturingAudioRecordHotkey
        {
            get => _isCapturingAudioRecordHotkey;
            set { _isCapturingAudioRecordHotkey = value; OnPropertyChanged(); }
        }

        public bool IsGifRecordingMode
        {
            get => _isGifRecordingMode;
            set
            {
                if (_isGifRecordingMode == value)
                {
                    return;
                }

                _isGifRecordingMode = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(RecordingModeText));
                OnPropertyChanged(nameof(RecordingModeHint));
            }
        }

        public string RecordingModeText => IsGifRecordingMode ? "GIF" : "MP4";

        public string RecordingModeHint => IsGifRecordingMode
            ? "GIF 模式不录声音，适合短操作演示。"
            : "MP4 模式录制画面，并尽量合成系统声音。";

        public IReadOnlyList<string> RecordingFrameRateOptions { get; } = new[] { "15 FPS", "30 FPS", "60 FPS" };

        public IReadOnlyList<string> RecordingQualityOptions { get; } = new[] { "体积优先", "均衡", "高清" };

        public string RecordingFrameRateOption
        {
            get => _recordingFrameRateOption;
            set
            {
                var option = string.IsNullOrWhiteSpace(value) || !RecordingFrameRateOptions.Contains(value) ? "30 FPS" : value;
                if (string.Equals(_recordingFrameRateOption, option, StringComparison.Ordinal))
                {
                    return;
                }

                _recordingFrameRateOption = option;
                OnPropertyChanged();
            }
        }

        public string RecordingQualityOption
        {
            get => _recordingQualityOption;
            set
            {
                var option = string.IsNullOrWhiteSpace(value) || !RecordingQualityOptions.Contains(value) ? "均衡" : value;
                if (string.Equals(_recordingQualityOption, option, StringComparison.Ordinal))
                {
                    return;
                }

                _recordingQualityOption = option;
                OnPropertyChanged();
            }
        }

        public ICommand ShowScreenshotCommand { get; }
        public ICommand TakeScreenshotNowCommand { get; }
        public ICommand StartVideoRecordingCommand { get; }
        public ICommand ToggleAudioRecordingCommand { get; }
        public ICommand SelectRecordingOutputFolderCommand { get; }
        public ICommand SelectAudioOutputFolderCommand { get; }
        public ICommand StartCaptureHotkeyCommand { get; }
        public ICommand CancelCaptureHotkeyCommand { get; }
        public ICommand StartCaptureVideoRecordHotkeyCommand { get; }
        public ICommand CancelCaptureVideoRecordHotkeyCommand { get; }
        public ICommand StartCaptureAudioRecordHotkeyCommand { get; }
        public ICommand CancelCaptureAudioRecordHotkeyCommand { get; }
        public ICommand SaveRecordingHotkeySettingsCommand { get; }
        public ICommand SaveScreenshotSettingsCommand { get; }
        public ICommand EditClipboardImageCommand { get; }
        public ICommand OcrClipboardImageCommand { get; }
        public ICommand CopyScreenshotHistoryItemCommand { get; }
        public ICommand EditScreenshotHistoryItemCommand { get; }
        public ICommand OpenScreenshotHistoryItemCommand { get; }
        public ICommand DeleteScreenshotHistoryItemCommand { get; }
        public ICommand OpenImageViewerFileCommand { get; }
        public ICommand CopyImageViewerImageCommand { get; }
        public ICommand SaveImageViewerImageAsCommand { get; }
        public ICommand EditImageViewerImageCommand { get; }
        public ICommand OpenPreviousImageViewerFileCommand { get; }
        public ICommand OpenNextImageViewerFileCommand { get; }
        public ICommand ImageViewerZoomInCommand { get; }
        public ICommand ImageViewerZoomOutCommand { get; }
        public ICommand ImageViewerZoomActualCommand { get; }
        public ICommand RotateImageViewerLeftCommand { get; }
        public ICommand RotateImageViewerRightCommand { get; }
        public ICommand FlipImageViewerHorizontalCommand { get; }
        public ICommand FlipImageViewerVerticalCommand { get; }
        public ICommand ToggleImageViewerGrayscaleCommand { get; }
        public ICommand ResetImageViewerAdjustmentsCommand { get; }
        public ICommand ImageViewerFitToFrameCommand { get; }
        public ObservableCollection<ImageFolderNode> ImageFolderRoots { get; } = ImageFolderNode.CreateRoots();
        public ICommand OpenVideoViewerFileCommand { get; }
        public ICommand OpenVideoViewerInExternalPlayerCommand { get; }
        public ICommand CaptureVideoViewerFrameCommand { get; }
        public ICommand GenerateVideoViewerWaveformCommand { get; }
        public ICommand ToggleVideoViewerMuteCommand { get; }
        public ICommand OpenPreviousVideoViewerPlaylistItemCommand { get; }
        public ICommand OpenNextVideoViewerPlaylistItemCommand { get; }
        public ICommand OpenVideoViewerPlaylistItemCommand { get; }
        public ICommand RemoveVideoViewerPlaylistItemCommand { get; }
        public ICommand CopyVideoViewerPlaylistCommand { get; }
        public ICommand CleanInvalidVideoViewerPlaylistItemsCommand { get; }
        public ICommand ClearVideoViewerPlaylistCommand { get; }
        public ICommand SaveVideoViewerPlaylistCommand { get; }
        public ICommand LoadVideoViewerPlaylistCommand { get; }
        public ICommand OpenRecentVideoViewerPlaylistCommand { get; }
        public ICommand FavoriteVideoViewerPlaylistCommand { get; }
        public ICommand OpenFavoriteVideoViewerPlaylistCommand { get; }
        public ICommand RemoveFavoriteVideoViewerPlaylistCommand { get; }
        public ICommand SetVideoViewerLoopStartCommand { get; }
        public ICommand SetVideoViewerLoopEndCommand { get; }
        public ICommand ClearVideoViewerLoopCommand { get; }
        public ICommand ToggleVideoViewerLoopCommand { get; }
        public ICommand SaveVideoViewerLoopRangeCommand { get; }
        public ICommand OpenVideoViewerLoopRangeCommand { get; }
        public ICommand RemoveVideoViewerLoopRangeCommand { get; }
        public ICommand AddVideoViewerBookmarkCommand { get; }
        public ICommand OpenVideoViewerBookmarkCommand { get; }
        public ICommand RemoveVideoViewerBookmarkCommand { get; }
        public ICommand LoadVideoViewerSubtitleCommand { get; }
        public ICommand ClearVideoViewerSubtitleCommand { get; }
        public ICommand ApplyCodexProfileCommand { get; }
        public ICommand SwitchCodexProfileCommand { get; }
        public ICommand ExportCodexProfileCommand { get; }
        public ICommand PreviewCodexProfileDiffCommand { get; }
        public ICommand ImportCodexProfileCommand { get; }
        public ICommand ImportCodexCpaTokenCommand { get; }
        public ICommand RefreshCodexProfileCommand { get; }
        public ICommand RenameCodexProfileCommand { get; }
        public ICommand EditCodexProfileNoteCommand { get; }
        public ICommand RestoreLastCodexBackupCommand { get; }
        public ICommand ExportCodexProfilesEncBoxCommand { get; }
        public ICommand ImportCodexProfilesEncBoxCommand { get; }
        public ICommand DeleteCodexProfileCommand { get; }
        public ICommand EditCodexConfigTomlCommand { get; }
        public ICommand EditCodexAuthJsonCommand { get; }
        public ICommand RotateToNextCodexProfileCommand { get; }
        public ICommand RestartCodexDesktopCommand { get; }
        public ICommand ToggleCodexProfileRotationCommand { get; }

        public string CodexNextSwitchPreview
        {
            get => _codexNextSwitchPreview;
            private set
            {
                if (string.Equals(_codexNextSwitchPreview, value, StringComparison.Ordinal)) return;
                _codexNextSwitchPreview = value ?? "无可用轮换目标";
                OnPropertyChanged();
            }
        }

        public CodexRotationSettings CodexRotationSettings
        {
            get => _codexRotationSettings;
            set
            {
                if (ReferenceEquals(_codexRotationSettings, value)) return;
                _codexRotationSettings = value ?? new CodexRotationSettings();
                OnPropertyChanged();
            }
        }

        public bool IsCodexRotationAvailable
        {
            get
            {
                if (CodexProfiles == null) return false;
                var current = CodexProfiles.FirstOrDefault(p => p != null && p.IsActive);
                if (current == null) return false;

                return CodexProfiles.Any(p => p != null
                    && p.EnableRotation
                    && p.Status != CodexProfileLibraryService.StatusExpired
                    && !string.Equals(p.DisplayName, current.DisplayName, StringComparison.Ordinal));
            }
        }

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

        public string RecordingOutputFolderText
        {
            get => _recordingOutputFolderText;
            private set
            {
                _recordingOutputFolderText = string.IsNullOrWhiteSpace(value) ? "未设置" : value;
                OnPropertyChanged();
            }
        }

        public string AudioOutputFolderText
        {
            get => _audioOutputFolderText;
            private set
            {
                _audioOutputFolderText = string.IsNullOrWhiteSpace(value) ? "未设置" : value;
                OnPropertyChanged();
            }
        }

        public BitmapSource ImageViewerPreviewImage
        {
            get => _imageViewerPreviewImage;
            private set
            {
                if (ReferenceEquals(_imageViewerPreviewImage, value))
                {
                    return;
                }

                _imageViewerPreviewImage = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasImageViewerImage));
                OnPropertyChanged(nameof(ImageViewerImageInfo));
                TriggerCommandRequery();
            }
        }

        public bool HasImageViewerImage => ImageViewerPreviewImage != null;

        public string ImageViewerFileName => string.IsNullOrWhiteSpace(_imageViewerFilePath)
            ? "未打开图片"
            : Path.GetFileName(_imageViewerFilePath);

        public string ImageViewerFilePath
        {
            get => _imageViewerFilePath;
            private set
            {
                if (string.Equals(_imageViewerFilePath, value, StringComparison.Ordinal))
                {
                    return;
                }

                _imageViewerFilePath = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ImageViewerFileName));
            }
        }

        public string ImageViewerImageInfo
        {
            get
            {
                if (ImageViewerPreviewImage == null)
                {
                    return "未载入";
                }

                return $"{ImageViewerPreviewImage.PixelWidth} x {ImageViewerPreviewImage.PixelHeight}px";
            }
        }

        public string ImageViewerDirectoryPositionText
        {
            get
            {
                if (_imageViewerDirectoryFiles == null || _imageViewerDirectoryFiles.Count == 0 || _imageViewerDirectoryIndex < 0)
                {
                    return "目录：未载入";
                }

                return $"目录：{_imageViewerDirectoryIndex + 1} / {_imageViewerDirectoryFiles.Count}";
            }
        }

        public string ImageViewerStatusMessage
        {
            get => _imageViewerStatusMessage;
            private set
            {
                if (string.Equals(_imageViewerStatusMessage, value, StringComparison.Ordinal))
                {
                    return;
                }

                _imageViewerStatusMessage = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        public double ImageViewerZoom
        {
            get => _imageViewerZoom;
            set
            {
                var zoom = Math.Max(0.1, Math.Min(4.0, Math.Round(value, 2)));
                if (Math.Abs(_imageViewerZoom - zoom) < 0.001)
                {
                    return;
                }

                _imageViewerZoom = zoom;
                if (_imageViewerFitToFrame)
                {
                    _imageViewerFitToFrame = false;
                    OnPropertyChanged(nameof(ImageViewerFitToFrame));
                }
                OnPropertyChanged();
                OnPropertyChanged(nameof(ImageViewerZoomText));
            }
        }

        public string ImageViewerZoomText => $"{Math.Round(ImageViewerZoom * 100):0}%";

        public bool ImageViewerFitToFrame
        {
            get => _imageViewerFitToFrame;
            private set { _imageViewerFitToFrame = value; OnPropertyChanged(); }
        }

        public bool ImageViewerIsGrayscale
        {
            get => _imageViewerIsGrayscale;
            private set
            {
                if (_imageViewerIsGrayscale == value)
                {
                    return;
                }

                _imageViewerIsGrayscale = value;
                OnPropertyChanged();
            }
        }

        public IReadOnlyList<string> ImageViewerCropModes { get; } = new[] { "原图", "1:1", "4:3", "16:9" };

        public string ImageViewerCropMode
        {
            get => _imageViewerCropMode;
            set
            {
                var mode = string.IsNullOrWhiteSpace(value) ? "原图" : value;
                if (!ImageViewerCropModes.Contains(mode))
                {
                    mode = "原图";
                }

                if (string.Equals(_imageViewerCropMode, mode, StringComparison.Ordinal))
                {
                    return;
                }

                _imageViewerCropMode = mode;
                OnPropertyChanged();
                UpdateImageViewerPreview(mode == "原图" ? "已取消裁剪。" : "已应用中心裁剪：" + mode);
            }
        }

        public int ImageViewerBrightness
        {
            get => _imageViewerBrightness;
            set
            {
                var normalized = Math.Max(-100, Math.Min(100, value));
                if (_imageViewerBrightness == normalized)
                {
                    return;
                }

                _imageViewerBrightness = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ImageViewerBrightnessText));
                UpdateImageViewerPreview("已调整亮度。");
            }
        }

        public string ImageViewerBrightnessText => _imageViewerBrightness == 0
            ? "0"
            : (_imageViewerBrightness > 0 ? "+" : string.Empty) + _imageViewerBrightness.ToString();

        public int ImageViewerContrast
        {
            get => _imageViewerContrast;
            set
            {
                var normalized = Math.Max(-100, Math.Min(100, value));
                if (_imageViewerContrast == normalized)
                {
                    return;
                }

                _imageViewerContrast = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ImageViewerContrastText));
                UpdateImageViewerPreview("已调整对比度。");
            }
        }

        public string ImageViewerContrastText => _imageViewerContrast == 0
            ? "0"
            : (_imageViewerContrast > 0 ? "+" : string.Empty) + _imageViewerContrast.ToString();

        public int ImageViewerSharpenAmount
        {
            get => _imageViewerSharpenAmount;
            set
            {
                var normalized = Math.Max(0, Math.Min(100, value));
                if (_imageViewerSharpenAmount == normalized)
                {
                    return;
                }

                _imageViewerSharpenAmount = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ImageViewerSharpenText));
                UpdateImageViewerPreview("已调整锐化。");
            }
        }

        public string ImageViewerSharpenText => _imageViewerSharpenAmount.ToString();

        public string ImageViewerTransformText
        {
            get
            {
                var parts = new List<string>();
                if (_imageViewerCropMode != "原图") parts.Add("裁剪 " + _imageViewerCropMode);
                if (_imageViewerRotationDegrees != 0) parts.Add($"旋转 {_imageViewerRotationDegrees}°");
                if (_imageViewerFlipHorizontal) parts.Add("水平翻转");
                if (_imageViewerFlipVertical) parts.Add("垂直翻转");
                if (_imageViewerIsGrayscale) parts.Add("灰度");
                if (_imageViewerBrightness != 0) parts.Add("亮度 " + ImageViewerBrightnessText);
                if (_imageViewerContrast != 0) parts.Add("对比度 " + ImageViewerContrastText);
                if (_imageViewerSharpenAmount != 0) parts.Add("锐化 " + _imageViewerSharpenAmount);
                return parts.Count == 0 ? "原图" : string.Join(" / ", parts);
            }
        }

        public Uri VideoViewerSource
        {
            get => _videoViewerSource;
            private set
            {
                if (Equals(_videoViewerSource, value))
                {
                    return;
                }

                _videoViewerSource = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasVideoViewerVideo));
                TriggerCommandRequery();
            }
        }

        public bool HasVideoViewerVideo => VideoViewerSource != null;
        public bool IsVideoViewerCapturingFrame
        {
            get => _isVideoViewerCapturingFrame;
            private set
            {
                if (_isVideoViewerCapturingFrame == value)
                {
                    return;
                }

                _isVideoViewerCapturingFrame = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerFrameCaptureText));
                TriggerCommandRequery();
            }
        }

        public string VideoViewerFrameCaptureText => IsVideoViewerCapturingFrame ? "截帧中..." : "截帧";

        public bool IsVideoViewerGeneratingWaveform
        {
            get => _isVideoViewerGeneratingWaveform;
            private set
            {
                if (_isVideoViewerGeneratingWaveform == value)
                {
                    return;
                }

                _isVideoViewerGeneratingWaveform = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerWaveformText));
                TriggerCommandRequery();
            }
        }

        public string VideoViewerWaveformText => IsVideoViewerGeneratingWaveform ? "生成中..." : "波形";

        public BitmapImage VideoViewerWaveformImage
        {
            get => _videoViewerWaveformImage;
            private set
            {
                _videoViewerWaveformImage = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasVideoViewerWaveform));
            }
        }

        public string VideoViewerWaveformPath
        {
            get => _videoViewerWaveformPath;
            private set
            {
                _videoViewerWaveformPath = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        public bool HasVideoViewerWaveform => VideoViewerWaveformImage != null;

        public string VideoViewerFileName => string.IsNullOrWhiteSpace(_videoViewerFilePath)
            ? "未打开视频"
            : Path.GetFileName(_videoViewerFilePath);

        public string VideoViewerFilePath
        {
            get => _videoViewerFilePath;
            private set
            {
                if (string.Equals(_videoViewerFilePath, value, StringComparison.Ordinal))
                {
                    return;
                }

                _videoViewerFilePath = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerFileName));
            }
        }

        public string VideoViewerStatusMessage
        {
            get => _videoViewerStatusMessage;
            set
            {
                if (string.Equals(_videoViewerStatusMessage, value, StringComparison.Ordinal))
                {
                    return;
                }

                _videoViewerStatusMessage = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        public bool IsVideoViewerPlaying
        {
            get => _isVideoViewerPlaying;
            set
            {
                if (_isVideoViewerPlaying == value)
                {
                    return;
                }

                _isVideoViewerPlaying = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerPlaybackText));
            }
        }

        public string VideoViewerPlaybackText => IsVideoViewerPlaying ? "暂停" : "播放";

        public ObservableCollection<VideoPlaylistItem> VideoViewerPlaylist => _videoViewerPlaylist;

        public string VideoViewerPlaylistSearchText
        {
            get => _videoViewerPlaylistSearchText;
            set
            {
                var normalized = value ?? string.Empty;
                if (string.Equals(_videoViewerPlaylistSearchText, normalized, StringComparison.Ordinal))
                {
                    return;
                }

                _videoViewerPlaylistSearchText = normalized;
                OnPropertyChanged();
                FilteredVideoViewerPlaylistView?.Refresh();
                OnPropertyChanged(nameof(VideoViewerPlaylistFilterText));
            }
        }

        public string VideoViewerPlaylistFilterText
        {
            get
            {
                if (_videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0)
                {
                    return "未载入";
                }

                if (string.IsNullOrWhiteSpace(VideoViewerPlaylistSearchText))
                {
                    return $"{_videoViewerPlaylist.Count} 项";
                }

                var count = FilteredVideoViewerPlaylistView == null
                    ? 0
                    : FilteredVideoViewerPlaylistView.Cast<object>().Count();
                return $"筛选 {count} / {_videoViewerPlaylist.Count}";
            }
        }

        public string VideoViewerPlaylistText
        {
            get
            {
                if (_videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0 || _videoViewerPlaylistIndex < 0)
                {
                    return "列表：未载入";
                }

                return $"列表：{_videoViewerPlaylistIndex + 1} / {_videoViewerPlaylist.Count}";
            }
        }

        public IReadOnlyList<string> VideoViewerPlayModes { get; } = new[] { "顺序", "列表循环", "单项循环", "随机" };

        public string VideoViewerPlayMode
        {
            get => _videoViewerPlayMode;
            set
            {
                var mode = string.IsNullOrWhiteSpace(value) ? "顺序" : value;
                if (!VideoViewerPlayModes.Contains(mode))
                {
                    mode = "顺序";
                }

                if (string.Equals(_videoViewerPlayMode, mode, StringComparison.Ordinal))
                {
                    return;
                }

                _videoViewerPlayMode = mode;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerPlayModeText));
                OnPropertyChanged(nameof(VideoViewerAutoAdvanceText));
            }
        }

        public string VideoViewerPlayModeText => "模式：" + VideoViewerPlayMode;
        public string VideoViewerAutoAdvanceText => VideoViewerPlayMode == "单项循环"
            ? "正在单项循环。"
            : VideoViewerPlayMode == "随机"
                ? "正在随机播放下一项。"
                : "正在播放下一项。";

        public double VideoViewerLoopStartSeconds
        {
            get => _videoViewerLoopStartSeconds;
            private set
            {
                var normalized = NormalizeVideoPointSeconds(value);
                if (Math.Abs(_videoViewerLoopStartSeconds - normalized) < 0.1)
                {
                    return;
                }

                _videoViewerLoopStartSeconds = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerLoopStartText));
                OnPropertyChanged(nameof(VideoViewerLoopRangeText));
                OnPropertyChanged(nameof(HasVideoViewerLoopRange));
                OnPropertyChanged(nameof(HasVideoViewerWaveformLoopRange));
                TriggerCommandRequery();
            }
        }

        public double VideoViewerLoopEndSeconds
        {
            get => _videoViewerLoopEndSeconds;
            private set
            {
                var normalized = NormalizeVideoPointSeconds(value);
                if (Math.Abs(_videoViewerLoopEndSeconds - normalized) < 0.1)
                {
                    return;
                }

                _videoViewerLoopEndSeconds = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerLoopEndText));
                OnPropertyChanged(nameof(VideoViewerLoopRangeText));
                OnPropertyChanged(nameof(HasVideoViewerLoopRange));
                OnPropertyChanged(nameof(HasVideoViewerWaveformLoopRange));
                TriggerCommandRequery();
            }
        }

        public bool VideoViewerIsLoopEnabled
        {
            get => _videoViewerIsLoopEnabled;
            set
            {
                var enabled = value && HasVideoViewerLoopRange;
                if (_videoViewerIsLoopEnabled == enabled)
                {
                    return;
                }

                _videoViewerIsLoopEnabled = enabled;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerLoopStatusText));
                TriggerCommandRequery();
            }
        }

        public bool HasVideoViewerLoopRange => VideoViewerLoopStartSeconds >= 0
            && VideoViewerLoopEndSeconds > VideoViewerLoopStartSeconds + 0.2;
        public bool HasVideoViewerWaveformLoopRange => HasVideoViewerWaveform && HasVideoViewerLoopRange;
        public string VideoViewerLoopStartText => VideoViewerLoopStartSeconds < 0 ? "--:--" : FormatVideoTime(VideoViewerLoopStartSeconds);
        public string VideoViewerLoopEndText => VideoViewerLoopEndSeconds < 0 ? "--:--" : FormatVideoTime(VideoViewerLoopEndSeconds);
        public string VideoViewerLoopRangeText => HasVideoViewerLoopRange
            ? $"{VideoViewerLoopStartText} - {VideoViewerLoopEndText}"
            : "未设置 A/B";
        public string VideoViewerLoopStatusText => VideoViewerIsLoopEnabled ? "A/B 循环开" : "A/B 循环关";

        public ObservableCollection<VideoLoopRangeItem> VideoViewerLoopRanges => _videoViewerLoopRanges;
        public bool HasVideoViewerLoopRanges => VideoViewerLoopRanges != null && VideoViewerLoopRanges.Count > 0;

        public ObservableCollection<VideoBookmarkItem> VideoViewerBookmarks => _videoViewerBookmarks;
        public bool HasVideoViewerBookmarks => VideoViewerBookmarks != null && VideoViewerBookmarks.Count > 0;

        public double VideoViewerSpeedRatio
        {
            get => _videoViewerSpeedRatio;
            set
            {
                var normalized = Math.Max(0.5, Math.Min(2.0, value));
                if (Math.Abs(_videoViewerSpeedRatio - normalized) < 0.001)
                {
                    return;
                }

                _videoViewerSpeedRatio = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerSpeedText));
            }
        }

        public double VideoViewerPositionSeconds
        {
            get => _videoViewerPositionSeconds;
            set
            {
                var normalized = Math.Max(0, value);
                if (Math.Abs(_videoViewerPositionSeconds - normalized) < 0.1)
                {
                    return;
                }

                _videoViewerPositionSeconds = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerPositionText));
            }
        }

        public double VideoViewerDurationSeconds
        {
            get => _videoViewerDurationSeconds;
            set
            {
                var normalized = Math.Max(0, value);
                if (Math.Abs(_videoViewerDurationSeconds - normalized) < 0.1)
                {
                    return;
                }

                _videoViewerDurationSeconds = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerDurationText));
            }
        }

        public double VideoViewerVolume
        {
            get => _videoViewerVolume;
            set
            {
                var normalized = Math.Max(0, Math.Min(1, value));
                if (Math.Abs(_videoViewerVolume - normalized) < 0.001)
                {
                    return;
                }

                _videoViewerVolume = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerVolumeText));
            }
        }

        public bool VideoViewerIsMuted
        {
            get => _videoViewerIsMuted;
            set
            {
                if (_videoViewerIsMuted == value)
                {
                    return;
                }

                _videoViewerIsMuted = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerMuteText));
            }
        }

        public string VideoViewerPositionText => FormatVideoTime(VideoViewerPositionSeconds);
        public string VideoViewerDurationText => FormatVideoTime(VideoViewerDurationSeconds);
        public string VideoViewerVolumeText => $"{Math.Round(VideoViewerVolume * 100):0}%";
        public string VideoViewerMuteText => VideoViewerIsMuted ? "取消静音" : "静音";
        public string VideoViewerSpeedText => $"{VideoViewerSpeedRatio:0.##}x";
        public bool HasVideoViewerSubtitle => _videoViewerSubtitles != null && _videoViewerSubtitles.Count > 0;
        public bool HasVideoViewerSubtitleText => !string.IsNullOrWhiteSpace(VideoViewerSubtitleText);

        public string VideoViewerSubtitleText
        {
            get => _videoViewerSubtitleText;
            private set
            {
                if (string.Equals(_videoViewerSubtitleText, value, StringComparison.Ordinal))
                {
                    return;
                }

                _videoViewerSubtitleText = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasVideoViewerSubtitleText));
            }
        }

        public string VideoViewerSubtitleStatus
        {
            get => _videoViewerSubtitleStatus;
            private set
            {
                if (string.Equals(_videoViewerSubtitleStatus, value, StringComparison.Ordinal))
                {
                    return;
                }

                _videoViewerSubtitleStatus = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        public string VideoViewerSubtitleFileName => string.IsNullOrWhiteSpace(_videoViewerSubtitleFilePath)
            ? "未载入"
            : Path.GetFileName(_videoViewerSubtitleFilePath);

        public ObservableCollection<RecentPlaylistItem> VideoViewerRecentPlaylists
        {
            get => _videoViewerRecentPlaylists;
            set
            {
                _videoViewerRecentPlaylists = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasVideoViewerRecentPlaylists));
            }
        }

        public bool HasVideoViewerRecentPlaylists => VideoViewerRecentPlaylists != null && VideoViewerRecentPlaylists.Count > 0;

        public ObservableCollection<FavoritePlaylistItem> VideoViewerFavoritePlaylists
        {
            get => _videoViewerFavoritePlaylists;
            set
            {
                _videoViewerFavoritePlaylists = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasVideoViewerFavoritePlaylists));
            }
        }

        public bool HasVideoViewerFavoritePlaylists => VideoViewerFavoritePlaylists != null && VideoViewerFavoritePlaylists.Count > 0;

        public double VideoViewerSubtitleOffsetSeconds
        {
            get => _videoViewerSubtitleOffsetSeconds;
            set
            {
                var normalized = Math.Max(-10, Math.Min(10, Math.Round(value, 1)));
                if (Math.Abs(_videoViewerSubtitleOffsetSeconds - normalized) < 0.01)
                {
                    return;
                }

                _videoViewerSubtitleOffsetSeconds = normalized;
                _videoViewerSubtitleIndex = -1;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerSubtitleOffsetText));
                UpdateVideoViewerSubtitle(VideoViewerPositionSeconds);
            }
        }

        public string VideoViewerSubtitleOffsetText => _videoViewerSubtitleOffsetSeconds == 0
            ? "0.0 秒"
            : (_videoViewerSubtitleOffsetSeconds > 0 ? "+" : string.Empty) + _videoViewerSubtitleOffsetSeconds.ToString("0.0") + " 秒";

        public double VideoViewerSubtitleFontSize
        {
            get => _videoViewerSubtitleFontSize;
            set
            {
                var normalized = Math.Max(12, Math.Min(28, Math.Round(value)));
                if (Math.Abs(_videoViewerSubtitleFontSize - normalized) < 0.1)
                {
                    return;
                }

                _videoViewerSubtitleFontSize = normalized;
                OnPropertyChanged();
                OnPropertyChanged(nameof(VideoViewerSubtitleLineHeight));
                OnPropertyChanged(nameof(VideoViewerSubtitleFontSizeText));
            }
        }

        public double VideoViewerSubtitleLineHeight => Math.Round(VideoViewerSubtitleFontSize * 1.35);
        public string VideoViewerSubtitleFontSizeText => $"{VideoViewerSubtitleFontSize:0}px";

        public string CodexProfilesStatusMessage
        {
            get => _codexProfilesStatusMessage;
            set { _codexProfilesStatusMessage = value; OnPropertyChanged(); }
        }

        public ICommand LockWin10Command { get; }
        public ICommand ExitCommand { get; }
        public ICommand RestoreCommand { get; }
        public ICommand OpenLogFolderCommand { get; }
        public ICommand ToggleAutoStartCommand { get; }
        public ICommand ComputeFileHashCommand { get; }
        public ICommand ExportFileHashListCommand { get; }
        public ICommand ImportFileHashListCommand { get; }
        public ICommand CopyBatchFileHashColumnCommand { get; }

        public string FileHashResult
        {
            get => _fileHashResult;
            set { _fileHashResult = value; OnPropertyChanged(); }
        }

        public string ExpectedFileHash
        {
            get => _expectedFileHash;
            set
            {
                _expectedFileHash = value ?? string.Empty;
                OnPropertyChanged();
                UpdateHashCompareFromCurrentResult();
            }
        }

        public string FileHashCompareResult
        {
            get => _fileHashCompareResult;
            set
            {
                _fileHashCompareResult = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasFileHashCompareResult));
            }
        }

        public bool HasFileHashCompareResult => !string.IsNullOrWhiteSpace(FileHashCompareResult);

        public ObservableCollection<FileHashResult> BatchFileHashResults => _batchFileHashResults;
        public bool HasBatchFileHashResults => BatchFileHashResults.Count > 0;

        public bool IsFileHashBusy
        {
            get => _isFileHashBusy;
            set { _isFileHashBusy = value; OnPropertyChanged(); }
        }

        public string FileHashStatusMessage
        {
            get => _fileHashStatusMessage;
            set { _fileHashStatusMessage = value; OnPropertyChanged(); }
        }

        public bool IsAutoStartEnabled
        {
            get => _isAutoStartEnabled;
            set { _isAutoStartEnabled = value; OnPropertyChanged(); }
        }
        public ICommand TestSqlConnectionCommand { get; }
        public ICommand ExportSqlTableCommand { get; }
        public ICommand CancelSqlExportCommand { get; }
        public ICommand ApplySqlRecentConnectionCommand { get; }
        public ICommand ToggleDefenderCommand { get; }
        public ICommand ToggleAutoUpdateCommand { get; }
        public ICommand TriggerUpdateNowCommand { get; }
        public ICommand RefreshSystemStatusCommand { get; }
        public ICommand ApplyAlwaysOnPowerPolicyCommand { get; }
        public ICommand RefreshInstalledProgramsCommand { get; }
        public ICommand UninstallProgramCommand { get; }
        public ICommand SelectFilteredInstalledProgramsCommand { get; }
        public ICommand ClearInstalledProgramSelectionCommand { get; }
        public ICommand ExportInstalledProgramListCommand { get; }
        public ICommand StartAutoOptimizeCommand { get; }
        public ICommand StartJunkScanCommand { get; }
        public ICommand RunJunkCleanupCommand { get; }
        public ICommand ExportJunkCleanupPlanCommand { get; }
        public ICommand ScanWeChatCleanupCommand { get; }
        public ICommand StartWeChatCleanupCommand { get; }
        public ICommand StartWeChatBackupCommand { get; }
        public ICommand StartWeChatRestoreCommand { get; }
        public ICommand DeleteSelectedReportsCommand { get; }
        public ICommand SelectWeChatBackupOutputFolderCommand { get; }
        public ICommand SelectWeChatRestoreZipCommand { get; }
        public ICommand SelectWeChatRestoreTargetRootCommand { get; }
        public ICommand OpenRecentWeChatBackupCommand { get; }
        public ICommand ShowReportDetailsCommand { get; }

        public void AddCodexProfileFolders(IEnumerable<string> folderPaths)
        {
            SafeFireAndForget(AddCodexProfileFoldersAsync(folderPaths));
        }

        private static readonly string[] SystemSectionKeys = { "Network", "Startup", "Uninstall", "SystemInfo", "Benchmark", "SystemSettings" };

        private void ShowSystemSection(string section)
        {
            if (string.Equals(section, "Optimization", StringComparison.Ordinal))
            {
                SwitchModule("Optimization");
                return;
            }
            var normalized = NormalizeSystemSection(section);
            SwitchModule(normalized);
            CurrentSystemSection = normalized;
            var index = Array.IndexOf(SystemSectionKeys, normalized);
            if (index < 0) index = 0;
            if (_selectedSystemSectionIndex != index)
            {
                _selectedSystemSectionIndex = index;
                OnPropertyChanged(nameof(SelectedSystemSectionIndex));
            }

            LoadSystemSection(normalized);
        }

        private static string NormalizeSystemSection(string section)
        {
            if (SystemSectionKeys.Contains(section)) return section;
            return "Startup";
        }

        private string BuildCurrentNavigationText()
        {
            switch (CurrentModule)
            {
                case "Home":
                    return "首页";
                case "Network":
                    return "当前网络";
                case "Startup":
                    return "启动管理";
                case "Uninstall":
                    return "程序卸载";
                case "Optimization":
                    return "系统优化";
                case "SystemInfo":
                    return "系统信息";
                case "Benchmark":
                    return "性能测试";
                case "SystemSettings":
                    return "系统设置";
                case "SqlExport":
                    return "SQL 导出";
                case "Frp":
                    return "内网穿透";
                case "Multimedia":
                    return "多媒体";
                case "Screenshot":
                    return "截图录屏";
                case "CodexProfiles":
                    return "Codex 配置";
                case "FileVerify":
                    return "文件校验";
                case "Schedule":
                    return "排班管理";
                case "ImageViewer":
                    return "图片查看";
                case "VideoViewer":
                    return "音视频播放";
                case "Convert":
                    return "格式转换";
                default:
                    return string.IsNullOrWhiteSpace(CurrentModule) ? "首页" : CurrentModule;
            }
        }

        private bool IsSystemSectionVisible(string section)
        {
            return string.Equals(CurrentModule, "System", StringComparison.Ordinal)
                   && string.Equals(CurrentSystemSection, section, StringComparison.Ordinal);
        }

        private void NotifySystemSectionVisibilityChanged()
        {
            OnPropertyChanged(nameof(IsSystemOptimizationVisible));
            OnPropertyChanged(nameof(IsSystemNetworkVisible));
            OnPropertyChanged(nameof(IsSystemStartupVisible));
            OnPropertyChanged(nameof(IsSystemUninstallVisible));
            OnPropertyChanged(nameof(IsSystemInfoVisible));
            OnPropertyChanged(nameof(IsSystemBenchmarkVisible));
            OnPropertyChanged(nameof(IsSystemSettingsVisible));
        }

        private void LoadSystemSection(string section)
        {
            if (string.Equals(section, "Network", StringComparison.Ordinal))
            {
                Refresh();
            }
            else if (string.Equals(section, "Startup", StringComparison.Ordinal))
            {
                Refresh();
            }
            else if (string.Equals(section, "Uninstall", StringComparison.Ordinal))
            {
                SafeFireAndForget(LoadInstalledProgramsAsync());
            }
            else if (string.Equals(section, "SystemInfo", StringComparison.Ordinal))
            {
                EnsureSystemInfoSnapshotLoading();
                SafeFireAndForget(LoadSystemInfoAsync());
            }
            else if (string.Equals(section, "Optimization", StringComparison.Ordinal))
            {
                EnsureSystemOptimizationDataLoading();
                EnsureWeChatStartupDataLoading();
                RefreshSystemStatus();
            }
        }

        private void EnsureSqlStartupDataLoading()
        {
            if (_sqlHistoryLoadRequested)
            {
                return;
            }

            _sqlHistoryLoadRequested = true;
            SafeFireAndForget(LoadSqlConnectionHistoryAsync());
        }

        private void EnsureScreenshotStartupDataLoading()
        {
            if (_screenshotStartupLoadRequested)
            {
                return;
            }

            _screenshotStartupLoadRequested = true;
            SafeFireAndForget(LoadRecordingOutputFoldersAsync());
            EnsureScreenshotHotkeySettingsLoading();
            Application.Current?.Dispatcher.BeginInvoke(
                DispatcherPriority.Background,
                new Action(LoadScreenshotHistory));
        }

        public void ScheduleStartupHotkeyRegistration()
        {
            var dispatcher = Application.Current?.Dispatcher ?? Dispatcher.CurrentDispatcher;
            dispatcher.BeginInvoke(
                DispatcherPriority.ApplicationIdle,
                new Action(EnsureScreenshotHotkeySettingsLoading));
        }

        private void EnsureScreenshotHotkeySettingsLoading()
        {
            if (_screenshotHotkeysLoadRequested)
            {
                return;
            }

            _screenshotHotkeysLoadRequested = true;
            SafeFireAndForget(LoadScreenshotSettingsAsync());
        }

        private void EnsureVideoViewerStartupDataLoading()
        {
            if (_videoViewerStartupLoadRequested)
            {
                return;
            }

            _videoViewerStartupLoadRequested = true;
            SafeFireAndForget(LoadRecentVideoViewerPlaylistsAsync());
            SafeFireAndForget(LoadFavoriteVideoViewerPlaylistsAsync());
        }

        private void EnsureCodexProfilesLoading()
        {
            if (_codexProfilesLoadRequested)
            {
                return;
            }

            _codexProfilesLoadRequested = true;
            SafeFireAndForget(LoadCodexProfilesAsync());
        }

        private void EnsureSystemOptimizationDataLoading()
        {
            if (_systemOptimizationLoadRequested)
            {
                return;
            }

            _systemOptimizationLoadRequested = true;
            SafeFireAndForget(LoadOptimizationReportsAsync());
        }

        private void EnsureWeChatStartupDataLoading()
        {
            if (_weChatStartupLoadRequested)
            {
                return;
            }

            _weChatStartupLoadRequested = true;
            SafeFireAndForget(LoadWeChatRootsAsync());
            SafeFireAndForget(LoadRecentWeChatBackupsAsync());
        }

        private void EnsureFrpConfigLoading()
        {
            if (_frpConfigLoadRequested)
            {
                return;
            }

            _frpConfigLoadRequested = true;
            SafeFireAndForget(Frp.LoadConfigAsync());
        }

        private void EnsureSystemInfoSnapshotLoading()
        {
            SafeFireAndForget(LoadSystemInfoSnapshotAsync());
        }

        private void SwitchModule(string module)
        {
            if (!string.Equals(module, "Home", StringComparison.Ordinal))
            {
                DeferredUiResourceService.EnsureLoaded();
            }

            CurrentModule = module;

            if (string.Equals(module, "Schedule", StringComparison.Ordinal))
            {
                AddHomeRecentItem("排班管理", "最近进入排班模块", module, string.Empty, "打开");
            }
            else if (string.Equals(module, "Multimedia", StringComparison.Ordinal))
            {
                AddHomeRecentItem("多媒体", "最近进入多媒体模块", module, string.Empty, "打开");
            }
            else if (string.Equals(module, "Frp", StringComparison.Ordinal))
            {
                AddHomeRecentItem("隧道穿透", "最近进入穿透模块", module, string.Empty, "打开");
            }

            if (string.Equals(module, "Network", StringComparison.Ordinal))
            {
                Refresh();
            }
            else if (string.Equals(module, "Startup", StringComparison.Ordinal))
            {
                Refresh();
            }
            else if (string.Equals(module, "Uninstall", StringComparison.Ordinal))
            {
                SafeFireAndForget(LoadInstalledProgramsAsync());
            }
            else if (string.Equals(module, "SystemInfo", StringComparison.Ordinal))
            {
                EnsureSystemInfoSnapshotLoading();
                SafeFireAndForget(LoadSystemInfoAsync());
            }
            else if (string.Equals(module, "Optimization", StringComparison.Ordinal))
            {
                EnsureSystemOptimizationDataLoading();
                EnsureWeChatStartupDataLoading();
                RefreshSystemStatus();
            }
            else if (string.Equals(module, "Frp", StringComparison.Ordinal))
            {
                EnsureFrpConfigLoading();
            }
            else if (string.Equals(module, "SqlExport", StringComparison.Ordinal))
            {
                EnsureSqlStartupDataLoading();
            }
            else if (string.Equals(module, "Screenshot", StringComparison.Ordinal))
            {
                EnsureScreenshotStartupDataLoading();
            }
            else if (string.Equals(module, "CodexProfiles", StringComparison.Ordinal))
            {
                EnsureCodexProfilesLoading();
            }
        }

        public void SwitchToHomeFromMultimedia()
        {
            SwitchModule("Home");
        }

        private void ShowMultimedia(MultimediaPreferredFilter preferredFilter)
        {
            SwitchModule("Multimedia");
            if (preferredFilter == MultimediaPreferredFilter.AudioVideo)
            {
                EnsureVideoViewerStartupDataLoading();
            }

            Multimedia.PreferredFilter = preferredFilter;
            if (preferredFilter == MultimediaPreferredFilter.All)
            {
                DetectFfmpeg();
            }

            Application.Current?.Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, new Action(() =>
            {
                SafeFireAndForget(Multimedia.InitializeOnEnterAsync());
            }));
        }
        private void InitializeHomeCommandItems()
        {
            HomeCommandItems.Clear();

            AddHomeCommand("多媒体", "图片查看 + 音视频播放 + 批量格式转换", "多媒体 图片 视频 音频 转换 multimedia image video audio convert", ShowMultimediaCommand);
            AddHomeCommand("截图 / 录像 / 录音", "打开截图工具、区域录像、系统声音录音", "截图 截屏 录像 录屏 录音 声音 热键 capture record audio", ShowScreenshotCommand);
            AddHomeCommand("SQL 导出 / 查询", "连接数据库、查询、导出 Excel 或 CSV", "sql 数据库 查询 导出 excel csv mysql postgresql", ShowSqlExportCommand);
            AddHomeCommand("隧道穿透", "配置 frp 服务器和端口映射，一键启动或停止 frpc", "frp 穿透 隧道 端口 映射 公网 tunnel proxy", ShowFrpCommand);
            AddHomeCommand("排班管理", "维护人员、生成排班、冲突检查、导出 Excel", "排班 班次 人员 休息 冲突 excel schedule", ShowScheduleCommand);
            AddHomeCommand("文件哈希校验", "计算 MD5、SHA-1、SHA-256、CRC32", "哈希 校验 md5 sha crc 文件 verify hash", ShowFileVerifyCommand);
            AddHomeCommand("系统", "系统优化、当前网络、启动管理、程序卸载、系统信息、性能测试和系统设置", "系统 优化 清理 网络 启动 卸载 信息 性能 设置 backup settings network startup uninstall benchmark", ShowSystemCommand);
            AddHomeCommand("Codex 配置", "导入、切换、备份 Codex 配置资料", "codex 配置 profile config auth", ShowCodexProfilesCommand);
        }

        private void AddHomeCommand(string title, string subtitle, string keywords, ICommand command)
        {
            HomeCommandItems.Add(new HomeCommandItem
            {
                Title = title,
                Subtitle = subtitle,
                Keywords = keywords,
                Command = command
            });
        }

        private bool FilterHomeCommandItem(object value)
        {
            var item = value as HomeCommandItem;
            if (item == null)
            {
                return false;
            }

            var query = HomeCommandSearchText;
            return string.IsNullOrWhiteSpace(query) || item.Matches(query);
        }

        private bool FilterVideoViewerPlaylistItem(object value)
        {
            var item = value as VideoPlaylistItem;
            if (item == null)
            {
                return false;
            }

            return item.Matches(VideoViewerPlaylistSearchText);
        }

        private void ExecuteHomeCommandItem(object parameter)
        {
            var item = parameter as HomeCommandItem;
            if (item == null || item.Command == null || !item.Command.CanExecute(null))
            {
                return;
            }

            item.Command.Execute(null);
        }

        private void OpenHomeRecentItem(object parameter)
        {
            var item = parameter as HomeRecentItem;
            if (item == null)
            {
                return;
            }

            if (!string.IsNullOrWhiteSpace(item.FilePath) && !File.Exists(item.FilePath))
            {
                HomeRecentItems.Remove(item);
                AddHomeRecentItem(item.Title, "文件不存在：" + item.FilePath, item.Module, string.Empty, "跳转");
                SwitchModule(item.Module);
                return;
            }

            if (string.Equals(item.Module, "ImageViewer", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(item.FilePath))
            {
                TryOpenImageViewerFile(item.FilePath);
                return;
            }

            if (string.Equals(item.Module, "VideoViewer", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(item.FilePath))
            {
                TryOpenVideoViewerFile(item.FilePath);
                return;
            }

            if (string.Equals(item.Module, "Convert", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(item.FilePath))
            {
                SwitchModule("Convert");
                OpenPathInExplorer(item.FilePath);
                return;
            }

            SwitchModule(item.Module);
        }

        private void OpenPathInExplorer(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = "/select,\"" + path + "\"",
                        UseShellExecute = true
                    });
                    return;
                }

                if (!string.IsNullOrWhiteSpace(path) && Directory.Exists(path))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = path,
                        UseShellExecute = true
                    });
                }
            }
            catch (OperationCanceledException)
            {
                SqlStatusMessage = "导出已取消。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Open recent path failed: {Msg}", ex.Message);
                SystemStatusMessage = "打开位置失败：" + ex.Message;
            }
        }

        private void AddHomeRecentItem(string title, string subtitle, string module, string filePath, string actionText)
        {
            if (HomeRecentItems == null)
            {
                return;
            }

            var normalizedPath = string.IsNullOrWhiteSpace(filePath) ? string.Empty : Path.GetFullPath(filePath);
            var normalizedModule = module ?? string.Empty;
            var existing = HomeRecentItems.FirstOrDefault(item =>
                string.Equals(item.FilePath ?? string.Empty, normalizedPath, StringComparison.OrdinalIgnoreCase)
                && string.Equals(item.Module ?? string.Empty, normalizedModule, StringComparison.OrdinalIgnoreCase)
                && (normalizedPath.Length > 0 || string.Equals(item.Title ?? string.Empty, title ?? string.Empty, StringComparison.OrdinalIgnoreCase)));

            if (existing != null)
            {
                HomeRecentItems.Remove(existing);
            }

            HomeRecentItems.Insert(0, new HomeRecentItem
            {
                Title = string.IsNullOrWhiteSpace(title) ? "最近使用" : title,
                Subtitle = string.IsNullOrWhiteSpace(subtitle) ? normalizedPath : subtitle,
                Module = normalizedModule,
                FilePath = normalizedPath,
                ActionText = string.IsNullOrWhiteSpace(actionText) ? "打开" : actionText,
                LastUsedAt = DateTime.Now
            });

            while (HomeRecentItems.Count > 8)
            {
                HomeRecentItems.RemoveAt(HomeRecentItems.Count - 1);
            }

            OnPropertyChanged(nameof(HasHomeRecentItems));
        }

        public bool TryOpenImageViewerFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return false;
            }

            var extension = Path.GetExtension(filePath);
            if (!IsSupportedImageViewerFile(extension))
            {
                return false;
            }

            SwitchModule("ImageViewer");
            LoadImageViewerFile(filePath);
            return true;
        }

        public bool TryOpenVideoViewerFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return false;
            }

            var extension = Path.GetExtension(filePath);
            if (!IsSupportedVideoViewerFile(extension))
            {
                return false;
            }

            SwitchModule("VideoViewer");
            LoadVideoViewerPlaylist(new[] { filePath }, 0);
            return true;
        }

        public bool TryOpenVideoViewerFiles(IEnumerable<string> filePaths)
        {
            var files = (filePaths ?? Enumerable.Empty<string>())
                .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path) && IsSupportedVideoViewerFile(Path.GetExtension(path)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (files.Count == 0)
            {
                return false;
            }

            SwitchModule("VideoViewer");
            LoadVideoViewerPlaylist(files, 0);
            return true;
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

                    var profileItem = CreateCodexProfileItem(fullPath, null, null, null, null, protectedConfig, protectedAuth);
                    profileItem.StatusMessage = "已保存配置内容。";
                    AddCodexProfileItem(profileItem);
                    addedCount++;
                }
                catch (Exception ex)
                {
                    failedCount++;
                    AppLogService.Error(new InvalidOperationException(ex.Message), "Adding Codex profile folder failed for {FolderName} with {ErrorType}", Path.GetFileName(fullPath), ex.GetType().Name);
                }
            }

            if (addedCount > 0 || updatedCount > 0)
            {
                SwitchModule("CodexProfiles");
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
                var file = await CodexProfileLibraryService.LoadAsync(CancellationToken.None);
                var active = await CodexProfileLibraryService.LoadActiveAsync(CancellationToken.None);
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    foreach (var item in CodexProfiles)
                    {
                        item.PropertyChanged -= CodexProfileItem_OnPropertyChanged;
                    }

                    CodexProfiles.Clear();
                    foreach (var item in file.items ?? new List<CodexProfileItem>())
                    {
                        if (item == null)
                        {
                            continue;
                        }

                        ApplyCodexProfileMetadata(item);
                        item.IsActive = !string.IsNullOrWhiteSpace(active.ActiveDisplayName)
                            && string.Equals(item.DisplayName, active.ActiveDisplayName, StringComparison.OrdinalIgnoreCase);
                        AddCodexProfileItem(item);
                    }

                    SortCodexProfilesByLastApplied();
                    CodexProfilesStatusMessage = CodexProfiles.Count == 0
                        ? "未保存 Codex 账号档案。请先在 Codex CLI 登录一次，然后点击导入当前账号。"
                        : $"已加载 {CodexProfiles.Count} 个 Codex 账号档案。本机：{Environment.MachineName}";
                    UpdateCodexRotationState();
                });
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Loading Codex config profiles failed with {ErrorType}", ex.GetType().Name);
                CodexProfilesStatusMessage = "读取 Codex 配置记录失败。";
            }
        }

        private CodexProfileItem CreateCodexProfileItem(
            string folderPath,
            string name,
            string remark,
            string tags,
            DateTime? lastAppliedAt,
            string configTomlContentProtected,
            string authJsonContentProtected)
        {
            var normalizedPath = NormalizeFolderPath(folderPath);
            var defaultName = ResolveCodexProfileName(name, normalizedPath);
            var item = new CodexProfileItem
            {
                DisplayName = defaultName,
                Name = defaultName,
                Note = string.IsNullOrWhiteSpace(remark) ? string.Empty : remark.Trim(),
                Remark = string.IsNullOrWhiteSpace(remark) ? defaultName : remark.Trim(),
                Tags = tags ?? string.Empty,
                LastAppliedAt = lastAppliedAt,
                LastImportedAt = DateTime.UtcNow,
                FolderPath = normalizedPath,
                ProtectedConfigTomlBase64 = configTomlContentProtected ?? string.Empty,
                ProtectedAuthJsonBase64 = authJsonContentProtected ?? string.Empty,
                ConfigTomlContentProtected = configTomlContentProtected ?? string.Empty,
                AuthJsonContentProtected = authJsonContentProtected ?? string.Empty,
                StatusMessage = string.Empty
            };

            ApplyCodexProfileMetadata(item);
            return item;
        }

        private void ApplyCodexProfileMetadata(CodexProfileItem item)
        {
            if (item == null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(item.DisplayName))
            {
                item.DisplayName = string.IsNullOrWhiteSpace(item.Name) ? "Codex 账号" : item.Name.Trim();
            }

            item.Name = item.DisplayName;
            if (string.IsNullOrWhiteSpace(item.Note))
            {
                item.Note = item.Remark ?? string.Empty;
            }

            item.Remark = string.IsNullOrWhiteSpace(item.Note) ? item.DisplayName : item.Note;
            if (item.LastImportedAt == default(DateTime))
            {
                item.LastImportedAt = item.LastAppliedAt ?? DateTime.UtcNow;
            }

            if (string.IsNullOrWhiteSpace(item.ProtectedConfigTomlBase64))
            {
                item.ProtectedConfigTomlBase64 = item.ConfigTomlContentProtected ?? string.Empty;
            }

            if (string.IsNullOrWhiteSpace(item.ProtectedAuthJsonBase64))
            {
                item.ProtectedAuthJsonBase64 = item.AuthJsonContentProtected ?? string.Empty;
            }

            item.ConfigTomlContentProtected = item.ProtectedConfigTomlBase64;
            item.AuthJsonContentProtected = item.ProtectedAuthJsonBase64;
            var authBytes = CodexConfigProfileService.UnprotectBytesFromBase64(item.ProtectedAuthJsonBase64);
            item.AccountEmail = string.IsNullOrWhiteSpace(item.AccountEmail)
                ? CodexProfileLibraryService.ParseAccountEmail(authBytes)
                : item.AccountEmail;
            item.AccessTokenExpiresAt = item.AccessTokenExpiresAt ?? CodexProfileLibraryService.ParseAccessTokenExp(authBytes);
            item.RefreshTokenExpiresAt = null;
            item.Status = CodexProfileLibraryService.ComputeStatus(item.AccessTokenExpiresAt);
        }

        private void AddCodexProfileItem(CodexProfileItem item)
        {
            ApplyCodexProfileMetadata(item);
            item.PropertyChanged += CodexProfileItem_OnPropertyChanged;
            CodexProfiles.Add(item);
        }

        private void SortCodexProfilesByLastApplied()
        {
            var ordered = CodexProfiles
                .OrderByDescending(item => item.IsActive)
                .ThenByDescending(item => item.LastAppliedAt ?? item.LastImportedAt)
                .ThenBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            foreach (var item in CodexProfiles)
            {
                item.PropertyChanged -= CodexProfileItem_OnPropertyChanged;
            }

            CodexProfiles.Clear();
            foreach (var item in ordered)
            {
                AddCodexProfileItem(item);
            }
        }

        private void CodexProfileItem_OnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == nameof(CodexProfileItem.Remark)
                || e.PropertyName == nameof(CodexProfileItem.Tags)
                || e.PropertyName == nameof(CodexProfileItem.Note))
            {
                SafeFireAndForget(SaveCodexProfilesAsync());
            }

            if (e.PropertyName == nameof(CodexProfileItem.EnableRotation)
                || e.PropertyName == nameof(CodexProfileItem.RotationPriority))
            {
                UpdateCodexRotationState();
                SafeFireAndForget(SaveCodexProfilesAsync());
            }
        }

        private async Task SaveCodexProfilesAsync()
        {
            var file = BuildCodexProfilesFileFromCollection();
            await CodexProfileLibraryService.SaveAsync(file, CancellationToken.None);
        }

        private CodexProfilesFile BuildCodexProfilesFileFromCollection()
        {
            var file = new CodexProfilesFile
            {
                schemaVersion = CodexProfileLibraryService.CurrentSchemaVersion,
                machineName = Environment.MachineName,
                createdAtUtc = DateTime.UtcNow,
                items = CodexProfiles
                    .Where(item => item != null && !string.IsNullOrWhiteSpace(item.DisplayName))
                    .Select(item =>
                    {
                        ApplyCodexProfileMetadata(item);
                        return new CodexProfileItem
                        {
                            DisplayName = item.DisplayName,
                            Name = item.DisplayName,
                            AccountEmail = item.AccountEmail ?? string.Empty,
                            Note = item.Note ?? string.Empty,
                            Remark = item.Note ?? string.Empty,
                            Tags = item.Tags ?? string.Empty,
                            FolderPath = NormalizeFolderPath(item.FolderPath),
                            LastAppliedAt = item.LastAppliedAt,
                            LastImportedAt = item.LastImportedAt,
                            AccessTokenExpiresAt = item.AccessTokenExpiresAt,
                            RefreshTokenExpiresAt = null,
                            ProtectedConfigTomlBase64 = item.ProtectedConfigTomlBase64 ?? item.ConfigTomlContentProtected,
                            ProtectedAuthJsonBase64 = item.ProtectedAuthJsonBase64 ?? item.AuthJsonContentProtected,
                            ConfigTomlContentProtected = item.ProtectedConfigTomlBase64 ?? item.ConfigTomlContentProtected,
                            AuthJsonContentProtected = item.ProtectedAuthJsonBase64 ?? item.AuthJsonContentProtected,
                            Status = item.Status ?? CodexProfileLibraryService.StatusUnknown,
                            EnableRotation = item.EnableRotation,
                            RotationPriority = item.RotationPriority
                        };
                    })
                    .ToList()
            };
            return file;
        }

        private async Task ApplyCodexProfileAsync(object parameter)
        {
            if (!(parameter is CodexProfileItem item))
            {
                return;
            }

            var confirm = MessageBox.Show(
                $"即将切换到「{item.DisplayName}」，当前 ~/.codex 将自动备份。继续？",
                "切换 Codex 账号",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Question);
            if (confirm != MessageBoxResult.OK)
            {
                return;
            }

            try
            {
                item.IsApplying = true;
                item.StatusMessage = "正在切换...";

                var configTomlBytes = CodexConfigProfileService.UnprotectBytesFromBase64(item.ProtectedConfigTomlBase64 ?? item.ConfigTomlContentProtected);
                var authJsonBytes = CodexConfigProfileService.UnprotectBytesFromBase64(item.ProtectedAuthJsonBase64 ?? item.AuthJsonContentProtected);

                if (configTomlBytes == null || authJsonBytes == null)
                {
                    var fallbackFolderPath = NormalizeFolderPath(item.FolderPath);
                    if (string.IsNullOrWhiteSpace(fallbackFolderPath) || !Directory.Exists(fallbackFolderPath))
                    {
                        throw new InvalidOperationException("该记录未保存配置内容，且来源文件夹已不存在，请重新导入当前账号。 ");
                    }

                    var sourceFiles = await CodexConfigProfileService.ReadProfileFromFolderAsync(fallbackFolderPath, CancellationToken.None);
                    configTomlBytes = sourceFiles.ConfigTomlBytes;
                    authJsonBytes = sourceFiles.AuthJsonBytes;
                    item.ProtectedConfigTomlBase64 = CodexConfigProfileService.ProtectBytesToBase64(configTomlBytes);
                    item.ProtectedAuthJsonBase64 = CodexConfigProfileService.ProtectBytesToBase64(authJsonBytes);
                    item.ConfigTomlContentProtected = item.ProtectedConfigTomlBase64;
                    item.AuthJsonContentProtected = item.ProtectedAuthJsonBase64;
                }

                var previousActive = CodexProfiles.FirstOrDefault(profile => profile.IsActive)?.DisplayName ?? "codex";
                var backupPath = await CodexProfileLibraryService.BackupCurrentCodexFolderAsync(previousActive, CancellationToken.None);
                var result = await CodexConfigProfileService.ApplyAsync(configTomlBytes, authJsonBytes, CancellationToken.None);
                item.LastAppliedAt = DateTime.Now;
                item.LastImportedAt = item.LastImportedAt == default(DateTime) ? DateTime.UtcNow : item.LastImportedAt;
                item.AccountEmail = CodexProfileLibraryService.ParseAccountEmail(authJsonBytes);
                item.AccessTokenExpiresAt = CodexProfileLibraryService.ParseAccessTokenExp(authJsonBytes);
                item.Status = CodexProfileLibraryService.ComputeStatus(item.AccessTokenExpiresAt);
                item.StatusMessage = $"已切换：{item.LastAppliedAt:yyyy-MM-dd HH:mm:ss}";

                foreach (var profile in CodexProfiles)
                {
                    profile.IsActive = ReferenceEquals(profile, item);
                }

                await CodexProfileLibraryService.SaveActiveAsync(new CodexActiveFile
                {
                    ActiveDisplayName = item.DisplayName,
                    SwitchedAtUtc = DateTime.UtcNow
                }, CancellationToken.None);

                CodexProfilesStatusMessage = string.IsNullOrWhiteSpace(backupPath)
                    ? $"已切换到「{item.DisplayName}」。当前 ~/.codex 原本无可备份文件。"
                    : $"已切换到「{item.DisplayName}」。已备份切换前配置。";
                SortCodexProfilesByLastApplied();
                await SaveCodexProfilesAsync();
                AppLogService.Information("Switched Codex profile to {DisplayName}, backup at {BackupPath}", SafeCodexLogName(item.DisplayName), backupPath ?? string.Empty);
                UpdateCodexRotationState();
                MessageBox.Show(
                    $"已成功切换到「{item.DisplayName}」。\n\n目标目录：{result.TargetFolderPath}\n请重启 Codex 或重新打开终端后使用。",
                    "Codex 账号切换成功",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Switching Codex profile failed for {ProfileName} with {ErrorType}", SafeCodexLogName(item.DisplayName), ex.GetType().Name);
                item.StatusMessage = "切换失败：" + ex.Message;
                MessageBox.Show(ex.Message, "Codex 账号切换失败", MessageBoxButton.OK, MessageBoxImage.Error);
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
                    AppLogService.Information("Exported Codex profile {Name} to folder {FolderName}", SafeCodexLogName(item.Name), Path.GetFileName(targetFolder));
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Exporting Codex profile failed for {Name} with {ErrorType}", SafeCodexLogName(item.Name), ex.GetType().Name);
                MessageBox.Show(ex.Message, "导出失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ImportCodexProfileAsync()
        {
            try
            {
                var sourceFiles = await CodexProfileLibraryService.ReadCurrentCodexFilesAsync(CancellationToken.None);
                var email = CodexProfileLibraryService.ParseAccountEmail(sourceFiles.AuthJsonBytes);
                var fallbackName = string.IsNullOrWhiteSpace(email) ? $"Codex账号_{DateTime.Now:yyyyMMdd_HHmmss}" : email;
                var inputName = Interaction.InputBox("请输入账号档案别名", "导入当前 Codex 账号", fallbackName);
                var name = string.IsNullOrWhiteSpace(inputName) ? fallbackName : inputName.Trim();
                name = EnsureUniqueCodexDisplayName(name, null);

                var importSummary = BuildCodexProfileImportSummary(name, sourceFiles.ConfigTomlBytes, sourceFiles.AuthJsonBytes);
                var confirm = MessageBox.Show(
                    importSummary + "\n\n是否保存当前 ~/.codex 为这条账号档案？",
                    "导入当前 Codex 账号预览",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Information);
                if (confirm != MessageBoxResult.OK)
                {
                    CodexProfilesStatusMessage = "已取消导入。";
                    return;
                }

                var item = CreateCodexProfileItem(
                    sourceFiles.SourceFolderPath,
                    name,
                    string.Empty,
                    string.Empty,
                    null,
                    CodexConfigProfileService.ProtectBytesToBase64(sourceFiles.ConfigTomlBytes),
                    CodexConfigProfileService.ProtectBytesToBase64(sourceFiles.AuthJsonBytes));
                item.DisplayName = name;
                item.Name = name;
                item.Note = string.Empty;
                item.Remark = name;
                item.AccountEmail = email;
                item.LastImportedAt = DateTime.UtcNow;
                item.AccessTokenExpiresAt = CodexProfileLibraryService.ParseAccessTokenExp(sourceFiles.AuthJsonBytes);
                item.Status = CodexProfileLibraryService.ComputeStatus(item.AccessTokenExpiresAt);
                item.StatusMessage = "已导入当前账号。";
                AddCodexProfileItem(item);

                await SaveCodexProfilesAsync();
                CodexProfilesStatusMessage = $"已导入当前账号为「{item.DisplayName}」。";
                AppLogService.Information("Imported Codex profile {Name} from current Codex folder", SafeCodexLogName(item.DisplayName));
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Importing current Codex profile failed with {ErrorType}", ex.GetType().Name);
                MessageBox.Show(ex.Message, "导入失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ImportCodexCpaTokenAsync()
        {
            try
            {
                var dialog = new OpenFileDialog
                {
                    Title = "导入 CPA Token JSON",
                    Filter = "JSON 文件 (*.json)|*.json|所有文件 (*.*)|*.*"
                };
                if (dialog.ShowDialog() != true)
                {
                    return;
                }

                var tokenText = Encoding.UTF8.GetString(await ReadAllBytesAsync(dialog.FileName, CancellationToken.None).ConfigureAwait(true));
                var tokenJson = JObject.Parse(tokenText);
                var accessToken = SelectJsonString(tokenJson, "access_token", "accessToken");
                var rawRefreshToken = SelectJsonString(tokenJson, "refresh_token", "refreshToken");
                var sessionToken = SelectJsonString(tokenJson, "session_token", "sessionToken");
                var refreshToken = ResolveCpaRefreshToken(rawRefreshToken, sessionToken);
                var idToken = NormalizeCpaIdToken(SelectJsonString(tokenJson, "id_token", "idToken"));
                var usesSessionTokenAsRefreshToken = string.IsNullOrWhiteSpace(rawRefreshToken) && !string.IsNullOrWhiteSpace(refreshToken);
                if (string.IsNullOrWhiteSpace(accessToken))
                {
                    throw new InvalidOperationException("CPA Token 文件缺少 access_token，无法生成实验档案。");
                }

                var email = ResolveCpaTokenEmail(tokenJson, accessToken, idToken);
                var fallbackName = string.IsNullOrWhiteSpace(email)
                    ? $"CPA导入_{DateTime.Now:yyyyMMdd_HHmmss}"
                    : "CPA导入-" + CodexProfileLibraryService.MaskEmail(email);
                var inputName = Interaction.InputBox("请输入实验档案别名", "导入 CPA Token", fallbackName);
                var name = string.IsNullOrWhiteSpace(inputName) ? fallbackName : inputName.Trim();
                name = EnsureUniqueCodexDisplayName(name, null);

                var configBytes = await ResolveCodexConfigForCpaImportAsync().ConfigureAwait(true);
                var authBytes = BuildCodexAuthJsonFromCpaToken(tokenJson, accessToken, refreshToken, idToken, email);
                var summary = BuildCodexCpaImportSummary(name, dialog.FileName, tokenJson, authBytes, accessToken, refreshToken, idToken, usesSessionTokenAsRefreshToken);
                var confirm = MessageBox.Show(
                    summary + "\n\n该功能不会刷新 token，只生成实验档案。保存后可点击“切换”写入 ~/.codex 测试 Codex App 是否识别。是否继续？",
                    "导入 CPA Token 预览",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Information);
                if (confirm != MessageBoxResult.OK)
                {
                    CodexProfilesStatusMessage = "已取消 CPA Token 导入。";
                    return;
                }

                var item = CreateCodexProfileItem(
                    Path.GetDirectoryName(dialog.FileName) ?? string.Empty,
                    name,
                    "由 CPA token 文件生成，未验证可登录。",
                    "cpa,experiment",
                    null,
                    CodexConfigProfileService.ProtectBytesToBase64(configBytes),
                    CodexConfigProfileService.ProtectBytesToBase64(authBytes));
                item.DisplayName = name;
                item.Name = name;
                item.Note = "由 CPA token 文件生成，未验证可登录。";
                item.Remark = item.Note;
                item.AccountEmail = email;
                item.LastImportedAt = DateTime.UtcNow;
                item.AccessTokenExpiresAt = CodexProfileLibraryService.ParseAccessTokenExp(authBytes);
                item.Status = CodexProfileLibraryService.ComputeStatus(item.AccessTokenExpiresAt);
                item.StatusMessage = usesSessionTokenAsRefreshToken
                    ? "已导入 CPA 实验档案；使用 session_token 兜底，可能无法刷新。"
                    : string.IsNullOrWhiteSpace(refreshToken)
                    ? "已导入 CPA 实验档案；缺少 refresh_token。"
                    : "已导入 CPA 实验档案；未验证可登录。";
                AddCodexProfileItem(item);
                SortCodexProfilesByLastApplied();
                await SaveCodexProfilesAsync().ConfigureAwait(true);
                CodexProfilesStatusMessage = $"已导入 CPA Token 实验档案「{item.DisplayName}」。点击“切换”后可测试 Codex App 是否识别。";
                AppLogService.Information("Imported CPA token experiment profile {DisplayName} from {FileName}, has refresh token = {HasRefreshToken}, used session fallback = {UsedSessionFallback}", SafeCodexLogName(item.DisplayName), Path.GetFileName(dialog.FileName), !string.IsNullOrWhiteSpace(refreshToken), usesSessionTokenAsRefreshToken);
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Importing CPA token experiment profile failed with {ErrorType}", ex.GetType().Name);
                MessageBox.Show(ex.Message, "导入 CPA Token 失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private async Task RefreshCodexProfileAsync(object parameter)
        {
            if (!(parameter is CodexProfileItem item))
            {
                return;
            }

            try
            {
                var sourceFiles = await CodexProfileLibraryService.ReadCurrentCodexFilesAsync(CancellationToken.None);
                item.ProtectedConfigTomlBase64 = CodexConfigProfileService.ProtectBytesToBase64(sourceFiles.ConfigTomlBytes);
                item.ProtectedAuthJsonBase64 = CodexConfigProfileService.ProtectBytesToBase64(sourceFiles.AuthJsonBytes);
                item.ConfigTomlContentProtected = item.ProtectedConfigTomlBase64;
                item.AuthJsonContentProtected = item.ProtectedAuthJsonBase64;
                item.AccountEmail = CodexProfileLibraryService.ParseAccountEmail(sourceFiles.AuthJsonBytes);
                item.AccessTokenExpiresAt = CodexProfileLibraryService.ParseAccessTokenExp(sourceFiles.AuthJsonBytes);
                item.RefreshTokenExpiresAt = null;
                item.LastImportedAt = DateTime.UtcNow;
                item.Status = CodexProfileLibraryService.ComputeStatus(item.AccessTokenExpiresAt);
                item.StatusMessage = "已刷新当前 token。";
                await SaveCodexProfilesAsync();
                CodexProfilesStatusMessage = $"已刷新「{item.DisplayName}」。";
                AppLogService.Information("Refreshed Codex profile {DisplayName}, expires at {AccessTokenExpiresAt}", SafeCodexLogName(item.DisplayName), item.AccessTokenExpiresAt);
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Refreshing Codex profile failed for {DisplayName} with {ErrorType}", SafeCodexLogName(item.DisplayName), ex.GetType().Name);
                MessageBox.Show(ex.Message, "刷新失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task RenameCodexProfileAsync(object parameter)
        {
            if (!(parameter is CodexProfileItem item))
            {
                return;
            }

            var input = Interaction.InputBox("请输入新的账号档案别名", "重命名 Codex 档案", item.DisplayName ?? item.Name ?? string.Empty);
            if (string.IsNullOrWhiteSpace(input))
            {
                return;
            }

            var name = input.Trim();
            if (CodexProfiles.Any(profile => !ReferenceEquals(profile, item) && string.Equals(profile.DisplayName, name, StringComparison.OrdinalIgnoreCase)))
            {
                MessageBox.Show("已存在同名 Codex 档案，请换一个别名。", "重命名 Codex 档案", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            item.DisplayName = name;
            item.Name = name;
            item.Remark = string.IsNullOrWhiteSpace(item.Note) ? name : item.Note;
            if (item.IsActive)
            {
                await CodexProfileLibraryService.SaveActiveAsync(new CodexActiveFile
                {
                    ActiveDisplayName = item.DisplayName,
                    SwitchedAtUtc = DateTime.UtcNow
                }, CancellationToken.None);
            }

            await SaveCodexProfilesAsync();
            CodexProfilesStatusMessage = $"已重命名为「{item.DisplayName}」。";
        }

        private async Task EditCodexProfileNoteAsync(object parameter)
        {
            if (!(parameter is CodexProfileItem item))
            {
                return;
            }

            var input = Interaction.InputBox("请输入备注（最多 200 字）", "编辑 Codex 档案备注", item.Note ?? string.Empty);
            if (input == null)
            {
                return;
            }

            item.Note = input.Length > 200 ? input.Substring(0, 200) : input;
            item.Remark = string.IsNullOrWhiteSpace(item.Note) ? item.DisplayName : item.Note;
            await SaveCodexProfilesAsync();
            CodexProfilesStatusMessage = $"已更新「{item.DisplayName}」备注。";
        }

        private async Task RestoreLastCodexBackupAsync()
        {
            var confirm = MessageBox.Show(
                "将回滚到最近一次切换前的 ~/.codex 备份。继续？",
                "回滚 Codex 备份",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.OK)
            {
                return;
            }

            try
            {
                var path = await CodexProfileLibraryService.RestoreLatestBackupAsync(CancellationToken.None);
                foreach (var profile in CodexProfiles)
                {
                    profile.IsActive = false;
                }

                await CodexProfileLibraryService.SaveActiveAsync(new CodexActiveFile(), CancellationToken.None);
                CodexProfilesStatusMessage = "已回滚最近一次 Codex 切换备份，请重启 Codex。";
                AppLogService.Information("Restored latest Codex backup from {BackupPath}", path ?? string.Empty);
                MessageBox.Show("已回滚最近一次 Codex 切换备份，请重启 Codex。", "回滚完成", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Restoring Codex backup failed with {ErrorType}", ex.GetType().Name);
                MessageBox.Show(ex.Message, "回滚失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ExportCodexProfilesEncBoxAsync()
        {
            try
            {
                if (CodexProfiles.Count == 0)
                {
                    MessageBox.Show("当前没有可导出的 Codex 档案。", "导出加密包", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var dialog = new SaveFileDialog
                {
                    Title = "导出 Codex 加密档案包",
                    Filter = "Codex 加密档案包 (*.codexbox)|*.codexbox",
                    DefaultExt = ".codexbox",
                    FileName = $"CodexProfiles_{DateTime.Now:yyyyMMdd_HHmmss}.codexbox"
                };
                if (dialog.ShowDialog() != true)
                {
                    return;
                }

                var passwordDialog = new MyTools.Views.PasswordInputDialog("导出 Codex 加密档案包", "请输入导出口令。该口令只用于本次 .codexbox 文件加密，不会被保存。")
                {
                    Owner = Application.Current?.MainWindow
                };
                if (passwordDialog.ShowDialog() != true || string.IsNullOrEmpty(passwordDialog.Password))
                {
                    return;
                }

                var confirmPasswordDialog = new MyTools.Views.PasswordInputDialog("确认导出口令", "请再次输入导出口令。")
                {
                    Owner = Application.Current?.MainWindow
                };
                if (confirmPasswordDialog.ShowDialog() != true)
                {
                    return;
                }

                var password = passwordDialog.Password;
                var password2 = confirmPasswordDialog.Password;
                if (!string.Equals(password, password2, StringComparison.Ordinal))
                {
                    MessageBox.Show("两次输入的口令不一致。", "导出加密包", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var file = BuildCodexProfilesFileFromCollection();
                await CodexProfileLibraryService.ExportBoxAsync(file, dialog.FileName, password, CancellationToken.None);
                CodexProfilesStatusMessage = $"已导出 {file.items.Count} 个 Codex 档案到加密包。";
                AppLogService.Information("Exported Codex profile encbox {FileName}, item count = {Count}", Path.GetFileName(dialog.FileName), file.items.Count);
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Exporting Codex encbox failed with {ErrorType}", ex.GetType().Name);
                MessageBox.Show(ex.Message, "导出加密包失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ImportCodexProfilesEncBoxAsync()
        {
            try
            {
                var dialog = new OpenFileDialog
                {
                    Title = "导入 Codex 加密档案包",
                    Filter = "Codex 加密档案包 (*.codexbox)|*.codexbox|所有文件 (*.*)|*.*"
                };
                if (dialog.ShowDialog() != true)
                {
                    return;
                }

                var passwordDialog = new MyTools.Views.PasswordInputDialog("导入 Codex 加密档案包", "请输入 .codexbox 文件口令。")
                {
                    Owner = Application.Current?.MainWindow
                };
                if (passwordDialog.ShowDialog() != true || string.IsNullOrEmpty(passwordDialog.Password))
                {
                    return;
                }

                var importFile = await CodexProfileLibraryService.ImportBoxAsync(dialog.FileName, passwordDialog.Password, CancellationToken.None);
                var added = 0;
                var updated = 0;
                var skipped = 0;
                foreach (var imported in importFile.items ?? new List<CodexProfileItem>())
                {
                    if (imported == null)
                    {
                        skipped++;
                        continue;
                    }

                    ApplyCodexProfileMetadata(imported);
                    var existing = CodexProfiles.FirstOrDefault(item => string.Equals(item.DisplayName, imported.DisplayName, StringComparison.OrdinalIgnoreCase));
                    if (existing == null)
                    {
                        imported.DisplayName = EnsureUniqueCodexDisplayName(imported.DisplayName, null);
                        imported.Name = imported.DisplayName;
                        AddCodexProfileItem(imported);
                        added++;
                        continue;
                    }

                    var result = MessageBox.Show(
                        $"已存在同名档案「{imported.DisplayName}」。\n\n是=覆盖，否=重命名导入，取消=跳过。",
                        "导入冲突",
                        MessageBoxButton.YesNoCancel,
                        MessageBoxImage.Question);
                    if (result == MessageBoxResult.Cancel)
                    {
                        skipped++;
                        continue;
                    }

                    if (result == MessageBoxResult.Yes)
                    {
                        imported.IsActive = existing.IsActive;
                        var index = CodexProfiles.IndexOf(existing);
                        existing.PropertyChanged -= CodexProfileItem_OnPropertyChanged;
                        CodexProfiles[index] = imported;
                        imported.PropertyChanged += CodexProfileItem_OnPropertyChanged;
                        updated++;
                        continue;
                    }

                    imported.DisplayName = EnsureUniqueCodexDisplayName(imported.DisplayName, existing);
                    imported.Name = imported.DisplayName;
                    AddCodexProfileItem(imported);
                    added++;
                }

                SortCodexProfilesByLastApplied();
                await SaveCodexProfilesAsync();
                CodexProfilesStatusMessage = $"已导入加密包：新增 {added}，覆盖 {updated}，跳过 {skipped}。";
                AppLogService.Information("Imported Codex profile encbox {FileName}, added = {Added}, updated = {Updated}, skipped = {Skipped}", Path.GetFileName(dialog.FileName), added, updated, skipped);
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Importing Codex encbox failed with {ErrorType}", ex.GetType().Name);
                MessageBox.Show(ex.Message, "导入加密包失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task DeleteCodexProfileAsync(object parameter)
        {
            if (!(parameter is CodexProfileItem item))
            {
                return;
            }

            var result = MessageBox.Show(
                $"确定要删除档案「{item.DisplayName}」？删除后无法恢复（除非有加密导出包）。",
                "删除确认",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Warning);
            if (result != MessageBoxResult.OK)
            {
                return;
            }

            item.PropertyChanged -= CodexProfileItem_OnPropertyChanged;
            var wasActive = item.IsActive;
            CodexProfiles.Remove(item);
            if (wasActive)
            {
                await CodexProfileLibraryService.SaveActiveAsync(new CodexActiveFile(), CancellationToken.None);
            }

            CodexProfilesStatusMessage = $"已删除「{item.DisplayName}」。";
            await SaveCodexProfilesAsync();
            AppLogService.Information("Deleted Codex profile {DisplayName}", SafeCodexLogName(item.DisplayName));
            UpdateCodexRotationState();
        }

        private async Task RotateToNextCodexProfileAsync()
        {
            var current = CodexProfiles?.FirstOrDefault(p => p != null && p.IsActive);
            if (current == null)
            {
                CodexProfilesStatusMessage = "未找到当前激活的 Codex 账号";
                return;
            }

            await SaveCodexProfilesAsync();
            CodexProfilesStatusMessage = $"正在切换 Codex 账号...";
            var result = await CodexRotationService.RotateToNextAsync(
                current, CodexRotationSettings.NotifyOnSwitch, CancellationToken.None);

            if (result.Success)
            {
                current.IsActive = false;
                var next = CodexProfiles?.FirstOrDefault(p => p != null && p.DisplayName == result.ToProfile);
                if (next != null)
                {
                    next.IsActive = true;
                    next.LastAppliedAt = DateTime.Now;
                    next.StatusMessage = $"已轮换：{next.LastAppliedAt:yyyy-MM-dd HH:mm:ss}";
                }

                CodexProfilesStatusMessage = $"已写入：{result.FromProfile} → {result.ToProfile}。请重启 Codex App 后使用。";
                SortCodexProfilesByLastApplied();
                await SaveCodexProfilesAsync();
                UpdateCodexRotationState();
                var restartNow = MessageBox.Show(
                    $"已把 Codex 配置写入为「{result.ToProfile}」。\n\n当前已运行的 Codex App 后端不会热加载 auth.json。是否现在温和重启 Codex App？\n\n已保存的对话历史通常会保留；正在生成的回复或正在执行的命令会被中断。请确认当前 Codex 没有正在工作。",
                    "Codex 账号轮换已写入",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Information);
                if (restartNow == MessageBoxResult.Yes)
                {
                    await RestartCodexDesktopCoreAsync();
                }
            }
            else
            {
                CodexProfilesStatusMessage = $"轮换失败：{result.Message}";
            }
        }

        private async Task RestartCodexDesktopAsync()
        {
            var confirm = MessageBox.Show(
                "将关闭并重新打开 Codex App。\n\n工具会先请求 Codex 正常退出；如果仍有遗留的 Codex 后端进程未退出，会结束这些进程后再重新打开。\n\n已保存的对话历史通常会保留；正在生成的回复或正在执行的命令会被中断。请确认当前 Codex 没有正在生成内容。现在重启吗？",
                "重启 Codex App",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes)
            {
                return;
            }

            await RestartCodexDesktopCoreAsync();
        }

        private async Task RestartCodexDesktopCoreAsync()
        {
            CodexProfilesStatusMessage = "正在重启 Codex App...";
            var result = await CodexDesktopService.RestartAsync(CancellationToken.None);
            CodexProfilesStatusMessage = result.Message;
            MessageBox.Show(
                result.Message,
                result.Success ? "Codex App 已重启" : "Codex App 未完成重启",
                MessageBoxButton.OK,
                result.Success ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        private async Task ToggleCodexProfileRotationAsync(object parameter)
        {
            if (!(parameter is CodexProfileItem item)) return;
            item.EnableRotation = !item.EnableRotation;
            await SaveCodexProfilesAsync();
            UpdateCodexRotationState();
        }

        private void UpdateCodexNextSwitchPreview()
        {
            var current = CodexProfiles?.FirstOrDefault(p => p != null && p.IsActive);
            if (current == null)
            {
                CodexNextSwitchPreview = "无可用轮换目标";
                return;
            }

            var next = CodexProfiles?
                .Where(p => p != null
                            && p.EnableRotation
                            && p.Status != CodexProfileLibraryService.StatusExpired
                            && !string.Equals(p.DisplayName, current.DisplayName, StringComparison.Ordinal))
                .OrderBy(p => p.RotationPriority)
                .ThenBy(p => p.LastAppliedAt ?? DateTime.MinValue)
                .FirstOrDefault();

            CodexNextSwitchPreview = next == null
                ? "无可用轮换目标"
                : $"当前：{current.DisplayName} → 切换至：{next.DisplayName}";
        }

        private void UpdateCodexRotationState()
        {
            OnPropertyChanged(nameof(IsCodexRotationAvailable));
            UpdateCodexNextSwitchPreview();
            _rotateToNextCodexProfileCommand?.RaiseCanExecuteChanged();
        }

        private static string SafeCodexLogName(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return value.IndexOf('@') >= 0
                ? CodexProfileLibraryService.MaskEmail(value)
                : value;
        }
        private string EnsureUniqueCodexDisplayName(string requestedName, CodexProfileItem ignoreItem)
        {
            var baseName = string.IsNullOrWhiteSpace(requestedName) ? "Codex 账号" : requestedName.Trim();
            var candidate = baseName;
            var index = 2;
            while (CodexProfiles.Any(item => !ReferenceEquals(item, ignoreItem) && string.Equals(item.DisplayName, candidate, StringComparison.OrdinalIgnoreCase)))
            {
                candidate = baseName + "_" + index;
                index++;
            }

            return candidate;
        }
        private async Task EditCodexFileAsync(object parameter, string fileName)
        {
            if (!(parameter is CodexProfileItem item))
            {
                return;
            }

            try
            {
                bool isConfigToml = string.Equals(fileName, CodexConfigProfileService.ConfigFileName, StringComparison.OrdinalIgnoreCase);
                var protectedContent = isConfigToml ? item.ConfigTomlContentProtected : item.AuthJsonContentProtected;

                // 解密；若无内容则从文件夹回退
                byte[] currentBytes = CodexConfigProfileService.UnprotectBytesFromBase64(protectedContent);
                if (currentBytes == null)
                {
                    var fallbackFolderPath = NormalizeFolderPath(item.FolderPath);
                    var fallbackFile = string.IsNullOrWhiteSpace(fallbackFolderPath)
                        ? null
                        : Path.Combine(fallbackFolderPath, fileName);
                    if (!string.IsNullOrWhiteSpace(fallbackFile) && File.Exists(fallbackFile))
                    {
                        currentBytes = await ReadAllBytesAsync(fallbackFile, CancellationToken.None).ConfigureAwait(true);
                    }
                    else
                    {
                        currentBytes = new byte[0];
                    }
                }

                var initialText = Encoding.UTF8.GetString(currentBytes);
                var dlg = new MyTools.Views.CodexFileEditorDialog(fileName, item.Name, initialText)
                {
                    Owner = Application.Current?.MainWindow
                };
                if (dlg.ShowDialog() != true) return; // 关闭/取消

                var newBytes = Encoding.UTF8.GetBytes(dlg.EditedText ?? string.Empty);
                var newProtected = CodexConfigProfileService.ProtectBytesToBase64(newBytes);
                if (isConfigToml)
                {
                    item.ProtectedConfigTomlBase64 = newProtected;
                    item.ConfigTomlContentProtected = newProtected;
                }
                else
                {
                    item.ProtectedAuthJsonBase64 = newProtected;
                    item.AuthJsonContentProtected = newProtected;
                    item.AccountEmail = CodexProfileLibraryService.ParseAccountEmail(newBytes);
                    item.AccessTokenExpiresAt = CodexProfileLibraryService.ParseAccessTokenExp(newBytes);
                    item.Status = CodexProfileLibraryService.ComputeStatus(item.AccessTokenExpiresAt);
                }

                item.LastImportedAt = DateTime.UtcNow;
                item.StatusMessage = $"已保存 {fileName}：{DateTime.Now:HH:mm:ss}";
                CodexProfilesStatusMessage = $"已更新「{item.Name}」的 {fileName}。";
                await SaveCodexProfilesAsync().ConfigureAwait(true);
                AppLogService.Information("Edited Codex {File} for profile {Name}", fileName, SafeCodexLogName(item.Name));
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Editing Codex {File} failed for {Name} with {ErrorType}", fileName, SafeCodexLogName(item.Name), ex.GetType().Name);
                MessageBox.Show(ex.Message, "编辑失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task<byte[]> ResolveCodexConfigForCpaImportAsync()
        {
            var currentConfigPath = Path.Combine(CodexProfileLibraryService.CodexFolderPath, CodexConfigProfileService.ConfigFileName);
            if (File.Exists(currentConfigPath))
            {
                return await ReadAllBytesAsync(currentConfigPath, CancellationToken.None).ConfigureAwait(true);
            }

            return Encoding.UTF8.GetBytes(string.Empty);
        }

        private static byte[] BuildCodexAuthJsonFromCpaToken(JObject source, string accessToken, string refreshToken, string idToken, string email)
        {
            var accountId = ResolveCpaTokenAccountId(source, accessToken);
            var tokens = new JObject();
            AddJsonString(tokens, "id_token", idToken);
            AddJsonString(tokens, "access_token", accessToken);
            AddJsonString(tokens, "refresh_token", refreshToken);
            AddJsonString(tokens, "account_id", accountId);

            var auth = new JObject
            {
                ["auth_mode"] = "chatgpt",
                ["OPENAI_API_KEY"] = null,
                ["tokens"] = tokens,
                ["last_refresh"] = ResolveCpaLastRefresh(source)
            };

            return Encoding.UTF8.GetBytes(auth.ToString(Formatting.Indented));
        }

        private static string ResolveCpaRefreshToken(string refreshToken, string sessionToken)
        {
            if (!string.IsNullOrWhiteSpace(refreshToken))
            {
                return refreshToken.Trim();
            }

            return !string.IsNullOrWhiteSpace(sessionToken) ? sessionToken.Trim() : string.Empty;
        }

        private static string NormalizeCpaIdToken(string idToken)
        {
            if (string.IsNullOrWhiteSpace(idToken))
            {
                return string.Empty;
            }

            var normalized = idToken.Trim();
            return normalized.EndsWith(".", StringComparison.Ordinal) ? normalized + "AA" : normalized;
        }

        private static string BuildCodexCpaImportSummary(string profileName, string sourcePath, JObject source, byte[] authBytes, string accessToken, string refreshToken, string idToken, bool usesSessionTokenAsRefreshToken)
        {
            var builder = new StringBuilder();
            builder.AppendLine("即将导入 CPA Token 实验档案");
            builder.AppendLine("记录名：" + (profileName ?? string.Empty));
            builder.AppendLine("来源文件：" + Path.GetFileName(sourcePath));
            builder.AppendLine("auth.json：" + FileSizeFormatter.Format(authBytes?.LongLength ?? 0));
            builder.AppendLine("access_token：" + (string.IsNullOrWhiteSpace(accessToken) ? "缺失" : $"存在（长度 {accessToken.Length}）"));
            builder.AppendLine("refresh_token：" + (string.IsNullOrWhiteSpace(refreshToken) ? "缺失" : $"存在（长度 {refreshToken.Length}）"));
            if (usesSessionTokenAsRefreshToken)
            {
                builder.AppendLine("refresh_token 来源：session_token 兜底");
                builder.AppendLine("兜底限制：Codex 可能识别为已登录，但服务端作废 access_token 后仍会要求重新登录。");
            }
            builder.AppendLine("id_token：" + (string.IsNullOrWhiteSpace(idToken) ? "缺失" : $"存在（长度 {idToken.Length}）"));
            var expiresAt = CodexProfileLibraryService.ParseAccessTokenExp(authBytes);
            builder.AppendLine("access_token 到期：" + (expiresAt.HasValue ? expiresAt.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") : "未知"));
            var email = CodexProfileLibraryService.ParseAccountEmail(authBytes);
            builder.AppendLine("账号：" + (string.IsNullOrWhiteSpace(email) ? "未知" : CodexProfileLibraryService.MaskEmail(email)));
            builder.AppendLine("顶层字段：" + string.Join(", ", source.Properties().Select(property => property.Name).Take(24)));
            builder.AppendLine();
            builder.Append("敏感值已脱敏，不会在此窗口显示明文。");
            return builder.ToString();
        }

        private static string ResolveCpaLastRefresh(JObject source)
        {
            var raw = SelectJsonString(source, "last_refresh", "lastRefresh");
            if (!string.IsNullOrWhiteSpace(raw))
            {
                DateTime parsed;
                if (DateTime.TryParse(raw, null, System.Globalization.DateTimeStyles.AssumeUniversal | System.Globalization.DateTimeStyles.AdjustToUniversal, out parsed))
                {
                    return parsed.ToUniversalTime().ToString("O");
                }
            }

            return DateTime.UtcNow.ToString("O");
        }
        private static string ResolveCpaTokenEmail(JObject source, string accessToken, string idToken)
        {
            var email = SelectJsonString(source, "email", "profile.email", "account.email");
            if (!string.IsNullOrWhiteSpace(email))
            {
                return email.Trim();
            }

            var idPayload = TryReadJwtPayloadForCpa(idToken);
            email = SelectJsonString(idPayload, "email", "profile.email");
            if (!string.IsNullOrWhiteSpace(email))
            {
                return email.Trim();
            }

            var accessPayload = TryReadJwtPayloadForCpa(accessToken);
            email = SelectJsonString(accessPayload, "email", "profile.email");
            if (!string.IsNullOrWhiteSpace(email))
            {
                return email.Trim();
            }

            return SelectOpenAiJwtClaimString(accessPayload, "https://api.openai.com/profile", "email");
        }

        private static string ResolveCpaTokenAccountId(JObject source, string accessToken)
        {
            var accountId = SelectJsonString(source, "account_id", "accountId", "chatgpt_account_id");
            if (!string.IsNullOrWhiteSpace(accountId))
            {
                return accountId.Trim();
            }

            var accessPayload = TryReadJwtPayloadForCpa(accessToken);
            accountId = SelectOpenAiJwtClaimString(accessPayload, "https://api.openai.com/auth", "chatgpt_account_id");
            if (!string.IsNullOrWhiteSpace(accountId))
            {
                return accountId.Trim();
            }

            accountId = SelectOpenAiJwtClaimString(accessPayload, "https://api.openai.com/auth", "user_id");
            return !string.IsNullOrWhiteSpace(accountId) ? accountId.Trim() : SelectJsonString(accessPayload, "sub");
        }

        private static string SelectOpenAiJwtClaimString(JObject payload, string claimName, string propertyName)
        {
            if (payload == null || string.IsNullOrWhiteSpace(claimName) || string.IsNullOrWhiteSpace(propertyName))
            {
                return string.Empty;
            }

            var claim = payload[claimName] as JObject;
            return claim?[propertyName]?.Value<string>()?.Trim() ?? string.Empty;
        }
        private static string SelectJsonString(JToken token, params string[] paths)
        {
            if (token == null || paths == null)
            {
                return string.Empty;
            }

            foreach (var path in paths)
            {
                if (string.IsNullOrWhiteSpace(path))
                {
                    continue;
                }

                string value = null;
                try
                {
                    value = token.SelectToken(path)?.Value<string>();
                }
                catch
                {
                    value = null;
                }

                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return string.Empty;
        }

        private static void AddJsonString(JObject target, string name, string value)
        {
            if (target == null || string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(value))
            {
                return;
            }

            target[name] = value;
        }

        private static JObject TryReadJwtPayloadForCpa(string jwt)
        {
            if (string.IsNullOrWhiteSpace(jwt))
            {
                return null;
            }

            var parts = jwt.Split('.');
            if (parts.Length < 2)
            {
                return null;
            }

            try
            {
                var json = Encoding.UTF8.GetString(Base64UrlDecode(parts[1]));
                return JObject.Parse(json);
            }
            catch
            {
                return null;
            }
        }

        private static byte[] Base64UrlDecode(string value)
        {
            var normalized = (value ?? string.Empty).Replace('-', '+').Replace('_', '/');
            switch (normalized.Length % 4)
            {
                case 2:
                    normalized += "==";
                    break;
                case 3:
                    normalized += "=";
                    break;
            }

            return Convert.FromBase64String(normalized);
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
            if (TryGetClipboardImage(out var image, out var clipboardError))
            {
                ShowScreenshotEditorWindow(image);
                return;
            }

            if (clipboardError != null)
            {
                AppLogService.Error(clipboardError, "Reading clipboard image failed.");
                SystemStatusMessage = "读取剪贴板失败，请重新复制图片后再试。";
                MessageBox.Show("剪贴板图片数据无效，请重新复制图片后再试。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            MessageBox.Show("剪贴板中没有图片", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private async Task PreviewCodexProfileDiffAsync(object parameter)
        {
            if (!(parameter is CodexProfileItem item))
            {
                return;
            }

            try
            {
                var targetFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    ".codex");
                var builder = new StringBuilder();
                builder.AppendLine("Codex 配置差异预览");
                builder.AppendLine("配置：" + (item.Name ?? string.Empty));
                builder.AppendLine("目标目录：" + targetFolder);
                builder.AppendLine();

                builder.AppendLine(await BuildCodexFileDiffSummaryAsync(item, CodexConfigProfileService.ConfigFileName, targetFolder));
                builder.AppendLine();
                builder.AppendLine(await BuildCodexFileDiffSummaryAsync(item, CodexConfigProfileService.AuthFileName, targetFolder));

                MessageBox.Show(builder.ToString(), "Codex 配置差异预览", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(new InvalidOperationException(ex.Message), "Previewing Codex profile diff failed for {Name} with {ErrorType}", SafeCodexLogName(item.Name), ex.GetType().Name);
                MessageBox.Show(ex.Message, "差异预览失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task<string> BuildCodexFileDiffSummaryAsync(CodexProfileItem item, string fileName, string targetFolder)
        {
            var profileBytes = await ResolveCodexProfileFileBytesAsync(item, fileName).ConfigureAwait(true);
            var targetPath = Path.Combine(targetFolder, fileName);
            var targetBytes = File.Exists(targetPath)
                ? await ReadAllBytesAsync(targetPath, CancellationToken.None).ConfigureAwait(true)
                : null;

            var builder = new StringBuilder();
            builder.AppendLine(fileName + "：");
            if (profileBytes == null)
            {
                builder.AppendLine("  - 配置记录中没有可比对内容。");
                return builder.ToString().TrimEnd();
            }

            if (targetBytes == null)
            {
                builder.AppendLine("  - 当前 Codex 目录中不存在该文件，将新增。");
                builder.AppendLine("  - 记录大小：" + FileSizeFormatter.Format(profileBytes.LongLength));
                return builder.ToString().TrimEnd();
            }

            if (profileBytes.SequenceEqual(targetBytes))
            {
                builder.AppendLine("  - 内容一致。");
                builder.AppendLine("  - 大小：" + FileSizeFormatter.Format(profileBytes.LongLength));
                return builder.ToString().TrimEnd();
            }

            var profileText = Encoding.UTF8.GetString(profileBytes);
            var targetText = Encoding.UTF8.GetString(targetBytes);
            var profileLines = SplitLines(profileText);
            var targetLines = SplitLines(targetText);
            var changedLines = CountChangedLinePositions(profileLines, targetLines);

            builder.AppendLine("  - 内容不同。");
            builder.AppendLine("  - 当前大小：" + FileSizeFormatter.Format(targetBytes.LongLength));
            builder.AppendLine("  - 记录大小：" + FileSizeFormatter.Format(profileBytes.LongLength));
            builder.AppendLine($"  - 行数：当前 {targetLines.Length} 行 / 记录 {profileLines.Length} 行 / 不同位置 {changedLines} 行。");
            return builder.ToString().TrimEnd();
        }

        private async Task<byte[]> ResolveCodexProfileFileBytesAsync(CodexProfileItem item, string fileName)
        {
            var protectedContent = string.Equals(fileName, CodexConfigProfileService.ConfigFileName, StringComparison.OrdinalIgnoreCase)
                ? item.ConfigTomlContentProtected
                : item.AuthJsonContentProtected;
            var bytes = CodexConfigProfileService.UnprotectBytesFromBase64(protectedContent);
            if (bytes != null)
            {
                return bytes;
            }

            var folder = NormalizeFolderPath(item.FolderPath);
            var path = string.IsNullOrWhiteSpace(folder) ? null : Path.Combine(folder, fileName);
            return !string.IsNullOrWhiteSpace(path) && File.Exists(path)
                ? await ReadAllBytesAsync(path, CancellationToken.None).ConfigureAwait(true)
                : null;
        }

        private static string BuildCodexProfileImportSummary(string profileName, byte[] configBytes, byte[] authBytes)
        {
            var builder = new StringBuilder();
            builder.AppendLine("即将导入 Codex 配置记录");
            builder.AppendLine("记录名：" + (profileName ?? string.Empty));
            builder.AppendLine("config.toml：" + FileSizeFormatter.Format(configBytes?.LongLength ?? 0));

            var configLines = ExtractCodexConfigSummaryLines(configBytes).ToList();
            if (configLines.Count == 0)
            {
                builder.AppendLine("- 未识别到常用配置键。");
            }
            else
            {
                foreach (var line in configLines.Take(10))
                {
                    builder.AppendLine("- " + line);
                }
            }

            builder.AppendLine();
            builder.AppendLine("auth.json：" + FileSizeFormatter.Format(authBytes?.LongLength ?? 0));
            var authLines = ExtractCodexAuthSummaryLines(authBytes).ToList();
            if (authLines.Count == 0)
            {
                builder.AppendLine("- 未识别到认证字段，或 auth.json 不是标准 JSON。");
            }
            else
            {
                foreach (var line in authLines.Take(12))
                {
                    builder.AppendLine("- " + line);
                }
            }

            builder.AppendLine();
            builder.Append("敏感值已脱敏，不会在此窗口显示明文。");
            return builder.ToString();
        }

        private static IEnumerable<string> ExtractCodexConfigSummaryLines(byte[] bytes)
        {
            var text = bytes == null ? string.Empty : Encoding.UTF8.GetString(bytes);
            var keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "model",
                "model_provider",
                "approval_policy",
                "sandbox_mode",
                "base_url",
                "experimental_use_exec_command_tool",
                "network_access"
            };

            foreach (var rawLine in SplitLines(text))
            {
                var line = (rawLine ?? string.Empty).Trim();
                if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                var index = line.IndexOf('=');
                if (index <= 0)
                {
                    continue;
                }

                var key = line.Substring(0, index).Trim();
                if (!keys.Contains(key) && !IsSensitiveCodexKey(key))
                {
                    continue;
                }

                var value = line.Substring(index + 1).Trim();
                yield return key + " = " + MaskCodexValue(key, value);
            }
        }

        private static IEnumerable<string> ExtractCodexAuthSummaryLines(byte[] bytes)
        {
            if (bytes == null || bytes.Length == 0)
            {
                yield break;
            }

            JToken token;
            try
            {
                token = JToken.Parse(Encoding.UTF8.GetString(bytes));
            }
            catch
            {
                yield break;
            }

            foreach (var item in EnumerateJsonFields(token, string.Empty).Take(24))
            {
                yield return item;
            }
        }

        private static IEnumerable<string> EnumerateJsonFields(JToken token, string path)
        {
            if (token is JObject obj)
            {
                foreach (var property in obj.Properties())
                {
                    var childPath = string.IsNullOrWhiteSpace(path) ? property.Name : path + "." + property.Name;
                    foreach (var item in EnumerateJsonFields(property.Value, childPath))
                    {
                        yield return item;
                    }
                }

                yield break;
            }

            if (token is JArray array)
            {
                yield return (string.IsNullOrWhiteSpace(path) ? "$" : path) + $"：数组（{array.Count} 项）";
                yield break;
            }

            yield return (string.IsNullOrWhiteSpace(path) ? "$" : path) + "：" + DescribeJsonToken(token);
        }

        private static string DescribeJsonToken(JToken token)
        {
            if (token == null || token.Type == JTokenType.Null)
            {
                return "空";
            }

            switch (token.Type)
            {
                case JTokenType.Boolean:
                    return "布尔值";
                case JTokenType.Integer:
                case JTokenType.Float:
                    return "数字";
                case JTokenType.Date:
                    return "日期";
                default:
                    return "已设置（已脱敏）";
            }
        }

        private static string MaskCodexValue(string key, string value)
        {
            if (IsSensitiveCodexKey(key))
            {
                return "***";
            }

            var normalized = value ?? string.Empty;
            if (normalized.Length <= 80)
            {
                return normalized;
            }

            return normalized.Substring(0, 77) + "...";
        }

        private static bool IsSensitiveCodexKey(string key)
        {
            var value = key ?? string.Empty;
            return value.IndexOf("key", StringComparison.OrdinalIgnoreCase) >= 0
                   || value.IndexOf("token", StringComparison.OrdinalIgnoreCase) >= 0
                   || value.IndexOf("secret", StringComparison.OrdinalIgnoreCase) >= 0
                   || value.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0
                   || value.IndexOf("credential", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string[] SplitLines(string text)
        {
            return (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        }

        private static int CountChangedLinePositions(string[] left, string[] right)
        {
            var length = Math.Max(left?.Length ?? 0, right?.Length ?? 0);
            var count = 0;
            for (var i = 0; i < length; i++)
            {
                var a = left != null && i < left.Length ? left[i] : string.Empty;
                var b = right != null && i < right.Length ? right[i] : string.Empty;
                if (!string.Equals(a, b, StringComparison.Ordinal))
                {
                    count++;
                }
            }

            return count;
        }

        private void OpenImageViewerFile()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择要查看的图片",
                Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif;*.tif;*.tiff|所有文件|*.*"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            LoadImageViewerFile(dialog.FileName);
        }

        private void LoadImageViewerFile(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    MessageBox.Show("图片文件不存在。", "图片查看", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var image = new BitmapImage();
                image.BeginInit();
                image.CacheOption = BitmapCacheOption.OnLoad;
                image.CreateOptions = BitmapCreateOptions.PreservePixelFormat;
                image.UriSource = new Uri(filePath, UriKind.Absolute);
                image.EndInit();
                image.Freeze();

                _imageViewerOriginalImage = image;
                ImageViewerFilePath = filePath;
                AddHomeRecentItem(Path.GetFileName(filePath), filePath, "ImageViewer", filePath, "看图");
                RefreshImageViewerDirectoryFiles(filePath);
                _imageViewerRotationDegrees = 0;
                _imageViewerFlipHorizontal = false;
                _imageViewerFlipVertical = false;
                _imageViewerCropMode = "原图";
                _imageViewerBrightness = 0;
                _imageViewerContrast = 0;
                _imageViewerSharpenAmount = 0;
                ImageViewerIsGrayscale = false;
                ImageViewerZoom = 1.0;
                OnPropertyChanged(nameof(ImageViewerCropMode));
                OnPropertyChanged(nameof(ImageViewerBrightness));
                OnPropertyChanged(nameof(ImageViewerBrightnessText));
                OnPropertyChanged(nameof(ImageViewerContrast));
                OnPropertyChanged(nameof(ImageViewerContrastText));
                OnPropertyChanged(nameof(ImageViewerSharpenAmount));
                OnPropertyChanged(nameof(ImageViewerSharpenText));

                UpdateImageViewerPreview($"已打开：{Path.GetFileName(filePath)}");
                AppLogService.Information("Image viewer opened {File}, {W}x{H}.", Path.GetFileName(filePath), image.PixelWidth, image.PixelHeight);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Image viewer open failed.");
                ImageViewerStatusMessage = "打开失败：" + ex.Message;
                MessageBox.Show("打开图片失败：" + ex.Message, "图片查看", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CopyImageViewerImage()
        {
            var image = CreateCurrentImageViewerBitmapSource();
            if (image == null)
            {
                return;
            }

            try
            {
                ScreenshotService.SetClipboardCompatible(image);
                ImageViewerStatusMessage = "已复制到剪贴板。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Image viewer copy failed.");
                ImageViewerStatusMessage = "复制失败：" + ex.Message;
                MessageBox.Show("复制图片失败：" + ex.Message, "图片查看", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void SaveImageViewerImageAs()
        {
            var image = CreateCurrentImageViewerBitmapSource();
            if (image == null)
            {
                return;
            }

            var baseName = string.IsNullOrWhiteSpace(ImageViewerFileName)
                ? "image"
                : Path.GetFileNameWithoutExtension(ImageViewerFileName);
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "另存图片",
                FileName = $"{baseName}_edited.png",
                Filter = "PNG 图片|*.png|JPEG 图片|*.jpg|BMP 图片|*.bmp|TIFF 图片|*.tif"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                SaveBitmapSource(image, dialog.FileName);
                ImageViewerStatusMessage = "已保存：" + dialog.FileName;
                AppLogService.Information("Image viewer saved {File}.", Path.GetFileName(dialog.FileName));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Image viewer save failed.");
                ImageViewerStatusMessage = "保存失败：" + ex.Message;
                MessageBox.Show("保存图片失败：" + ex.Message, "图片查看", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void EditImageViewerImage()
        {
            var image = CreateCurrentImageViewerBitmapSource();
            if (image == null)
            {
                return;
            }

            ShowScreenshotEditorWindow(image);
            ImageViewerStatusMessage = "已打开标注编辑器。";
        }

        private bool CanOpenImageViewerSibling(int delta)
        {
            if (_imageViewerDirectoryFiles == null || _imageViewerDirectoryFiles.Count <= 1 || _imageViewerDirectoryIndex < 0)
            {
                return false;
            }

            var nextIndex = _imageViewerDirectoryIndex + delta;
            return nextIndex >= 0 && nextIndex < _imageViewerDirectoryFiles.Count;
        }

        private void OpenImageViewerSibling(int delta)
        {
            if (!CanOpenImageViewerSibling(delta))
            {
                return;
            }

            LoadImageViewerFile(_imageViewerDirectoryFiles[_imageViewerDirectoryIndex + delta]);
        }

        private void RefreshImageViewerDirectoryFiles(string filePath)
        {
            try
            {
                var folder = Path.GetDirectoryName(filePath);
                if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                {
                    _imageViewerDirectoryFiles = new List<string>();
                    _imageViewerDirectoryIndex = -1;
                    OnPropertyChanged(nameof(ImageViewerDirectoryPositionText));
                    TriggerCommandRequery();
                    return;
                }

                _imageViewerDirectoryFiles = Directory.EnumerateFiles(folder)
                    .Where(path => IsSupportedImageViewerFile(Path.GetExtension(path)))
                    .OrderBy(path => Path.GetFileName(path), StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
                _imageViewerDirectoryIndex = _imageViewerDirectoryFiles.FindIndex(path =>
                    string.Equals(path, filePath, StringComparison.OrdinalIgnoreCase));
                OnPropertyChanged(nameof(ImageViewerDirectoryPositionText));
                TriggerCommandRequery();
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Refresh image directory list failed: {Msg}", ex.Message);
                _imageViewerDirectoryFiles = new List<string>();
                _imageViewerDirectoryIndex = -1;
                OnPropertyChanged(nameof(ImageViewerDirectoryPositionText));
                TriggerCommandRequery();
            }
        }

        private void ToggleImageViewerFitToFrame()
        {
            _imageViewerFitToFrame = !_imageViewerFitToFrame;
            if (_imageViewerFitToFrame)
            {
                _imageViewerZoom = 1.0;
                OnPropertyChanged(nameof(ImageViewerZoom));
                OnPropertyChanged(nameof(ImageViewerZoomText));
            }
            OnPropertyChanged(nameof(ImageViewerFitToFrame));
            UpdateImageViewerPreview(_imageViewerFitToFrame ? "已满框显示。" : "已还原实际大小。");
        }

        public void OnImageFolderTreeSelected(ImageFolderNode node)
        {
            if (node == null || string.IsNullOrEmpty(node.FullPath) || !Directory.Exists(node.FullPath))
                return;
            try
            {
                var files = Directory.EnumerateFiles(node.FullPath)
                    .Where(f => IsSupportedImageViewerFile(Path.GetExtension(f)))
                    .OrderBy(f => Path.GetFileName(f), StringComparer.CurrentCultureIgnoreCase)
                    .ToList();
                if (files.Count == 0)
                {
                    ImageViewerStatusMessage = $"文件夹「{node.Name}」中没有图片。";
                    return;
                }
                LoadImageViewerFile(files[0]);
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Image folder tree select failed: {Msg}", ex.Message);
                ImageViewerStatusMessage = "读取文件夹失败：" + ex.Message;
            }
        }

        private void RotateImageViewer(int deltaDegrees)
        {
            _imageViewerRotationDegrees = NormalizeRotation(_imageViewerRotationDegrees + deltaDegrees);
            UpdateImageViewerPreview(deltaDegrees < 0 ? "已向左旋转。" : "已向右旋转。");
        }

        private void ResetImageViewerAdjustments()
        {
            _imageViewerCropMode = "原图";
            _imageViewerBrightness = 0;
            _imageViewerContrast = 0;
            _imageViewerSharpenAmount = 0;
            ImageViewerIsGrayscale = false;
            OnPropertyChanged(nameof(ImageViewerCropMode));
            OnPropertyChanged(nameof(ImageViewerBrightness));
            OnPropertyChanged(nameof(ImageViewerBrightnessText));
            OnPropertyChanged(nameof(ImageViewerContrast));
            OnPropertyChanged(nameof(ImageViewerContrastText));
            OnPropertyChanged(nameof(ImageViewerSharpenAmount));
            OnPropertyChanged(nameof(ImageViewerSharpenText));
            UpdateImageViewerPreview("已重置轻编辑效果。");
        }

        private void UpdateImageViewerPreview(string statusMessage = null)
        {
            try
            {
                ImageViewerPreviewImage = CreateCurrentImageViewerBitmapSource();
                if (!string.IsNullOrWhiteSpace(statusMessage))
                {
                    ImageViewerStatusMessage = statusMessage;
                }

                OnPropertyChanged(nameof(ImageViewerTransformText));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Image viewer preview update failed.");
                ImageViewerStatusMessage = "处理失败：" + ex.Message;
                MessageBox.Show("处理图片失败：" + ex.Message, "图片查看", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private BitmapSource CreateCurrentImageViewerBitmapSource()
        {
            if (_imageViewerOriginalImage == null)
            {
                return null;
            }

            using (var source = BitmapSourceToBitmap(_imageViewerOriginalImage))
            {
                var bitmap = new Drawing.Bitmap(source.Width, source.Height, DrawingImaging.PixelFormat.Format32bppArgb);
                Drawing.Bitmap working = bitmap;
                try
                {
                    using (var graphics = Drawing.Graphics.FromImage(bitmap))
                    {
                        graphics.CompositingQuality = Drawing.Drawing2D.CompositingQuality.HighQuality;
                        graphics.InterpolationMode = Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        graphics.SmoothingMode = Drawing.Drawing2D.SmoothingMode.HighQuality;
                        graphics.DrawImage(source, 0, 0, source.Width, source.Height);
                    }

                    ApplyImageViewerTransforms(bitmap);
                    working = ApplyImageViewerCrop(bitmap);
                    ApplyImageAdjustments(working);
                    if (ImageViewerIsGrayscale)
                    {
                        ApplyGrayscale(working);
                    }

                    return ScreenshotService.ConvertToBitmapSource(working);
                }
                finally
                {
                    if (!ReferenceEquals(working, bitmap))
                    {
                        working.Dispose();
                    }

                    bitmap.Dispose();
                }
            }
        }

        private void ApplyImageViewerTransforms(Drawing.Bitmap bitmap)
        {
            switch (NormalizeRotation(_imageViewerRotationDegrees))
            {
                case 90:
                    bitmap.RotateFlip(Drawing.RotateFlipType.Rotate90FlipNone);
                    break;
                case 180:
                    bitmap.RotateFlip(Drawing.RotateFlipType.Rotate180FlipNone);
                    break;
                case 270:
                    bitmap.RotateFlip(Drawing.RotateFlipType.Rotate270FlipNone);
                    break;
            }

            if (_imageViewerFlipHorizontal)
            {
                bitmap.RotateFlip(Drawing.RotateFlipType.RotateNoneFlipX);
            }

            if (_imageViewerFlipVertical)
            {
                bitmap.RotateFlip(Drawing.RotateFlipType.RotateNoneFlipY);
            }
        }

        private Drawing.Bitmap ApplyImageViewerCrop(Drawing.Bitmap bitmap)
        {
            if (bitmap == null || string.Equals(_imageViewerCropMode, "原图", StringComparison.Ordinal))
            {
                return bitmap;
            }

            double ratio;
            switch (_imageViewerCropMode)
            {
                case "1:1":
                    ratio = 1.0;
                    break;
                case "4:3":
                    ratio = 4.0 / 3.0;
                    break;
                case "16:9":
                    ratio = 16.0 / 9.0;
                    break;
                default:
                    return bitmap;
            }

            var sourceRatio = bitmap.Width / (double)bitmap.Height;
            var width = bitmap.Width;
            var height = bitmap.Height;
            if (sourceRatio > ratio)
            {
                width = Math.Max(1, (int)Math.Round(bitmap.Height * ratio));
            }
            else if (sourceRatio < ratio)
            {
                height = Math.Max(1, (int)Math.Round(bitmap.Width / ratio));
            }
            else
            {
                return bitmap;
            }

            var rect = new Drawing.Rectangle(
                Math.Max(0, (bitmap.Width - width) / 2),
                Math.Max(0, (bitmap.Height - height) / 2),
                Math.Min(width, bitmap.Width),
                Math.Min(height, bitmap.Height));

            return bitmap.Clone(rect, DrawingImaging.PixelFormat.Format32bppArgb);
        }

        private void ApplyImageAdjustments(Drawing.Bitmap bitmap)
        {
            if (bitmap == null)
            {
                return;
            }


            if (_imageViewerBrightness != 0 || _imageViewerContrast != 0)
            {
                ApplyBrightnessContrast(bitmap, _imageViewerBrightness, _imageViewerContrast);
            }

            if (_imageViewerSharpenAmount > 0)
            {
                ApplySharpen(bitmap, _imageViewerSharpenAmount);
            }
        }

        private static Drawing.Bitmap BitmapSourceToBitmap(BitmapSource source)
        {
            using (var stream = new MemoryStream())
            {
                var encoder = new PngBitmapEncoder();
                encoder.Frames.Add(BitmapFrame.Create(source));
                encoder.Save(stream);
                stream.Position = 0;
                using (var temp = new Drawing.Bitmap(stream))
                {
                    return new Drawing.Bitmap(temp);
                }
            }
        }

        private static void ApplyGrayscale(Drawing.Bitmap bitmap)
        {
            var rect = new Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(rect, DrawingImaging.ImageLockMode.ReadWrite, DrawingImaging.PixelFormat.Format32bppArgb);
            try
            {
                var byteCount = Math.Abs(data.Stride) * bitmap.Height;
                var pixels = new byte[byteCount];
                Marshal.Copy(data.Scan0, pixels, 0, byteCount);

                for (var y = 0; y < bitmap.Height; y++)
                {
                    var row = y * data.Stride;
                    for (var x = 0; x < bitmap.Width; x++)
                    {
                        var index = row + x * 4;
                        var blue = pixels[index];
                        var green = pixels[index + 1];
                        var red = pixels[index + 2];
                        var gray = (byte)Math.Min(255, (int)(red * 0.299 + green * 0.587 + blue * 0.114));
                        pixels[index] = gray;
                        pixels[index + 1] = gray;
                        pixels[index + 2] = gray;
                    }
                }

                Marshal.Copy(pixels, 0, data.Scan0, byteCount);
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }

        private static void ApplyBrightnessContrast(Drawing.Bitmap bitmap, int brightness, int contrast)
        {
            var brightnessValue = Math.Max(-100, Math.Min(100, brightness)) * 255 / 100;
            var contrastValue = Math.Max(-100, Math.Min(100, contrast));
            var factor = (259.0 * (contrastValue + 255.0)) / (255.0 * (259.0 - contrastValue));
            var rect = new Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(rect, DrawingImaging.ImageLockMode.ReadWrite, DrawingImaging.PixelFormat.Format32bppArgb);
            try
            {
                var byteCount = Math.Abs(data.Stride) * bitmap.Height;
                var pixels = new byte[byteCount];
                Marshal.Copy(data.Scan0, pixels, 0, byteCount);

                for (var y = 0; y < bitmap.Height; y++)
                {
                    var row = y * data.Stride;
                    for (var x = 0; x < bitmap.Width; x++)
                    {
                        var index = row + x * 4;
                        pixels[index] = AdjustColorChannel(pixels[index], brightnessValue, factor);
                        pixels[index + 1] = AdjustColorChannel(pixels[index + 1], brightnessValue, factor);
                        pixels[index + 2] = AdjustColorChannel(pixels[index + 2], brightnessValue, factor);
                    }
                }

                Marshal.Copy(pixels, 0, data.Scan0, byteCount);
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }

        private static byte AdjustColorChannel(byte value, int brightness, double contrastFactor)
        {
            var adjusted = contrastFactor * (value - 128) + 128 + brightness;
            if (adjusted < 0) return 0;
            if (adjusted > 255) return 255;
            return (byte)adjusted;
        }

        private static void ApplySharpen(Drawing.Bitmap bitmap, int amount)
        {
            var strength = Math.Max(0, Math.Min(100, amount)) / 100.0;
            if (strength <= 0)
            {
                return;
            }

            var rect = new Drawing.Rectangle(0, 0, bitmap.Width, bitmap.Height);
            var data = bitmap.LockBits(rect, DrawingImaging.ImageLockMode.ReadWrite, DrawingImaging.PixelFormat.Format32bppArgb);
            try
            {
                var stride = data.Stride;
                var byteCount = Math.Abs(stride) * bitmap.Height;
                var source = new byte[byteCount];
                var target = new byte[byteCount];
                Marshal.Copy(data.Scan0, source, 0, byteCount);
                Buffer.BlockCopy(source, 0, target, 0, byteCount);

                for (var y = 1; y < bitmap.Height - 1; y++)
                {
                    for (var x = 1; x < bitmap.Width - 1; x++)
                    {
                        var index = y * stride + x * 4;
                        for (var channel = 0; channel < 3; channel++)
                        {
                            var center = source[index + channel];
                            var blur = (
                                source[(y - 1) * stride + x * 4 + channel]
                                + source[(y + 1) * stride + x * 4 + channel]
                                + source[y * stride + (x - 1) * 4 + channel]
                                + source[y * stride + (x + 1) * 4 + channel]) / 4.0;
                            var sharpened = center + (center - blur) * strength;
                            target[index + channel] = ClampToByte(sharpened);
                        }
                    }
                }

                Marshal.Copy(target, 0, data.Scan0, byteCount);
            }
            finally
            {
                bitmap.UnlockBits(data);
            }
        }

        private static byte ClampToByte(double value)
        {
            if (value < 0) return 0;
            if (value > 255) return 255;
            return (byte)value;
        }

        private static void SaveBitmapSource(BitmapSource image, string outputPath)
        {
            var extension = Path.GetExtension(outputPath)?.ToLowerInvariant();
            BitmapEncoder encoder;
            BitmapSource frameSource = image;

            switch (extension)
            {
                case ".jpg":
                case ".jpeg":
                    encoder = new JpegBitmapEncoder { QualityLevel = 92 };
                    frameSource = new FormatConvertedBitmap(image, Media.PixelFormats.Bgr24, null, 0);
                    break;
                case ".bmp":
                    encoder = new BmpBitmapEncoder();
                    break;
                case ".tif":
                case ".tiff":
                    encoder = new TiffBitmapEncoder();
                    break;
                default:
                    encoder = new PngBitmapEncoder();
                    break;
            }

            encoder.Frames.Add(BitmapFrame.Create(frameSource));
            using (var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                encoder.Save(stream);
            }
        }

        private static int NormalizeRotation(int degrees)
        {
            var normalized = degrees % 360;
            if (normalized < 0)
            {
                normalized += 360;
            }

            return normalized;
        }

        private static bool IsSupportedImageViewerFile(string extension)
        {
            switch ((extension ?? string.Empty).ToLowerInvariant())
            {
                case ".png":
                case ".jpg":
                case ".jpeg":
                case ".bmp":
                case ".gif":
                case ".tif":
                case ".tiff":
                    return true;
                default:
                    return false;
            }
        }

        private void OpenVideoViewerFile()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择要播放的视频",
                Filter = "音视频文件|*.mp4;*.m4v;*.mov;*.wmv;*.avi;*.mkv;*.webm;*.mpg;*.mpeg;*.flv;*.mp3;*.wav;*.wma;*.m4a;*.aac;*.flac;*.ogg;*.opus|所有文件|*.*",
                Multiselect = true
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            LoadVideoViewerPlaylist(dialog.FileNames ?? new[] { dialog.FileName }, 0);
        }

        private void LoadVideoViewerPlaylist(IEnumerable<string> filePaths, int index)
        {
            var files = (filePaths ?? Enumerable.Empty<string>())
                .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path) && IsSupportedVideoViewerFile(Path.GetExtension(path)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (files.Count == 0)
            {
                VideoViewerStatusMessage = "未找到可播放的音视频文件。";
                return;
            }

            _videoViewerPlaylist.Clear();
            for (var i = 0; i < files.Count; i++)
            {
                _videoViewerPlaylist.Add(new VideoPlaylistItem(files[i], i + 1));
            }

            _videoViewerPlaylistIndex = Math.Max(0, Math.Min(index, _videoViewerPlaylist.Count - 1));
            RefreshVideoViewerPlaylistState();
            FilteredVideoViewerPlaylistView?.Refresh();
            OnPropertyChanged(nameof(VideoViewerPlaylistText));
            OnPropertyChanged(nameof(VideoViewerPlaylistFilterText));
            LoadVideoViewerFile(_videoViewerPlaylist[_videoViewerPlaylistIndex].FilePath);
            TriggerCommandRequery();
        }

        private void SaveVideoViewerPlaylist()
        {
            if (_videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0)
            {
                return;
            }

            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "保存播放列表",
                Filter = "M3U8 播放列表 (*.m3u8)|*.m3u8|M3U 播放列表 (*.m3u)|*.m3u|所有文件|*.*",
                DefaultExt = ".m3u8",
                AddExtension = true,
                FileName = "MyTools_Playlist.m3u8",
                OverwritePrompt = true
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                var baseDirectory = Path.GetDirectoryName(dialog.FileName);
                var builder = new StringBuilder();
                var exportedLoopRangeCount = 0;
                builder.AppendLine("#EXTM3U");
                builder.AppendLine("#PLAYLIST:MyTools");
                foreach (var item in _videoViewerPlaylist)
                {
                    exportedLoopRangeCount += AppendVideoViewerLoopRangePlaylistComments(builder, item.FilePath);
                    builder.AppendLine("#EXTINF:-1," + item.FileName);
                    builder.AppendLine(MakePlaylistPath(item.FilePath, baseDirectory));
                }

                File.WriteAllText(dialog.FileName, builder.ToString(), new UTF8Encoding(false));
                VideoViewerStatusMessage = exportedLoopRangeCount > 0
                    ? $"播放列表已保存：{dialog.FileName}（含 {exportedLoopRangeCount} 个 A/B 区间备注）"
                    : "播放列表已保存：" + dialog.FileName;
                SafeFireAndForget(AddRecentVideoViewerPlaylistAsync(dialog.FileName));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Save video playlist failed.");
                VideoViewerStatusMessage = "保存播放列表失败：" + ex.Message;
                MessageBox.Show("保存播放列表失败：" + ex.Message, "音视频播放", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private int AppendVideoViewerLoopRangePlaylistComments(StringBuilder builder, string filePath)
        {
            if (builder == null ||
                VideoViewerLoopRanges == null ||
                VideoViewerLoopRanges.Count == 0 ||
                string.IsNullOrWhiteSpace(filePath))
            {
                return 0;
            }

            var ranges = VideoViewerLoopRanges
                .Where(range => range != null &&
                                range.EndSeconds > range.StartSeconds + 0.2 &&
                                string.Equals(range.FilePath, filePath, StringComparison.OrdinalIgnoreCase))
                .OrderBy(range => range.StartSeconds)
                .ThenBy(range => range.EndSeconds)
                .ToList();
            foreach (var range in ranges)
            {
                builder.AppendLine(BuildVideoViewerLoopRangePlaylistComment(range));
            }

            return ranges.Count;
        }

        private static string BuildVideoViewerLoopRangePlaylistComment(VideoLoopRangeItem range)
        {
            if (range == null)
            {
                return "#MYTOOLS-AB:";
            }

            return string.Format(
                CultureInfo.InvariantCulture,
                "#MYTOOLS-AB:START={0:0.###};END={1:0.###};RANGE={2};DURATION={3}",
                range.StartSeconds,
                range.EndSeconds,
                EscapePlaylistCommentValue(range.RangeText),
                EscapePlaylistCommentValue(range.DurationText));
        }

        private static string EscapePlaylistCommentValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return value
                .Replace('\r', ' ')
                .Replace('\n', ' ')
                .Replace(';', '，')
                .Trim();
        }

        private void LoadVideoViewerPlaylistFromFile()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "载入播放列表",
                Filter = "播放列表文件|*.m3u;*.m3u8|所有文件|*.*",
                Multiselect = false
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                var files = ReadPlaylistFile(dialog.FileName)
                    .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path) && IsSupportedVideoViewerFile(Path.GetExtension(path)))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (files.Count == 0)
                {
                    VideoViewerStatusMessage = "播放列表中没有可播放的有效文件。";
                    MessageBox.Show("播放列表中没有可播放的有效文件。", "音视频播放", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                LoadVideoViewerPlaylist(files, 0);
                VideoViewerStatusMessage = $"已载入播放列表：{Path.GetFileName(dialog.FileName)}（{files.Count} 项）";
                SafeFireAndForget(AddRecentVideoViewerPlaylistAsync(dialog.FileName));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Load video playlist failed.");
                VideoViewerStatusMessage = "载入播放列表失败：" + ex.Message;
                MessageBox.Show("载入播放列表失败：" + ex.Message, "音视频播放", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenRecentVideoViewerPlaylist(object parameter)
        {
            var item = parameter as RecentPlaylistItem;
            if (item == null || string.IsNullOrWhiteSpace(item.FilePath))
            {
                return;
            }

            if (!File.Exists(item.FilePath))
            {
                VideoViewerRecentPlaylists.Remove(item);
                OnPropertyChanged(nameof(HasVideoViewerRecentPlaylists));
                SafeFireAndForget(SaveRecentVideoViewerPlaylistsAsync());
                VideoViewerStatusMessage = "最近播放列表文件不存在，已移除。";
                return;
            }

            try
            {
                var files = ReadPlaylistFile(item.FilePath)
                    .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path) && IsSupportedVideoViewerFile(Path.GetExtension(path)))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (files.Count == 0)
                {
                    VideoViewerStatusMessage = "最近播放列表中没有可播放的有效文件。";
                    return;
                }

                LoadVideoViewerPlaylist(files, 0);
                VideoViewerStatusMessage = $"已载入最近播放列表：{item.Title}（{files.Count} 项）";
                SafeFireAndForget(AddRecentVideoViewerPlaylistAsync(item.FilePath));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Open recent video playlist failed.");
                VideoViewerStatusMessage = "载入最近播放列表失败：" + ex.Message;
            }
        }

        private async Task LoadFavoriteVideoViewerPlaylistsAsync()
        {
            var settings = await AppSettingsService.LoadAsync();
            var items = (settings.FavoritePlaylists ?? new List<FavoritePlaylistSettings>())
                .Where(item => item != null && item.FilePaths != null && item.FilePaths.Count > 0)
                .OrderByDescending(item => item.LastUsedAt == default(DateTime) ? item.CreatedAt : item.LastUsedAt)
                .Take(8)
                .Select(item => new FavoritePlaylistItem(
                    string.IsNullOrWhiteSpace(item.Name) ? "收藏播放列表" : item.Name,
                    item.FilePaths,
                    item.CreatedAt,
                    item.LastUsedAt))
                .ToList();

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                ReplaceItems(VideoViewerFavoritePlaylists, items);
                OnPropertyChanged(nameof(HasVideoViewerFavoritePlaylists));
            });
        }

        private void FavoriteVideoViewerPlaylist()
        {
            if (_videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0)
            {
                return;
            }

            var titleSource = _videoViewerPlaylistIndex >= 0 && _videoViewerPlaylistIndex < _videoViewerPlaylist.Count
                ? _videoViewerPlaylist[_videoViewerPlaylistIndex].FileName
                : _videoViewerPlaylist[0].FileName;
            var defaultName = $"{Path.GetFileNameWithoutExtension(titleSource)} 等 {_videoViewerPlaylist.Count} 项";
            var name = Interaction.InputBox("请输入收藏名称：", "收藏播放列表", defaultName);
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            var filePaths = _videoViewerPlaylist
                .Select(item => item.FilePath)
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (filePaths.Count == 0)
            {
                VideoViewerStatusMessage = "当前播放列表没有可收藏的文件。";
                return;
            }

            var now = DateTime.Now;
            var favorite = new FavoritePlaylistItem(name.Trim(), filePaths, now, now);
            var existing = VideoViewerFavoritePlaylists.FirstOrDefault(item => item.HasSameFiles(filePaths));
            if (existing != null)
            {
                VideoViewerFavoritePlaylists.Remove(existing);
            }

            VideoViewerFavoritePlaylists.Insert(0, favorite);
            while (VideoViewerFavoritePlaylists.Count > 8)
            {
                VideoViewerFavoritePlaylists.RemoveAt(VideoViewerFavoritePlaylists.Count - 1);
            }

            OnPropertyChanged(nameof(HasVideoViewerFavoritePlaylists));
            SafeFireAndForget(SaveFavoriteVideoViewerPlaylistsAsync());
            VideoViewerStatusMessage = $"已收藏播放列表：{favorite.Title}（{favorite.CountText}）";
        }

        private void OpenFavoriteVideoViewerPlaylist(object parameter)
        {
            var item = parameter as FavoritePlaylistItem;
            if (item == null || item.FilePaths == null || item.FilePaths.Count == 0)
            {
                return;
            }

            var files = item.FilePaths
                .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path) && IsSupportedVideoViewerFile(Path.GetExtension(path)))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (files.Count == 0)
            {
                VideoViewerStatusMessage = "收藏播放列表中没有可播放的有效文件。";
                MessageBox.Show("收藏播放列表中没有可播放的有效文件。", "音视频播放", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            item.LastUsedAt = DateTime.Now;
            var currentIndex = VideoViewerFavoritePlaylists.IndexOf(item);
            if (currentIndex > 0)
            {
                VideoViewerFavoritePlaylists.Move(currentIndex, 0);
            }

            LoadVideoViewerPlaylist(files, 0);
            VideoViewerStatusMessage = $"已载入收藏播放列表：{item.Title}（{files.Count} 项）";
            SafeFireAndForget(SaveFavoriteVideoViewerPlaylistsAsync());
        }

        private void RemoveFavoriteVideoViewerPlaylist(object parameter)
        {
            var item = parameter as FavoritePlaylistItem;
            if (item == null)
            {
                return;
            }

            VideoViewerFavoritePlaylists.Remove(item);
            OnPropertyChanged(nameof(HasVideoViewerFavoritePlaylists));
            SafeFireAndForget(SaveFavoriteVideoViewerPlaylistsAsync());
            VideoViewerStatusMessage = "已移除收藏播放列表：" + item.Title;
        }

        private async Task SaveFavoriteVideoViewerPlaylistsAsync()
        {
            var snapshot = VideoViewerFavoritePlaylists
                .Select(item => new FavoritePlaylistSettings
                {
                    Name = item.Title,
                    FilePaths = item.FilePaths.ToList(),
                    CreatedAt = item.CreatedAt,
                    LastUsedAt = item.LastUsedAt
                })
                .ToList();
            await AppSettingsService.UpdateAsync(settings => settings.FavoritePlaylists = snapshot);
        }

        private void SetVideoViewerLoopPoint(bool isStart)
        {
            if (!HasVideoViewerVideo)
            {
                return;
            }

            var seconds = ClampVideoPointSeconds(VideoViewerPositionSeconds);
            if (isStart)
            {
                VideoViewerLoopStartSeconds = seconds;
                if (VideoViewerLoopEndSeconds >= 0 && VideoViewerLoopEndSeconds <= seconds + 0.2)
                {
                    VideoViewerLoopEndSeconds = -1;
                }

                VideoViewerStatusMessage = "已设置 A 点：" + VideoViewerLoopStartText;
            }
            else
            {
                if (VideoViewerLoopStartSeconds >= 0 && seconds <= VideoViewerLoopStartSeconds + 0.2)
                {
                    VideoViewerStatusMessage = "B 点需要晚于 A 点。";
                    return;
                }

                VideoViewerLoopEndSeconds = seconds;
                VideoViewerStatusMessage = "已设置 B 点：" + VideoViewerLoopEndText;
            }

            if (HasVideoViewerLoopRange)
            {
                VideoViewerIsLoopEnabled = true;
                VideoViewerStatusMessage = "已启用 A/B 循环：" + VideoViewerLoopRangeText;
            }
        }

        private void ClearVideoViewerLoop()
        {
            VideoViewerIsLoopEnabled = false;
            VideoViewerLoopStartSeconds = -1;
            VideoViewerLoopEndSeconds = -1;
            VideoViewerStatusMessage = "已清除 A/B 循环。";
        }

        public bool SetVideoViewerLoopRange(double startSeconds, double endSeconds, bool enableLoop)
        {
            if (!HasVideoViewerVideo)
            {
                return false;
            }

            var start = ClampVideoPointSeconds(Math.Min(startSeconds, endSeconds));
            var end = ClampVideoPointSeconds(Math.Max(startSeconds, endSeconds));
            if (end <= start + 0.2)
            {
                VideoViewerStatusMessage = "波形选区太短，至少需要 0.2 秒。";
                return false;
            }

            VideoViewerLoopStartSeconds = start;
            VideoViewerLoopEndSeconds = end;
            VideoViewerIsLoopEnabled = enableLoop;
            VideoViewerStatusMessage = enableLoop
                ? "已按波形选区启用 A/B 循环：" + VideoViewerLoopRangeText
                : "已按波形选区设置 A/B：" + VideoViewerLoopRangeText;
            return true;
        }

        public bool ShouldLoopVideoViewerAt(double positionSeconds, out double targetSeconds)
        {
            targetSeconds = 0;
            if (!VideoViewerIsLoopEnabled || !HasVideoViewerLoopRange)
            {
                return false;
            }

            if (positionSeconds >= VideoViewerLoopEndSeconds)
            {
                targetSeconds = VideoViewerLoopStartSeconds;
                return true;
            }

            return false;
        }

        private void SaveVideoViewerLoopRange()
        {
            if (!HasVideoViewerLoopRange)
            {
                return;
            }

            var start = VideoViewerLoopStartSeconds;
            var end = VideoViewerLoopEndSeconds;
            var existing = VideoViewerLoopRanges
                .FirstOrDefault(range => Math.Abs(range.StartSeconds - start) < 0.5 && Math.Abs(range.EndSeconds - end) < 0.5);
            if (existing != null)
            {
                VideoViewerLoopRanges.Remove(existing);
            }

            var item = new VideoLoopRangeItem(start, end, VideoViewerFileName, VideoViewerFilePath);
            VideoViewerLoopRanges.Add(item);
            var ordered = VideoViewerLoopRanges
                .OrderBy(range => range.StartSeconds)
                .ThenBy(range => range.EndSeconds)
                .ToList();
            ReplaceItems(VideoViewerLoopRanges, ordered);
            OnPropertyChanged(nameof(HasVideoViewerLoopRanges));
            VideoViewerStatusMessage = "已保存 A/B 区间：" + item.RangeText;
            TriggerCommandRequery();
        }

        private void OpenVideoViewerLoopRange(object parameter)
        {
            if (!(parameter is VideoLoopRangeItem item) || !HasVideoViewerVideo)
            {
                return;
            }

            if (!SetVideoViewerLoopRange(item.StartSeconds, item.EndSeconds, true))
            {
                return;
            }

            VideoViewerPositionSeconds = ClampVideoPointSeconds(item.StartSeconds);
            UpdateVideoViewerSubtitle(VideoViewerPositionSeconds);
            VideoViewerStatusMessage = "已套用 A/B 区间：" + item.RangeText;
        }

        private void RemoveVideoViewerLoopRange(object parameter)
        {
            var item = parameter as VideoLoopRangeItem;
            if (item == null)
            {
                return;
            }

            VideoViewerLoopRanges.Remove(item);
            OnPropertyChanged(nameof(HasVideoViewerLoopRanges));
            VideoViewerStatusMessage = "已移除 A/B 区间：" + item.RangeText;
            TriggerCommandRequery();
        }

        private void AddVideoViewerBookmark()
        {
            if (!HasVideoViewerVideo)
            {
                return;
            }

            var seconds = ClampVideoPointSeconds(VideoViewerPositionSeconds);
            var existing = VideoViewerBookmarks
                .FirstOrDefault(item => Math.Abs(item.PositionSeconds - seconds) < 0.5);
            if (existing != null)
            {
                VideoViewerBookmarks.Remove(existing);
            }

            VideoViewerBookmarks.Add(new VideoBookmarkItem(seconds, VideoViewerFileName, VideoViewerFilePath));
            var ordered = VideoViewerBookmarks.OrderBy(item => item.PositionSeconds).ToList();
            ReplaceItems(VideoViewerBookmarks, ordered);
            OnPropertyChanged(nameof(HasVideoViewerBookmarks));
            VideoViewerStatusMessage = "已添加书签：" + FormatVideoTime(seconds);
        }

        private void OpenVideoViewerBookmark(object parameter)
        {
            if (!(parameter is VideoBookmarkItem item) || !HasVideoViewerVideo)
            {
                return;
            }

            VideoViewerPositionSeconds = ClampVideoPointSeconds(item.PositionSeconds);
            UpdateVideoViewerSubtitle(VideoViewerPositionSeconds);
            VideoViewerStatusMessage = "已跳转到书签：" + item.TimeText;
        }

        private void RemoveVideoViewerBookmark(object parameter)
        {
            var item = parameter as VideoBookmarkItem;
            if (item == null)
            {
                return;
            }

            VideoViewerBookmarks.Remove(item);
            OnPropertyChanged(nameof(HasVideoViewerBookmarks));
            VideoViewerStatusMessage = "已移除书签：" + item.TimeText;
            TriggerCommandRequery();
        }

        private void ResetVideoViewerLoopAndBookmarks()
        {
            VideoViewerIsLoopEnabled = false;
            VideoViewerLoopStartSeconds = -1;
            VideoViewerLoopEndSeconds = -1;
            VideoViewerLoopRanges.Clear();
            OnPropertyChanged(nameof(HasVideoViewerLoopRanges));
            VideoViewerBookmarks.Clear();
            OnPropertyChanged(nameof(HasVideoViewerBookmarks));
        }

        private double ClampVideoPointSeconds(double seconds)
        {
            if (double.IsNaN(seconds) || double.IsInfinity(seconds))
            {
                seconds = 0;
            }

            if (VideoViewerDurationSeconds > 0)
            {
                seconds = Math.Min(VideoViewerDurationSeconds, seconds);
            }

            return Math.Max(0, seconds);
        }

        private static double NormalizeVideoPointSeconds(double seconds)
        {
            if (double.IsNaN(seconds) || double.IsInfinity(seconds) || seconds < 0)
            {
                return -1;
            }

            return Math.Round(seconds, 1);
        }

        private async Task LoadRecentVideoViewerPlaylistsAsync()
        {
            var settings = await AppSettingsService.LoadAsync();
            var items = (settings.RecentPlaylists ?? new List<RecentPlaylistSettings>())
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.FilePath))
                .OrderByDescending(item => item.LastUsedAt)
                .Take(6)
                .Select(item => new RecentPlaylistItem(item.FilePath, item.LastUsedAt))
                .ToList();

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                ReplaceItems(VideoViewerRecentPlaylists, items);
                OnPropertyChanged(nameof(HasVideoViewerRecentPlaylists));
            });
        }

        private async Task AddRecentVideoViewerPlaylistAsync(string playlistPath)
        {
            if (string.IsNullOrWhiteSpace(playlistPath))
            {
                return;
            }

            var normalizedPath = Path.GetFullPath(playlistPath);
            var now = DateTime.Now;
            var existing = VideoViewerRecentPlaylists
                .FirstOrDefault(item => string.Equals(item.FilePath, normalizedPath, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                VideoViewerRecentPlaylists.Remove(existing);
            }

            VideoViewerRecentPlaylists.Insert(0, new RecentPlaylistItem(normalizedPath, now));
            while (VideoViewerRecentPlaylists.Count > 6)
            {
                VideoViewerRecentPlaylists.RemoveAt(VideoViewerRecentPlaylists.Count - 1);
            }

            OnPropertyChanged(nameof(HasVideoViewerRecentPlaylists));
            await SaveRecentVideoViewerPlaylistsAsync();
        }

        private async Task SaveRecentVideoViewerPlaylistsAsync()
        {
            var snapshot = VideoViewerRecentPlaylists
                .Select(item => new RecentPlaylistSettings
                {
                    FilePath = item.FilePath,
                    LastUsedAt = item.LastUsedAt
                })
                .ToList();
            await AppSettingsService.UpdateAsync(settings => settings.RecentPlaylists = snapshot);
        }

        private void LoadVideoViewerFile(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    MessageBox.Show("视频文件不存在。", "视频查看", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                if (!IsSupportedVideoViewerFile(Path.GetExtension(filePath)))
                {
                    MessageBox.Show("暂不支持该音视频扩展名。", "视频查看", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                VideoViewerFilePath = filePath;
                VideoViewerSource = new Uri(filePath, UriKind.Absolute);
                IsVideoViewerPlaying = false;
                VideoViewerPositionSeconds = 0;
                VideoViewerDurationSeconds = 0;
                VideoViewerSpeedRatio = 1.0;
                ResetVideoViewerLoopAndBookmarks();
                ClearVideoViewerWaveform();
                VideoViewerStatusMessage = "已打开：" + Path.GetFileName(filePath);
                AddHomeRecentItem(Path.GetFileName(filePath), filePath, "VideoViewer", filePath, "播放");
                RefreshVideoViewerPlaylistState();
                OnPropertyChanged(nameof(VideoViewerPlaylistText));
                OnPropertyChanged(nameof(VideoViewerPlaylistFilterText));
                _ = TryLoadSiblingSubtitleAsync(filePath);
                TriggerCommandRequery();
                AppLogService.Information("Video viewer opened {File}.", Path.GetFileName(filePath));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Video viewer open failed.");
                VideoViewerStatusMessage = "打开失败：" + ex.Message;
                MessageBox.Show("打开视频失败：" + ex.Message, "视频查看", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool CanOpenVideoViewerPlaylistSibling(int delta)
        {
            if (_videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0 || _videoViewerPlaylistIndex < 0)
            {
                return false;
            }

            if (VideoViewerPlayMode == "列表循环" || VideoViewerPlayMode == "随机")
            {
                return _videoViewerPlaylist.Count > 1;
            }

            var next = _videoViewerPlaylistIndex + delta;
            return next >= 0 && next < _videoViewerPlaylist.Count;
        }

        private void OpenVideoViewerPlaylistSibling(int delta)
        {
            if (!TryMoveVideoViewerPlaylist(delta, false))
            {
                return;
            }
        }

        private void OpenVideoViewerPlaylistItem(object parameter)
        {
            if (!(parameter is VideoPlaylistItem item) || _videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0)
            {
                return;
            }

            var index = _videoViewerPlaylist.IndexOf(item);
            if (index < 0)
            {
                return;
            }

            _videoViewerPlaylistIndex = index;
            LoadVideoViewerFile(item.FilePath);
        }

        private void RemoveVideoViewerPlaylistItem(object parameter)
        {
            if (!(parameter is VideoPlaylistItem item) || _videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0)
            {
                return;
            }

            var index = _videoViewerPlaylist.IndexOf(item);
            if (index < 0)
            {
                return;
            }

            var removedName = item.FileName;
            var wasCurrent = index == _videoViewerPlaylistIndex;
            _videoViewerPlaylist.RemoveAt(index);

            if (_videoViewerPlaylist.Count == 0)
            {
                ResetVideoViewerPlaybackState("播放列表已清空。");
                return;
            }

            if (wasCurrent)
            {
                _videoViewerPlaylistIndex = Math.Min(index, _videoViewerPlaylist.Count - 1);
                var nextItem = _videoViewerPlaylist[_videoViewerPlaylistIndex];
                LoadVideoViewerFile(nextItem.FilePath);
                VideoViewerStatusMessage = $"已移除当前项，切换到：{nextItem.FileName}";
                return;
            }

            if (index < _videoViewerPlaylistIndex)
            {
                _videoViewerPlaylistIndex--;
            }

            RefreshVideoViewerPlaylistState();
            FilteredVideoViewerPlaylistView?.Refresh();
            OnPropertyChanged(nameof(VideoViewerPlaylistText));
            OnPropertyChanged(nameof(VideoViewerPlaylistFilterText));
            VideoViewerStatusMessage = "已从播放列表移除：" + removedName;
            TriggerCommandRequery();
        }

        private void CopyVideoViewerPlaylist()
        {
            if (_videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0)
            {
                return;
            }

            try
            {
                var text = string.Join(Environment.NewLine, _videoViewerPlaylist.Select(item => item.FilePath));
                Clipboard.SetText(text);
                VideoViewerStatusMessage = $"已复制播放列表路径（{_videoViewerPlaylist.Count} 项）。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Copy video playlist failed.");
                VideoViewerStatusMessage = "复制播放列表失败：" + ex.Message;
                MessageBox.Show("复制播放列表失败：" + ex.Message, "音视频播放", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void CleanInvalidVideoViewerPlaylistItems()
        {
            if (_videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0)
            {
                return;
            }

            var currentPath = _videoViewerPlaylistIndex >= 0 && _videoViewerPlaylistIndex < _videoViewerPlaylist.Count
                ? _videoViewerPlaylist[_videoViewerPlaylistIndex].FilePath
                : VideoViewerFilePath;
            var removedCount = 0;
            var removedCurrent = false;

            for (var i = _videoViewerPlaylist.Count - 1; i >= 0; i--)
            {
                var item = _videoViewerPlaylist[i];
                if (!File.Exists(item.FilePath) || !IsSupportedVideoViewerFile(Path.GetExtension(item.FilePath)))
                {
                    if (string.Equals(item.FilePath, currentPath, StringComparison.OrdinalIgnoreCase))
                    {
                        removedCurrent = true;
                    }

                    _videoViewerPlaylist.RemoveAt(i);
                    removedCount++;
                }
            }

            if (removedCount == 0)
            {
                RefreshVideoViewerPlaylistState();
                FilteredVideoViewerPlaylistView?.Refresh();
                OnPropertyChanged(nameof(VideoViewerPlaylistText));
                OnPropertyChanged(nameof(VideoViewerPlaylistFilterText));
                VideoViewerStatusMessage = "播放列表没有失效项。";
                TriggerCommandRequery();
                return;
            }

            if (_videoViewerPlaylist.Count == 0)
            {
                ResetVideoViewerPlaybackState($"已清理 {removedCount} 个失效项，播放列表已清空。");
                return;
            }

            _videoViewerPlaylistIndex = _videoViewerPlaylist
                .Select((playlistItem, index) => new { playlistItem, index })
                .Where(entry => string.Equals(entry.playlistItem.FilePath, currentPath, StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.index)
                .DefaultIfEmpty(Math.Min(Math.Max(_videoViewerPlaylistIndex, 0), _videoViewerPlaylist.Count - 1))
                .First();

            if (removedCurrent)
            {
                var nextItem = _videoViewerPlaylist[_videoViewerPlaylistIndex];
                LoadVideoViewerFile(nextItem.FilePath);
                VideoViewerStatusMessage = $"已清理 {removedCount} 个失效项，切换到：{nextItem.FileName}";
                return;
            }

            RefreshVideoViewerPlaylistState();
            FilteredVideoViewerPlaylistView?.Refresh();
            OnPropertyChanged(nameof(VideoViewerPlaylistText));
            OnPropertyChanged(nameof(VideoViewerPlaylistFilterText));
            VideoViewerStatusMessage = $"已清理 {removedCount} 个失效项。";
            TriggerCommandRequery();
        }

        private void ClearVideoViewerPlaylist()
        {
            if (_videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0)
            {
                return;
            }

            _videoViewerPlaylist.Clear();
            ResetVideoViewerPlaybackState("播放列表已清空。");
        }

        private void ResetVideoViewerPlaybackState(string status)
        {
            _videoViewerPlaylistIndex = -1;
            VideoViewerSource = null;
            VideoViewerFilePath = string.Empty;
            IsVideoViewerPlaying = false;
            VideoViewerPositionSeconds = 0;
            VideoViewerDurationSeconds = 0;
            ResetVideoViewerLoopAndBookmarks();
            ClearVideoViewerWaveform();
            RefreshVideoViewerPlaylistState();
            FilteredVideoViewerPlaylistView?.Refresh();
            OnPropertyChanged(nameof(VideoViewerPlaylistText));
            OnPropertyChanged(nameof(VideoViewerPlaylistFilterText));
            ClearVideoViewerSubtitle("未载入字幕");
            VideoViewerStatusMessage = string.IsNullOrWhiteSpace(status) ? "播放列表已清空。" : status;
            TriggerCommandRequery();
        }

        public void MoveVideoViewerPlaylistItem(VideoPlaylistItem item, VideoPlaylistItem target)
        {
            if (item == null || target == null || ReferenceEquals(item, target))
            {
                return;
            }

            var oldIndex = _videoViewerPlaylist.IndexOf(item);
            var newIndex = _videoViewerPlaylist.IndexOf(target);
            if (oldIndex < 0 || newIndex < 0 || oldIndex == newIndex)
            {
                return;
            }

            var currentPath = _videoViewerPlaylistIndex >= 0 && _videoViewerPlaylistIndex < _videoViewerPlaylist.Count
                ? _videoViewerPlaylist[_videoViewerPlaylistIndex].FilePath
                : VideoViewerFilePath;
            _videoViewerPlaylist.Move(oldIndex, newIndex);
            RefreshVideoViewerPlaylistNumbers();
            _videoViewerPlaylistIndex = _videoViewerPlaylist
                .Select((playlistItem, index) => new { playlistItem, index })
                .Where(entry => string.Equals(entry.playlistItem.FilePath, currentPath, StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.index)
                .DefaultIfEmpty(Math.Min(newIndex, _videoViewerPlaylist.Count - 1))
                .First();
            RefreshVideoViewerPlaylistState();
            FilteredVideoViewerPlaylistView?.Refresh();
            OnPropertyChanged(nameof(VideoViewerPlaylistText));
            OnPropertyChanged(nameof(VideoViewerPlaylistFilterText));
            VideoViewerStatusMessage = "已调整播放列表顺序。";
            TriggerCommandRequery();
        }

        public bool OpenNextVideoViewerPlaylistItem()
        {
            return TryMoveVideoViewerPlaylist(1, true);
        }

        private bool TryMoveVideoViewerPlaylist(int delta, bool isAutoAdvance)
        {
            if (_videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0 || _videoViewerPlaylistIndex < 0)
            {
                return false;
            }

            var next = ResolveNextVideoViewerPlaylistIndex(delta, isAutoAdvance);
            if (next < 0 || next >= _videoViewerPlaylist.Count)
            {
                return false;
            }

            _videoViewerPlaylistIndex = next;
            LoadVideoViewerFile(_videoViewerPlaylist[_videoViewerPlaylistIndex].FilePath);
            return true;
        }

        private int ResolveNextVideoViewerPlaylistIndex(int delta, bool isAutoAdvance)
        {
            if (_videoViewerPlaylist == null || _videoViewerPlaylist.Count == 0)
            {
                return -1;
            }

            if (VideoViewerPlayMode == "单项循环" && isAutoAdvance)
            {
                return _videoViewerPlaylistIndex;
            }

            if (VideoViewerPlayMode == "随机" && _videoViewerPlaylist.Count > 1)
            {
                int next;
                do
                {
                    next = _videoViewerRandom.Next(_videoViewerPlaylist.Count);
                } while (next == _videoViewerPlaylistIndex);
                return next;
            }

            var candidate = _videoViewerPlaylistIndex + delta;
            if (candidate >= 0 && candidate < _videoViewerPlaylist.Count)
            {
                return candidate;
            }

            if (VideoViewerPlayMode == "列表循环" && _videoViewerPlaylist.Count > 1)
            {
                return candidate < 0 ? _videoViewerPlaylist.Count - 1 : 0;
            }

            return -1;
        }

        private void RefreshVideoViewerPlaylistState()
        {
            RefreshVideoViewerPlaylistNumbers();
            for (var i = 0; i < _videoViewerPlaylist.Count; i++)
            {
                _videoViewerPlaylist[i].IsCurrent = i == _videoViewerPlaylistIndex;
            }
        }

        private void RefreshVideoViewerPlaylistNumbers()
        {
            for (var i = 0; i < _videoViewerPlaylist.Count; i++)
            {
                _videoViewerPlaylist[i].IndexText = (i + 1).ToString("00");
            }
        }

        private static IEnumerable<string> ReadPlaylistFile(string playlistPath)
        {
            var baseDirectory = Path.GetDirectoryName(playlistPath) ?? string.Empty;
            foreach (var rawLine in File.ReadLines(playlistPath, Encoding.UTF8))
            {
                var line = (rawLine ?? string.Empty).Trim();
                if (line.Length == 0 || line.StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                var candidate = line.Trim('"');
                if (!Path.IsPathRooted(candidate))
                {
                    candidate = Path.GetFullPath(Path.Combine(baseDirectory, candidate));
                }

                yield return candidate;
            }
        }

        private static string MakePlaylistPath(string filePath, string baseDirectory)
        {
            if (string.IsNullOrWhiteSpace(filePath) || string.IsNullOrWhiteSpace(baseDirectory))
            {
                return filePath ?? string.Empty;
            }

            try
            {
                var baseUri = new Uri(AppendDirectorySeparatorChar(baseDirectory));
                var fileUri = new Uri(filePath);
                var relative = Uri.UnescapeDataString(baseUri.MakeRelativeUri(fileUri).ToString())
                    .Replace('/', Path.DirectorySeparatorChar);
                return string.IsNullOrWhiteSpace(relative) || relative.StartsWith("..", StringComparison.Ordinal)
                    ? filePath
                    : relative;
            }
            catch
            {
                return filePath;
            }
        }

        private static string AppendDirectorySeparatorChar(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || path.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal))
            {
                return path;
            }

            return path + Path.DirectorySeparatorChar;
        }

        private void OpenVideoViewerInExternalPlayer()
        {
            if (string.IsNullOrWhiteSpace(VideoViewerFilePath) || !File.Exists(VideoViewerFilePath))
            {
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = VideoViewerFilePath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Open video in external player failed.");
                VideoViewerStatusMessage = "外部打开失败：" + ex.Message;
                MessageBox.Show("外部打开失败：" + ex.Message, "视频查看", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task CaptureVideoViewerFrameAsync()
        {
            if (!HasVideoViewerVideo || string.IsNullOrWhiteSpace(VideoViewerFilePath) || !File.Exists(VideoViewerFilePath))
            {
                return;
            }

            var ffmpegPath = MediaConvertService.FindFfmpeg();
            if (string.IsNullOrWhiteSpace(ffmpegPath))
            {
                VideoViewerStatusMessage = "未检测到 ffmpeg.exe，无法截帧。";
                MessageBox.Show("未检测到 ffmpeg.exe，请将 FFmpeg 放到程序目录或加入系统 PATH 后再截帧。", "音视频播放", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            IsVideoViewerCapturingFrame = true;
            VideoViewerStatusMessage = "正在截取当前帧...";
            try
            {
                var outputDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "VideoFrames");
                var result = await MediaConvertService.CaptureVideoFrameAsync(
                    ffmpegPath,
                    VideoViewerFilePath,
                    VideoViewerPositionSeconds,
                    outputDirectory,
                    CancellationToken.None).ConfigureAwait(true);

                if (!result.Success)
                {
                    VideoViewerStatusMessage = "截帧失败：" + result.Message;
                    MessageBox.Show("截帧失败：" + result.Message, "音视频播放", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                LoadImageViewerFile(result.OutputPath);
                CurrentModule = "ImageViewer";
                VideoViewerStatusMessage = "已截帧并打开图片查看：" + Path.GetFileName(result.OutputPath);
                ImageViewerStatusMessage = $"来自音视频截帧：{Path.GetFileName(VideoViewerFilePath)} @ {VideoViewerPositionText}";
                AppLogService.Information("Video frame captured {File}.", Path.GetFileName(result.OutputPath));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Video frame capture failed.");
                VideoViewerStatusMessage = "截帧失败：" + ex.Message;
                MessageBox.Show("截帧失败：" + ex.Message, "音视频播放", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsVideoViewerCapturingFrame = false;
            }
        }

        private async Task GenerateVideoViewerWaveformAsync()
        {
            if (!HasVideoViewerVideo || string.IsNullOrWhiteSpace(VideoViewerFilePath) || !File.Exists(VideoViewerFilePath))
            {
                return;
            }

            var ffmpegPath = MediaConvertService.FindFfmpeg();
            if (string.IsNullOrWhiteSpace(ffmpegPath))
            {
                VideoViewerStatusMessage = "未检测到 ffmpeg.exe，无法生成波形。";
                MessageBox.Show("未检测到 ffmpeg.exe，请将 FFmpeg 放到程序目录或加入系统 PATH 后再生成波形。", "音视频播放", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            IsVideoViewerGeneratingWaveform = true;
            VideoViewerStatusMessage = "正在生成音频波形...";
            try
            {
                var outputDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "AudioWaveforms");
                var result = await MediaConvertService.GenerateAudioWaveformAsync(
                    ffmpegPath,
                    VideoViewerFilePath,
                    outputDirectory,
                    1280,
                    180,
                    CancellationToken.None).ConfigureAwait(true);

                if (!result.Success)
                {
                    VideoViewerStatusMessage = "波形生成失败：" + result.Message;
                    MessageBox.Show("波形生成失败：" + result.Message, "音视频播放", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                LoadVideoViewerWaveformImage(result.OutputPath);
                VideoViewerStatusMessage = "已生成音频波形：" + Path.GetFileName(result.OutputPath);
                AppLogService.Information("Audio waveform generated {File}.", Path.GetFileName(result.OutputPath));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Generate audio waveform failed.");
                VideoViewerStatusMessage = "波形生成失败：" + ex.Message;
                MessageBox.Show("波形生成失败：" + ex.Message, "音视频播放", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsVideoViewerGeneratingWaveform = false;
            }
        }

        private void LoadVideoViewerWaveformImage(string filePath)
        {
            var image = new BitmapImage();
            image.BeginInit();
            image.CacheOption = BitmapCacheOption.OnLoad;
            image.UriSource = new Uri(filePath, UriKind.Absolute);
            image.EndInit();
            image.Freeze();
            VideoViewerWaveformImage = image;
            VideoViewerWaveformPath = filePath;
            OnPropertyChanged(nameof(HasVideoViewerWaveformLoopRange));
        }

        private void ClearVideoViewerWaveform()
        {
            VideoViewerWaveformImage = null;
            VideoViewerWaveformPath = string.Empty;
            OnPropertyChanged(nameof(HasVideoViewerWaveformLoopRange));
        }

        private async Task LoadVideoViewerSubtitleAsync()
        {
            if (!HasVideoViewerVideo)
            {
                return;
            }

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择 SRT 字幕",
                Filter = "SRT 字幕文件|*.srt|所有文件|*.*",
                Multiselect = false
            };

            if (!string.IsNullOrWhiteSpace(VideoViewerFilePath))
            {
                var directory = Path.GetDirectoryName(VideoViewerFilePath);
                if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
                {
                    dialog.InitialDirectory = directory;
                }
            }

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            await LoadVideoViewerSubtitleFromPathAsync(dialog.FileName, "已载入字幕：").ConfigureAwait(false);
        }

        private async Task TryLoadSiblingSubtitleAsync(string mediaPath)
        {
            ClearVideoViewerSubtitle("未载入字幕");
            var subtitlePath = SubtitleService.FindSiblingSrt(mediaPath);
            if (string.IsNullOrWhiteSpace(subtitlePath))
            {
                return;
            }

            await LoadVideoViewerSubtitleFromPathAsync(subtitlePath, "已自动载入同名字幕：").ConfigureAwait(false);
        }

        private async Task LoadVideoViewerSubtitleFromPathAsync(string filePath, string statusPrefix)
        {
            try
            {
                var cues = await SubtitleService.LoadSrtAsync(filePath, CancellationToken.None);
                if (cues == null || cues.Count == 0)
                {
                    ClearVideoViewerSubtitle("字幕为空或格式无法识别。");
                    return;
                }

                _videoViewerSubtitleFilePath = filePath;
                _videoViewerSubtitles = cues;
                _videoViewerSubtitleIndex = -1;
                VideoViewerSubtitleStatus = (statusPrefix ?? "已载入字幕：") + Path.GetFileName(filePath) + $"（{cues.Count} 条）";
                VideoViewerSubtitleText = string.Empty;
                OnPropertyChanged(nameof(VideoViewerSubtitleFileName));
                OnPropertyChanged(nameof(HasVideoViewerSubtitle));
                OnPropertyChanged(nameof(HasVideoViewerSubtitleText));
                TriggerCommandRequery();
                UpdateVideoViewerSubtitle(VideoViewerPositionSeconds);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Subtitle load failed.");
                ClearVideoViewerSubtitle("字幕载入失败：" + ex.Message);
            }
        }

        private void ClearVideoViewerSubtitle(string status)
        {
            _videoViewerSubtitleFilePath = string.Empty;
            _videoViewerSubtitles = new List<SubtitleCue>();
            _videoViewerSubtitleIndex = -1;
            VideoViewerSubtitleText = string.Empty;
            VideoViewerSubtitleStatus = string.IsNullOrWhiteSpace(status) ? "未载入字幕" : status;
            OnPropertyChanged(nameof(VideoViewerSubtitleFileName));
            OnPropertyChanged(nameof(HasVideoViewerSubtitle));
            OnPropertyChanged(nameof(HasVideoViewerSubtitleText));
            TriggerCommandRequery();
        }

        public void UpdateVideoViewerSubtitle(double positionSeconds)
        {
            if (!HasVideoViewerSubtitle)
            {
                VideoViewerSubtitleText = string.Empty;
                return;
            }

            if (double.IsNaN(positionSeconds) || double.IsInfinity(positionSeconds) || positionSeconds < 0)
            {
                positionSeconds = 0;
            }

            var adjustedSeconds = Math.Max(0, positionSeconds + VideoViewerSubtitleOffsetSeconds);
            var position = TimeSpan.FromSeconds(adjustedSeconds);
            if (_videoViewerSubtitleIndex >= 0 && _videoViewerSubtitleIndex < _videoViewerSubtitles.Count)
            {
                var current = _videoViewerSubtitles[_videoViewerSubtitleIndex];
                if (position >= current.Start && position <= current.End)
                {
                    VideoViewerSubtitleText = current.Text;
                    return;
                }
            }

            for (var i = 0; i < _videoViewerSubtitles.Count; i++)
            {
                var cue = _videoViewerSubtitles[i];
                if (position >= cue.Start && position <= cue.End)
                {
                    _videoViewerSubtitleIndex = i;
                    VideoViewerSubtitleText = cue.Text;
                    return;
                }
            }

            _videoViewerSubtitleIndex = -1;
            VideoViewerSubtitleText = string.Empty;
        }

        private static bool IsSupportedVideoViewerFile(string extension)
        {
            return MediaFileAssociationCore.IsSupportedMediaExtension(extension);
        }

        private static string FormatVideoTime(double seconds)
        {
            if (seconds <= 0 || double.IsNaN(seconds) || double.IsInfinity(seconds))
            {
                return "00:00";
            }

            var time = TimeSpan.FromSeconds(seconds);
            return time.TotalHours >= 1
                ? $"{(int)time.TotalHours:00}:{time.Minutes:00}:{time.Seconds:00}"
                : $"{time.Minutes:00}:{time.Seconds:00}";
        }

        private async Task OcrClipboardImageAsync()
        {
            if (!OcrService.IsSupported)
            {
                MessageBox.Show("WindowsOCR 仅支持 Windows 10 1903 及以上系统。", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (!TryGetClipboardImage(out var image, out var clipboardError))
            {
                if (clipboardError != null)
                {
                    AppLogService.Error(clipboardError, "OCR: reading clipboard image failed.");
                    MessageBox.Show("剪贴板图片数据无效，请重新复制图片后再试。", "OCR", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                MessageBox.Show("剪贴板中没有图片。请先截图或复制一张图片，再点击此按钮。", "OCR", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            try
            {
                AppLogService.Information("OCR: starting recognize, size={W}x{H}", image.PixelWidth, image.PixelHeight);
                var text = await OcrService.RecognizeAsync(image);
                if (string.IsNullOrWhiteSpace(text))
                {
                    MessageBox.Show("未识别到任何文字。", "OCR", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                try { Clipboard.SetText(text); } catch { }
                AppLogService.Information("OCR: recognized {Len} characters", text.Length);

                var preview = text.Length > 800 ? text.Substring(0, 800) + "..." : text;
                MessageBox.Show($"识别完成，已复制到剪贴板（{text.Length} 字符）：\n\n{preview}",
                    "OCR", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "OCR recognize failed");
                MessageBox.Show("OCR 识别失败：" + ex.Message, "OCR", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ComputeFileHashAsync()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择要计算哈希的文件",
                Filter = "所有文件 (*.*)|*.*"
            };

            if (dialog.ShowDialog() != true) return;

            IsFileHashBusy = true;
            FileHashStatusMessage = "正在计算…";
            _currentFileHashResult = null;
            FileHashResult = string.Empty;
            FileHashCompareResult = string.Empty;
            BatchFileHashResults.Clear();
            OnPropertyChanged(nameof(HasBatchFileHashResults));

            try
            {
                var progress = new Progress<string>(msg => FileHashStatusMessage = msg);
                var result = await FileHashService.ComputeAsync(dialog.FileName, progress, CancellationToken.None);

                _currentFileHashResult = result;
                FileHashResult = FormatFileHashResult(result, includePath: false);
                FileHashCompareResult = BuildHashCompareResult(result);
                ApplyImportedFileHashMatch(result);

                FileHashStatusMessage = "计算完成。点击结果可复制到剪贴板。";
                AppLogService.Information("File hash computed for {File}", result.FileName);
            }
            catch (OperationCanceledException)
            {
                FileHashStatusMessage = "已取消。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "File hash computation failed");
                FileHashStatusMessage = "计算失败：" + ex.Message;
                MessageBox.Show("哈希计算失败：" + ex.Message, "提示", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsFileHashBusy = false;
            }
        }

        private static bool TryGetClipboardImage(out System.Windows.Media.Imaging.BitmapSource image, out Exception error)
        {
            image = null;
            error = null;
            const int maxAttempts = 3;

            for (var attempt = 1; attempt <= maxAttempts; attempt++)
            {
                try
                {
                    if (!Clipboard.ContainsImage())
                    {
                        return false;
                    }

                    image = Clipboard.GetImage();
                    return image != null;
                }
                catch (COMException ex)
                {
                    error = ex;
                }
                catch (ExternalException ex)
                {
                    error = ex;
                }

                if (attempt < maxAttempts)
                {
                    Thread.Sleep(80 * attempt);
                }
            }

            return false;
        }

        public void ReRegisterHotkey()
        {
            if (_pendingKey != 0)
                SafeFireAndForget(TryRegisterHotkeyAsync(_pendingModifiers, _pendingKey, false));
            if (_pendingVideoRecordKey != 0)
                SafeFireAndForget(TryRegisterHotkeyByIdAsync(HotkeyService.VideoRecordHotkeyId, _pendingVideoRecordModifiers, _pendingVideoRecordKey, false));
            if (_pendingAudioRecordKey != 0)
                SafeFireAndForget(TryRegisterHotkeyByIdAsync(HotkeyService.AudioRecordHotkeyId, _pendingAudioRecordModifiers, _pendingAudioRecordKey, false));
        }

        public void ApplyPendingVideoRecordHotkey(uint modifiers, uint key)
        {
            var prevMod = _pendingVideoRecordModifiers;
            var prevKey = _pendingVideoRecordKey;
            var prevText = VideoRecordHotkeyText;
            _pendingVideoRecordModifiers = modifiers;
            _pendingVideoRecordKey = key;
            VideoRecordHotkeyText = HotkeyService.BuildDisplayText(modifiers, key);
            SafeFireAndForget(TryRegisterHotkeyByIdAsync(HotkeyService.VideoRecordHotkeyId, modifiers, key, true, () =>
            {
                _pendingVideoRecordModifiers = prevMod;
                _pendingVideoRecordKey = prevKey;
                VideoRecordHotkeyText = prevText;
            }));
            IsCapturingVideoRecordHotkey = false;
        }

        public void ApplyPendingAudioRecordHotkey(uint modifiers, uint key)
        {
            var prevMod = _pendingAudioRecordModifiers;
            var prevKey = _pendingAudioRecordKey;
            var prevText = AudioRecordHotkeyText;
            _pendingAudioRecordModifiers = modifiers;
            _pendingAudioRecordKey = key;
            AudioRecordHotkeyText = HotkeyService.BuildDisplayText(modifiers, key);
            SafeFireAndForget(TryRegisterHotkeyByIdAsync(HotkeyService.AudioRecordHotkeyId, modifiers, key, true, () =>
            {
                _pendingAudioRecordModifiers = prevMod;
                _pendingAudioRecordKey = prevKey;
                AudioRecordHotkeyText = prevText;
            }));
            IsCapturingAudioRecordHotkey = false;
        }

        public void ApplyPendingHotkey(uint modifiers, uint key)
        {
            var previousModifiers = _pendingModifiers;
            var previousKey = _pendingKey;
            var previousDisplayText = ScreenshotHotkeyText;

            _pendingModifiers = modifiers;
            _pendingKey = key;
            ScreenshotHotkeyText = HotkeyService.BuildDisplayText(modifiers, key);

            SafeFireAndForget(TryRegisterHotkeyAsync(modifiers, key, true, () =>
            {
                _pendingModifiers = previousModifiers;
                _pendingKey = previousKey;
                ScreenshotHotkeyText = previousDisplayText;
            }));

            IsCapturingHotkey = false;
        }

        // ===== Window-mode capture: P/Invoke =====
        [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("dwmapi.dll")] private static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);
        [DllImport("user32.dll")] private static extern int GetSystemMetrics(int nIndex);
        private const int SM_XVIRTUALSCREEN_FOR_CROP = 76;
        private const int SM_YVIRTUALSCREEN_FOR_CROP = 77;
        private const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        public async Task TriggerScreenshotAsync()
        {
            if (_isScreenshotBusy)
            {
                return;
            }

            _isScreenshotBusy = true;

            try
            {
                AppLogService.Information("Screenshot capture started in mode {Mode}.", ScreenshotMode);

                // Window 模式：抓取当前前景窗口（即使是 MainWindow 本身也允许，符合"保持当前屏幕样子"的预期）
                IntPtr targetWindow = IntPtr.Zero;
                if (ScreenshotMode == "Window")
                {
                    targetWindow = GetForegroundWindow();
                }

                // 不再隐藏 MainWindow，也不再延迟——立即按下即按下，画面所见即所得
                var fullScreenshot = await Task.Run(() => ScreenshotService.CaptureFullScreen());
                BitmapSource screenshot = fullScreenshot;

                if (ScreenshotMode == "Window" && targetWindow != IntPtr.Zero)
                {
                    screenshot = CropToWindow(fullScreenshot, targetWindow) ?? fullScreenshot;
                }
                else if (ScreenshotMode == "Region")
                {
                    var cropped = await ShowRegionSelectorAsync(fullScreenshot);
                    if (cropped == null)
                    {
                        // 用户取消
                        AppLogService.Information("Screenshot region selection cancelled.");
                        return;
                    }
                    screenshot = cropped;
                }

                AddScreenshotHistory(screenshot);

                if (ShowEditorAfterCapture)
                {
                    // 弹编辑器：剪贴板不立即写入，由编辑器关闭/保存时再写入最终图
                    ShowScreenshotEditorWindow(screenshot);
                }
                else
                {
                    // 直接进剪贴板
                    var dispatcher = Application.Current?.Dispatcher;
                    if (dispatcher == null || dispatcher.CheckAccess())
                    {
                        ScreenshotService.SetClipboardCompatible(screenshot);
                    }
                    else
                    {
                        await dispatcher.InvokeAsync(() => ScreenshotService.SetClipboardCompatible(screenshot));
                    }
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
                _isScreenshotBusy = false;
            }
        }

        /// <summary>把全屏截图按指定窗口的物理像素边界裁剪。优先使用 DWM 扩展边界，避免 Windows 10+ 不可见阴影留白。</summary>
        private static BitmapSource CropToWindow(BitmapSource fullSnapshot, IntPtr hwnd)
        {
            try
            {
                RECT rect;
                int hr = DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, out rect, Marshal.SizeOf(typeof(RECT)));
                if (hr != 0 || rect.Right <= rect.Left || rect.Bottom <= rect.Top)
                {
                    if (!GetWindowRect(hwnd, out rect)) return null;
                }

                int virtLeft = GetSystemMetrics(SM_XVIRTUALSCREEN_FOR_CROP);
                int virtTop = GetSystemMetrics(SM_YVIRTUALSCREEN_FOR_CROP);

                int x = Math.Max(0, rect.Left - virtLeft);
                int y = Math.Max(0, rect.Top - virtTop);
                int w = Math.Min(rect.Right - rect.Left, fullSnapshot.PixelWidth - x);
                int h = Math.Min(rect.Bottom - rect.Top, fullSnapshot.PixelHeight - y);
                if (w < 1 || h < 1) return null;

                var cropped = new System.Windows.Media.Imaging.CroppedBitmap(fullSnapshot, new Int32Rect(x, y, w, h));
                cropped.Freeze();
                return cropped;
            }
            catch (Exception ex)
            {
                AppLogService.Warning("CropToWindow failed: {Msg}", ex.Message);
                return null;
            }
        }

        /// <summary>弹出全屏区域选择窗口，返回裁剪后的图；用户取消时返回 null。</summary>
        private async Task<BitmapSource> ShowRegionSelectorAsync(BitmapSource fullSnapshot)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null) return null;

            return await dispatcher.InvokeAsync(() =>
            {
                var win = new MyTools.Views.RegionSelectorWindow(fullSnapshot);
                var ok = win.ShowDialog() == true;
                if (!ok) return (BitmapSource)null;
                var rect = win.SelectedRectPx;
                rect = ClampBitmapRect(rect, fullSnapshot.PixelWidth, fullSnapshot.PixelHeight);
                if (rect.Width < 1 || rect.Height < 1) return null;
                var cropped = new System.Windows.Media.Imaging.CroppedBitmap(fullSnapshot, rect);
                cropped.Freeze();
                return (BitmapSource)cropped;
            });
        }

        private static Int32Rect ClampBitmapRect(Int32Rect rect, int pixelWidth, int pixelHeight)
        {
            if (pixelWidth < 1 || pixelHeight < 1)
            {
                return Int32Rect.Empty;
            }

            int x = Math.Max(0, Math.Min(rect.X, pixelWidth - 1));
            int y = Math.Max(0, Math.Min(rect.Y, pixelHeight - 1));
            int right = Math.Max(x + 1, Math.Min(rect.X + rect.Width, pixelWidth));
            int bottom = Math.Max(y + 1, Math.Min(rect.Y + rect.Height, pixelHeight));
            return new Int32Rect(x, y, right - x, bottom - y);
        }

        private void LoadScreenshotHistory()
        {
            try
            {
                ScreenshotHistoryItems.Clear();
                var folder = GetScreenshotHistoryFolder();
                if (!Directory.Exists(folder))
                {
                    RefreshScreenshotHistoryState();
                    return;
                }

                var files = Directory.GetFiles(folder, "Screenshot_*.png")
                    .Select(path => new FileInfo(path))
                    .OrderByDescending(file => file.CreationTime)
                    .Take(20)
                    .ToList();

                foreach (var file in files)
                {
                    var item = CreateScreenshotHistoryItem(file.FullName, file.CreationTime, file.Length);
                    if (item != null)
                    {
                        ScreenshotHistoryItems.Add(item);
                    }
                }

                RefreshScreenshotHistoryState();
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Load screenshot history failed: {Msg}", ex.Message);
                RefreshScreenshotHistoryState();
            }
        }

        private void AddScreenshotHistory(BitmapSource screenshot)
        {
            if (screenshot == null)
            {
                return;
            }

            try
            {
                var folder = GetScreenshotHistoryFolder();
                Directory.CreateDirectory(folder);
                var path = BuildUniqueScreenshotHistoryPath(folder);
                SaveBitmapSource(screenshot, path);

                var fileInfo = new FileInfo(path);
                var item = CreateScreenshotHistoryItem(path, fileInfo.CreationTime, fileInfo.Length);
                if (item != null)
                {
                    ScreenshotHistoryItems.Insert(0, item);
                    AddHomeRecentItem(item.Title, item.FilePath, "ImageViewer", item.FilePath, "看图");
                }

                TrimScreenshotHistory();
                RefreshScreenshotHistoryState();
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Save screenshot history failed: {Msg}", ex.Message);
            }
        }

        private ScreenshotHistoryItem CreateScreenshotHistoryItem(string filePath, DateTime createdAt, long bytes)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    return null;
                }

                var image = LoadBitmapSourceFromFile(filePath, 180);
                return new ScreenshotHistoryItem
                {
                    Title = Path.GetFileName(filePath),
                    FilePath = filePath,
                    Thumbnail = image,
                    CreatedAt = createdAt,
                    SizeText = FormatFileSize(bytes),
                    DimensionsText = image != null ? $"{image.PixelWidth} x {image.PixelHeight}" : string.Empty
                };
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Create screenshot history item failed: {Msg}", ex.Message);
                return null;
            }
        }

        private static BitmapSource LoadBitmapSourceFromFile(string filePath, int decodePixelWidth = 0)
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.CreateOptions = BitmapCreateOptions.PreservePixelFormat;
            if (decodePixelWidth > 0)
            {
                bitmap.DecodePixelWidth = decodePixelWidth;
            }
            bitmap.UriSource = new Uri(filePath, UriKind.Absolute);
            bitmap.EndInit();
            bitmap.Freeze();
            return bitmap;
        }

        private void CopyScreenshotHistoryItem(object parameter)
        {
            var item = parameter as ScreenshotHistoryItem;
            if (!CanUseScreenshotHistoryItem(item))
            {
                return;
            }

            try
            {
                var image = LoadBitmapSourceFromFile(item.FilePath);
                ScreenshotService.SetClipboardCompatible(image);
                SystemStatusMessage = "已复制截图历史到剪贴板。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Copy screenshot history failed: {Msg}", ex.Message);
                MessageBox.Show("复制截图失败：" + ex.Message, "截图历史", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void EditScreenshotHistoryItem(object parameter)
        {
            var item = parameter as ScreenshotHistoryItem;
            if (!CanUseScreenshotHistoryItem(item))
            {
                return;
            }

            try
            {
                var image = LoadBitmapSourceFromFile(item.FilePath);
                ShowScreenshotEditorWindow(image);
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Edit screenshot history failed: {Msg}", ex.Message);
                MessageBox.Show("打开编辑器失败：" + ex.Message, "截图历史", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenScreenshotHistoryItem(object parameter)
        {
            var item = parameter as ScreenshotHistoryItem;
            if (!CanUseScreenshotHistoryItem(item))
            {
                return;
            }

            TryOpenImageViewerFile(item.FilePath);
        }

        private void DeleteScreenshotHistoryItem(object parameter)
        {
            var item = parameter as ScreenshotHistoryItem;
            if (item == null)
            {
                return;
            }

            try
            {
                ScreenshotHistoryItems.Remove(item);
                if (!string.IsNullOrWhiteSpace(item.FilePath) && File.Exists(item.FilePath))
                {
                    File.Delete(item.FilePath);
                }
                RefreshScreenshotHistoryState();
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Delete screenshot history failed: {Msg}", ex.Message);
                MessageBox.Show("删除截图历史失败：" + ex.Message, "截图历史", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool CanUseScreenshotHistoryItem(object parameter)
        {
            var item = parameter as ScreenshotHistoryItem;
            return item != null && !string.IsNullOrWhiteSpace(item.FilePath) && File.Exists(item.FilePath);
        }

        private void TrimScreenshotHistory()
        {
            while (ScreenshotHistoryItems.Count > 20)
            {
                var last = ScreenshotHistoryItems[ScreenshotHistoryItems.Count - 1];
                ScreenshotHistoryItems.RemoveAt(ScreenshotHistoryItems.Count - 1);
                TryDeleteFile(last.FilePath);
            }

            try
            {
                var folder = GetScreenshotHistoryFolder();
                if (!Directory.Exists(folder))
                {
                    return;
                }

                foreach (var file in Directory.GetFiles(folder, "Screenshot_*.png")
                    .Select(path => new FileInfo(path))
                    .OrderByDescending(file => file.CreationTime)
                    .Skip(20))
                {
                    TryDeleteFile(file.FullName);
                }
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Trim screenshot history failed: {Msg}", ex.Message);
            }
        }

        private static void TryDeleteFile(string filePath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Delete screenshot history file failed: {Msg}", ex.Message);
            }
        }

        private void RefreshScreenshotHistoryState()
        {
            OnPropertyChanged(nameof(HasScreenshotHistoryItems));
            OnPropertyChanged(nameof(ScreenshotHistorySummary));
            CommandManager.InvalidateRequerySuggested();
        }

        private static string GetScreenshotHistoryFolder()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Screenshots");
        }

        private static string BuildUniqueScreenshotHistoryPath(string folder)
        {
            var baseName = "Screenshot_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var path = Path.Combine(folder, baseName + ".png");
            var index = 2;
            while (File.Exists(path))
            {
                path = Path.Combine(folder, baseName + "_" + index.ToString("00") + ".png");
                index++;
            }

            return path;
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
                    _suppressScreenshotAutoSave = true;
                    _pendingModifiers = settings.ScreenshotHotkey.Modifiers;
                    _pendingKey = settings.ScreenshotHotkey.Key;
                    ShowEditorAfterCapture = settings.ShowEditorAfterCapture;
                    ScreenshotMode = settings.ScreenshotMode;
                    _suppressScreenshotAutoSave = false;
                    ScreenshotHotkeyText = string.IsNullOrWhiteSpace(settings.ScreenshotHotkey.DisplayText)
                        ? HotkeyService.BuildDisplayText(_pendingModifiers, _pendingKey)
                        : settings.ScreenshotHotkey.DisplayText;

                    _pendingVideoRecordModifiers = settings.VideoRecordHotkey.Modifiers;
                    _pendingVideoRecordKey = settings.VideoRecordHotkey.Key;
                    VideoRecordHotkeyText = _pendingVideoRecordKey != 0
                        ? (string.IsNullOrWhiteSpace(settings.VideoRecordHotkey.DisplayText)
                            ? HotkeyService.BuildDisplayText(_pendingVideoRecordModifiers, _pendingVideoRecordKey)
                            : settings.VideoRecordHotkey.DisplayText)
                        : "未设置";

                    _pendingAudioRecordModifiers = settings.AudioRecordHotkey.Modifiers;
                    _pendingAudioRecordKey = settings.AudioRecordHotkey.Key;
                    AudioRecordHotkeyText = _pendingAudioRecordKey != 0
                        ? (string.IsNullOrWhiteSpace(settings.AudioRecordHotkey.DisplayText)
                            ? HotkeyService.BuildDisplayText(_pendingAudioRecordModifiers, _pendingAudioRecordKey)
                            : settings.AudioRecordHotkey.DisplayText)
                        : "未设置";
                });

                await TryRegisterHotkeyAsync(_pendingModifiers, _pendingKey, false);
                if (_pendingVideoRecordKey != 0)
                    await TryRegisterHotkeyByIdAsync(HotkeyService.VideoRecordHotkeyId, _pendingVideoRecordModifiers, _pendingVideoRecordKey, false);
                if (_pendingAudioRecordKey != 0)
                    await TryRegisterHotkeyByIdAsync(HotkeyService.AudioRecordHotkeyId, _pendingAudioRecordModifiers, _pendingAudioRecordKey, false);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "LoadScreenshotSettings failed");
            }
        }

        private bool _suppressScreenshotAutoSave;
        /// <summary>静默持久化截图行为（模式 / 是否打开编辑器），不弹"已保存"对话框。</summary>
        private async Task PersistScreenshotBehaviorAsync()
        {
            if (_suppressScreenshotAutoSave) return;
            try
            {
                await AppSettingsService.UpdateAsync(settings =>
                {
                    settings.ShowEditorAfterCapture = ShowEditorAfterCapture;
                    settings.ScreenshotMode = ScreenshotMode;
                });
            }
            catch (Exception ex)
            {
                AppLogService.Warning("PersistScreenshotBehavior failed: {Msg}", ex.Message);
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
                    settings.ScreenshotMode = ScreenshotMode;
                });

                var registered = await TryRegisterHotkeyAsync(_pendingModifiers, _pendingKey, true);
                if (!registered)
                {
                    return;
                }
                // Use BeginInvoke / InvokeAsync to avoid sync block from background save path (perf + reentrancy safety)
                _ = Application.Current?.Dispatcher.BeginInvoke(new Action(() =>
                    MessageBox.Show("设置已保存", "完成", MessageBoxButton.OK, MessageBoxImage.Information)));
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
            recordWindow.Closed += RecordRegionWindow_OnClosed;

            _recordRegionWindow = recordWindow;
            recordWindow.Show();
            recordWindow.Activate();
        }

        private async void RecordRegionWindow_OnClosed(object sender, EventArgs e)
        {
            if (sender is RecordRegionWindow window)
            {
                window.ToggleRecordingRequested -= RecordRegionWindow_OnToggleRecordingRequested;
                window.Closed -= RecordRegionWindow_OnClosed;
            }

            if (ReferenceEquals(_recordRegionWindow, sender))
            {
                _recordRegionWindow = null;
            }

            if (IsVideoRecording)
            {
                await StopVideoRecordingInternalAsync(showMessage: !App.IsExiting);
            }

            if (!App.IsExiting)
            {
                RestoreWindow();
            }
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
                if (!await window.RunStartCountdownAsync(CancellationToken.None))
                {
                    return;
                }

                var gifMode = IsGifRecordingMode;
                _activeVideoOutputPath = BuildRecordingOutputPath(outputFolder, gifMode ? "GIF录像" : "录像", gifMode ? ".gif" : ".mp4");
                await Recording.StartVideoRecordingAsync(region, _activeVideoOutputPath, gifMode, BuildRecordingOptions(), CancellationToken.None);
                IsVideoRecording = true;
                window.SetRecordingState(true);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Starting video recording failed.");
                MessageBox.Show(ex.Message, "录像失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private RecordingOptions BuildRecordingOptions()
        {
            var frameRate = RecordingFrameRateOption.StartsWith("60", StringComparison.Ordinal) ? 60
                : RecordingFrameRateOption.StartsWith("15", StringComparison.Ordinal) ? 15
                : 30;

            var crf = 23;
            var preset = "veryfast";
            if (string.Equals(RecordingQualityOption, "体积优先", StringComparison.Ordinal))
            {
                crf = 28;
                preset = "veryfast";
            }
            else if (string.Equals(RecordingQualityOption, "高清", StringComparison.Ordinal))
            {
                crf = 20;
                preset = "fast";
            }

            return new RecordingOptions
            {
                FrameRate = frameRate,
                Crf = crf,
                Preset = preset
            };
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
                await Recording.StartAudioOnlyAsync(_activeAudioOutputPath, CancellationToken.None);
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
                var result = await Recording.StopVideoRecordingAsync();
                IsVideoRecording = false;
                if (!showMessage)
                {
                    return;
                }

                var outputPath = !string.IsNullOrWhiteSpace(result.OutputPath)
                    ? result.OutputPath
                    : _activeVideoOutputPath;
                if (result.TimedOut || result.FileSizeBytes <= 0 || !File.Exists(outputPath))
                {
                    AppLogService.Warning("Video recording output missing or empty: {Path}, timedOut={TimedOut}, bytes={Bytes}",
                        outputPath ?? string.Empty,
                        result.TimedOut,
                        result.FileSizeBytes);
                    MessageBox.Show(
                        "录像没有生成有效文件。\n\n"
                        + "实际尝试保存到：\n"
                        + outputPath
                        + "\n\n已记录 ffmpeg 日志。请检查音频设备是否被占用；程序已在新版本中自动降级为仅录画面。",
                        "录像",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
                else
                {
                    var isGifOutput = string.Equals(Path.GetExtension(outputPath), ".gif", StringComparison.OrdinalIgnoreCase);
                    var audioNote = isGifOutput ? "GIF 动图" : result.VideoHasAudio ? "包含音频" : "仅画面";
                    AppLogService.Information("Video recording saved: {Path}, bytes={Bytes}, audio={HasAudio}",
                        outputPath,
                        result.FileSizeBytes,
                        result.VideoHasAudio);
                    OpenFileInExplorer(outputPath);
                    MessageBox.Show(
                        $"录像已保存（{audioNote}）：\n{outputPath}\n\n已在文件夹中为你定位该文件。",
                        "录像",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
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
                var result = await Recording.StopAudioOnlyAsync();
                _audioRecordingTimer.Stop();
                IsAudioRecording = false;
                AudioRecordingIndicator = string.Empty;
                if (!showMessage)
                {
                    return;
                }

                var outputPath = !string.IsNullOrWhiteSpace(result.OutputPath)
                    ? result.OutputPath
                    : _activeAudioOutputPath;
                if (result.TimedOut || result.FileSizeBytes <= 0 || !File.Exists(outputPath))
                {
                    MessageBox.Show(
                        $"录音文件可能未正确写入，请检查 {outputPath}。",
                        "录音",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
                else
                {
                    OpenFileInExplorer(outputPath);
                    MessageBox.Show($"录音已保存到 {outputPath}", "录音", MessageBoxButton.OK, MessageBoxImage.Information);
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

        private void AudioRecordingTimer_OnTick(object sender, EventArgs e)
        {
            UpdateAudioRecordingIndicator();
        }

        private bool EnsureFfmpegAvailable()
        {
            if (Recording.TryGetFfmpegPath(out _))
            {
                return true;
            }

            var expectedPath = Recording.ExpectedFfmpegPath;
            var expectedFolder = Path.GetDirectoryName(expectedPath) ?? AppDomain.CurrentDomain.BaseDirectory;
            var result = MessageBox.Show(
                "录像和录音需要 ffmpeg.exe。\n\n"
                + "请将 ffmpeg.exe 放到：\n"
                + expectedPath
                + "\n\n也可以把 ffmpeg.exe 所在目录加入系统 PATH。\n\n是否现在打开应放置的文件夹？",
                "缺少 ffmpeg",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    Directory.CreateDirectory(expectedFolder);
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = expectedFolder,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    AppLogService.Error(ex, "Opening ffmpeg folder failed.");
                    MessageBox.Show("无法打开目录：" + ex.Message, "打开目录失败", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }

            return false;
        }

        private async Task<string> EnsureRecordingOutputFolderAsync()
        {
            var settings = await AppSettingsService.LoadAsync();
            if (!string.IsNullOrWhiteSpace(settings.RecordingOutputFolder) && Directory.Exists(settings.RecordingOutputFolder))
            {
                RecordingOutputFolderText = settings.RecordingOutputFolder;
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
                RecordingOutputFolderText = dialog.SelectedPath;
                return dialog.SelectedPath;
            }
        }

        private async Task<string> EnsureAudioOutputFolderAsync()
        {
            var settings = await AppSettingsService.LoadAsync();
            if (!string.IsNullOrWhiteSpace(settings.AudioOutputFolder) && Directory.Exists(settings.AudioOutputFolder))
            {
                AudioOutputFolderText = settings.AudioOutputFolder;
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
                AudioOutputFolderText = dialog.SelectedPath;
                return dialog.SelectedPath;
            }
        }

        private async Task LoadRecordingOutputFoldersAsync()
        {
            var settings = await AppSettingsService.LoadAsync();
            RecordingOutputFolderText = string.IsNullOrWhiteSpace(settings.RecordingOutputFolder) ? "未设置" : settings.RecordingOutputFolder;
            AudioOutputFolderText = string.IsNullOrWhiteSpace(settings.AudioOutputFolder) ? "未设置" : settings.AudioOutputFolder;
        }

        private async Task SelectRecordingOutputFolderAsync()
        {
            var folder = SelectOutputFolder("选择录像保存目录", RecordingOutputFolderText);
            if (string.IsNullOrWhiteSpace(folder))
            {
                return;
            }

            await AppSettingsService.UpdateAsync(current => current.RecordingOutputFolder = folder);
            RecordingOutputFolderText = folder;
            SystemStatusMessage = "录像目录已更新。";
        }

        private async Task SelectAudioOutputFolderAsync()
        {
            var folder = SelectOutputFolder("选择录音保存目录", AudioOutputFolderText);
            if (string.IsNullOrWhiteSpace(folder))
            {
                return;
            }

            await AppSettingsService.UpdateAsync(current => current.AudioOutputFolder = folder);
            AudioOutputFolderText = folder;
            SystemStatusMessage = "录音目录已更新。";
        }

        private static string SelectOutputFolder(string description, string currentFolder)
        {
            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = description;
                dialog.ShowNewFolderButton = true;
                if (!string.IsNullOrWhiteSpace(currentFolder) && Directory.Exists(currentFolder))
                {
                    dialog.SelectedPath = currentFolder;
                }

                return dialog.ShowDialog() == WinForms.DialogResult.OK ? dialog.SelectedPath : string.Empty;
            }
        }

        private static string BuildRecordingOutputPath(string outputFolder, string prefix, string extension)
        {
            Directory.CreateDirectory(outputFolder);
            return Path.Combine(outputFolder, $"{prefix}_{DateTime.Now:yyyyMMdd_HHmmss}{extension}");
        }

        private static void OpenFileInExplorer(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
            {
                return;
            }

            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "explorer.exe",
                    Arguments = "/select,\"" + filePath + "\"",
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                AppLogService.Warning("OpenFileInExplorer failed for {Path}: {Msg}", filePath, ex.Message);
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
                var registered = HotkeyService.Register(HotkeyService.ScreenshotHotkeyId, modifiers, key);
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

        private Task<bool> TryRegisterHotkeyByIdAsync(int id, uint modifiers, uint key, bool showMessageOnFailure, Action onRegisterFailed = null)
        {
            var dispatcher = Application.Current?.Dispatcher;
            if (dispatcher == null) return System.Threading.Tasks.Task.FromResult(false);
            return dispatcher.InvokeAsync(() =>
            {
                var registered = HotkeyService.Register(id, modifiers, key);
                if (registered || !showMessageOnFailure) return registered;
                onRegisterFailed?.Invoke();
                MessageBox.Show(
                    BuildHotkeyRegistrationErrorMessage(modifiers, key, HotkeyService.LastWin32ErrorCode),
                    "快捷键不可用",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning);
                return false;
            }).Task;
        }

        private async System.Threading.Tasks.Task SaveRecordingHotkeySettingsAsync()
        {
            try
            {
                await AppSettingsService.UpdateAsync(settings =>
                {
                    settings.VideoRecordHotkey = new HotkeySettings { Modifiers = _pendingVideoRecordModifiers, Key = _pendingVideoRecordKey, DisplayText = VideoRecordHotkeyText };
                    settings.AudioRecordHotkey = new HotkeySettings { Modifiers = _pendingAudioRecordModifiers, Key = _pendingAudioRecordKey, DisplayText = AudioRecordHotkeyText };
                });
                if (_pendingVideoRecordKey != 0)
                    await TryRegisterHotkeyByIdAsync(HotkeyService.VideoRecordHotkeyId, _pendingVideoRecordModifiers, _pendingVideoRecordKey, true);
                if (_pendingAudioRecordKey != 0)
                    await TryRegisterHotkeyByIdAsync(HotkeyService.AudioRecordHotkeyId, _pendingAudioRecordModifiers, _pendingAudioRecordKey, true);
                // Use BeginInvoke to avoid sync block from background save path (perf + reentrancy safety)
                _ = Application.Current?.Dispatcher.BeginInvoke(new Action(() => MessageBox.Show("设置已保存", "完成", MessageBoxButton.OK, MessageBoxImage.Information)));
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "SaveRecordingHotkeySettings failed");
            }
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
                FilteredStartupView?.Refresh();
                OnPropertyChanged(nameof(HasNoStartupItems));
            }
            else if (CurrentModule == "SqlExport")
            {
                if (SelectedSqlDatabase != null && SqlTableList.Count == 0)
                {
                    SafeFireAndForget(LoadTablesForSelectedDatabaseAsync());
                }
            }
            else if (CurrentModule == "System")
            {
                if (string.Equals(CurrentSystemSection, "Network", StringComparison.Ordinal))
                {
                    var data = NetworkService.GetAllNetworkDetails();
                    NetworkList.Clear();
                    foreach (var item in data)
                    {
                        NetworkList.Add(item);
                    }
                }
                else if (string.Equals(CurrentSystemSection, "Startup", StringComparison.Ordinal))
                {
                    var data = StartupService.GetStartupItems();
                    StartupList.Clear();
                    foreach (var item in data)
                    {
                        StartupList.Add(item);
                    }
                    FilteredStartupView?.Refresh();
                    OnPropertyChanged(nameof(HasNoStartupItems));
                }
                else
                {
                    RefreshSystemStatus();
                }
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
            CancellationTokenSource exportCancellationTokenSource = null;
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
                exportCancellationTokenSource = new CancellationTokenSource();
                _sqlExportCancellationTokenSource = exportCancellationTokenSource;
                SqlStatusMessage = "正在检查数据量并导出 Excel...";
                var progress = new Progress<SqlExportProgress>(p => SqlStatusMessage = FormatSqlExportProgress(p));

                var provider = SqlExportProviderFactory.GetProvider(options.ProviderKind);
                var exportResult = await provider.ExportTableAsync(
                    options,
                    SelectedSqlDatabase.Name,
                    SelectedSqlTable,
                    dialog.FileName,
                    exportCancellationTokenSource.Token,
                    progress);

                SqlStatusMessage = FormatSqlExportResult("导出完成", exportResult);
                MessageBox.Show(
                    $"导出成功。\n行数：{exportResult.RowCount:N0}\n耗时：{FormatDuration(exportResult.Duration)}\n大小：{FormatFileSize(exportResult.FileSizeBytes)}\n文件路径：{exportResult.FilePath}",
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
                if (ReferenceEquals(_sqlExportCancellationTokenSource, exportCancellationTokenSource))
                {
                    _sqlExportCancellationTokenSource = null;
                }

                exportCancellationTokenSource?.Dispose();
                IsSqlBusy = false;
            }
        }

        private void CancelSqlExport()
        {
            try
            {
                _sqlExportCancellationTokenSource?.Cancel();
                SqlStatusMessage = "正在取消导出...";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Cancel SQL export failed: {Msg}", ex.Message);
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
                    ReplaceItems(SqlRecentConnections, history.RecentConnections);
                    OnPropertyChanged(nameof(HasSqlRecentConnections));

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
                    ReplaceItems(SqlRecentConnections, history.RecentConnections);
                    OnPropertyChanged(nameof(HasSqlRecentConnections));
                }
                finally
                {
                    _isApplyingSqlHistory = false;
                }
            });
        }

        private void ApplySqlRecentConnection(object parameter)
        {
            if (!(parameter is SqlConnectionHistoryItem item) || IsSqlBusy)
            {
                return;
            }

            _isApplyingSqlHistory = true;
            try
            {
                SqlServerAddress = item.ServerAddress ?? string.Empty;
                SqlPort = string.IsNullOrWhiteSpace(item.Port) ? GetDefaultSqlPort(SelectedSqlProvider) : item.Port;
                SqlUsername = item.Username ?? string.Empty;
                SqlPassword = item.Password ?? string.Empty;
                _hasUserModifiedSqlConnectionInputs = true;
                _activeSqlConnectionOptions = null;
            }
            finally
            {
                _isApplyingSqlHistory = false;
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
            SqlStatusMessage = "已套用最近连接，请测试连接。";
            TriggerCommandRequery();
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

        private void AttachInstalledProgramSelectionHandlers()
        {
            foreach (var program in InstalledPrograms)
            {
                program.PropertyChanged -= InstalledProgramOnPropertyChanged;
                program.PropertyChanged += InstalledProgramOnPropertyChanged;
            }

            OnPropertyChanged(nameof(SelectedInstalledProgramsCountText));
            TriggerCommandRequery();
        }

        private void InstalledProgramOnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            if (!string.Equals(e.PropertyName, nameof(InstalledProgram.IsSelected), StringComparison.Ordinal))
            {
                return;
            }

            OnPropertyChanged(nameof(SelectedInstalledProgramsCount));
            OnPropertyChanged(nameof(SelectedInstalledProgramsCountText));
            TriggerCommandRequery();
        }

        private async Task LoadInstalledProgramsAsync()
        {
            if (IsInstalledProgramsBusy)
            {
                return;
            }

            try
            {
                IsInstalledProgramsBusy = true;
                InstalledProgramsStatusMessage = "正在读取可卸载程序列表...";
                var programs = await Task.Run(() => InstalledProgramService.GetUninstallablePrograms()).ConfigureAwait(false);
                await Application.Current.Dispatcher.InvokeAsync(() =>
                {
                    ReplaceItems(InstalledPrograms, programs);
                    AttachInstalledProgramSelectionHandlers();
                    FilteredInstalledProgramsView?.Refresh();
                    OnPropertyChanged(nameof(HasNoInstalledPrograms));
                    OnPropertyChanged(nameof(InstalledProgramsCountText));
                });

                InstalledProgramsStatusMessage = programs.Count == 0
                    ? "未发现带卸载命令的程序。"
                    : $"已加载 {programs.Count} 个可卸载程序。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Loading installed programs failed.");
                InstalledProgramsStatusMessage = "读取可卸载程序失败：" + ex.Message;
                MessageBox.Show(ex.Message, "读取失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsInstalledProgramsBusy = false;
            }
        }

        private void SelectFilteredInstalledPrograms()
        {
            if (FilteredInstalledProgramsView == null)
            {
                return;
            }

            var selectedCount = 0;
            foreach (var program in FilteredInstalledProgramsView.Cast<InstalledProgram>())
            {
                if (!program.IsSelected)
                {
                    program.IsSelected = true;
                }

                selectedCount++;
            }

            InstalledProgramsStatusMessage = $"已选择当前筛选结果：{selectedCount} 个程序。";
            OnPropertyChanged(nameof(SelectedInstalledProgramsCountText));
            TriggerCommandRequery();
        }

        private void ClearInstalledProgramSelection()
        {
            foreach (var program in InstalledPrograms.Where(item => item.IsSelected).ToList())
            {
                program.IsSelected = false;
            }

            InstalledProgramsStatusMessage = "已清除程序选择。";
            OnPropertyChanged(nameof(SelectedInstalledProgramsCountText));
            TriggerCommandRequery();
        }

        private void ExportInstalledProgramList()
        {
            var programs = GetInstalledProgramsForExport();
            if (programs.Count == 0)
            {
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "导出程序清单",
                FileName = $"卸载程序清单_{DateTime.Now:yyyyMMdd_HHmmss}.txt",
                Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                DefaultExt = ".txt",
                AddExtension = true,
                OverwritePrompt = true
            };
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                File.WriteAllText(dialog.FileName, BuildInstalledProgramListText(programs), new UTF8Encoding(false));
                InstalledProgramsStatusMessage = $"程序清单已导出：{programs.Count} 个。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Export installed program list failed");
                InstalledProgramsStatusMessage = "导出程序清单失败：" + ex.Message;
            }
        }

        private List<InstalledProgram> GetInstalledProgramsForExport()
        {
            var selectedPrograms = InstalledPrograms.Where(program => program.IsSelected).ToList();
            if (selectedPrograms.Count > 0)
            {
                return selectedPrograms;
            }

            return FilteredInstalledProgramsView == null
                ? new List<InstalledProgram>()
                : FilteredInstalledProgramsView.Cast<InstalledProgram>().ToList();
        }

        private string BuildInstalledProgramListText(IList<InstalledProgram> programs)
        {
            var builder = new StringBuilder();
            builder.AppendLine("MyTools 程序卸载清单");
            builder.AppendLine("生成时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            builder.AppendLine("导出范围：" + (SelectedInstalledProgramsCount > 0 ? "已选择程序" : "当前筛选结果"));
            builder.AppendLine("程序数量：" + (programs?.Count ?? 0));
            builder.AppendLine();

            foreach (var program in programs ?? Enumerable.Empty<InstalledProgram>())
            {
                builder.AppendLine("程序：" + (program.DisplayName ?? string.Empty));
                builder.AppendLine("版本：" + program.VersionDisplay);
                builder.AppendLine("发布者：" + program.PublisherDisplay);
                builder.AppendLine("安装日期：" + (program.InstallDateDisplay ?? "-"));
                builder.AppendLine("估算大小：" + (program.EstimatedSizeDisplay ?? "-"));
                builder.AppendLine("来源：" + (program.Source ?? "-"));
                builder.AppendLine("安装位置：" + program.InstallLocationDisplay);
                builder.AppendLine("权限提示：" + program.RequiresAdminDisplay);
                builder.AppendLine("静默候选：" + program.SilentUninstallDisplay);
                builder.AppendLine("静默说明：" + program.SilentUninstallDetail);
                builder.AppendLine("卸载命令：" + (program.UninstallString ?? string.Empty));
                builder.AppendLine("静默命令：" + (string.IsNullOrWhiteSpace(program.QuietUninstallString) ? "-" : program.QuietUninstallString));
                builder.AppendLine();
            }

            return builder.ToString().TrimEnd();
        }

        private async Task UninstallProgramAsync(object parameter)
        {
            if (!(parameter is InstalledProgram program))
            {
                return;
            }

            // 预检：若注册表已不存在该程序（可能此前已被卸载，列表未刷新），直接从列表移除
            InstalledProgramsStatusMessage = $"正在校验「{program.DisplayName}」...";
            var stillInstalledBefore = await Task.Run(() => InstalledProgramService.IsStillInstalled(program)).ConfigureAwait(true);
            if (!stillInstalledBefore)
            {
                InstalledPrograms.Remove(program);
                FilteredInstalledProgramsView?.Refresh();
                OnPropertyChanged(nameof(HasNoInstalledPrograms));
                OnPropertyChanged(nameof(InstalledProgramsCountText));
                OnPropertyChanged(nameof(SelectedInstalledProgramsCountText));
                InstalledProgramsStatusMessage = $"「{program.DisplayName}」已不在系统中，已从列表移除。";
                return;
            }

            var confirm = MessageBox.Show(
                BuildUninstallConfirmationText(program),
                "确认卸载",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes)
            {
                InstalledProgramsStatusMessage = "已取消卸载。";
                return;
            }

            Process process = null;
            try
            {
                process = InstalledProgramService.StartUninstall(program);
                InstalledProgramsStatusMessage = $"卸载程序运行中，请在弹出的向导中完成「{program.DisplayName}」的卸载...";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Starting uninstall failed for {ProgramName}", program.DisplayName ?? string.Empty);
                InstalledProgramsStatusMessage = "启动卸载失败：" + ex.Message;
                MessageBox.Show(ex.Message, "卸载失败", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 等待卸载进程退出（不阻塞 UI 线程）
            try
            {
                await Task.Run(() =>
                {
                    try { process.WaitForExit(); }
                    catch { /* ignore */ }
                }).ConfigureAwait(true);
            }
            finally
            {
                try { process.Dispose(); } catch { }
            }

            // 复查注册表
            var stillInstalledAfter = await Task.Run(() => InstalledProgramService.IsStillInstalled(program)).ConfigureAwait(true);
            if (!stillInstalledAfter)
            {
                InstalledPrograms.Remove(program);
                FilteredInstalledProgramsView?.Refresh();
                OnPropertyChanged(nameof(HasNoInstalledPrograms));
                OnPropertyChanged(nameof(InstalledProgramsCountText));
                OnPropertyChanged(nameof(SelectedInstalledProgramsCountText));
                InstalledProgramsStatusMessage = $"已成功卸载「{program.DisplayName}」，并从列表移除。";
            }
            else
            {
                InstalledProgramsStatusMessage = $"卸载进程已结束，但「{program.DisplayName}」仍存在于系统中，列表保留。";
            }
        }

        private static string BuildUninstallConfirmationText(InstalledProgram program)
        {
            var builder = new StringBuilder();
            builder.AppendLine("将启动该软件自带的卸载程序，不会由 MyTools 直接删除文件。");
            builder.AppendLine();
            builder.AppendLine("程序：" + (program.DisplayName ?? string.Empty));
            builder.AppendLine("发布者：" + program.PublisherDisplay);
            builder.AppendLine("版本：" + program.VersionDisplay);
            builder.AppendLine("安装位置：" + program.InstallLocationDisplay);
            builder.AppendLine("来源：" + (program.Source ?? "-"));
            builder.AppendLine("权限：" + program.RequiresAdminDisplay);
            builder.AppendLine("静默候选：" + program.SilentUninstallDisplay);
            builder.AppendLine("静默说明：" + program.SilentUninstallDetail);
            builder.AppendLine();
            builder.AppendLine("卸载命令：");
            builder.AppendLine(string.IsNullOrWhiteSpace(program.UninstallString) ? "-" : program.UninstallString);
            if (!string.IsNullOrWhiteSpace(program.QuietUninstallString))
            {
                builder.AppendLine();
                builder.AppendLine("注册表静默命令（仅供确认，不会自动批量执行）：");
                builder.AppendLine(program.QuietUninstallString);
            }
            builder.AppendLine();
            builder.Append("是否现在启动卸载向导？");
            return builder.ToString();
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

        private bool FilterStartupItem(object item)
        {
            if (!(item is StartupItem startup))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(StartupSearchText))
            {
                return true;
            }

            var keyword = StartupSearchText.Trim();
            return ContainsIgnoreCase(startup.Name, keyword)
                || ContainsIgnoreCase(startup.Command, keyword)
                || ContainsIgnoreCase(startup.Location, keyword)
                || ContainsIgnoreCase(startup.SourceCategory, keyword)
                || ContainsIgnoreCase(startup.SourceLocationDisplay, keyword)
                || ContainsIgnoreCase(startup.Publisher, keyword)
                || ContainsIgnoreCase(startup.ExecutablePath, keyword)
                || ContainsIgnoreCase(startup.ExecutableStatusDisplay, keyword)
                || ContainsIgnoreCase(startup.SignatureStatusDisplay, keyword)
                || ContainsIgnoreCase(startup.SignatureSubject, keyword)
                || ContainsIgnoreCase(startup.SignatureTrustStatus, keyword);
        }

        private bool FilterInstalledProgram(object item)
        {
            if (!(item is InstalledProgram program))
            {
                return false;
            }

            if (!MatchesInstalledProgramSizeFilter(program) || !MatchesInstalledProgramDateFilter(program))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(InstalledProgramSearchText))
            {
                return true;
            }

            var keyword = InstalledProgramSearchText.Trim();
            return ContainsIgnoreCase(program.DisplayName, keyword)
                || ContainsIgnoreCase(program.Publisher, keyword)
                || ContainsIgnoreCase(program.DisplayVersion, keyword)
                || ContainsIgnoreCase(program.InstallLocation, keyword)
                || ContainsIgnoreCase(program.Source, keyword)
                || ContainsIgnoreCase(program.InstallDateDisplay, keyword)
                || ContainsIgnoreCase(program.EstimatedSizeDisplay, keyword)
                || ContainsIgnoreCase(program.SilentUninstallDisplay, keyword)
                || ContainsIgnoreCase(program.SilentUninstallReason, keyword)
                || ContainsIgnoreCase(program.QuietUninstallString, keyword);
        }

        private bool MatchesInstalledProgramSizeFilter(InstalledProgram program)
        {
            var sizeKb = program?.EstimatedSizeKb ?? 0;
            switch (InstalledProgramSizeFilter)
            {
                case "未知大小":
                    return sizeKb <= 0;
                case "小于 100 MB":
                    return sizeKb > 0 && sizeKb < 100 * 1024;
                case "100 MB - 1 GB":
                    return sizeKb >= 100 * 1024 && sizeKb < 1024 * 1024;
                case "大于 1 GB":
                    return sizeKb >= 1024 * 1024;
                default:
                    return true;
            }
        }

        private bool MatchesInstalledProgramDateFilter(InstalledProgram program)
        {
            var installDate = program?.InstallDate;
            switch (InstalledProgramDateFilter)
            {
                case "未知日期":
                    return !installDate.HasValue;
                case "最近 30 天":
                    return installDate.HasValue && installDate.Value.Date >= DateTime.Today.AddDays(-30);
                case "最近 90 天":
                    return installDate.HasValue && installDate.Value.Date >= DateTime.Today.AddDays(-90);
                case "最近 1 年":
                    return installDate.HasValue && installDate.Value.Date >= DateTime.Today.AddYears(-1);
                default:
                    return true;
            }
        }

        private bool HasInstalledProgramViewFilter()
        {
            return !string.IsNullOrWhiteSpace(InstalledProgramSearchText)
                || !string.Equals(InstalledProgramSizeFilter, "全部大小", StringComparison.Ordinal)
                || !string.Equals(InstalledProgramDateFilter, "全部日期", StringComparison.Ordinal);
        }

        private void RefreshInstalledProgramView()
        {
            FilteredInstalledProgramsView?.Refresh();
            OnPropertyChanged(nameof(HasNoInstalledPrograms));
            OnPropertyChanged(nameof(InstalledProgramsCountText));
            TriggerCommandRequery();
        }

        private void ApplyInstalledProgramSort()
        {
            if (FilteredInstalledProgramsView == null)
            {
                return;
            }

            using (FilteredInstalledProgramsView.DeferRefresh())
            {
                FilteredInstalledProgramsView.SortDescriptions.Clear();
                switch (InstalledProgramSortMode)
                {
                    case "安装日期新到旧":
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.InstallDate), ListSortDirection.Descending));
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.DisplayName), ListSortDirection.Ascending));
                        break;
                    case "安装日期旧到新":
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.InstallDate), ListSortDirection.Ascending));
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.DisplayName), ListSortDirection.Ascending));
                        break;
                    case "大小从大到小":
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.EstimatedSizeKb), ListSortDirection.Descending));
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.DisplayName), ListSortDirection.Ascending));
                        break;
                    case "大小从小到大":
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.EstimatedSizeKb), ListSortDirection.Ascending));
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.DisplayName), ListSortDirection.Ascending));
                        break;
                    case "发布者 A-Z":
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.Publisher), ListSortDirection.Ascending));
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.DisplayName), ListSortDirection.Ascending));
                        break;
                    default:
                        FilteredInstalledProgramsView.SortDescriptions.Add(new SortDescription(nameof(InstalledProgram.DisplayName), ListSortDirection.Ascending));
                        break;
                }
            }

            OnPropertyChanged(nameof(InstalledProgramsCountText));
        }

        private static bool ContainsIgnoreCase(string value, string keyword)
        {
            return !string.IsNullOrWhiteSpace(value)
                && value.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private bool CanExportSqlTable()
        {
            return !IsSqlBusy && SelectedSqlDatabase != null && SelectedSqlTable != null;
        }

        private static string FormatSqlExportProgress(SqlExportProgress progress)
        {
            if (progress == null)
            {
                return "正在导出...";
            }

            var stage = string.IsNullOrWhiteSpace(progress.Stage) ? "正在导出" : progress.Stage;
            var rowsText = progress.TotalRows.HasValue && progress.TotalRows.Value > 0
                ? $"{progress.ProcessedRows:N0} / {progress.TotalRows.Value:N0} 行"
                : $"{progress.ProcessedRows:N0} 行";
            var message = $"{stage}：{rowsText} · {FormatDuration(progress.Elapsed)}";
            if (progress.FileSizeBytes > 0)
            {
                message += $" · {FormatFileSize(progress.FileSizeBytes)}";
            }

            return message;
        }

        private static string FormatSqlExportResult(string title, ExportResult result)
        {
            if (result == null)
            {
                return title;
            }

            return $"{title}：{result.RowCount:N0} 行 · {FormatDuration(result.Duration)} · {FormatFileSize(result.FileSizeBytes)}。";
        }

        private static string FormatDuration(TimeSpan duration)
        {
            if (duration < TimeSpan.Zero)
            {
                duration = TimeSpan.Zero;
            }

            return duration.TotalHours >= 1
                ? duration.ToString(@"h\:mm\:ss")
                : duration.ToString(@"m\:ss");
        }

        private static string FormatFileSize(long bytes)
        {
            if (bytes <= 0)
            {
                return "0 B";
            }

            const double kb = 1024d;
            const double mb = kb * 1024d;
            const double gb = mb * 1024d;
            var value = (double)bytes;

            if (value >= gb)
            {
                return (value / gb).ToString("F2") + " GB";
            }

            if (value >= mb)
            {
                return (value / mb).ToString("F2") + " MB";
            }

            if (value >= kb)
            {
                return (value / kb).ToString("F2") + " KB";
            }

            return bytes + " B";
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
            Dispose();
            Application.Current.Shutdown();
        }

        private static bool ReadAutoStartStatus()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", false))
                {
                    return key?.GetValue("MyTools") != null;
                }
            }
            catch
            {
                return false;
            }
        }

        private void ToggleAutoStart()
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Run", true))
                {
                    if (key == null) return;
                    if (IsAutoStartEnabled)
                    {
                        key.DeleteValue("MyTools", false);
                        IsAutoStartEnabled = false;
                        SystemStatusMessage = "已取消开机自启。";
                    }
                    else
                    {
                        var exePath = System.Diagnostics.Process.GetCurrentProcess().MainModule?.FileName;
                        if (!string.IsNullOrWhiteSpace(exePath))
                        {
                            key.SetValue("MyTools", $"\"{exePath}\"");
                            IsAutoStartEnabled = true;
                            SystemStatusMessage = "已设置开机自启。";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Toggle auto-start failed");
                MessageBox.Show("设置开机自启失败：" + ex.Message, "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void OpenLogFolder()
        {
            try
            {
                var baseDir = AppDomain.CurrentDomain.BaseDirectory;
                var logsDir = System.IO.Path.Combine(baseDir, "logs");
                var targetDir = System.IO.Directory.Exists(logsDir) ? logsDir : baseDir;
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = targetDir,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Failed to open log folder");
                MessageBox.Show("无法打开日志目录：" + ex.Message, "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ==================== SystemInfo Module ====================
        private HardwareSummary _systemInfo;
        private bool _isSystemInfoBusy;
        private string _systemInfoStatusMessage = "点击下方刷新读取硬件信息。";
        private bool _isSensorsRunning;
        private string _sensorStatusMessage = "正在启动传感器…";
        private DispatcherTimer _sensorTimer;
        private HardwareSensorService _sensorService;
        private string _sensorRefreshMode = "2 秒";
        private string _homeSensorRiskText = "传感器未启用，进入“系统信息”后可开启采集。";
        private string _cpuTemp = "—";
        private string _gpuTemp = "—";
        private string _motherboardTemp = "—";
        private string _fanRpm = "—";
        private string _cpuLoad = "—";

        public ICommand ShowSystemInfoCommand { get; }
        public ICommand LoadSystemInfoCommand { get; }
        public ICommand ToggleHardwareSensorsCommand { get; }
        public ICommand RefreshSensorsOnceCommand { get; }

        public ObservableCollection<SensorReading> SensorReadings { get; } = new ObservableCollection<SensorReading>();

        public string CpuTemp
        {
            get => _cpuTemp;
            private set { _cpuTemp = value; OnPropertyChanged(); }
        }

        public string GpuTemp
        {
            get => _gpuTemp;
            private set { _gpuTemp = value; OnPropertyChanged(); }
        }

        public string MotherboardTemp
        {
            get => _motherboardTemp;
            private set { _motherboardTemp = value; OnPropertyChanged(); }
        }

        public string FanRpm
        {
            get => _fanRpm;
            private set { _fanRpm = value; OnPropertyChanged(); }
        }

        public string CpuLoad
        {
            get => _cpuLoad;
            private set { _cpuLoad = value; OnPropertyChanged(); }
        }

        private DispatcherTimer SensorTimer
        {
            get
            {
                if (_sensorTimer == null)
                {
                    _sensorTimer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(2) };
                }

                return _sensorTimer;
            }
        }

        public HardwareSummary HardwareSummaryInfo
        {
            get => _systemInfo;
            set { _systemInfo = value; OnPropertyChanged(); OnPropertyChanged(nameof(HasHardwareSummary)); }
        }

        public bool HasHardwareSummary => _systemInfo != null;

        public bool IsSystemInfoBusy
        {
            get => _isSystemInfoBusy;
            set { _isSystemInfoBusy = value; OnPropertyChanged(); }
        }

        public string SystemInfoStatusMessage
        {
            get => _systemInfoStatusMessage;
            set { _systemInfoStatusMessage = value; OnPropertyChanged(); }
        }

        public bool IsSensorsRunning
        {
            get => _isSensorsRunning;
            set { _isSensorsRunning = value; OnPropertyChanged(); OnPropertyChanged(nameof(SensorsToggleLabel)); }
        }

        public string SensorsToggleLabel => _isSensorsRunning ? "停止采集" : "启用传感器";

        public string SensorStatusMessage
        {
            get => _sensorStatusMessage;
            set { _sensorStatusMessage = value; OnPropertyChanged(); }
        }

        public string HomeSensorRiskText
        {
            get => _homeSensorRiskText;
            private set { _homeSensorRiskText = value; OnPropertyChanged(); }
        }

        public IReadOnlyList<string> SensorRefreshModes { get; } = new[] { "手动", "2 秒", "5 秒" };

        public string SensorRefreshMode
        {
            get => _sensorRefreshMode;
            set
            {
                var mode = string.IsNullOrWhiteSpace(value) || !SensorRefreshModes.Contains(value) ? "2 秒" : value;
                if (string.Equals(_sensorRefreshMode, mode, StringComparison.Ordinal))
                {
                    return;
                }

                _sensorRefreshMode = mode;
                OnPropertyChanged();
                ApplySensorRefreshMode();
            }
        }

        private async Task LoadSystemInfoAsync()
        {
            if (_isSystemInfoBusy) return;
            IsSystemInfoBusy = true;
            SystemInfoStatusMessage = "正在读取硬件信息…";
            try
            {
                var summary = await HardwareInfoService.GetSummaryAsync().ConfigureAwait(true);
                HardwareSummaryInfo = summary;
                SystemInfoStatusMessage = $"已读取 · {summary.Cpus.Count} CPU · {summary.Gpus.Count} GPU · {summary.MemoryModules.Count} 内存条 · {summary.Disks.Count} 硬盘";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "LoadSystemInfo failed");
                SystemInfoStatusMessage = "读取失败：" + ex.Message;
            }
            finally
            {
                IsSystemInfoBusy = false;
            }
        }

        private void ToggleHardwareSensors()
        {
            SensorTimer.Stop();
            SensorTimer.Tick -= SensorTimer_OnTick;
            _sensorService?.Dispose();
            _sensorService = null;
            _isSensorsRunning = false;
            OnPropertyChanged(nameof(IsSensorsRunning));
            SensorReadings.Clear();
            SensorStatusMessage = "已停止采集。";
            HomeSensorRiskText = "传感器已停止，进入“系统信息”后可重新开启采集。";
        }

        private void ApplySensorRefreshMode()
        {
            if (string.Equals(SensorRefreshMode, "手动", StringComparison.Ordinal))
            {
                SensorTimer.Stop();
                if (_isSensorsRunning)
                {
                    SensorStatusMessage = $"已切换为手动刷新 · {SensorReadings.Count} 项";
                }
                return;
            }

            SensorTimer.Interval = string.Equals(SensorRefreshMode, "5 秒", StringComparison.Ordinal)
                ? TimeSpan.FromSeconds(5)
                : TimeSpan.FromSeconds(2);
            if (_isSensorsRunning || _sensorService != null)
            {
                SensorTimer.Start();
            }
        }

        private void SensorTimer_OnTick(object sender, EventArgs e)
        {
            try
            {
                var readings = _sensorService?.ReadAll();
                if (readings == null) return;
                SensorReadings.Clear();
                foreach (var r in readings.OrderBy(r => r.HardwareKind).ThenBy(r => r.HardwareName).ThenBy(r => r.SensorKind))
                {
                    SensorReadings.Add(r);
                }

                ExtractSensorSummary(readings);

                var warning = BuildSensorWarning(readings);
                if (!string.IsNullOrWhiteSpace(warning))
                {
                    SensorStatusMessage = warning;
                    HomeSensorRiskText = warning;
                }
                else if (_isSensorsRunning && SensorReadings.Count > 0)
                {
                    SensorStatusMessage = string.Equals(SensorRefreshMode, "手动", StringComparison.Ordinal)
                        ? $"已刷新 · {SensorReadings.Count} 项 · 手动"
                        : $"已刷新 · {SensorReadings.Count} 项 · 每 {SensorRefreshMode}";
                    HomeSensorRiskText = $"传感器正常：已读取 {SensorReadings.Count} 项，未发现高温或高负载。";
                }

                // Keep status fresh on each tick if user is running as non-admin
                if (_isSensorsRunning && SensorReadings.Count == 0 && !HardwareSensorService.IsRunningAsAdmin)
                {
                    SensorStatusMessage = "未读到传感器：请以管理员身份重启 MyTools。";
                    HomeSensorRiskText = "未读到传感器：建议以管理员身份重启后查看温度风险。";
                }
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Sensor read failed: {Msg}", ex.Message);
                HomeSensorRiskText = "传感器读取失败：" + ex.Message;
            }
        }

        private static readonly HashSet<string> GpuHardwareKinds = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "GpuNvidia", "GpuAmd", "GpuIntel"
        };

        private void ExtractSensorSummary(IReadOnlyList<SensorReading> readings)
        {
            string cpuTempVal = null, gpuTempVal = null, mbTempVal = null, fanVal = null, cpuLoadVal = null;

            foreach (var r in readings)
            {
                if (string.Equals(r.SensorKind, "Temperature", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.Equals(r.HardwareKind, "Cpu", StringComparison.OrdinalIgnoreCase))
                        cpuTempVal = r.Value + r.Unit;
                    else if (GpuHardwareKinds.Contains(r.HardwareKind))
                        gpuTempVal = r.Value + r.Unit;
                    else if (string.Equals(r.HardwareKind, "Motherboard", StringComparison.OrdinalIgnoreCase))
                        mbTempVal = r.Value + r.Unit;
                    else
                        AppLogService.InformationIfInitialized("Sensor: Temp HardwareKind={HK} SensorName={SN} Value={V}", r.HardwareKind, r.SensorName, r.Value);
                }
                else if (string.Equals(r.SensorKind, "Fan", StringComparison.OrdinalIgnoreCase))
                {
                    AppLogService.InformationIfInitialized("Sensor: Fan HardwareKind={HK} SensorName={SN} Value={V}", r.HardwareKind, r.SensorName, r.Value);
                    if (fanVal == null) fanVal = r.Value + r.Unit;
                }
                else if (string.Equals(r.SensorKind, "Load", StringComparison.OrdinalIgnoreCase))
                {
                    if (string.Equals(r.HardwareKind, "Cpu", StringComparison.OrdinalIgnoreCase)
                        && (cpuLoadVal == null || r.SensorName.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0 || r.SensorName.IndexOf("Package", StringComparison.OrdinalIgnoreCase) >= 0))
                    {
                        cpuLoadVal = r.Value + r.Unit;
                    }
                }
            }

            if (cpuTempVal != null) CpuTemp = cpuTempVal;
            if (gpuTempVal != null) GpuTemp = gpuTempVal;
            if (mbTempVal != null) MotherboardTemp = mbTempVal;
            if (fanVal != null) FanRpm = fanVal;
            if (cpuLoadVal != null) CpuLoad = cpuLoadVal;
            AppLogService.InformationIfInitialized("Sensor summary: cpuTemp={CT} gpuTemp={GT} mbTemp={MT} fan={FR} cpuLoad={CL}", cpuTempVal, gpuTempVal, mbTempVal, fanVal, cpuLoadVal);
        }

        private static string BuildSensorWarning(IEnumerable<SensorReading> readings)
        {
            var warning = (readings ?? Enumerable.Empty<SensorReading>())
                .Select(reading => new { Reading = reading, Value = ParseSensorValue(reading?.Value) })
                .FirstOrDefault(item =>
                    item.Value.HasValue &&
                    ((string.Equals(item.Reading.SensorKind, "Temperature", StringComparison.OrdinalIgnoreCase) && item.Value.Value >= 85)
                     || (string.Equals(item.Reading.SensorKind, "Load", StringComparison.OrdinalIgnoreCase) && item.Value.Value >= 95)));
            if (warning == null)
            {
                return string.Empty;
            }

            return $"异常提示：{warning.Reading.HardwareName} / {warning.Reading.SensorName} {warning.Reading.Value}{warning.Reading.Unit}";
        }

        private static double? ParseSensorValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var result)
                ? result
                : (double?)null;
        }

        private void RestartAsAdmin()
        {
            try
            {
                var exe = System.Reflection.Assembly.GetEntryAssembly()?.Location;
                if (string.IsNullOrEmpty(exe) || !System.IO.File.Exists(exe))
                {
                    SensorStatusMessage = "无法定位 MyTools.exe 路径。";
                    return;
                }
                var psi = new System.Diagnostics.ProcessStartInfo(exe)
                {
                    UseShellExecute = true,
                    Verb = "runas",
                    WorkingDirectory = System.IO.Path.GetDirectoryName(exe) ?? string.Empty
                };
                System.Diagnostics.Process.Start(psi);
                System.Windows.Application.Current?.Shutdown();
            }
            catch (System.ComponentModel.Win32Exception)
            {
                // User cancelled UAC
                SensorStatusMessage = "已取消管理员授权。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("RestartAsAdmin failed: {Msg}", ex.Message);
                SensorStatusMessage = "重启失败：" + ex.Message;
            }
        }

        public bool IsRunningAsAdmin => HardwareSensorService.IsRunningAsAdmin;
        public ICommand RestartAsAdminCommand { get; }

        // ==================== Convert Module ====================
        private bool _isConvertBusy;
        private string _convertStatusMessage = "选择图片或音视频文件进行格式转换。";
        private string _convertResult = string.Empty;
        private string _convertOutputPath = string.Empty;
        private string _convertOutputFolder = string.Empty;
        private string _convertOutputMode = "同目录";
        private string _convertCustomOutputFolder = string.Empty;
        private string _ffmpegPath;
        private bool _isFfmpegAvailable;
        private string _imageOutputFormat = "jpg";
        private int _imageMaxWidth;
        private int _imageMaxHeight;
        private int _imageQuality = 85;
        private string _mediaOutputFormat = "mp3";
        private string _mediaExtraArgs = string.Empty;
        private bool _isConvertPaused;
        private CancellationTokenSource _convertCancellationTokenSource;
        private ObservableCollection<ConvertQueueItem> _convertQueueItems = new ObservableCollection<ConvertQueueItem>();

        public ICommand ShowConvertCommand { get; }
        public ICommand ConvertImageCommand { get; }
        public ICommand ConvertMediaCommand { get; }
        public ICommand CancelConvertCommand { get; }
        public ICommand OpenConvertOutputFolderCommand { get; }
        public ICommand SelectConvertOutputFolderCommand { get; }
        public ICommand ClearConvertQueueCommand { get; }
        public ICommand OpenConvertQueueOutputCommand { get; }
        public ICommand CopyConvertQueueMessageCommand { get; }
        public ICommand RetryConvertQueueItemCommand { get; }
        public ICommand RemoveConvertQueueItemCommand { get; }
        public ICommand MoveConvertQueueItemUpCommand { get; }
        public ICommand MoveConvertQueueItemDownCommand { get; }
        public ICommand ToggleConvertPauseCommand { get; }
        public ICommand ApplyImageCompressPresetCommand { get; }
        public ICommand ApplyImageAvatarPresetCommand { get; }
        public ICommand ApplyMediaMp4PresetCommand { get; }
        public ICommand ApplyMediaMp3PresetCommand { get; }
        public ObservableCollection<ConvertQueueItem> ConvertQueueItems => _convertQueueItems;
        public bool HasConvertQueueItems => ConvertQueueItems.Count > 0;

        public bool IsConvertBusy
        {
            get => _isConvertBusy;
            set
            {
                _isConvertBusy = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ConvertPauseButtonText));
                TriggerCommandRequery();
            }
        }

        public bool IsConvertPaused
        {
            get => _isConvertPaused;
            private set
            {
                if (_isConvertPaused == value)
                {
                    return;
                }

                _isConvertPaused = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ConvertPauseButtonText));
                TriggerCommandRequery();
            }
        }

        public string ConvertPauseButtonText => IsConvertPaused ? "继续" : "暂停";

        public string ConvertStatusMessage
        {
            get => _convertStatusMessage;
            set { _convertStatusMessage = value; OnPropertyChanged(); }
        }

        public string ConvertResult
        {
            get => _convertResult;
            set { _convertResult = value; OnPropertyChanged(); }
        }

        public string ConvertOutputPath
        {
            get => _convertOutputPath;
            set
            {
                _convertOutputPath = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasConvertOutputPath));
                OnPropertyChanged(nameof(HasConvertOutputTarget));
            }
        }

        public bool HasConvertOutputPath => !string.IsNullOrWhiteSpace(_convertOutputPath) && File.Exists(_convertOutputPath);

        public string ConvertOutputFolder
        {
            get => _convertOutputFolder;
            set
            {
                _convertOutputFolder = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasConvertOutputFolder));
                OnPropertyChanged(nameof(HasConvertOutputTarget));
            }
        }

        public bool HasConvertOutputFolder => !string.IsNullOrWhiteSpace(_convertOutputFolder) && Directory.Exists(_convertOutputFolder);
        public bool HasConvertOutputTarget => HasConvertOutputPath || HasConvertOutputFolder;

        public string ConvertOutputMode
        {
            get => _convertOutputMode;
            set
            {
                _convertOutputMode = string.IsNullOrWhiteSpace(value) ? "同目录" : value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasConvertCustomOutputFolder));
                OnPropertyChanged(nameof(ConvertOutputTargetText));
            }
        }

        public string ConvertCustomOutputFolder
        {
            get => _convertCustomOutputFolder;
            set
            {
                _convertCustomOutputFolder = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ConvertOutputTargetText));
            }
        }

        public bool HasConvertCustomOutputFolder => string.Equals(ConvertOutputMode, "指定目录", StringComparison.Ordinal);

        public string ConvertOutputTargetText
        {
            get
            {
                switch (ConvertOutputMode)
                {
                    case "桌面":
                        return "输出到桌面";
                    case "指定目录":
                        return string.IsNullOrWhiteSpace(ConvertCustomOutputFolder)
                            ? "请选择输出目录"
                            : ConvertCustomOutputFolder;
                    default:
                        return "输出到源文件同目录";
                }
            }
        }

        public bool IsFfmpegAvailable
        {
            get => _isFfmpegAvailable;
            set { _isFfmpegAvailable = value; OnPropertyChanged(); }
        }

        public string ImageOutputFormat
        {
            get => _imageOutputFormat;
            set { _imageOutputFormat = value; OnPropertyChanged(); }
        }

        public int ImageMaxWidth
        {
            get => _imageMaxWidth;
            set { _imageMaxWidth = value; OnPropertyChanged(); }
        }

        public int ImageMaxHeight
        {
            get => _imageMaxHeight;
            set { _imageMaxHeight = value; OnPropertyChanged(); }
        }

        public int ImageQuality
        {
            get => _imageQuality;
            set { _imageQuality = value; OnPropertyChanged(); }
        }

        public string MediaOutputFormat
        {
            get => _mediaOutputFormat;
            set { _mediaOutputFormat = value; OnPropertyChanged(); }
        }

        public string MediaExtraArgs
        {
            get => _mediaExtraArgs;
            set { _mediaExtraArgs = value; OnPropertyChanged(); }
        }

        private void DetectFfmpeg()
        {
            Task.Run(() =>
            {
                var path = MediaConvertService.FindFfmpeg();
                Application.Current?.Dispatcher?.Invoke(() =>
                {
                    _ffmpegPath = path;
                    IsFfmpegAvailable = path != null;
                    if (path != null)
                        ConvertStatusMessage = "已就绪（图片内置 + FFmpeg 已检测到）。";
                    else
                        ConvertStatusMessage = "已就绪（图片内置）。FFmpeg 未检测到，音视频转换不可用。";
                });
            });
        }

        private void SetConvertOutputPath(string outputPath)
        {
            ConvertOutputPath = string.IsNullOrWhiteSpace(outputPath) ? string.Empty : outputPath;
            ConvertOutputFolder = !string.IsNullOrWhiteSpace(outputPath) && File.Exists(outputPath)
                ? Path.GetDirectoryName(outputPath)
                : string.Empty;
            if (!string.IsNullOrWhiteSpace(outputPath) && File.Exists(outputPath))
            {
                AddHomeRecentItem(Path.GetFileName(outputPath), outputPath, "Convert", outputPath, "定位");
            }
            CommandManager.InvalidateRequerySuggested();
        }

        private void SetConvertOutputFolder(string folder)
        {
            ConvertOutputPath = string.Empty;
            ConvertOutputFolder = string.IsNullOrWhiteSpace(folder) ? string.Empty : folder;
            CommandManager.InvalidateRequerySuggested();
        }

        private void OpenConvertOutputFolder()
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(ConvertOutputPath) && File.Exists(ConvertOutputPath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = "/select,\"" + ConvertOutputPath + "\"",
                        UseShellExecute = true
                    });
                    return;
                }

                if (!string.IsNullOrWhiteSpace(ConvertOutputFolder) && Directory.Exists(ConvertOutputFolder))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = ConvertOutputFolder,
                        UseShellExecute = true
                    });
                    return;
                }

                if (string.IsNullOrWhiteSpace(ConvertOutputPath) || !File.Exists(ConvertOutputPath))
                {
                    ConvertStatusMessage = "输出文件不存在或已被移动。";
                    return;
                }
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Open convert output folder failed: {Msg}", ex.Message);
                ConvertStatusMessage = "打开位置失败：" + ex.Message;
            }
        }

        private void ClearConvertQueue()
        {
            if (IsConvertBusy)
            {
                return;
            }

            ConvertQueueItems.Clear();
            OnPropertyChanged(nameof(HasConvertQueueItems));
            ConvertStatusMessage = "转换队列已清空。";
            CommandManager.InvalidateRequerySuggested();
        }

        private void OpenConvertQueueOutput(object parameter)
        {
            if (!(parameter is ConvertQueueItem item))
            {
                return;
            }

            try
            {
                var outputPath = item.OutputPath ?? string.Empty;
                if (File.Exists(outputPath))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = "explorer.exe",
                        Arguments = "/select,\"" + outputPath + "\"",
                        UseShellExecute = true
                    });
                    return;
                }

                var outputFolder = string.IsNullOrWhiteSpace(outputPath) ? string.Empty : Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrWhiteSpace(outputFolder) && Directory.Exists(outputFolder))
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = outputFolder,
                        UseShellExecute = true
                    });
                    return;
                }

                ConvertStatusMessage = "该队列项没有可打开的输出文件。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Open convert queue output failed: {Msg}", ex.Message);
                ConvertStatusMessage = "打开队列输出失败：" + ex.Message;
            }
        }

        private static bool CanOpenConvertQueueOutput(object parameter)
        {
            return parameter is ConvertQueueItem item
                && !string.IsNullOrWhiteSpace(item.OutputPath)
                && (File.Exists(item.OutputPath)
                    || Directory.Exists(Path.GetDirectoryName(item.OutputPath) ?? string.Empty));
        }

        private void CopyConvertQueueMessage(object parameter)
        {
            if (!(parameter is ConvertQueueItem item))
            {
                return;
            }

            var text = !string.IsNullOrWhiteSpace(item.Message)
                ? item.Message
                : item.OutputPath;
            if (string.IsNullOrWhiteSpace(text))
            {
                text = $"{item.FileName}：{item.Status}";
            }

            try
            {
                Clipboard.SetText(text);
                ConvertStatusMessage = "队列项消息已复制。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Copy convert queue message failed: {Msg}", ex.Message);
                ConvertStatusMessage = "复制队列消息失败：" + ex.Message;
            }
        }

        private static bool CanCopyConvertQueueMessage(object parameter)
        {
            return parameter is ConvertQueueItem item
                && (!string.IsNullOrWhiteSpace(item.Message)
                    || !string.IsNullOrWhiteSpace(item.OutputPath)
                    || !string.IsNullOrWhiteSpace(item.Status));
        }

        private bool CanRetryConvertQueueItem(object parameter)
        {
            if (IsConvertBusy || !(parameter is ConvertQueueItem item))
            {
                return false;
            }

            return !string.IsNullOrWhiteSpace(item.SourcePath)
                && File.Exists(item.SourcePath)
                && (string.Equals(item.Kind, "图片", StringComparison.Ordinal)
                    || string.Equals(item.Kind, "音视频", StringComparison.Ordinal));
        }

        private async Task RetryConvertQueueItemAsync(object parameter)
        {
            if (!CanRetryConvertQueueItem(parameter))
            {
                return;
            }

            var item = (ConvertQueueItem)parameter;
            var isImage = string.Equals(item.Kind, "图片", StringComparison.Ordinal);
            if (!isImage && !_isFfmpegAvailable)
            {
                ConvertStatusMessage = "FFmpeg 未检测到，无法重试音视频转换。";
                MessageBox.Show("未检测到 ffmpeg.exe，请下载 FFmpeg 并加入系统 PATH 或放置在程序目录下。\n\n下载地址：https://ffmpeg.org/download.html",
                    "FFmpeg 未找到", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var outputDirectory = ResolveConvertOutputDirectory();
            if (string.Equals(ConvertOutputMode, "指定目录", StringComparison.Ordinal) && string.IsNullOrWhiteSpace(outputDirectory))
            {
                ConvertStatusMessage = "已取消：未选择输出目录。";
                return;
            }

            var cancellationTokenSource = BeginConvertOperation();
            var cancellationToken = cancellationTokenSource.Token;
            ConvertResult = string.Empty;
            SetConvertOutputPath(string.Empty);
            MarkConvertQueueItemRunning(item);

            try
            {
                ConvertResult result;
                if (isImage)
                {
                    ConvertStatusMessage = "正在重试图片：" + item.FileName;
                    result = await MediaConvertService.ConvertImageAsync(
                        item.SourcePath,
                        _imageOutputFormat,
                        _imageMaxWidth,
                        _imageMaxHeight,
                        _imageQuality,
                        outputDirectory,
                        cancellationToken);
                }
                else
                {
                    ConvertStatusMessage = "正在重试音视频：" + item.FileName;
                    var progress = new Progress<string>(msg => ConvertStatusMessage = "重试音视频：" + msg);
                    result = await MediaConvertService.ConvertMediaAsync(
                        _ffmpegPath,
                        item.SourcePath,
                        _mediaOutputFormat,
                        _mediaExtraArgs ?? string.Empty,
                        outputDirectory,
                        progress,
                        cancellationToken);
                }

                ApplyConvertQueueResult(item, result);
                ApplyConvertResults(new[] { result }, isImage ? "图片转换" : "音视频转换");
            }
            catch (OperationCanceledException)
            {
                item.Status = "已取消";
                item.Message = "用户取消。";
                ConvertStatusMessage = "已取消转换。";
                ConvertResult = "转换已取消。";
                CommandManager.InvalidateRequerySuggested();
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Retry convert queue item failed");
                item.Status = "失败";
                item.OutputPath = string.Empty;
                item.Message = ex.Message;
                ConvertStatusMessage = "重试失败：" + ex.Message;
                CommandManager.InvalidateRequerySuggested();
            }
            finally
            {
                EndConvertOperation(cancellationTokenSource);
            }
        }

        private void RemoveConvertQueueItem(object parameter)
        {
            if (IsConvertBusy || !(parameter is ConvertQueueItem item))
            {
                return;
            }

            if (ConvertQueueItems.Remove(item))
            {
                OnPropertyChanged(nameof(HasConvertQueueItems));
                ConvertStatusMessage = "已移除队列项：" + item.FileName;
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public void MoveConvertQueueItem(ConvertQueueItem item, ConvertQueueItem target)
        {
            if (IsConvertBusy ||
                item == null ||
                target == null ||
                ReferenceEquals(item, target))
            {
                return;
            }

            var oldIndex = ConvertQueueItems.IndexOf(item);
            var newIndex = ConvertQueueItems.IndexOf(target);
            if (oldIndex < 0 || newIndex < 0 || oldIndex == newIndex)
            {
                return;
            }

            ConvertQueueItems.Move(oldIndex, newIndex);
            RefreshConvertQueueNumbers();
            ConvertStatusMessage = $"已调整队列顺序：{item.FileName}";
            CommandManager.InvalidateRequerySuggested();
        }

        private void MoveConvertQueueItem(object parameter, int delta)
        {
            if (IsConvertBusy || !(parameter is ConvertQueueItem item))
            {
                return;
            }

            var oldIndex = ConvertQueueItems.IndexOf(item);
            var newIndex = oldIndex + delta;
            if (oldIndex < 0 || newIndex < 0 || newIndex >= ConvertQueueItems.Count)
            {
                return;
            }

            ConvertQueueItems.Move(oldIndex, newIndex);
            RefreshConvertQueueNumbers();
            ConvertStatusMessage = $"已调整队列顺序：{item.FileName}";
            CommandManager.InvalidateRequerySuggested();
        }

        private bool CanMoveConvertQueueItemUp(object parameter)
        {
            return !IsConvertBusy &&
                parameter is ConvertQueueItem item &&
                ConvertQueueItems.IndexOf(item) > 0;
        }

        private bool CanMoveConvertQueueItemDown(object parameter)
        {
            return !IsConvertBusy &&
                parameter is ConvertQueueItem item &&
                ConvertQueueItems.IndexOf(item) >= 0 &&
                ConvertQueueItems.IndexOf(item) < ConvertQueueItems.Count - 1;
        }

        private void SelectConvertOutputFolder()
        {
            using (var dialog = new WinForms.FolderBrowserDialog
            {
                Description = "选择转换输出目录",
                ShowNewFolderButton = true
            })
            {
                if (!string.IsNullOrWhiteSpace(ConvertCustomOutputFolder) && Directory.Exists(ConvertCustomOutputFolder))
                {
                    dialog.SelectedPath = ConvertCustomOutputFolder;
                }

                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    ConvertCustomOutputFolder = dialog.SelectedPath;
                    ConvertOutputMode = "指定目录";
                    ConvertStatusMessage = "输出目录已设置：" + dialog.SelectedPath;
                }
            }
        }

        private string ResolveConvertOutputDirectory()
        {
            switch (ConvertOutputMode)
            {
                case "桌面":
                    return Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                case "指定目录":
                    if (string.IsNullOrWhiteSpace(ConvertCustomOutputFolder))
                    {
                        SelectConvertOutputFolder();
                    }

                    return string.IsNullOrWhiteSpace(ConvertCustomOutputFolder) ? null : ConvertCustomOutputFolder;
                default:
                    return null;
            }
        }

        private void CancelConvert()
        {
            if (!IsConvertBusy)
            {
                return;
            }

            try
            {
                _convertCancellationTokenSource?.Cancel();
                ConvertStatusMessage = "正在取消转换…";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Cancel convert failed: {Msg}", ex.Message);
            }
        }

        private void ToggleConvertPause()
        {
            if (!IsConvertBusy)
            {
                return;
            }

            IsConvertPaused = !IsConvertPaused;
            ConvertStatusMessage = IsConvertPaused
                ? "已暂停队列：当前文件完成后暂停后续转换。"
                : "已继续队列。";
        }

        private async Task WaitIfConvertPausedAsync(CancellationToken cancellationToken)
        {
            while (IsConvertPaused)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ConvertStatusMessage = "队列已暂停，点击“继续”恢复。";
                await Task.Delay(250, cancellationToken);
            }
        }

        private CancellationTokenSource BeginConvertOperation()
        {
            _convertCancellationTokenSource?.Dispose();
            _convertCancellationTokenSource = new CancellationTokenSource();
            IsConvertPaused = false;
            IsConvertBusy = true;
            return _convertCancellationTokenSource;
        }

        private void EndConvertOperation(CancellationTokenSource cancellationTokenSource)
        {
            if (ReferenceEquals(_convertCancellationTokenSource, cancellationTokenSource))
            {
                _convertCancellationTokenSource.Dispose();
                _convertCancellationTokenSource = null;
            }

            IsConvertBusy = false;
            IsConvertPaused = false;
        }

        private void PrepareConvertQueue(IEnumerable<string> filePaths, string kind)
        {
            ConvertQueueItems.Clear();
            var number = 1;
            foreach (var filePath in filePaths ?? Enumerable.Empty<string>())
            {
                ConvertQueueItems.Add(new ConvertQueueItem
                {
                    Number = number++,
                    FileName = Path.GetFileName(filePath),
                    SourcePath = filePath,
                    Kind = kind,
                    Status = "等待",
                    Message = string.Empty,
                    OutputPath = string.Empty
                });
            }

            OnPropertyChanged(nameof(HasConvertQueueItems));
            CommandManager.InvalidateRequerySuggested();
        }

        private void RefreshConvertQueueNumbers()
        {
            for (var i = 0; i < ConvertQueueItems.Count; i++)
            {
                ConvertQueueItems[i].Number = i + 1;
            }
        }

        private static void MarkConvertQueueItemRunning(ConvertQueueItem item)
        {
            if (item == null)
            {
                return;
            }

            item.Status = "转换中";
            item.Message = string.Empty;
            item.OutputPath = string.Empty;
        }

        private static void ApplyConvertQueueResult(ConvertQueueItem item, ConvertResult result)
        {
            if (item == null)
            {
                return;
            }

            if (result != null && result.Success)
            {
                item.Status = "完成";
                item.OutputPath = result.OutputPath ?? string.Empty;
                item.Message = result.Message ?? string.Empty;
                CommandManager.InvalidateRequerySuggested();
                return;
            }

            item.Status = "失败";
            item.OutputPath = string.Empty;
            item.Message = result == null || string.IsNullOrWhiteSpace(result.Message)
                ? "未知错误"
                : result.Message.Replace(Environment.NewLine, " ");
            CommandManager.InvalidateRequerySuggested();
        }

        private void MarkConvertQueueCancelled(int startIndex)
        {
            for (var i = Math.Max(0, startIndex); i < ConvertQueueItems.Count; i++)
            {
                var item = ConvertQueueItems[i];
                if (string.Equals(item.Status, "等待", StringComparison.Ordinal)
                    || string.Equals(item.Status, "转换中", StringComparison.Ordinal))
                {
                    item.Status = "已取消";
                    item.Message = "用户取消。";
                }
            }

            CommandManager.InvalidateRequerySuggested();
        }

        private void MarkWaitingConvertQueueCancelled()
        {
            foreach (var item in ConvertQueueItems)
            {
                if (string.Equals(item.Status, "等待", StringComparison.Ordinal)
                    || string.Equals(item.Status, "转换中", StringComparison.Ordinal))
                {
                    item.Status = "已取消";
                    item.Message = "用户取消。";
                }
            }

            CommandManager.InvalidateRequerySuggested();
        }

        private void ApplyImageCompressPreset()
        {
            ImageOutputFormat = "jpg";
            ImageQuality = 75;
            ImageMaxWidth = 1920;
            ImageMaxHeight = 1920;
            ConvertStatusMessage = "已应用预设：图片压缩。";
        }

        private void ApplyImageAvatarPreset()
        {
            ImageOutputFormat = "jpg";
            ImageQuality = 90;
            ImageMaxWidth = 512;
            ImageMaxHeight = 512;
            ConvertStatusMessage = "已应用预设：头像 512。";
        }

        private void ApplyMediaMp4Preset()
        {
            MediaOutputFormat = "mp4";
            MediaExtraArgs = "-c:v libx264 -preset veryfast -crf 23 -c:a aac -b:a 160k";
            ConvertStatusMessage = "已应用预设：视频转 MP4。";
        }

        private void ApplyMediaMp3Preset()
        {
            MediaOutputFormat = "mp3";
            MediaExtraArgs = "-vn -b:a 192k";
            ConvertStatusMessage = "已应用预设：音频转 MP3。";
        }

        public async Task ConvertMultimediaFilesAsync(IList<MediaFileItem> targets, MediaConvertParameters parameters)
        {
            if (targets == null || targets.Count == 0 || parameters == null)
            {
                ConvertStatusMessage = "没有可转换的媒体文件。";
                return;
            }

            ImageOutputFormat = parameters.ImageFormat;
            ImageMaxWidth = parameters.ImageMaxWidth;
            ImageMaxHeight = parameters.ImageMaxHeight;
            ImageQuality = parameters.ImageQuality;
            MediaOutputFormat = parameters.MediaFormat;
            MediaExtraArgs = parameters.MediaExtraArgs ?? string.Empty;
            ConvertOutputMode = parameters.OutputMode;
            ConvertCustomOutputFolder = parameters.OutputFolder;

            var imageTargets = targets.Where(item => item.Kind == MediaKind.Image && File.Exists(item.Path)).ToList();
            var mediaTargets = targets.Where(item => (item.Kind == MediaKind.Audio || item.Kind == MediaKind.Video) && File.Exists(item.Path)).ToList();
            var outputDirectory = ResolveConvertOutputDirectory();
            var results = new List<ConvertResult>();

            if (mediaTargets.Count > 0 && !_isFfmpegAvailable)
            {
                DetectFfmpeg();
                if (!_isFfmpegAvailable)
                {
                    ConvertStatusMessage = "FFmpeg 未检测到，无法转换音视频。";
                    return;
                }
            }

            var allPaths = imageTargets.Select(item => item.Path).Concat(mediaTargets.Select(item => item.Path)).ToList();
            PrepareConvertQueue(allPaths, imageTargets.Count > 0 && mediaTargets.Count == 0 ? "图片" : mediaTargets.Count > 0 && imageTargets.Count == 0 ? "音视频" : "多媒体");
            var cancellationTokenSource = BeginConvertOperation();
            var cancellationToken = cancellationTokenSource.Token;
            ConvertResult = string.Empty;
            SetConvertOutputPath(string.Empty);

            try
            {
                var queueSnapshot = ConvertQueueItems.ToList();
                for (var i = 0; i < queueSnapshot.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    await WaitIfConvertPausedAsync(cancellationToken);
                    var queueItem = queueSnapshot[i];
                    if (queueItem == null) continue;
                    MarkConvertQueueItemRunning(queueItem);
                    var sourceKind = MediaFileTypeHelper.Classify(Path.GetExtension(queueItem.SourcePath));
                    ConvertStatusMessage = $"正在转换 {i + 1}/{queueSnapshot.Count}：{queueItem.FileName}";
                    ConvertResult result;
                    if (sourceKind == MediaKind.Image)
                    {
                        result = await MediaConvertService.ConvertImageAsync(queueItem.SourcePath, _imageOutputFormat, _imageMaxWidth, _imageMaxHeight, _imageQuality, outputDirectory, cancellationToken);
                    }
                    else
                    {
                        var index = i + 1;
                        var progress = new Progress<string>(msg => ConvertStatusMessage = $"音视频 {index}/{queueSnapshot.Count}：{msg}");
                        result = await MediaConvertService.ConvertMediaAsync(_ffmpegPath, queueItem.SourcePath, _mediaOutputFormat, _mediaExtraArgs ?? string.Empty, outputDirectory, progress, cancellationToken);
                    }
                    results.Add(result);
                    ApplyConvertQueueResult(queueItem, result);
                }

                ApplyConvertResults(results, "多媒体转换");
            }
            catch (OperationCanceledException)
            {
                ConvertStatusMessage = "已取消转换。";
                ConvertResult = "转换已取消。";
                MarkWaitingConvertQueueCancelled();
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Multimedia convert failed");
                ConvertStatusMessage = "转换失败：" + ex.Message;
            }
            finally
            {
                EndConvertOperation(cancellationTokenSource);
            }
        }
        private async Task ConvertImageAsync()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择要转换的图片",
                Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif|所有文件|*.*",
                Multiselect = true
            };
            if (dialog.ShowDialog() != true) return;

            var files = dialog.FileNames ?? new[] { dialog.FileName };
            var cancellationTokenSource = BeginConvertOperation();
            var cancellationToken = cancellationTokenSource.Token;
            ConvertStatusMessage = files.Length > 1 ? $"正在批量转换图片 0/{files.Length}…" : "正在转换图片…";
            ConvertResult = string.Empty;
            SetConvertOutputPath(string.Empty);
            PrepareConvertQueue(files, "图片");
            try
            {
                var outputDirectory = ResolveConvertOutputDirectory();
                if (string.Equals(ConvertOutputMode, "指定目录", StringComparison.Ordinal) && string.IsNullOrWhiteSpace(outputDirectory))
                {
                    ConvertStatusMessage = "已取消：未选择输出目录。";
                    return;
                }

                var results = new List<ConvertResult>();
                var queueSnapshot = ConvertQueueItems.ToList();
                for (var i = 0; i < queueSnapshot.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    await WaitIfConvertPausedAsync(cancellationToken);
                    var queueItem = queueSnapshot[i];
                    if (queueItem == null)
                    {
                        continue;
                    }

                    var file = queueItem.SourcePath;
                    MarkConvertQueueItemRunning(queueItem);
                    ConvertStatusMessage = queueSnapshot.Count > 1
                        ? $"正在批量转换图片 {i + 1}/{queueSnapshot.Count}：{Path.GetFileName(file)}"
                        : "正在转换图片…";
                    var result = await MediaConvertService.ConvertImageAsync(
                        file, _imageOutputFormat, _imageMaxWidth, _imageMaxHeight, _imageQuality, outputDirectory, cancellationToken);
                    results.Add(result);
                    ApplyConvertQueueResult(queueItem, result);
                }

                ApplyConvertResults(results, "图片转换");
            }
            catch (OperationCanceledException)
            {
                ConvertStatusMessage = "已取消转换。";
                ConvertResult = "转换已取消，未继续处理后续文件。";
                MarkWaitingConvertQueueCancelled();
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Image convert failed");
                ConvertStatusMessage = "转换失败：" + ex.Message;
            }
            finally { EndConvertOperation(cancellationTokenSource); }
        }

        private async Task ConvertMediaAsync()
        {
            if (!_isFfmpegAvailable)
            {
                MessageBox.Show("未检测到 ffmpeg.exe，请下载 FFmpeg 并加入系统 PATH 或放置在程序目录下。\n\n下载地址：https://ffmpeg.org/download.html",
                    "FFmpeg 未找到", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择要转换的音频/视频文件",
                Filter = "音视频文件|*.mp4;*.mkv;*.avi;*.mov;*.flv;*.wmv;*.mp3;*.wav;*.aac;*.flac;*.ogg;*.m4a|所有文件|*.*",
                Multiselect = true
            };
            if (dialog.ShowDialog() != true) return;

            var files = dialog.FileNames ?? new[] { dialog.FileName };
            var cancellationTokenSource = BeginConvertOperation();
            var cancellationToken = cancellationTokenSource.Token;
            ConvertResult = string.Empty;
            SetConvertOutputPath(string.Empty);
            PrepareConvertQueue(files, "音视频");
            try
            {
                var outputDirectory = ResolveConvertOutputDirectory();
                if (string.Equals(ConvertOutputMode, "指定目录", StringComparison.Ordinal) && string.IsNullOrWhiteSpace(outputDirectory))
                {
                    ConvertStatusMessage = "已取消：未选择输出目录。";
                    return;
                }

                var results = new List<ConvertResult>();
                var queueSnapshot = ConvertQueueItems.ToList();
                for (var i = 0; i < queueSnapshot.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    await WaitIfConvertPausedAsync(cancellationToken);
                    var queueItem = queueSnapshot[i];
                    if (queueItem == null)
                    {
                        continue;
                    }

                    var file = queueItem.SourcePath;
                    var index = i + 1;
                    MarkConvertQueueItemRunning(queueItem);
                    var progress = new Progress<string>(msg =>
                    {
                        ConvertStatusMessage = queueSnapshot.Count > 1
                            ? $"音视频 {index}/{queueSnapshot.Count}：{msg}"
                            : msg;
                    });
                    var result = await MediaConvertService.ConvertMediaAsync(
                        _ffmpegPath, file, _mediaOutputFormat, _mediaExtraArgs ?? string.Empty, outputDirectory, progress, cancellationToken);
                    results.Add(result);
                    ApplyConvertQueueResult(queueItem, result);
                }

                ApplyConvertResults(results, "音视频转换");
            }
            catch (OperationCanceledException)
            {
                ConvertStatusMessage = "已取消转换。";
                ConvertResult = "转换已取消，已尝试停止正在运行的 FFmpeg。";
                MarkWaitingConvertQueueCancelled();
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Media convert failed");
                ConvertStatusMessage = "转换失败：" + ex.Message;
            }
            finally { EndConvertOperation(cancellationTokenSource); }
        }

        private void ApplyConvertResults(IList<ConvertResult> results, string title)
        {
            if (results == null || results.Count == 0)
            {
                ConvertResult = "未生成转换结果。";
                ConvertStatusMessage = "转换失败。";
                SetConvertOutputPath(string.Empty);
                return;
            }

            var successResults = results.Where(r => r != null && r.Success && File.Exists(r.OutputPath)).ToList();
            var failedResults = results.Where(r => r == null || !r.Success).ToList();

            if (successResults.Count == 1 && results.Count == 1)
            {
                var result = successResults[0];
                ConvertResult = result.Message + Environment.NewLine + result.OutputPath;
                ConvertStatusMessage = $"完成 → {result.OutputPath}";
                SetConvertOutputPath(result.OutputPath);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine($"{title}完成：成功 {successResults.Count} 个，失败 {failedResults.Count} 个。");
            if (successResults.Count > 1)
            {
                var outputDirs = successResults
                    .Select(r => Path.GetDirectoryName(r.OutputPath))
                    .Where(dir => !string.IsNullOrWhiteSpace(dir))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (outputDirs.Count == 1)
                {
                    sb.AppendLine("输出目录：" + outputDirs[0]);
                }
                else if (outputDirs.Count > 1)
                {
                    sb.AppendLine($"输出目录：共 {outputDirs.Count} 个，打开位置将进入第一个目录。");
                }
            }
            foreach (var result in successResults.Take(8))
            {
                sb.AppendLine("√ " + Path.GetFileName(result.OutputPath));
            }
            foreach (var result in failedResults.Take(5))
            {
                sb.AppendLine("× " + ((result == null || string.IsNullOrWhiteSpace(result.Message)) ? "未知错误" : result.Message.Replace(Environment.NewLine, " ")));
            }
            if (successResults.Count > 8 || failedResults.Count > 5)
            {
                sb.AppendLine("…其余结果已省略。");
            }

            ConvertResult = sb.ToString().TrimEnd();
            ConvertStatusMessage = failedResults.Count == 0
                ? $"{title}完成：成功 {successResults.Count} 个。"
                : $"{title}完成：成功 {successResults.Count} 个，失败 {failedResults.Count} 个。";

            if (successResults.Count == 1)
            {
                SetConvertOutputPath(successResults[0].OutputPath);
            }
            else if (successResults.Count > 1)
            {
                SetConvertOutputFolder(Path.GetDirectoryName(successResults[0].OutputPath));
            }
            else
            {
                SetConvertOutputPath(string.Empty);
            }

            if (failedResults.Count > 0)
            {
                MessageBox.Show(ConvertResult, title, MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        // ==================== Benchmark Module ====================
        private bool _isBenchmarkBusy;
        private string _benchmarkStatusMessage = "点击运行开始性能测试。";

        public ICommand ShowBenchmarkCommand { get; }
        public ICommand ShowScheduleCommand { get; }
        public ScheduleViewModel Schedule
        {
            get
            {
                if (_schedule == null)
                {
                    _schedule = new ScheduleViewModel();
                    OnPropertyChanged();
                }

                return _schedule;
            }
        }
        public ICommand ShowSystemSettingsCommand { get; }
        public SystemSettingsViewModel SystemSettings
        {
            get
            {
                if (_systemSettings == null)
                {
                    _systemSettings = new SystemSettingsViewModel();
                    OnPropertyChanged();
                }

                return _systemSettings;
            }
        }
        public MultimediaViewModel Multimedia
        {
            get
            {
                if (_multimedia == null)
                {
                    _multimedia = new MultimediaViewModel(this);
                    OnPropertyChanged();
                }

                return _multimedia;
            }
        }
        public FrpViewModel Frp
        {
            get
            {
                if (_frp == null)
                {
                    _frp = new FrpViewModel(this);
                    OnPropertyChanged();
                }

                return _frp;
            }
        }
        public ICommand RunAllBenchmarksCommand { get; }
        public ICommand RunSingleBenchmarkCommand { get; }
        public ICommand CopyBenchmarkResultsCommand { get; }
        public ICommand ExportBenchmarkResultsCommand { get; }

        public ObservableCollection<BenchmarkResult> BenchmarkResults { get; } = new ObservableCollection<BenchmarkResult>();
        public bool HasBenchmarkResults => BenchmarkResults.Count > 0;

        public bool IsBenchmarkBusy
        {
            get => _isBenchmarkBusy;
            set { _isBenchmarkBusy = value; OnPropertyChanged(); }
        }

        public string BenchmarkStatusMessage
        {
            get => _benchmarkStatusMessage;
            set { _benchmarkStatusMessage = value; OnPropertyChanged(); }
        }

        private async Task RunAllBenchmarksAsync()
        {
            if (_isBenchmarkBusy) return;
            IsBenchmarkBusy = true;
            BenchmarkResults.Clear();
            OnPropertyChanged(nameof(HasBenchmarkResults));
            BenchmarkStatusMessage = "正在运行全部测试…";
            try
            {
                var progress = new Progress<string>(msg => BenchmarkStatusMessage = msg);
                var results = await BenchmarkService.RunAllAsync(progress, CancellationToken.None);
                foreach (var r in results) BenchmarkResults.Add(r);
                OnPropertyChanged(nameof(HasBenchmarkResults));
                CommandManager.InvalidateRequerySuggested();
                BenchmarkStatusMessage = $"全部完成 · {results.Count} 项测试";
                AppLogService.Information("Benchmark completed: {Count} tests", results.Count);
            }
            catch (OperationCanceledException) { BenchmarkStatusMessage = "已取消。"; }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Benchmark failed");
                BenchmarkStatusMessage = "测试失败：" + ex.Message;
            }
            finally { IsBenchmarkBusy = false; }
        }

        private async Task RunSingleBenchmarkAsync(object param)
        {
            if (_isBenchmarkBusy) return;
            var testName = param as string;
            if (string.IsNullOrEmpty(testName)) return;

            IsBenchmarkBusy = true;
            BenchmarkStatusMessage = $"正在运行 {testName}…";
            try
            {
                var progress = new Progress<string>(msg => BenchmarkStatusMessage = msg);
                BenchmarkResult result;
                switch (testName)
                {
                    case "CpuSingle":
                        result = await BenchmarkService.RunCpuSingleThreadAsync(progress, CancellationToken.None);
                        break;
                    case "CpuMulti":
                        result = await BenchmarkService.RunCpuMultiThreadAsync(progress, CancellationToken.None);
                        break;
                    case "MemBandwidth":
                        result = await BenchmarkService.RunMemoryBandwidthAsync(progress, CancellationToken.None);
                        break;
                    case "MemLatency":
                        result = await BenchmarkService.RunMemoryLatencyAsync(progress, CancellationToken.None);
                        break;
                    case "GpuInfo":
                        result = await BenchmarkService.RunGpuInfoAsync(progress, CancellationToken.None);
                        break;
                    default:
                        return;
                }
                BenchmarkResults.Add(result);
                OnPropertyChanged(nameof(HasBenchmarkResults));
                CommandManager.InvalidateRequerySuggested();
                BenchmarkStatusMessage = $"完成 · {result.Name}：{result.Score}";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Single benchmark failed");
                BenchmarkStatusMessage = "测试失败：" + ex.Message;
            }
            finally { IsBenchmarkBusy = false; }
        }

        private void CopyBenchmarkResults()
        {
            if (!HasBenchmarkResults)
            {
                return;
            }

            try
            {
                Clipboard.SetText(BuildBenchmarkResultsText());
                BenchmarkStatusMessage = "测试结果已复制到剪贴板。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Copy benchmark results failed: {Msg}", ex.Message);
                BenchmarkStatusMessage = "复制失败：" + ex.Message;
            }
        }

        private void ExportBenchmarkResults()
        {
            if (!HasBenchmarkResults)
            {
                return;
            }

            var dialog = new SaveFileDialog
            {
                Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                FileName = $"Benchmark_{DateTime.Now:yyyyMMdd_HHmmss}.txt",
                DefaultExt = ".txt",
                AddExtension = true,
                OverwritePrompt = true
            };
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                File.WriteAllText(dialog.FileName, BuildBenchmarkResultsText(), new UTF8Encoding(false));
                BenchmarkStatusMessage = "测试结果已导出：" + dialog.FileName;
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Export benchmark results failed");
                BenchmarkStatusMessage = "导出失败：" + ex.Message;
            }
        }

        private string BuildBenchmarkResultsText()
        {
            var builder = new StringBuilder();
            builder.AppendLine("MyTools 性能测试结果");
            builder.AppendLine("生成时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            builder.AppendLine("系统：" + OsDisplayName);
            builder.AppendLine();
            foreach (var result in BenchmarkResults)
            {
                builder.AppendLine("项目：" + (result.Name ?? string.Empty));
                builder.AppendLine("得分：" + (result.Score ?? string.Empty));
                builder.AppendLine("耗时：" + result.Elapsed);
                if (!string.IsNullOrWhiteSpace(result.Detail))
                {
                    builder.AppendLine("详情：" + result.Detail);
                }
                builder.AppendLine();
            }

            return builder.ToString().TrimEnd();
        }

        // ==================== FileVerify Module ====================
        public ICommand VerifyFileCommand { get; }
        public ICommand ShowFileVerifyCommand { get; }

        private async Task VerifyFileAsync()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择要校验的文件",
                Filter = "所有文件|*.*",
                Multiselect = true
            };
            if (dialog.ShowDialog() != true) return;
            await VerifyFromPathsAsync(dialog.FileNames ?? new[] { dialog.FileName });
        }

        public Task VerifyFromPathsAsync(IEnumerable<string> filePaths)
        {
            var files = (filePaths ?? Enumerable.Empty<string>())
                .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (files.Count == 0)
            {
                return Task.CompletedTask;
            }

            return files.Count == 1
                ? VerifyFromPathAsync(files[0])
                : VerifyBatchFromPathsAsync(files);
        }

        private async Task VerifyBatchFromPathsAsync(IList<string> filePaths)
        {
            if (_isFileHashBusy) return;

            SwitchModule("FileVerify");
            IsFileHashBusy = true;
            FileHashResult = string.Empty;
            FileHashCompareResult = string.Empty;
            _currentFileHashResult = null;
            BatchFileHashResults.Clear();
            OnPropertyChanged(nameof(HasBatchFileHashResults));

            try
            {
                for (var i = 0; i < filePaths.Count; i++)
                {
                    var filePath = filePaths[i];
                    FileHashStatusMessage = $"正在批量校验 {i + 1}/{filePaths.Count}：{Path.GetFileName(filePath)}";
                    var result = await FileHashService.ComputeAsync(filePath, null, CancellationToken.None).ConfigureAwait(true);
                    BatchFileHashResults.Add(result);
                }

                FileHashStatusMessage = $"批量校验完成：{BatchFileHashResults.Count} 个文件。";
                ApplyImportedFileHashMatches();
                if (_importedFileHashEntries.Count > 0)
                {
                    FileHashStatusMessage = BuildImportedHashSummary(_importedFileHashEntries.Count, BatchFileHashResults);
                }
                OnPropertyChanged(nameof(HasBatchFileHashResults));
                CommandManager.InvalidateRequerySuggested();
            }
            catch (OperationCanceledException)
            {
                FileHashStatusMessage = "已取消。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Batch file verify failed");
                FileHashStatusMessage = "批量校验失败：" + ex.Message;
            }
            finally
            {
                IsFileHashBusy = false;
            }
        }

        /// <summary>校验指定路径的文件（供拖放等外部调用）。</summary>
        public async Task VerifyFromPathAsync(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) return;
            if (_isFileHashBusy) return; // 计算中再来一个就忽略
            if (!File.Exists(filePath))
            {
                FileHashStatusMessage = "文件不存在：" + filePath;
                return;
            }

            // 进入文件验证模块（用户可能从其它页拖入）
            SwitchModule("FileVerify");

            IsFileHashBusy = true;
            FileHashStatusMessage = "正在计算…";
            _currentFileHashResult = null;
            FileHashResult = string.Empty;
            FileHashCompareResult = string.Empty;
            BatchFileHashResults.Clear();
            OnPropertyChanged(nameof(HasBatchFileHashResults));
            try
            {
                var progress = new Progress<string>(msg => FileHashStatusMessage = msg);
                var r = await FileHashService.ComputeAsync(filePath, progress, CancellationToken.None);

                _currentFileHashResult = r;
                FileHashResult = FormatFileHashResult(r, includePath: true);
                FileHashCompareResult = BuildHashCompareResult(r);
                ApplyImportedFileHashMatch(r);
                FileHashStatusMessage = "计算完成。点击结果可复制到剪贴板。";
                AppLogService.Information("File verify computed for {File}", r.FileName);
            }
            catch (OperationCanceledException) { FileHashStatusMessage = "已取消。"; }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "File verify failed");
                FileHashStatusMessage = "计算失败：" + ex.Message;
            }
            finally { IsFileHashBusy = false; }
        }

        private void ExportFileHashList()
        {
            if (!HasBatchFileHashResults)
            {
                return;
            }

            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "导出校验清单",
                FileName = "文件校验清单.txt",
                Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                AddExtension = true,
                DefaultExt = ".txt"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                var builder = new StringBuilder();
                builder.AppendLine("MyTools 文件校验清单");
                builder.AppendLine("生成时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                builder.AppendLine();
                foreach (var result in BatchFileHashResults)
                {
                    builder.AppendLine(FormatFileHashResult(result, includePath: true));
                    builder.AppendLine();
                }

                File.WriteAllText(dialog.FileName, builder.ToString(), new UTF8Encoding(false));
                FileHashStatusMessage = "校验清单已导出：" + dialog.FileName;
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Export file hash list failed");
                FileHashStatusMessage = "导出失败：" + ex.Message;
            }
        }

        private void ImportFileHashList()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "导入校验清单",
                Filter = "校验清单 (*.txt;*.sha256;*.sha1;*.md5;*.crc32;*.sfv)|*.txt;*.sha256;*.sha1;*.md5;*.crc32;*.sfv|所有文件 (*.*)|*.*",
                Multiselect = false
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                var entries = ParseFileHashList(File.ReadAllLines(dialog.FileName));
                _importedFileHashEntries = BuildImportedHashLookup(entries);
                ApplyImportedFileHashMatches();
                FileHashStatusMessage = HasBatchFileHashResults
                    ? BuildImportedHashSummary(entries.Count, BatchFileHashResults)
                    : $"已导入校验清单 {entries.Count} 项，请选择或拖入文件后比对。";
                CommandManager.InvalidateRequerySuggested();
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Import file hash list failed");
                FileHashStatusMessage = "导入校验清单失败：" + ex.Message;
            }
        }

        private void CopyBatchFileHashColumn(object parameter)
        {
            var kind = GetHashKindFromParameter(parameter);
            if (kind == null || !HasBatchFileHashResults)
            {
                return;
            }

            try
            {
                Clipboard.SetText(BuildBatchFileHashColumnText(kind));
                FileHashStatusMessage = $"已复制 {kind} 列，共 {BatchFileHashResults.Count} 个文件。";
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Copy batch file hash column failed: {Msg}", ex.Message);
                FileHashStatusMessage = "复制失败：" + ex.Message;
            }
        }

        private string BuildBatchFileHashColumnText(string kind)
        {
            var builder = new StringBuilder();
            foreach (var result in BatchFileHashResults)
            {
                builder.Append(result.FileName ?? string.Empty)
                    .Append('\t')
                    .AppendLine(GetActualHash(result, kind) ?? string.Empty);
            }

            return builder.ToString().TrimEnd();
        }

        private void ApplyImportedFileHashMatches()
        {
            foreach (var result in BatchFileHashResults)
            {
                ApplyImportedFileHashMatch(result);
            }
        }

        private void ApplyImportedFileHashMatch(FileHashResult result)
        {
            if (result == null)
            {
                return;
            }

            result.ExpectedHashKind = string.Empty;
            result.ExpectedHash = string.Empty;
            result.CompareStatus = string.Empty;

            if (_importedFileHashEntries == null || _importedFileHashEntries.Count == 0)
            {
                return;
            }

            var entry = FindImportedHashEntryForResult(result);
            if (entry == null)
            {
                result.CompareStatus = "未找到";
                return;
            }

            var actual = GetActualHash(result, entry.Kind);
            result.ExpectedHashKind = entry.Kind;
            result.ExpectedHash = entry.Hash;
            result.CompareStatus = HashEquals(entry.Hash, actual) ? "匹配" : "不匹配";
        }

        private ImportedHashEntry FindImportedHashEntryForResult(FileHashResult result)
        {
            if (result == null)
            {
                return null;
            }

            var keys = new[]
            {
                NormalizeImportedHashKey(result.FilePath),
                NormalizeImportedHashKey(Path.GetFileName(result.FilePath)),
                NormalizeImportedHashKey(result.FileName)
            };

            foreach (var key in keys.Where(key => !string.IsNullOrWhiteSpace(key)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                if (_importedFileHashEntries.TryGetValue(key, out var entry))
                {
                    return entry;
                }
            }

            return null;
        }

        private static List<ImportedHashEntry> ParseFileHashList(IEnumerable<string> lines)
        {
            var entries = new List<ImportedHashEntry>();
            var currentFileName = string.Empty;
            var currentFilePath = string.Empty;

            foreach (var rawLine in lines ?? Enumerable.Empty<string>())
            {
                var line = (rawLine ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                if (TryParseLabeledHashFileLine(line, ref currentFileName, ref currentFilePath, entries))
                {
                    continue;
                }

                var entry = ParseLooseHashLine(line);
                if (entry != null)
                {
                    entries.Add(entry);
                }
            }

            return entries;
        }

        private static bool TryParseLabeledHashFileLine(
            string line,
            ref string currentFileName,
            ref string currentFilePath,
            ICollection<ImportedHashEntry> entries)
        {
            if (line.StartsWith("文件：", StringComparison.Ordinal))
            {
                currentFileName = ExtractFileHashListFileName(line.Substring("文件：".Length));
                return true;
            }

            if (line.StartsWith("路径：", StringComparison.Ordinal))
            {
                currentFilePath = line.Substring("路径：".Length).Trim();
                return true;
            }

            var separator = line.IndexOf('：');
            if (separator < 0)
            {
                separator = line.IndexOf(':');
            }

            if (separator <= 0)
            {
                return false;
            }

            var kind = GetHashKindFromParameter(line.Substring(0, separator));
            if (kind == null)
            {
                return false;
            }

            var hash = NormalizeHashText(line.Substring(separator + 1));
            if (GetHashKind(hash.Length) != kind || string.IsNullOrWhiteSpace(currentFileName))
            {
                return false;
            }

            entries.Add(new ImportedHashEntry(currentFileName, currentFilePath, kind, hash));
            return true;
        }

        private static ImportedHashEntry ParseLooseHashLine(string line)
        {
            var tokens = line
                .Split(new[] { ' ', '\t', ',', ';', '|', '*', '"' },
                    StringSplitOptions.RemoveEmptyEntries)
                .Select(token => token.Trim())
                .Where(token => token.Length > 0)
                .ToArray();
            if (tokens.Length < 2)
            {
                return null;
            }

            var hashToken = tokens
                .Select((token, index) => new { Token = NormalizeHashText(token), Index = index })
                .FirstOrDefault(item => GetHashKind(item.Token.Length) != null);
            if (hashToken == null)
            {
                return null;
            }

            var fileName = string.Join(" ", tokens.Where((token, index) => index != hashToken.Index)).Trim();
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return null;
            }

            var kind = GetHashKind(hashToken.Token.Length);
            return new ImportedHashEntry(Path.GetFileName(fileName), fileName, kind, hashToken.Token);
        }

        private static string ExtractFileHashListFileName(string value)
        {
            var text = (value ?? string.Empty).Trim();
            var bracketIndex = text.IndexOf('（');
            if (bracketIndex < 0)
            {
                bracketIndex = text.IndexOf('(');
            }

            return bracketIndex > 0 ? text.Substring(0, bracketIndex).Trim() : text;
        }

        private static Dictionary<string, ImportedHashEntry> BuildImportedHashLookup(IEnumerable<ImportedHashEntry> entries)
        {
            var lookup = new Dictionary<string, ImportedHashEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (var entry in entries ?? Enumerable.Empty<ImportedHashEntry>())
            {
                AddImportedHashLookup(lookup, entry.FilePath, entry);
                AddImportedHashLookup(lookup, entry.FileName, entry);
                AddImportedHashLookup(lookup, Path.GetFileName(entry.FilePath), entry);
            }

            return lookup;
        }

        private static void AddImportedHashLookup(IDictionary<string, ImportedHashEntry> lookup, string key, ImportedHashEntry entry)
        {
            var normalized = NormalizeImportedHashKey(key);
            if (string.IsNullOrWhiteSpace(normalized) || lookup.ContainsKey(normalized))
            {
                return;
            }

            lookup.Add(normalized, entry);
        }

        private static string NormalizeImportedHashKey(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? string.Empty : value.Trim().Trim('"');
        }

        private static string BuildImportedHashSummary(int importedCount, IEnumerable<FileHashResult> results)
        {
            var list = (results ?? Enumerable.Empty<FileHashResult>()).ToList();
            var matched = list.Count(result => string.Equals(result.CompareStatus, "匹配", StringComparison.Ordinal));
            var mismatched = list.Count(result => string.Equals(result.CompareStatus, "不匹配", StringComparison.Ordinal));
            var missing = list.Count(result => string.Equals(result.CompareStatus, "未找到", StringComparison.Ordinal));
            return $"已导入校验清单 {importedCount} 项；比对完成：匹配 {matched}，不匹配 {mismatched}，未找到 {missing}。";
        }

        private static string GetHashKindFromParameter(object parameter)
        {
            var value = Convert.ToString(parameter)?.Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            switch (value.ToUpperInvariant().Replace("_", "-"))
            {
                case "MD5":
                    return "MD5";
                case "SHA1":
                case "SHA-1":
                    return "SHA-1";
                case "SHA256":
                case "SHA-256":
                    return "SHA-256";
                case "CRC32":
                case "CRC-32":
                    return "CRC32";
                default:
                    return null;
            }
        }

        private void UpdateHashCompareFromCurrentResult()
        {
            if (_currentFileHashResult == null)
            {
                FileHashCompareResult = string.Empty;
                return;
            }

            FileHashCompareResult = BuildHashCompareResult(_currentFileHashResult);
        }

        private string BuildHashCompareResult(FileHashResult result)
        {
            if (result == null || string.IsNullOrWhiteSpace(ExpectedFileHash))
            {
                return string.Empty;
            }

            var expected = ExtractExpectedHash(ExpectedFileHash);
            if (string.IsNullOrWhiteSpace(expected))
            {
                return "无法识别：请粘贴 8/32/40/64 位十六进制校验值。";
            }

            var kind = GetHashKind(expected.Length);
            if (kind == null)
            {
                return "无法识别：请粘贴 8/32/40/64 位十六进制校验值。";
            }

            var actual = GetActualHash(result, kind);
            if (HashEquals(expected, actual))
            {
                return $"匹配：{kind} 与当前文件一致。";
            }

            return $"不匹配：输入值看起来是 {kind}，但与当前文件不一致。";
        }

        private static bool HashEquals(string expected, string actual)
        {
            return !string.IsNullOrWhiteSpace(actual)
                && string.Equals(expected, NormalizeHashText(actual), StringComparison.OrdinalIgnoreCase);
        }

        private static string GetActualHash(FileHashResult result, string kind)
        {
            switch (kind)
            {
                case "MD5": return result.Md5;
                case "SHA-1": return result.Sha1;
                case "SHA-256": return result.Sha256;
                case "CRC32": return result.Crc32;
                default: return string.Empty;
            }
        }

        private static string GetHashKind(int length)
        {
            switch (length)
            {
                case 8: return "CRC32";
                case 32: return "MD5";
                case 40: return "SHA-1";
                case 64: return "SHA-256";
                default: return null;
            }
        }

        private static string ExtractExpectedHash(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var tokens = value
                .Split(new[] { ' ', '\t', '\r', '\n', ':', '=', ',', ';', '|', '(', ')', '[', ']', '{', '}', '<', '>' },
                    StringSplitOptions.RemoveEmptyEntries)
                .Select(NormalizeHashText)
                .Where(token => token.Length == 64 || token.Length == 40 || token.Length == 32 || token.Length == 8)
                .OrderByDescending(token => token.Length)
                .ToArray();

            if (tokens.Length > 0)
            {
                return tokens[0];
            }

            var labelFree = value.ToUpperInvariant()
                .Replace("SHA256", string.Empty)
                .Replace("SHA-256", string.Empty)
                .Replace("SHA1", string.Empty)
                .Replace("SHA-1", string.Empty)
                .Replace("MD5", string.Empty)
                .Replace("CRC32", string.Empty)
                .Replace("CRC-32", string.Empty);
            var normalized = NormalizeHashText(labelFree);
            return normalized.Length == 64 || normalized.Length == 40 || normalized.Length == 32 || normalized.Length == 8
                ? normalized
                : string.Empty;
        }

        private static string NormalizeHashText(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            var sb = new StringBuilder(value.Length);
            foreach (var ch in value)
            {
                if (Uri.IsHexDigit(ch))
                {
                    sb.Append(char.ToUpperInvariant(ch));
                }
            }
            return sb.ToString();
        }

        private static string FormatFileHashResult(FileHashResult result, bool includePath)
        {
            var sizeText = FileSizeFormatter.Format(result.FileSize);
            var builder = new StringBuilder()
                .Append("文件：").Append(result.FileName).Append("（").Append(sizeText).AppendLine("）");

            if (includePath)
            {
                builder.Append("路径：").AppendLine(result.FilePath);
            }

            builder.Append("MD5：").AppendLine(result.Md5)
                .Append("SHA-1：").AppendLine(result.Sha1)
                .Append("SHA-256：").AppendLine(result.Sha256)
                .Append("CRC32：").Append(result.Crc32);
            return builder.ToString();
        }

        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            _isDisposed = true;
            CancelPendingTableLoad();
            _queryCts?.Cancel();
            _queryCts?.Dispose();
            _queryCts = null;
            _audioRecordingTimer.Stop();
            _audioRecordingTimer.Tick -= AudioRecordingTimer_OnTick;

            if (_sensorTimer != null)
            {
                _sensorTimer.Stop();
                try { _sensorTimer.Tick -= SensorTimer_OnTick; } catch { }
            }
            _sensorService?.Dispose();
            _sensorService = null;

            _frp?.Dispose();

            CloseOwnedWindowsForShutdown();
            StopActiveRecordingsForShutdown();
        }

        private void CloseOwnedWindowsForShutdown()
        {
            _screenshotEditorWindow?.Close();
            _screenshotEditorWindow = null;

            var recordWindow = _recordRegionWindow;
            if (recordWindow == null)
            {
                return;
            }

            _recordRegionWindow = null;
            recordWindow.ToggleRecordingRequested -= RecordRegionWindow_OnToggleRecordingRequested;
            recordWindow.Closed -= RecordRegionWindow_OnClosed;
            recordWindow.Close();
        }

        private void StopActiveRecordingsForShutdown()
        {
            try
            {
                if (IsVideoRecording)
                {
                    Recording.StopVideoRecordingAsync().GetAwaiter().GetResult();
                    IsVideoRecording = false;
                }

                if (IsAudioRecording)
                {
                    Recording.StopAudioOnlyAsync().GetAwaiter().GetResult();
                    IsAudioRecording = false;
                    AudioRecordingIndicator = string.Empty;
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Stopping active recording during shutdown failed.");
            }
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

        private async Task ApplyAlwaysOnPowerPolicyAsync()
        {
            var confirm = MessageBox.Show(
                "将把当前电源计划设置为：不自动关闭屏幕、不自动关闭硬盘、不自动睡眠/休眠，并关闭系统休眠文件。交流电和电池模式都会生效。\n\n是否继续？",
                "应用常亮电源策略",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (confirm != MessageBoxResult.Yes)
            {
                PowerPolicyStatusMessage = "已取消应用常亮电源策略。";
                return;
            }

            IsPowerPolicyBusy = true;
            PowerPolicyStatusMessage = "正在应用常亮电源策略，请在 UAC 弹窗中确认...";
            try
            {
                await PowerPolicyService.ApplyAlwaysOnPolicyAsync(CancellationToken.None).ConfigureAwait(true);
                PowerPolicyStatusMessage = "已应用：屏幕、硬盘、睡眠和休眠均已设置为不自动关闭。";
            }
            catch (OperationCanceledException)
            {
                PowerPolicyStatusMessage = "操作已取消（UAC 未授权）。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Applying always-on power policy failed.");
                PowerPolicyStatusMessage = "应用常亮电源策略失败：" + ex.Message;
                MessageBox.Show(ex.Message, "应用失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsPowerPolicyBusy = false;
            }
        }

        private void LockCurrentWindowsVersion()
        {
            try
            {
                var os = OsVersionService.MajorVersion;
                if (os == 0)
                {
                    MessageBox.Show("无法检测当前 Windows 版本。", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                string scriptPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "LockCurrentWindows.ps1");

                if (!System.IO.File.Exists(scriptPath))
                {
                    MessageBox.Show("锁定脚本不存在: " + scriptPath, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

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
            catch (Exception ex)
            {
                MessageBox.Show("执行失败: " + ex.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool CanExecuteSqlQuery()
            => !IsQueryBusy && SelectedSqlDatabase != null && !string.IsNullOrWhiteSpace(SqlQueryText);

        private bool CanExportQueryResult()
            => !IsQueryBusy && SqlQueryResult != null && SqlQueryResult.Count > 0;

        private CancellationTokenSource _queryCts;

        private async Task ExecuteSqlQueryAsync()
        {
            _queryCts?.Cancel();
            _queryCts?.Dispose();
            _queryCts = new CancellationTokenSource();
            var ct = _queryCts.Token;

            IsQueryBusy = true;
            QueryStatusMessage = "正在执行查询...";
            SqlQueryResult = null;
            _cancelQueryCommand?.RaiseCanExecuteChanged();
            try
            {
                var options = GetEffectiveSqlConnectionOptions();
                var provider = SqlExportProviderFactory.GetProvider(options.ProviderKind);
                var table = await provider.ExecuteQueryAsync(
                    options,
                    SelectedSqlDatabase.Name,
                    SqlQueryText,
                    ct);

                SqlQueryResult = table.DefaultView;
                QueryStatusMessage = $"共 {table.Rows.Count} 行，{table.Columns.Count} 列。";
            }
            catch (OperationCanceledException)
            {
                AppLogService.Information("SQL query cancelled by user");
                QueryStatusMessage = "查询已取消。";
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
                _cancelQueryCommand?.RaiseCanExecuteChanged();
            }
        }

        private void CancelSqlQuery()
        {
            try
            {
                _queryCts?.Cancel();
                QueryStatusMessage = "正在取消查询...";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Cancelling SQL query failed");
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
                var progress = new Progress<SqlExportProgress>(p => QueryStatusMessage = FormatSqlExportProgress(p));
                var result = await SqlExportService.ExportDataTableAsync(
                    table,
                    "QueryResult",
                    dialog.FileName,
                    CancellationToken.None,
                    progress);

                QueryStatusMessage = FormatSqlExportResult("导出完成", result);
                MessageBox.Show(
                    $"导出成功。\n行数：{result.RowCount:N0}\n耗时：{FormatDuration(result.Duration)}\n大小：{FormatFileSize(result.FileSizeBytes)}\n文件路径：{result.FilePath}",
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

        private async Task ExportQueryResultCsvAsync()
        {
            if (SqlQueryResult == null) return;
            var dialog = new SaveFileDialog
            {
                Filter = "CSV 文件 (*.csv)|*.csv",
                FileName = $"QueryResult_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
                DefaultExt = ".csv",
                AddExtension = true,
                OverwritePrompt = true
            };
            if (dialog.ShowDialog() != true) return;

            IsQueryBusy = true;
            QueryStatusMessage = "正在导出 CSV...";
            try
            {
                var table = SqlQueryResult.Table;
                var progress = new Progress<SqlExportProgress>(p => QueryStatusMessage = FormatSqlExportProgress(p));
                var result = await SqlExportService.ExportDataTableToCsvAsync(
                    table,
                    dialog.FileName,
                    CancellationToken.None,
                    progress);

                QueryStatusMessage = FormatSqlExportResult("CSV 导出完成", result);
                MessageBox.Show(
                    $"导出成功。\n行数：{result.RowCount:N0}\n耗时：{FormatDuration(result.Duration)}\n大小：{FormatFileSize(result.FileSizeBytes)}\n文件路径：{result.FilePath}",
                    "导出完成",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Query result CSV export failed");
                QueryStatusMessage = "CSV 导出失败：" + ex.Message;
                MessageBox.Show(ex.Message, "CSV 导出失败", MessageBoxButton.OK, MessageBoxImage.Error);
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
                var reports = await OptimizationReportsStore.LoadAllAsync();
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
                var roots = await Task.Run(() => WeChatLocator.LocateRoots());
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
                SystemOptimizer.AllowExplorerRestartForThumbnailCleanup = restartExplorerResult == MessageBoxResult.Yes;
                var progress = new Progress<string>(message => AutoOptimizeStatusMessage = message);
                var report = await Task.Run(() => SystemOptimizer.RunAsync(progress, CancellationToken.None));
                await OptimizationReportsStore.SaveAsync(report);

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
                var scanned = await Task.Run(() => JunkCleaner.ScanAsync(progress, CancellationToken.None));
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

        private bool CanExportJunkCleanupPlan()
        {
            return !IsJunkBusy && JunkCandidates.Any(x => x.IsSelected);
        }

        private void ExportJunkCleanupPlan()
        {
            var selected = JunkCandidates.Where(x => x.IsSelected).ToList();
            if (selected.Count == 0)
            {
                JunkStatusMessage = "请先扫描并选择要导出的清理项目。";
                return;
            }

            var dialog = new SaveFileDialog
            {
                Title = "导出清理前报告",
                FileName = $"垃圾清理预检报告_{DateTime.Now:yyyyMMdd_HHmmss}.txt",
                Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                DefaultExt = ".txt",
                AddExtension = true,
                OverwritePrompt = true
            };
            if (dialog.ShowDialog() != true)
            {
                return;
            }

            try
            {
                File.WriteAllText(dialog.FileName, BuildJunkCleanupPlanText(selected), new UTF8Encoding(false));
                JunkStatusMessage = $"清理前报告已导出：{selected.Count} 项。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Export junk cleanup plan failed.");
                JunkStatusMessage = "导出清理前报告失败：" + ex.Message;
            }
        }

        private static string BuildJunkCleanupPlanText(IList<JunkCandidate> candidates)
        {
            var items = candidates ?? new List<JunkCandidate>();
            var builder = new StringBuilder();
            builder.AppendLine("MyTools 垃圾清理预检报告");
            builder.AppendLine("生成时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            builder.AppendLine("项目数量：" + items.Count);
            builder.AppendLine("预计释放：" + FileSizeFormatter.Format(items.Sum(x => x.Bytes)));
            builder.AppendLine();
            builder.AppendLine("风险分布：");
            foreach (var group in items.GroupBy(x => x.RiskDisplay).OrderByDescending(x => x.Sum(item => item.Bytes)))
            {
                builder.AppendLine($"- {group.Key}：{group.Count()} 项，{FileSizeFormatter.Format(group.Sum(x => x.Bytes))}");
            }

            builder.AppendLine();
            builder.AppendLine("清理项目：");
            foreach (var item in items.OrderBy(x => x.CategoryDisplay).ThenByDescending(x => x.Bytes))
            {
                builder.AppendLine("类别：" + item.CategoryDisplay);
                builder.AppendLine("大小：" + item.BytesDisplay);
                builder.AppendLine("风险：" + item.RiskDisplay);
                builder.AppendLine("建议：" + item.AdviceDisplay);
                builder.AppendLine("说明：" + (item.Reason ?? string.Empty));
                builder.AppendLine("路径：" + (item.Path ?? string.Empty));
                builder.AppendLine();
            }

            return builder.ToString().TrimEnd();
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
                $"将清理 {selected.Count} 项，约 {FileSizeFormatter.Format(totalBytes)}。\n\n普通文件会先发送到回收站；系统缓存可能需要管理员授权并按系统缓存目录清理。\n\n是否继续？",
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
                var execution = await Task.Run(() => JunkCleaner.CleanupAsync(selected, progress, CancellationToken.None));
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

                await OptimizationReportsStore.SaveAsync(report);
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
                WeChatCleanupStatusMessage = "正在检测微信数据...";
                EnsureWeChatStartupDataLoading();
                await Task.Delay(500);
                if (SelectedWeChatRoot == null)
                {
                    WeChatCleanupStatusMessage = "未检测到本机微信数据，请确认已安装微信且有过登录记录。";
                    return;
                }
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
                var scanResult = await Task.Run(() => WeChatCleaner.ScanAsync(
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
                var execution = await Task.Run(() => WeChatCleaner.CleanupAsync(
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

                await OptimizationReportsStore.SaveAsync(report);
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

                var backupResult = await Task.Run(() => WeChatBackupStore.BackupAsync(options, progress, CancellationToken.None));
                await AddRecentWeChatBackupAsync(backupResult.ZipPath, backupResult.FileCount, backupResult.TotalBytes);
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

            if (BuildWeChatRestoreCategories().Count == 0)
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
                var restoreResult = await Task.Run(() => WeChatBackupStore.RestoreAsync(
                    new WeChatRestoreOptions
                    {
                        ZipPath = WeChatRestoreZipPath,
                        RestoreToOriginal = RestoreToOriginal,
                        CustomTargetRoot = RestoreToOriginal ? null : WeChatRestoreTargetRoot,
                        Categories = BuildWeChatRestoreCategories()
                    },
                    new Progress<string>(msg => WeChatRestoreStatusMessage = msg),
                    CancellationToken.None));

                var skippedText = restoreResult.SkippedByCategory > 0
                    ? $"，按类别跳过 {restoreResult.SkippedByCategory}"
                    : string.Empty;
                WeChatRestoreStatusMessage = $"恢复完成：成功 {restoreResult.Success}，失败 {restoreResult.Failed}{skippedText}。";
                if (restoreResult.Failed > 0)
                {
                    MessageBox.Show(
                        $"恢复完成：成功 {restoreResult.Success}，失败 {restoreResult.Failed}{skippedText}（详见日志）。",
                        "微信恢复",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                }
                else
                {
                    MessageBox.Show(
                        $"恢复完成：成功 {restoreResult.Success}，失败 {restoreResult.Failed}{skippedText}（详见日志）。",
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

        private HashSet<WeChatDataCategory> BuildWeChatRestoreCategories()
        {
            return WeChatCleanupService.BuildCategories(
                WeChatRestoreIncludeText,
                WeChatRestoreIncludeImage,
                WeChatRestoreIncludeVideo,
                WeChatRestoreIncludeVoice,
                WeChatRestoreIncludeFile,
                WeChatRestoreIncludeCache);
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
            SafeFireAndForget(LoadRestoreManifestSummaryAsync(dialog.FileName));
        }

        private void OpenRecentWeChatBackup(object parameter)
        {
            var item = parameter as RecentWeChatBackupItem;
            if (item == null)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(item.FilePath) || !File.Exists(item.FilePath))
            {
                RecentWeChatBackups.Remove(item);
                OnPropertyChanged(nameof(HasRecentWeChatBackups));
                SafeFireAndForget(SaveRecentWeChatBackupsAsync());
                WeChatRestoreStatusMessage = "最近备份文件不存在，已从列表移除。";
                return;
            }

            WeChatRestoreZipPath = item.FilePath;
            SafeFireAndForget(LoadRestoreManifestSummaryAsync(item.FilePath));
            SafeFireAndForget(AddRecentWeChatBackupAsync(item.FilePath, item.FileCount, item.TotalBytes));
        }

        private async Task LoadRecentWeChatBackupsAsync()
        {
            var settings = await AppSettingsService.LoadAsync();
            var items = (settings.RecentWeChatBackups ?? new List<RecentWeChatBackupSettings>())
                .Where(item => item != null && !string.IsNullOrWhiteSpace(item.FilePath))
                .OrderByDescending(item => item.LastUsedAt)
                .Take(6)
                .Select(item => new RecentWeChatBackupItem(item.FilePath, item.LastUsedAt, item.FileCount, item.TotalBytes))
                .ToList();

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                ReplaceItems(RecentWeChatBackups, items);
                OnPropertyChanged(nameof(HasRecentWeChatBackups));
            });
        }

        private async Task AddRecentWeChatBackupAsync(string zipPath, int fileCount, long totalBytes)
        {
            if (string.IsNullOrWhiteSpace(zipPath))
            {
                return;
            }

            var normalizedPath = Path.GetFullPath(zipPath);
            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                var existing = RecentWeChatBackups
                    .FirstOrDefault(item => string.Equals(item.FilePath, normalizedPath, StringComparison.OrdinalIgnoreCase));
                if (existing != null)
                {
                    RecentWeChatBackups.Remove(existing);
                }

                RecentWeChatBackups.Insert(0, new RecentWeChatBackupItem(normalizedPath, DateTime.Now, fileCount, totalBytes));
                while (RecentWeChatBackups.Count > 6)
                {
                    RecentWeChatBackups.RemoveAt(RecentWeChatBackups.Count - 1);
                }

                OnPropertyChanged(nameof(HasRecentWeChatBackups));
            });

            await SaveRecentWeChatBackupsAsync();
        }

        private async Task SaveRecentWeChatBackupsAsync()
        {
            var snapshot = RecentWeChatBackups
                .Select(item => new RecentWeChatBackupSettings
                {
                    FilePath = item.FilePath,
                    LastUsedAt = item.LastUsedAt,
                    FileCount = item.FileCount,
                    TotalBytes = item.TotalBytes
                })
                .ToList();
            await AppSettingsService.UpdateAsync(settings => settings.RecentWeChatBackups = snapshot);
        }

        private async Task LoadRestoreManifestSummaryAsync(string zipPath)
        {
            try
            {
                var manifest = await WeChatBackupStore.ReadManifestAsync(zipPath, CancellationToken.None);
                var count = manifest.Entries?.Count ?? 0;
                var total = manifest.Entries?.Sum(x => x.Size) ?? 0L;
                var categories = manifest.Categories == null || manifest.Categories.Count == 0
                    ? "无"
                    : string.Join(", ", manifest.Categories);
                var categorySummary = BuildWeChatManifestCategorySummary(manifest);

                var builder = new StringBuilder();
                builder.AppendLine($"微信版本：{manifest.WechatVariant}");
                builder.AppendLine($"wxId：{manifest.WxId}");
                builder.AppendLine($"时间范围：{manifest.DateRange?.Start} ~ {manifest.DateRange?.End}");
                builder.AppendLine($"类别：{categories}");
                if (!string.IsNullOrWhiteSpace(categorySummary))
                {
                    builder.AppendLine($"类别统计：{categorySummary}");
                }

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

        private static string BuildWeChatManifestCategorySummary(WeChatBackupManifest manifest)
        {
            var entries = manifest?.Entries ?? new List<WeChatBackupManifestEntry>();
            var groups = entries
                .Where(item => !string.IsNullOrWhiteSpace(item.Category))
                .GroupBy(item => item.Category, StringComparer.OrdinalIgnoreCase)
                .Select(group =>
                {
                    var bytes = group.Sum(item => item.Size);
                    return $"{GetWeChatCategoryDisplay(group.Key)} {group.Count()} / {FileSizeFormatter.Format(bytes)}";
                })
                .ToList();

            return groups.Count == 0 ? string.Empty : string.Join("；", groups);
        }

        private static string GetWeChatCategoryDisplay(string category)
        {
            WeChatDataCategory value;
            if (!Enum.TryParse(category ?? string.Empty, true, out value))
            {
                return category ?? string.Empty;
            }

            switch (value)
            {
                case WeChatDataCategory.Text:
                    return "文字";
                case WeChatDataCategory.Image:
                    return "图像";
                case WeChatDataCategory.Video:
                    return "视频";
                case WeChatDataCategory.Voice:
                    return "语音";
                case WeChatDataCategory.File:
                    return "文件";
                case WeChatDataCategory.Cache:
                    return "缓存";
                default:
                    return value.ToString();
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
                await OptimizationReportsStore.DeleteAsync(selected.Select(x => x.Id));
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
            (CancelSqlExportCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ApplySqlRecentConnectionCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            _executeQueryCommand?.RaiseCanExecuteChanged();
            _exportQueryResultCommand?.RaiseCanExecuteChanged();
            _exportQueryResultCsvCommand?.RaiseCanExecuteChanged();
            _startAutoOptimizeCommand?.RaiseCanExecuteChanged();
            _startJunkScanCommand?.RaiseCanExecuteChanged();
            _runJunkCleanupCommand?.RaiseCanExecuteChanged();
            _exportJunkCleanupPlanCommand?.RaiseCanExecuteChanged();
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
            (RefreshSensorsOnceCommand as RelayCommand)?.RaiseCanExecuteChanged();
            _refreshInstalledProgramsCommand?.RaiseCanExecuteChanged();
            _uninstallProgramCommand?.RaiseCanExecuteChanged();
            _importCodexProfileCommand?.RaiseCanExecuteChanged();
            _importCodexCpaTokenCommand?.RaiseCanExecuteChanged();
            _exportCodexProfileCommand?.RaiseCanExecuteChanged();
            _previewCodexProfileDiffCommand?.RaiseCanExecuteChanged();
            _copyBenchmarkResultsCommand?.RaiseCanExecuteChanged();
            _exportBenchmarkResultsCommand?.RaiseCanExecuteChanged();
            (RetryConvertQueueItemCommand as AsyncRelayParameterCommand)?.RaiseCanExecuteChanged();
            (RemoveConvertQueueItemCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            (MoveConvertQueueItemUpCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            (MoveConvertQueueItemDownCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            (ToggleConvertPauseCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (CopyBatchFileHashColumnCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            (SelectFilteredInstalledProgramsCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ClearInstalledProgramSelectionCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ExportInstalledProgramListCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (CaptureVideoViewerFrameCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
            (GenerateVideoViewerWaveformCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
            (LoadVideoViewerSubtitleCommand as AsyncRelayCommand)?.RaiseCanExecuteChanged();
            (ClearVideoViewerSubtitleCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (RemoveVideoViewerPlaylistItemCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            (CopyVideoViewerPlaylistCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (CleanInvalidVideoViewerPlaylistItemsCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ClearVideoViewerPlaylistCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (SaveVideoViewerPlaylistCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (LoadVideoViewerPlaylistCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (FavoriteVideoViewerPlaylistCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (OpenFavoriteVideoViewerPlaylistCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            (RemoveFavoriteVideoViewerPlaylistCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            (SetVideoViewerLoopStartCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (SetVideoViewerLoopEndCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ClearVideoViewerLoopCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (ToggleVideoViewerLoopCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (SaveVideoViewerLoopRangeCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (OpenVideoViewerLoopRangeCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            (RemoveVideoViewerLoopRangeCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            (AddVideoViewerBookmarkCommand as RelayCommand)?.RaiseCanExecuteChanged();
            (OpenVideoViewerBookmarkCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            (RemoveVideoViewerBookmarkCommand as RelayParameterCommand)?.RaiseCanExecuteChanged();
            CommandManager.InvalidateRequerySuggested();
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

        private static async void SafeFireAndForget(Task task, [CallerMemberName] string caller = null)
        {
            try
            {
                await task.ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Unobserved async exception in {Caller}", caller ?? "unknown");
            }
        }
    }

        public class CodexProfileItem : INotifyPropertyChanged
    {
        private string _displayName;
        private string _name;
        private string _accountEmail;
        private string _note;
        private string _remark;
        private string _tags;
        private DateTime? _lastAppliedAt;
        private DateTime _lastImportedAt;
        private DateTime? _accessTokenExpiresAt;
        private DateTime? _refreshTokenExpiresAt;
        private string _protectedConfigTomlBase64;
        private string _protectedAuthJsonBase64;
        private bool _isActive;
        private bool _isApplying;
        private string _status;
        private string _statusMessage;
        private string _configTomlContentProtected;
        private string _authJsonContentProtected;
        private bool _enableRotation;
        private int _rotationPriority;

        public string DisplayName
        {
            get => _displayName;
            set
            {
                if (string.Equals(_displayName, value, StringComparison.Ordinal)) return;
                _displayName = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(EffectiveDisplayName));
            }
        }

        public string Name
        {
            get => string.IsNullOrWhiteSpace(_name) ? DisplayName : _name;
            set
            {
                if (string.Equals(_name, value, StringComparison.Ordinal)) return;
                _name = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(EffectiveDisplayName));
            }
        }

        public string FolderPath { get; set; }

        public string AccountEmail
        {
            get => _accountEmail;
            set
            {
                if (string.Equals(_accountEmail, value, StringComparison.Ordinal)) return;
                _accountEmail = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(AccountEmailDisplay));
            }
        }

        public string Note
        {
            get => _note;
            set
            {
                var next = value ?? string.Empty;
                if (next.Length > 200) next = next.Substring(0, 200);
                if (string.Equals(_note, next, StringComparison.Ordinal)) return;
                _note = next;
                OnPropertyChanged();
                OnPropertyChanged(nameof(NoteDisplay));
            }
        }

        public string Remark
        {
            get => _remark;
            set
            {
                if (string.Equals(_remark, value, StringComparison.Ordinal)) return;
                _remark = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        public string Tags
        {
            get => _tags;
            set
            {
                if (string.Equals(_tags, value, StringComparison.Ordinal)) return;
                _tags = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(TagsDisplay));
            }
        }

        public DateTime? LastAppliedAt
        {
            get => _lastAppliedAt;
            set
            {
                if (_lastAppliedAt == value) return;
                _lastAppliedAt = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(LastAppliedText));
            }
        }

        public DateTime LastImportedAt
        {
            get => _lastImportedAt;
            set
            {
                if (_lastImportedAt == value) return;
                _lastImportedAt = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(LastImportedText));
            }
        }

        public DateTime? AccessTokenExpiresAt
        {
            get => _accessTokenExpiresAt;
            set
            {
                if (_accessTokenExpiresAt == value) return;
                _accessTokenExpiresAt = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(AccessTokenExpiresText));
                OnPropertyChanged(nameof(RemainingValidityText));
            }
        }

        public DateTime? RefreshTokenExpiresAt
        {
            get => _refreshTokenExpiresAt;
            set
            {
                if (_refreshTokenExpiresAt == value) return;
                _refreshTokenExpiresAt = value;
                OnPropertyChanged();
            }
        }

        public string ProtectedConfigTomlBase64
        {
            get => _protectedConfigTomlBase64;
            set
            {
                if (string.Equals(_protectedConfigTomlBase64, value, StringComparison.Ordinal)) return;
                _protectedConfigTomlBase64 = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasEmbeddedContent));
                OnPropertyChanged(nameof(ContentStorageSummary));
            }
        }

        public string ProtectedAuthJsonBase64
        {
            get => _protectedAuthJsonBase64;
            set
            {
                if (string.Equals(_protectedAuthJsonBase64, value, StringComparison.Ordinal)) return;
                _protectedAuthJsonBase64 = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasEmbeddedContent));
                OnPropertyChanged(nameof(ContentStorageSummary));
            }
        }

        [JsonIgnore]
        public bool IsActive
        {
            get => _isActive;
            set
            {
                if (_isActive == value) return;
                _isActive = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(ActivePrefix));
            }
        }

        [JsonIgnore]
        public bool IsApplying
        {
            get => _isApplying;
            set
            {
                if (_isApplying == value) return;
                _isApplying = value;
                OnPropertyChanged();
            }
        }

        public string Status
        {
            get => _status;
            set
            {
                if (string.Equals(_status, value, StringComparison.Ordinal)) return;
                _status = value ?? CodexProfileLibraryService.StatusUnknown;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsStatusOk));
                OnPropertyChanged(nameof(IsStatusWarn));
                OnPropertyChanged(nameof(IsStatusExpired));
                OnPropertyChanged(nameof(IsStatusUnknown));
            }
        }

        [JsonIgnore]
        public string StatusMessage
        {
            get => _statusMessage;
            set
            {
                if (string.Equals(_statusMessage, value, StringComparison.Ordinal)) return;
                _statusMessage = value ?? string.Empty;
                OnPropertyChanged();
            }
        }

        public string ConfigTomlContentProtected
        {
            get => _configTomlContentProtected;
            set
            {
                if (string.Equals(_configTomlContentProtected, value, StringComparison.Ordinal)) return;
                _configTomlContentProtected = value ?? string.Empty;
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
                if (string.Equals(_authJsonContentProtected, value, StringComparison.Ordinal)) return;
                _authJsonContentProtected = value ?? string.Empty;
                OnPropertyChanged();
                OnPropertyChanged(nameof(HasEmbeddedContent));
                OnPropertyChanged(nameof(ContentStorageSummary));
            }
        }

        public bool EnableRotation
        {
            get => _enableRotation;
            set
            {
                if (_enableRotation == value) return;
                _enableRotation = value;
                OnPropertyChanged();
            }
        }

        public int RotationPriority
        {
            get => _rotationPriority;
            set
            {
                if (_rotationPriority == value) return;
                _rotationPriority = value;
                OnPropertyChanged();
            }
        }

        [JsonIgnore]
        public string EffectiveDisplayName => string.IsNullOrWhiteSpace(DisplayName) ? Name : DisplayName;

        [JsonIgnore]
        public string AccountEmailDisplay => string.IsNullOrWhiteSpace(AccountEmail) ? "未识别邮箱" : CodexProfileLibraryService.MaskEmail(AccountEmail);

        [JsonIgnore]
        public string NoteDisplay => string.IsNullOrWhiteSpace(Note) ? "未设置备注" : Note;

        [JsonIgnore]
        public string ActivePrefix => IsActive ? "⚡ " : string.Empty;

        [JsonIgnore]
        public bool IsStatusOk => string.Equals(Status, CodexProfileLibraryService.StatusOk, StringComparison.Ordinal);

        [JsonIgnore]
        public bool IsStatusWarn => string.Equals(Status, CodexProfileLibraryService.StatusWarn, StringComparison.Ordinal);

        [JsonIgnore]
        public bool IsStatusExpired => string.Equals(Status, CodexProfileLibraryService.StatusExpired, StringComparison.Ordinal);

        [JsonIgnore]
        public bool IsStatusUnknown => !IsStatusOk && !IsStatusWarn && !IsStatusExpired;

        [JsonIgnore]
        public bool HasEmbeddedContent =>
            (!string.IsNullOrWhiteSpace(ProtectedConfigTomlBase64) || !string.IsNullOrWhiteSpace(ConfigTomlContentProtected))
            && (!string.IsNullOrWhiteSpace(ProtectedAuthJsonBase64) || !string.IsNullOrWhiteSpace(AuthJsonContentProtected));

        [JsonIgnore]
        public string ContentStorageSummary =>
            HasEmbeddedContent
                ? "已加密保存 config.toml 和 auth.json"
                : "未保存配置内容（请重新导入当前账号）";

        [JsonIgnore]
        public string TagsDisplay => string.IsNullOrWhiteSpace(Tags) ? "未设置标签" : Tags;

        [JsonIgnore]
        public string LastAppliedText =>
            LastAppliedAt.HasValue
                ? "最近切换：" + LastAppliedAt.Value.ToString("MM-dd HH:mm")
                : "最近切换：-";

        [JsonIgnore]
        public string LastImportedText =>
            LastImportedAt == default(DateTime)
                ? "最后更新：-"
                : "最后更新：" + LastImportedAt.ToLocalTime().ToString("yyyy-MM-dd HH:mm");

        [JsonIgnore]
        public string AccessTokenExpiresText =>
            AccessTokenExpiresAt.HasValue
                ? "Access 过期：" + AccessTokenExpiresAt.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm")
                : "Access 过期：未知";

        [JsonIgnore]
        public string RemainingValidityText
        {
            get
            {
                if (!AccessTokenExpiresAt.HasValue)
                {
                    return "剩余：未知";
                }

                var span = AccessTokenExpiresAt.Value - DateTime.UtcNow;
                if (span.TotalSeconds <= 0)
                {
                    span = DateTime.UtcNow - AccessTokenExpiresAt.Value;
                    return span.TotalHours >= 1
                        ? $"已过期 {Math.Floor(span.TotalHours)} 小时"
                        : $"已过期 {Math.Max(1, Math.Floor(span.TotalMinutes))} 分钟";
                }

                return span.TotalDays >= 1
                    ? $"剩余 {Math.Floor(span.TotalDays)} 天 {span.Hours} 小时"
                    : $"剩余 {Math.Max(1, Math.Floor(span.TotalHours))} 小时";
            }
        }

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

    public class HomeCommandItem
    {
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public string Keywords { get; set; }
        public ICommand Command { get; set; }

        public bool Matches(string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return true;
            }

            var value = query.Trim();
            return Contains(Title, value)
                   || Contains(Subtitle, value)
                   || Contains(Keywords, value);
        }

        private static bool Contains(string source, string value)
        {
            return (source ?? string.Empty).IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }

    public class HomeRecentItem
    {
        public string Title { get; set; }
        public string Subtitle { get; set; }
        public string Module { get; set; }
        public string FilePath { get; set; }
        public string ActionText { get; set; }
        public DateTime LastUsedAt { get; set; }

        public string TimeText => LastUsedAt == default(DateTime)
            ? string.Empty
            : LastUsedAt.ToString("MM-dd HH:mm");
    }

    public class ScreenshotHistoryItem
    {
        public string Title { get; set; }
        public string FilePath { get; set; }
        public BitmapSource Thumbnail { get; set; }
        public DateTime CreatedAt { get; set; }
        public string SizeText { get; set; }
        public string DimensionsText { get; set; }

        public string TimeText => CreatedAt == default(DateTime)
            ? string.Empty
            : CreatedAt.ToString("MM-dd HH:mm:ss");
    }

    public class VideoPlaylistItem : INotifyPropertyChanged
    {
        private string _indexText;
        private bool _isCurrent;

        public VideoPlaylistItem(string filePath, int index)
        {
            FilePath = filePath ?? string.Empty;
            FileName = Path.GetFileName(FilePath);
            DirectoryPath = Path.GetDirectoryName(FilePath) ?? string.Empty;
            IndexText = index.ToString("00");
        }

        public string FilePath { get; }
        public string FileName { get; }
        public string DirectoryPath { get; }
        public string SearchText => string.Join(" ", new[] { IndexText, FileName, DirectoryPath, FilePath });

        public string IndexText
        {
            get => _indexText;
            set { _indexText = value ?? string.Empty; OnPropertyChanged(); }
        }

        public bool IsCurrent
        {
            get => _isCurrent;
            set { _isCurrent = value; OnPropertyChanged(); }
        }

        public bool Matches(string query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return true;
            }

            return SearchText.IndexOf(query.Trim(), StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class RecentPlaylistItem
    {
        public RecentPlaylistItem(string filePath, DateTime lastUsedAt)
        {
            FilePath = filePath ?? string.Empty;
            Title = string.IsNullOrWhiteSpace(FilePath) ? "播放列表" : Path.GetFileName(FilePath);
            DirectoryPath = string.IsNullOrWhiteSpace(FilePath) ? string.Empty : Path.GetDirectoryName(FilePath) ?? string.Empty;
            LastUsedAt = lastUsedAt;
        }

        public string Title { get; }
        public string FilePath { get; }
        public string DirectoryPath { get; }
        public DateTime LastUsedAt { get; }

        public string TimeText => LastUsedAt == default(DateTime)
            ? string.Empty
            : LastUsedAt.ToString("MM-dd HH:mm");
    }

    public class RecentWeChatBackupItem
    {
        public RecentWeChatBackupItem(string filePath, DateTime lastUsedAt, int fileCount, long totalBytes)
        {
            FilePath = filePath ?? string.Empty;
            Title = string.IsNullOrWhiteSpace(FilePath) ? "微信备份" : Path.GetFileName(FilePath);
            DirectoryPath = string.IsNullOrWhiteSpace(FilePath) ? string.Empty : Path.GetDirectoryName(FilePath) ?? string.Empty;
            LastUsedAt = lastUsedAt;
            FileCount = fileCount;
            TotalBytes = totalBytes;
        }

        public string Title { get; }
        public string FilePath { get; }
        public string DirectoryPath { get; }
        public DateTime LastUsedAt { get; }
        public int FileCount { get; }
        public long TotalBytes { get; }
        public string SummaryText => $"{FileCount:N0} 个文件 · {FileSizeFormatter.Format(TotalBytes)}";
        public string TimeText => LastUsedAt == default(DateTime)
            ? string.Empty
            : LastUsedAt.ToString("MM-dd HH:mm");
    }

    public class FavoritePlaylistItem : INotifyPropertyChanged
    {
        private DateTime _lastUsedAt;

        public FavoritePlaylistItem(string title, IEnumerable<string> filePaths, DateTime createdAt, DateTime lastUsedAt)
        {
            Title = string.IsNullOrWhiteSpace(title) ? "收藏播放列表" : title.Trim();
            FilePaths = (filePaths ?? Enumerable.Empty<string>())
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            CreatedAt = createdAt == default(DateTime) ? DateTime.Now : createdAt;
            _lastUsedAt = lastUsedAt == default(DateTime) ? CreatedAt : lastUsedAt;
        }

        public string Title { get; }
        public IReadOnlyList<string> FilePaths { get; }
        public DateTime CreatedAt { get; }

        public DateTime LastUsedAt
        {
            get => _lastUsedAt;
            set
            {
                _lastUsedAt = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(TimeText));
            }
        }

        public string CountText => $"{FilePaths.Count} 项";
        public string PreviewText => string.Join(" / ", FilePaths.Take(2).Select(Path.GetFileName));
        public string FullPathText => string.Join(Environment.NewLine, FilePaths);
        public string TimeText => LastUsedAt.ToString("MM-dd HH:mm");

        public bool HasSameFiles(IReadOnlyList<string> filePaths)
        {
            if (filePaths == null || filePaths.Count != FilePaths.Count)
            {
                return false;
            }

            return !FilePaths
                .Except(filePaths, StringComparer.OrdinalIgnoreCase)
                .Any();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class VideoLoopRangeItem
    {
        public VideoLoopRangeItem(double startSeconds, double endSeconds, string fileName, string filePath)
        {
            StartSeconds = Math.Max(0, Math.Min(startSeconds, endSeconds));
            EndSeconds = Math.Max(StartSeconds, Math.Max(startSeconds, endSeconds));
            FileName = fileName ?? string.Empty;
            FilePath = filePath ?? string.Empty;
        }

        public double StartSeconds { get; }
        public double EndSeconds { get; }
        public string FileName { get; }
        public string FilePath { get; }
        public string StartText => FormatMediaTime(StartSeconds);
        public string EndText => FormatMediaTime(EndSeconds);
        public string DurationText => FormatMediaTime(Math.Max(0, EndSeconds - StartSeconds));
        public string RangeText => StartText + " - " + EndText;
        public string Title => RangeText + " · " + DurationText;

        private static string FormatMediaTime(double seconds)
        {
            if (seconds <= 0 || double.IsNaN(seconds) || double.IsInfinity(seconds))
            {
                return "00:00";
            }

            var time = TimeSpan.FromSeconds(seconds);
            return time.TotalHours >= 1
                ? $"{(int)time.TotalHours:00}:{time.Minutes:00}:{time.Seconds:00}"
                : $"{time.Minutes:00}:{time.Seconds:00}";
        }
    }

    public class VideoBookmarkItem
    {
        public VideoBookmarkItem(double positionSeconds, string fileName, string filePath)
        {
            PositionSeconds = Math.Max(0, positionSeconds);
            FileName = fileName ?? string.Empty;
            FilePath = filePath ?? string.Empty;
        }

        public double PositionSeconds { get; }
        public string FileName { get; }
        public string FilePath { get; }
        public string TimeText => FormatBookmarkTime(PositionSeconds);
        public string Title => string.IsNullOrWhiteSpace(FileName)
            ? TimeText
            : TimeText + " · " + FileName;

        private static string FormatBookmarkTime(double seconds)
        {
            if (seconds <= 0 || double.IsNaN(seconds) || double.IsInfinity(seconds))
            {
                return "00:00";
            }

            var time = TimeSpan.FromSeconds(seconds);
            return time.TotalHours >= 1
                ? $"{(int)time.TotalHours:00}:{time.Minutes:00}:{time.Seconds:00}"
                : $"{time.Minutes:00}:{time.Seconds:00}";
        }
    }

    public class ConvertQueueItem : INotifyPropertyChanged
    {
        private int _number;
        private string _fileName;
        private string _sourcePath;
        private string _kind;
        private string _status;
        private string _outputPath;
        private string _message;

        public int Number
        {
            get => _number;
            set
            {
                if (_number == value)
                {
                    return;
                }

                _number = value;
                OnPropertyChanged();
            }
        }

        public string FileName
        {
            get => _fileName;
            set { _fileName = value ?? string.Empty; OnPropertyChanged(); }
        }

        public string SourcePath
        {
            get => _sourcePath;
            set { _sourcePath = value ?? string.Empty; OnPropertyChanged(); }
        }

        public string Kind
        {
            get => _kind;
            set { _kind = value ?? string.Empty; OnPropertyChanged(); }
        }

        public string Status
        {
            get => _status;
            set { _status = value ?? string.Empty; OnPropertyChanged(); }
        }

        public string OutputPath
        {
            get => _outputPath;
            set { _outputPath = value ?? string.Empty; OnPropertyChanged(); }
        }

        public string Message
        {
            get => _message;
            set { _message = value ?? string.Empty; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    internal class ImportedHashEntry
    {
        public ImportedHashEntry(string fileName, string filePath, string kind, string hash)
        {
            FileName = fileName ?? string.Empty;
            FilePath = filePath ?? string.Empty;
            Kind = kind ?? string.Empty;
            Hash = hash ?? string.Empty;
        }

        public string FileName { get; }
        public string FilePath { get; }
        public string Kind { get; }
        public string Hash { get; }
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

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public void RaiseCanExecuteChanged()
        {
            CommandManager.InvalidateRequerySuggested();
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

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public void RaiseCanExecuteChanged()
        {
            CommandManager.InvalidateRequerySuggested();
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

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public void RaiseCanExecuteChanged()
        {
            CommandManager.InvalidateRequerySuggested();
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

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public void RaiseCanExecuteChanged()
        {
            CommandManager.InvalidateRequerySuggested();
        }
    }
}




