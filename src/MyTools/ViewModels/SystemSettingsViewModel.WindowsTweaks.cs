using System;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;
using MyTools.Services;
using MyTools.Shared;

namespace MyTools.ViewModels
{
    public partial class SystemSettingsViewModel
    {
        private bool _tweaksLoading;

        public ICommand RestartExplorerCommand { get; private set; }
        public ICommand RefreshTweaksCommand { get; private set; }

        public bool IsWindows11 => WindowsTweaksService.IsWindows11;
        public bool SupportsSecondsInClock => WindowsTweaksService.SupportsSecondsInClock;
        public string TweaksHint => IsWindows11
            ? "当前系统：" + OsVersionService.DisplayName + "。任务栏、托盘和右键菜单改动需重启资源管理器后生效。"
            : "当前系统：" + OsVersionService.DisplayName + "。Win11 专属项在当前系统不可用，其余项 Win10/11 均可。";

        // ========== 时钟秒 ==========
        private bool _showSecondsInClock;
        public bool ShowSecondsInClock
        {
            get => _showSecondsInClock;
            set
            {
                if (_showSecondsInClock == value) return;
                _showSecondsInClock = value;
                OnPropertyChanged();
                if (!_tweaksLoading) TryWriteTweak(() => WindowsTweaksService.SetShowSecondsInClock(value), nameof(ShowSecondsInClock));
            }
        }

        // ========== 托盘图标全部显示 ==========
        private bool _trayShowAll;
        public bool TrayShowAll
        {
            get => _trayShowAll;
            set
            {
                if (_trayShowAll == value) return;
                _trayShowAll = value;
                OnPropertyChanged();
                if (!_tweaksLoading) TryWriteTweak(() => WindowsTweaksService.SetTrayShowAll(value), nameof(TrayShowAll));
            }
        }

        // ========== 任务栏合并 三档 ==========
        private WindowsTweaksService.TaskbarGlom _taskbarGlom;
        public WindowsTweaksService.TaskbarGlom TaskbarGlom
        {
            get => _taskbarGlom;
            set
            {
                if (_taskbarGlom == value) return;
                _taskbarGlom = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsGlomAlways));
                OnPropertyChanged(nameof(IsGlomWhenFull));
                OnPropertyChanged(nameof(IsGlomNever));
                if (!_tweaksLoading) TryWriteTweak(() => WindowsTweaksService.SetTaskbarGlom(value), nameof(TaskbarGlom));
            }
        }
        public bool IsGlomAlways
        {
            get => _taskbarGlom == WindowsTweaksService.TaskbarGlom.AlwaysCombine;
            set { if (value) TaskbarGlom = WindowsTweaksService.TaskbarGlom.AlwaysCombine; }
        }
        public bool IsGlomWhenFull
        {
            get => _taskbarGlom == WindowsTweaksService.TaskbarGlom.CombineWhenFull;
            set { if (value) TaskbarGlom = WindowsTweaksService.TaskbarGlom.CombineWhenFull; }
        }
        public bool IsGlomNever
        {
            get => _taskbarGlom == WindowsTweaksService.TaskbarGlom.NeverCombine;
            set { if (value) TaskbarGlom = WindowsTweaksService.TaskbarGlom.NeverCombine; }
        }

        private bool _useClassicContextMenu;
        public bool UseClassicContextMenu
        {
            get => _useClassicContextMenu;
            set
            {
                if (_useClassicContextMenu == value) return;
                _useClassicContextMenu = value;
                OnPropertyChanged();
                if (!_tweaksLoading) TryWriteTweak(() => WindowsTweaksService.SetUseClassicContextMenu(value), nameof(UseClassicContextMenu));
            }
        }

        // ========== 桌面图标 ==========
        private bool _showComputer;
        public bool ShowComputer
        {
            get => _showComputer;
            set { if (_showComputer == value) return; _showComputer = value; OnPropertyChanged();
                if (!_tweaksLoading) TryWriteTweak(() => WindowsTweaksService.SetDesktopIconVisible(WindowsTweaksService.ClsidComputer, value), nameof(ShowComputer)); }
        }
        private bool _showRecycleBin;
        public bool ShowRecycleBin
        {
            get => _showRecycleBin;
            set { if (_showRecycleBin == value) return; _showRecycleBin = value; OnPropertyChanged();
                if (!_tweaksLoading) TryWriteTweak(() => WindowsTweaksService.SetDesktopIconVisible(WindowsTweaksService.ClsidRecycleBin, value), nameof(ShowRecycleBin)); }
        }
        private bool _showControlPanel;
        public bool ShowControlPanel
        {
            get => _showControlPanel;
            set { if (_showControlPanel == value) return; _showControlPanel = value; OnPropertyChanged();
                if (!_tweaksLoading) TryWriteTweak(() => WindowsTweaksService.SetDesktopIconVisible(WindowsTweaksService.ClsidControlPanel, value), nameof(ShowControlPanel)); }
        }
        private bool _showUserFiles;
        public bool ShowUserFiles
        {
            get => _showUserFiles;
            set { if (_showUserFiles == value) return; _showUserFiles = value; OnPropertyChanged();
                if (!_tweaksLoading) TryWriteTweak(() => WindowsTweaksService.SetDesktopIconVisible(WindowsTweaksService.ClsidUserFiles, value), nameof(ShowUserFiles)); }
        }
        private bool _showNetwork;
        public bool ShowNetwork
        {
            get => _showNetwork;
            set { if (_showNetwork == value) return; _showNetwork = value; OnPropertyChanged();
                if (!_tweaksLoading) TryWriteTweak(() => WindowsTweaksService.SetDesktopIconVisible(WindowsTweaksService.ClsidNetwork, value), nameof(ShowNetwork)); }
        }

        // ========== 状态消息 ==========
        private string _tweaksStatusMessage = "切换开关后即时写入注册表。";
        public string TweaksStatusMessage
        {
            get => _tweaksStatusMessage;
            set { _tweaksStatusMessage = value; OnPropertyChanged(); }
        }

        private void InitWindowsTweaks()
        {
            RestartExplorerCommand = new RelayCommand(RestartExplorer);
            RefreshTweaksCommand = new RelayCommand(LoadWindowsTweaks);
            LoadWindowsTweaks();
        }

        private void LoadWindowsTweaks()
        {
            try
            {
                _tweaksLoading = true;
                ShowSecondsInClock = WindowsTweaksService.GetShowSecondsInClock();
                TrayShowAll = WindowsTweaksService.GetTrayShowAll();
                TaskbarGlom = WindowsTweaksService.GetTaskbarGlom();
                ShowComputer = WindowsTweaksService.GetDesktopIconVisible(WindowsTweaksService.ClsidComputer);
                ShowRecycleBin = WindowsTweaksService.GetDesktopIconVisible(WindowsTweaksService.ClsidRecycleBin);
                ShowControlPanel = WindowsTweaksService.GetDesktopIconVisible(WindowsTweaksService.ClsidControlPanel);
                ShowUserFiles = WindowsTweaksService.GetDesktopIconVisible(WindowsTweaksService.ClsidUserFiles);
                ShowNetwork = WindowsTweaksService.GetDesktopIconVisible(WindowsTweaksService.ClsidNetwork);
                UseClassicContextMenu = WindowsTweaksService.GetUseClassicContextMenu();
                TweaksStatusMessage = "已读取当前注册表状态。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Load Windows tweaks failed.");
                TweaksStatusMessage = "读取失败：" + ex.Message;
            }
            finally { _tweaksLoading = false; }
        }

        private void TryWriteTweak(Action write, string fieldName)
        {
            try
            {
                write();
                TweaksStatusMessage = NeedsExplorerRestart(fieldName)
                    ? "已写入：" + GetTweakDisplayName(fieldName) + "。请点「重启资源管理器」后生效。"
                    : "已写入：" + GetTweakDisplayName(fieldName) + "。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Write Windows tweak failed: " + fieldName);
                TweaksStatusMessage = "写入失败：" + ex.Message;
                MessageBox.Show(ex.Message, "写入失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static bool NeedsExplorerRestart(string fieldName)
        {
            return fieldName == nameof(ShowSecondsInClock)
                   || fieldName == nameof(TrayShowAll)
                   || fieldName == nameof(TaskbarGlom)
                   || fieldName == nameof(UseClassicContextMenu);
        }

        private static string GetTweakDisplayName(string fieldName)
        {
            switch (fieldName)
            {
                case nameof(ShowSecondsInClock):
                    return "任务栏时钟显示秒";
                case nameof(TrayShowAll):
                    return "显示所有托盘图标";
                case nameof(TaskbarGlom):
                    return "任务栏按钮合并方式";
                case nameof(UseClassicContextMenu):
                    return "右键直接显示完整菜单";
                case nameof(ShowComputer):
                    return "桌面图标：计算机";
                case nameof(ShowRecycleBin):
                    return "桌面图标：回收站";
                case nameof(ShowControlPanel):
                    return "桌面图标：控制面板";
                case nameof(ShowUserFiles):
                    return "桌面图标：用户的文件";
                case nameof(ShowNetwork):
                    return "桌面图标：网络";
                default:
                    return fieldName;
            }
        }

        private void RestartExplorer()
        {
            var confirm = MessageBox.Show(
                "将重启 Windows 资源管理器（explorer.exe），任务栏会短暂消失再恢复。是否继续？",
                "重启资源管理器",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Question);
            if (confirm != MessageBoxResult.OK) return;

            try
            {
                WindowsTweaksService.RestartExplorer();
                TweaksStatusMessage = "已重启资源管理器。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Restart explorer failed.");
                TweaksStatusMessage = "重启失败：" + ex.Message;
                MessageBox.Show(ex.Message, "重启失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
