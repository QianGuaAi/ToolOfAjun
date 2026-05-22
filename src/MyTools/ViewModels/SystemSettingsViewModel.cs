using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using MyTools.Services;
using MyTools.Shared;
using WinForms = System.Windows.Forms;

namespace MyTools.ViewModels
{
    /// <summary>
    /// 系统设置：导出/导入程序数据与依赖文件。
    /// </summary>
    public partial class SystemSettingsViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public SystemSettingsViewModel()
        {
            PreviewExportCommand = new AsyncRelayCommand(PreviewExportAsync, CanRunExportOrPreview);
            ExportCommand = new AsyncRelayCommand(ExportAsync, CanRunExportOrPreview);
            PreviewImportCommand = new AsyncRelayCommand(PreviewImportAsync, CanRunImportOrPreview);
            ImportCommand = new AsyncRelayCommand(ImportAsync, CanRunImportOrPreview);
            VerifyBackupCommand = new AsyncRelayCommand(VerifyBackupAsync, () => !IsBusy);
            AssociateMediaFilesCommand = new AsyncRelayCommand(AssociateMediaFilesAsync, () => !IsBusy);
            AssociateVideoFilesCommand = new AsyncRelayCommand(() => AssociateMediaFilesAsync(MediaAssociationKind.Video), () => !IsBusy);
            AssociateAudioFilesCommand = new AsyncRelayCommand(() => AssociateMediaFilesAsync(MediaAssociationKind.Audio), () => !IsBusy);
            RestoreMediaAssociationCommand = new AsyncRelayCommand(RestoreMediaAssociationAsync, () => !IsBusy);
            RefreshMediaAssociationStatusCommand = new RelayCommand(RefreshMediaAssociationStatus, () => !IsBusy);
            OpenDefaultAppsSettingsCommand = new RelayCommand(OpenDefaultAppsSettings, () => !IsBusy);
            RefreshMediaAssociationStatus();
            InitWallpaperCommands();
            InitWindowsTweaks();
        }

        public ICommand PreviewExportCommand { get; }
        public ICommand ExportCommand { get; }
        public ICommand PreviewImportCommand { get; }
        public ICommand ImportCommand { get; }
        public ICommand VerifyBackupCommand { get; }
        public ICommand AssociateMediaFilesCommand { get; }
        public ICommand AssociateVideoFilesCommand { get; }
        public ICommand AssociateAudioFilesCommand { get; }
        public ICommand RestoreMediaAssociationCommand { get; }
        public ICommand RefreshMediaAssociationStatusCommand { get; }
        public ICommand OpenDefaultAppsSettingsCommand { get; }

        private bool _isBusy;
        public bool IsBusy
        {
            get => _isBusy;
            private set
            {
                _isBusy = value;
                OnPropertyChanged();
                CommandManager.InvalidateRequerySuggested();
            }
        }

        private string _statusMessage = "可将当前程序的所有数据与依赖（含 ffmpeg、Codex、SQL 历史、排班数据等）导出到文件夹，便于备份或迁移到其它机器。";
        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value; OnPropertyChanged(); }
        }

        private string _mediaAssociationStatusMessage = "正在读取文件关联状态...";
        public string MediaAssociationStatusMessage
        {
            get => _mediaAssociationStatusMessage;
            set { _mediaAssociationStatusMessage = value; OnPropertyChanged(); }
        }

        private bool _backupSettings = true;
        public bool BackupSettings
        {
            get => _backupSettings;
            set { SetOption(ref _backupSettings, value); }
        }

        private bool _backupLocalData = true;
        public bool BackupLocalData
        {
            get => _backupLocalData;
            set { SetOption(ref _backupLocalData, value); }
        }

        private bool _backupCodex = true;
        public bool BackupCodex
        {
            get => _backupCodex;
            set { SetOption(ref _backupCodex, value); }
        }

        private bool _backupNativeBinaries = true;
        public bool BackupNativeBinaries
        {
            get => _backupNativeBinaries;
            set { SetOption(ref _backupNativeBinaries, value); }
        }

        private bool _restoreSettings = true;
        public bool RestoreSettings
        {
            get => _restoreSettings;
            set { SetOption(ref _restoreSettings, value); }
        }

        private bool _restoreLocalData = true;
        public bool RestoreLocalData
        {
            get => _restoreLocalData;
            set { SetOption(ref _restoreLocalData, value); }
        }

        private bool _restoreCodex = true;
        public bool RestoreCodex
        {
            get => _restoreCodex;
            set { SetOption(ref _restoreCodex, value); }
        }

        private bool _restoreNativeBinaries = true;
        public bool RestoreNativeBinaries
        {
            get => _restoreNativeBinaries;
            set { SetOption(ref _restoreNativeBinaries, value); }
        }

        // ============================ Export ============================
        private bool CanRunExportOrPreview()
        {
            return !IsBusy && BuildBackupSections() != BackupSection.None;
        }

        private async Task PreviewExportAsync()
        {
            try
            {
                IsBusy = true;
                StatusMessage = "正在生成导出预检...";
                var plan = await SystemBackupService.BuildExportPlanAsync(BuildBackupSections()).ConfigureAwait(true);
                StatusMessage = $"预检完成：将导出 {plan.Files} 个文件，{SystemBackupService.FormatBytes(plan.TotalBytes)}。";
                MessageBox.Show(
                    BuildExportPlanMessage(plan),
                    "导出预检",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Preview system data export failed.");
                StatusMessage = "预检失败：" + ex.Message;
                MessageBox.Show(ex.Message, "预检失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task ExportAsync()
        {
            string targetFolder;
            using (var dlg = new WinForms.FolderBrowserDialog())
            {
                dlg.Description = "选择导出位置（将在所选文件夹内新建一个时间戳子目录）";
                dlg.ShowNewFolderButton = true;
                if (dlg.ShowDialog() != WinForms.DialogResult.OK) return;
                targetFolder = dlg.SelectedPath;
            }

            try
            {
                IsBusy = true;
                StatusMessage = "正在生成导出预检...";
                var sections = BuildBackupSections();
                var plan = await SystemBackupService.BuildExportPlanAsync(sections).ConfigureAwait(true);
                var confirm = MessageBox.Show(
                    BuildExportPlanMessage(plan) + "\n\n是否按以上范围导出？",
                    "确认导出",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Question);
                if (confirm != MessageBoxResult.OK)
                {
                    StatusMessage = "已取消导出。";
                    return;
                }

                StatusMessage = "正在导出...";
                var result = await SystemBackupService.ExportAsync(targetFolder, sections).ConfigureAwait(true);
                StatusMessage = $"导出完成：{result.FilesCopied} 个文件，{SystemBackupService.FormatBytes(result.TotalBytes)}。";

                var sectionDetail = result.Sections.Count == 0
                    ? "（未复制任何类别）"
                    : string.Join("\n", result.Sections);
                var openFolder = MessageBox.Show(
                    $"已导出到：\n{result.BackupRoot}\n\n{sectionDetail}\n\n共 {result.FilesCopied} 个文件，{SystemBackupService.FormatBytes(result.TotalBytes)}。\n\n是否打开导出文件夹？",
                    "导出成功", MessageBoxButton.YesNo, MessageBoxImage.Information);
                if (openFolder == MessageBoxResult.Yes)
                {
                    try { System.Diagnostics.Process.Start("explorer.exe", "\"" + result.BackupRoot + "\""); }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Export system data failed.");
                StatusMessage = "导出失败：" + ex.Message;
                MessageBox.Show(ex.Message, "导出失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { IsBusy = false; }
        }

        // ============================ Import ============================
        private bool CanRunImportOrPreview()
        {
            return !IsBusy && BuildRestoreSections() != BackupSection.None;
        }

        private async Task PreviewImportAsync()
        {
            var backupRoot = SelectBackupRootFromDialog();
            if (backupRoot == null)
            {
                return;
            }

            try
            {
                IsBusy = true;
                StatusMessage = "正在生成导入预览...";
                var preview = await SystemBackupService.BuildImportPreviewAsync(backupRoot, BuildRestoreSections()).ConfigureAwait(true);
                StatusMessage = $"导入预览完成：备份内 {preview.IncomingFiles} 个文件，预计覆盖 {preview.ExistingTargetFiles} 个同名文件。";
                MessageBox.Show(
                    BuildImportPreviewMessage(preview),
                    "导入预览",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Preview system data import failed.");
                StatusMessage = "预览失败：" + ex.Message;
                MessageBox.Show(ex.Message, "预览失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private async Task ImportAsync()
        {
            var backupRoot = SelectBackupRootFromDialog();
            if (backupRoot == null)
            {
                return;
            }

            try
            {
                IsBusy = true;
                StatusMessage = "正在生成导入预览...";
                var sections = BuildRestoreSections();
                var preview = await SystemBackupService.BuildImportPreviewAsync(backupRoot, sections).ConfigureAwait(true);

                var dpapiPreviewHint = (preview.HadDpapiData && !string.Equals(preview.SourceUserName, Environment.UserName, StringComparison.OrdinalIgnoreCase))
                    ? $"\n\n注意：备份来自用户「{preview.SourceUserName}」，与当前用户「{Environment.UserName}」不同，DPAPI 加密数据（密码 / Codex 凭据）将无法解密。"
                    : string.Empty;
                var confirm = MessageBox.Show(
                    BuildImportPreviewMessage(preview) + dpapiPreviewHint + "\n\n导入会覆盖所选类别中的同名数据，是否继续？",
                    "确认导入",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Warning);
                if (confirm != MessageBoxResult.OK)
                {
                    StatusMessage = "已取消导入。";
                    return;
                }

                StatusMessage = "正在导入...";
                var result = await SystemBackupService.ImportAsync(backupRoot, sections).ConfigureAwait(true);

                var sectionDetail = result.Sections.Count == 0 ? "（未发现可导入数据）" : string.Join("\n", result.Sections);
                var dpapiHint = (result.HadDpapiData && !string.Equals(result.SourceUserName, Environment.UserName, StringComparison.OrdinalIgnoreCase))
                    ? $"\n\n注意：备份来自用户「{result.SourceUserName}」，与当前用户「{Environment.UserName}」不同，DPAPI 加密数据（密码 / Codex 凭据）将无法解密。"
                    : string.Empty;

                StatusMessage = $"导入完成：{result.FilesCopied} 个文件，{SystemBackupService.FormatBytes(result.TotalBytes)}。建议重启程序使设置生效。";

                MessageBox.Show(
                    $"导入完成：{result.FilesCopied} 个文件，{SystemBackupService.FormatBytes(result.TotalBytes)}。\n\n{sectionDetail}{dpapiHint}\n\n建议重启程序以使设置完全生效。",
                    "导入成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Import system data failed.");
                StatusMessage = "导入失败：" + ex.Message;
                MessageBox.Show(ex.Message, "导入失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { IsBusy = false; }
        }

        private async Task VerifyBackupAsync()
        {
            var backupRoot = SelectBackupRootFromDialog();
            if (backupRoot == null)
            {
                return;
            }

            try
            {
                IsBusy = true;
                StatusMessage = "正在校验备份包...";
                var result = await SystemBackupService.VerifyBackupAsync(backupRoot).ConfigureAwait(true);
                StatusMessage = result.IsValid
                    ? $"备份校验通过：{result.VerifiedFiles}/{result.ExpectedFiles} 个文件，{SystemBackupService.FormatBytes(result.TotalBytes)}。"
                    : $"备份校验发现 {result.Problems.Count} 个问题。";

                MessageBox.Show(
                    BuildBackupVerifyMessage(result),
                    result.IsValid ? "备份校验通过" : "备份校验异常",
                    MessageBoxButton.OK,
                    result.IsValid ? MessageBoxImage.Information : MessageBoxImage.Warning);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Verify system backup failed.");
                StatusMessage = "校验失败：" + ex.Message;
                MessageBox.Show(ex.Message, "校验失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        // ============================ File associations ============================
        private async Task AssociateMediaFilesAsync()
        {
            await AssociateMediaFilesAsync(MediaAssociationKind.All).ConfigureAwait(true);
        }

        private async Task AssociateMediaFilesAsync(MediaAssociationKind kind)
        {
            var kindName = GetMediaAssociationKindName(kind);
            var confirm = MessageBox.Show(
                $"将把常见{kindName}文件关联到阿君的工具。\n\n"
                + BuildMediaAssociationExtensionHint(kind)
                + "\n\n继续？",
                "关联" + kindName + "文件",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Question);
            if (confirm != MessageBoxResult.OK)
            {
                return;
            }

            try
            {
                IsBusy = true;
                StatusMessage = "正在写入" + kindName + "文件关联...";
                var appPath = System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName;
                var count = await Task.Run(() => MediaFileAssociationCore.RegisterForCurrentUser(appPath, kind)).ConfigureAwait(true);
                RefreshMediaAssociationStatus();
                StatusMessage = $"已关联 {count} 种{kindName}扩展名。若资源管理器未立即更新图标，请刷新窗口或重新登录。";
                MessageBox.Show(
                    $"已完成{kindName}文件关联，共 {count} 种扩展名。\n\n以后双击这些文件会用阿君的工具打开。",
                    "关联完成",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Associate media files failed.");
                StatusMessage = "关联失败：" + ex.Message;
                MessageBox.Show(ex.Message, "关联失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { IsBusy = false; }
        }

        private async Task RestoreMediaAssociationAsync()
        {
            var confirm = MessageBox.Show(
                "将移除当前用户下 MyTools 写入的音视频文件默认关联。\n\n"
                + "Windows 可能会恢复原默认程序；如果系统无法判断，下次双击文件时会提示重新选择默认应用。\n\n继续？",
                "恢复系统默认关联",
                MessageBoxButton.OKCancel,
                MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.OK)
            {
                return;
            }

            try
            {
                IsBusy = true;
                StatusMessage = "正在恢复系统默认音视频关联...";
                var count = await Task.Run(() => MediaFileAssociationCore.RestoreSystemDefaultForCurrentUser(MediaAssociationKind.All)).ConfigureAwait(true);
                RefreshMediaAssociationStatus();
                StatusMessage = $"已移除 {count} 项 MyTools 当前用户默认关联。";
                MessageBox.Show(
                    "已移除 MyTools 当前用户音视频关联。\n\n如资源管理器仍显示旧图标，请刷新窗口、重新登录，或在 Windows“默认应用”里重新选择。",
                    "恢复完成",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Restore media file associations failed.");
                StatusMessage = "恢复关联失败：" + ex.Message;
                MessageBox.Show(ex.Message, "恢复失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                IsBusy = false;
            }
        }

        private void RefreshMediaAssociationStatus()
        {
            try
            {
                var video = MediaFileAssociationCore.GetCurrentUserStatus(MediaAssociationKind.Video);
                var audio = MediaFileAssociationCore.GetCurrentUserStatus(MediaAssociationKind.Audio);
                MediaAssociationStatusMessage = video.Summary + "；" + audio.Summary;
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Read media file association status failed: {Message}", ex.Message);
                MediaAssociationStatusMessage = "无法读取当前文件关联状态。";
            }
        }

        private void OpenDefaultAppsSettings()
        {
            try
            {
                if (OsVersionService.IsWindows10OrGreater)
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = "ms-settings:defaultapps",
                        UseShellExecute = true
                    });
                }
                else
                {
                    System.Diagnostics.Process.Start("control.exe", "/name Microsoft.DefaultPrograms");
                }

                StatusMessage = "已打开 Windows 默认应用设置。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Open default apps settings failed.");
                StatusMessage = "打开默认应用设置失败：" + ex.Message;
                MessageBox.Show(ex.Message, "打开失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static string GetMediaAssociationKindName(MediaAssociationKind kind)
        {
            switch (kind)
            {
                case MediaAssociationKind.Video:
                    return "视频";
                case MediaAssociationKind.Audio:
                    return "音频";
                default:
                    return "音视频";
            }
        }

        private static string BuildMediaAssociationExtensionHint(MediaAssociationKind kind)
        {
            var extensions = MediaFileAssociationCore.GetExtensions(kind);
            var preview = string.Join("、", extensions.Take(8).Select(item => item.TrimStart('.').ToUpperInvariant()));
            return $"包含 {preview} 等 {extensions.Length} 种扩展名。";
        }

        private string SelectBackupRootFromDialog()
        {
            string sourceFolder;
            using (var dlg = new WinForms.FolderBrowserDialog())
            {
                dlg.Description = "选择之前导出的备份文件夹（包含 manifest.json）";
                dlg.ShowNewFolderButton = false;
                if (dlg.ShowDialog() != WinForms.DialogResult.OK)
                {
                    return null;
                }

                sourceFolder = dlg.SelectedPath;
            }

            var backupRoot = SystemBackupService.ResolveBackupRoot(sourceFolder);
            if (backupRoot != null)
            {
                return backupRoot;
            }

            MessageBox.Show(
                "所选文件夹不是有效的备份目录。请选择以 MyToolsBackup_ 开头的文件夹（包含 manifest.json），\n或选择其上一级父文件夹（将自动选择最新的备份）。",
                "导入失败",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            return null;
        }

        private BackupSection BuildBackupSections()
        {
            return BuildSections(BackupSettings, BackupLocalData, BackupCodex, BackupNativeBinaries);
        }

        private BackupSection BuildRestoreSections()
        {
            return BuildSections(RestoreSettings, RestoreLocalData, RestoreCodex, RestoreNativeBinaries);
        }

        private static BackupSection BuildSections(bool settings, bool localData, bool codex, bool nativeBinaries)
        {
            var sections = BackupSection.None;
            if (settings) sections |= BackupSection.Settings;
            if (localData) sections |= BackupSection.LocalAppData;
            if (codex) sections |= BackupSection.Codex;
            if (nativeBinaries) sections |= BackupSection.NativeBinaries;
            return sections;
        }

        private static string BuildExportPlanMessage(SystemBackupService.BackupPlan plan)
        {
            var lines = new List<string>
            {
                "导出预检",
                $"预计导出：{plan.Files} 个文件，{SystemBackupService.FormatBytes(plan.TotalBytes)}",
                string.Empty,
                "明细："
            };

            foreach (var item in plan.Items)
            {
                var status = item.Exists
                    ? $"{item.Files} 个文件，{SystemBackupService.FormatBytes(item.TotalBytes)}"
                    : "未找到，将跳过";
                lines.Add($"- {item.SectionName}：{status}");
                lines.Add($"  {item.SourcePath}");
            }

            if (plan.Skipped.Count > 0)
            {
                lines.Add(string.Empty);
                lines.Add("跳过：");
                lines.AddRange(plan.Skipped.Select(item => "- " + item));
            }

            return string.Join(Environment.NewLine, lines);
        }

        private static string BuildImportPreviewMessage(SystemBackupService.ImportPreview preview)
        {
            var builder = new StringBuilder();
            builder.AppendLine("导入预览");
            builder.AppendLine("备份目录：" + preview.BackupRoot);
            builder.AppendLine("来源：" + (preview.MachineName ?? "-") + " / " + (preview.SourceUserName ?? "-"));
            if (preview.CreatedAt != default(DateTime))
            {
                builder.AppendLine("创建时间：" + preview.CreatedAt.ToString("yyyy-MM-dd HH:mm:ss"));
            }

            builder.AppendLine($"备份内文件：{preview.IncomingFiles} 个，{SystemBackupService.FormatBytes(preview.IncomingBytes)}");
            builder.AppendLine($"预计覆盖同名文件：{preview.ExistingTargetFiles} 个；新增：{preview.NewTargetFiles} 个。");
            builder.AppendLine();
            builder.AppendLine("明细：");
            foreach (var item in preview.Items)
            {
                var status = item.Exists
                    ? $"{item.Files} 个文件，{SystemBackupService.FormatBytes(item.TotalBytes)}，覆盖 {item.ExistingTargetFiles} 个"
                    : "备份中缺失，将跳过";
                builder.AppendLine("- " + item.SectionName + "：" + status);
                builder.AppendLine("  " + item.BackupPath);
            }

            if (preview.Skipped.Count > 0)
            {
                builder.AppendLine();
                builder.AppendLine("跳过：");
                foreach (var item in preview.Skipped)
                {
                    builder.AppendLine("- " + item);
                }
            }

            return builder.ToString().TrimEnd();
        }

        private static string BuildBackupVerifyMessage(SystemBackupService.BackupVerifyResult result)
        {
            var lines = new List<string>
            {
                result.IsValid ? "备份包校验通过" : "备份包校验发现异常",
                "备份目录：" + (result.BackupRoot ?? string.Empty),
                $"校验文件：{result.VerifiedFiles}/{result.ExpectedFiles}",
                "总大小：" + SystemBackupService.FormatBytes(result.TotalBytes)
            };

            if (result.Problems.Count > 0)
            {
                lines.Add(string.Empty);
                lines.Add("问题：");
                lines.AddRange(result.Problems.Take(20).Select(item => "- " + item));
                if (result.Problems.Count > 20)
                {
                    lines.Add($"- 还有 {result.Problems.Count - 20} 项未显示");
                }
            }

            return string.Join(Environment.NewLine, lines);
        }

        private void SetOption(ref bool field, bool value)
        {
            if (field == value)
            {
                return;
            }

            field = value;
            OnPropertyChanged();
            CommandManager.InvalidateRequerySuggested();
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
