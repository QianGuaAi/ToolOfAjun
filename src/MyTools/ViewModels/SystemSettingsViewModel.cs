using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using MyTools.Services;
using WinForms = System.Windows.Forms;

namespace MyTools.ViewModels
{
    /// <summary>
    /// 系统设置：导出/导入程序数据与依赖文件。
    /// </summary>
    public class SystemSettingsViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public SystemSettingsViewModel()
        {
            ExportCommand = new AsyncRelayCommand(ExportAsync, () => !IsBusy);
            ImportCommand = new AsyncRelayCommand(ImportAsync, () => !IsBusy);
        }

        public ICommand ExportCommand { get; }
        public ICommand ImportCommand { get; }

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

        // ============================ Export ============================
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
                StatusMessage = "正在导出...";
                var result = await SystemBackupService.ExportAsync(targetFolder).ConfigureAwait(true);
                StatusMessage = $"导出完成：{result.FilesCopied} 个文件，{SystemBackupService.FormatBytes(result.TotalBytes)}。";

                var openFolder = MessageBox.Show(
                    $"已导出到：\n{result.BackupRoot}\n\n共 {result.FilesCopied} 个文件，{SystemBackupService.FormatBytes(result.TotalBytes)}。\n\n是否打开导出文件夹？",
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
        private async Task ImportAsync()
        {
            string sourceFolder;
            using (var dlg = new WinForms.FolderBrowserDialog())
            {
                dlg.Description = "选择之前导出的备份文件夹（包含 manifest.json）";
                dlg.ShowNewFolderButton = false;
                if (dlg.ShowDialog() != WinForms.DialogResult.OK) return;
                sourceFolder = dlg.SelectedPath;
            }

            var backupRoot = SystemBackupService.ResolveBackupRoot(sourceFolder);
            if (backupRoot == null)
            {
                MessageBox.Show(
                    "所选文件夹不是有效的备份目录。请选择以 MyToolsBackup_ 开头的文件夹（包含 manifest.json），\n或选择其上一级父文件夹（将自动选择最新的备份）。",
                    "导入失败", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                $"将从以下位置导入备份并覆盖当前数据：\n{backupRoot}\n\n" +
                "提示：MyTools.settings.json / MyTools.sqlhistory.json 使用 Windows DPAPI 加密，\n" +
                "若导出时与当前 Windows 账户不同，相关密码 / Codex 凭据可能无法解密。\n\n继续导入？",
                "确认导入", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.OK) return;

            try
            {
                IsBusy = true;
                StatusMessage = "正在导入...";
                var result = await SystemBackupService.ImportAsync(backupRoot).ConfigureAwait(true);

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

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
