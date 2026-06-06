using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using MyTools.Services;
using MyTools.Shared;

namespace MyTools.ViewModels
{
    public partial class SystemSettingsViewModel
    {
        public ObservableCollection<WallpaperItem> Wallpapers { get; } = new ObservableCollection<WallpaperItem>();

        private WallpaperItem _selectedWallpaper;
        public WallpaperItem SelectedWallpaper
        {
            get => _selectedWallpaper;
            set { _selectedWallpaper = value; OnPropertyChanged(); CommandManager.InvalidateRequerySuggested(); }
        }

        private string _wallpaperStatusMessage = "图库目录：正在检测...";
        public string WallpaperStatusMessage
        {
            get => _wallpaperStatusMessage;
            set { _wallpaperStatusMessage = value; OnPropertyChanged(); }
        }

        // 显示方式 — 默认填充
        private WallpaperService.WallpaperStyle _wallpaperStyle = WallpaperService.WallpaperStyle.Fill;
        public WallpaperService.WallpaperStyle WallpaperStyle
        {
            get => _wallpaperStyle;
            set
            {
                if (_wallpaperStyle == value) return;
                _wallpaperStyle = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsStyleFill));
                OnPropertyChanged(nameof(IsStyleFit));
                OnPropertyChanged(nameof(IsStyleStretch));
                OnPropertyChanged(nameof(IsStyleTile));
                OnPropertyChanged(nameof(IsStyleCenter));
            }
        }
        public bool IsStyleFill { get => _wallpaperStyle == WallpaperService.WallpaperStyle.Fill; set { if (value) WallpaperStyle = WallpaperService.WallpaperStyle.Fill; } }
        public bool IsStyleFit { get => _wallpaperStyle == WallpaperService.WallpaperStyle.Fit; set { if (value) WallpaperStyle = WallpaperService.WallpaperStyle.Fit; } }
        public bool IsStyleStretch { get => _wallpaperStyle == WallpaperService.WallpaperStyle.Stretch; set { if (value) WallpaperStyle = WallpaperService.WallpaperStyle.Stretch; } }
        public bool IsStyleTile { get => _wallpaperStyle == WallpaperService.WallpaperStyle.Tile; set { if (value) WallpaperStyle = WallpaperService.WallpaperStyle.Tile; } }
        public bool IsStyleCenter { get => _wallpaperStyle == WallpaperService.WallpaperStyle.Center; set { if (value) WallpaperStyle = WallpaperService.WallpaperStyle.Center; } }

        public ICommand RefreshWallpaperLibraryCommand { get; private set; }
        public ICommand ImportWallpapersCommand { get; private set; }
        public ICommand SaveCurrentWallpaperCommand { get; private set; }
        public ICommand ApplyWallpaperCommand { get; private set; }
        public ICommand BrowseWallpaperInImageViewerCommand { get; private set; }
        public ICommand RemoveWallpaperCommand { get; private set; }
        public ICommand OpenWallpaperFolderCommand { get; private set; }

        private void InitWallpaperCommands()
        {
            RefreshWallpaperLibraryCommand = new RelayCommand(LoadWallpaperLibrary);
            ImportWallpapersCommand = new AsyncRelayCommand(ImportWallpapersAsync, () => !IsBusy);
            SaveCurrentWallpaperCommand = new AsyncRelayCommand(SaveCurrentWallpaperAsync, () => !IsBusy);
            ApplyWallpaperCommand = new RelayCommand(ApplyWallpaper, () => SelectedWallpaper != null);
            BrowseWallpaperInImageViewerCommand = new RelayCommand(BrowseInImageViewer, () => SelectedWallpaper != null);
            RemoveWallpaperCommand = new RelayCommand(RemoveSelectedWallpaper, () => SelectedWallpaper != null);
            OpenWallpaperFolderCommand = new RelayCommand(OpenWallpaperFolder);
            LoadWallpaperLibrary();
        }

        private void LoadWallpaperLibrary()
        {
            try
            {
                Wallpapers.Clear();
                foreach (var path in WallpaperService.ListLibrary())
                {
                    Wallpapers.Add(new WallpaperItem(path));
                }
                WallpaperStatusMessage = Wallpapers.Count == 0
                    ? "图库为空。点「导入图片」或「保存当前桌面到图库」来添加。目录：" + WallpaperService.LibraryDirectory
                    : $"图库共 {Wallpapers.Count} 张图片。目录：{WallpaperService.LibraryDirectory}";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Load wallpaper library failed.");
                WallpaperStatusMessage = "加载图库失败：" + ex.Message;
            }
        }

        private async Task ImportWallpapersAsync()
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Title = "选择要导入图库的图片",
                Filter = "图片文件|*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.tif|所有文件|*.*",
                Multiselect = true
            };
            if (dialog.ShowDialog() != true) return;

            try
            {
                IsBusy = true;
                WallpaperStatusMessage = $"正在导入 {dialog.FileNames.Length} 张图片...";
                var imported = await WallpaperService.ImportImagesAsync(dialog.FileNames).ConfigureAwait(true);
                LoadWallpaperLibrary();
                WallpaperStatusMessage = $"已导入 {imported.Count} 张图片。";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Import wallpapers failed.");
                WallpaperStatusMessage = "导入失败：" + ex.Message;
                MessageBox.Show(ex.Message, "导入失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { IsBusy = false; }
        }

        private async Task SaveCurrentWallpaperAsync()
        {
            try
            {
                IsBusy = true;
                WallpaperStatusMessage = "正在读取当前桌面壁纸...";
                var saved = await WallpaperService.SaveCurrentWallpaperToLibraryAsync().ConfigureAwait(true);
                LoadWallpaperLibrary();
                WallpaperStatusMessage = "已保存当前桌面壁纸：" + Path.GetFileName(saved) + "。目录：" + WallpaperService.LibraryDirectory;
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Save current wallpaper failed.");
                WallpaperStatusMessage = "保存失败：" + ex.Message;
                MessageBox.Show(ex.Message, "保存失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { IsBusy = false; }
        }

        private void ApplyWallpaper()
        {
            var item = SelectedWallpaper;
            if (item == null) return;
            try
            {
                WallpaperService.SetWallpaper(item.Path, WallpaperStyle);
                WallpaperStatusMessage = $"已设为桌面背景：{item.FileName}（{StyleLabel(WallpaperStyle)}）";
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Apply wallpaper failed.");
                WallpaperStatusMessage = "设置失败：" + ex.Message;
                MessageBox.Show(ex.Message, "设置失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void BrowseInImageViewer()
        {
            var item = SelectedWallpaper;
            if (item == null) return;
            try
            {
                var mainVm = Application.Current?.MainWindow?.DataContext as MainViewModel;
                if (mainVm == null)
                {
                    MessageBox.Show("无法找到主窗口，请稍后重试。", "提示", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                mainVm.TryOpenImageViewerFile(item.Path);
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Browse wallpaper in image viewer failed.");
                MessageBox.Show(ex.Message, "打开失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void RemoveSelectedWallpaper()
        {
            var item = SelectedWallpaper;
            if (item == null) return;
            var confirm = MessageBox.Show($"确定从图库移除「{item.FileName}」吗？文件将被删除。",
                "移除确认", MessageBoxButton.OKCancel, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.OK) return;
            try
            {
                WallpaperService.DeleteFromLibrary(item.Path);
                LoadWallpaperLibrary();
                WallpaperStatusMessage = "已移除：" + item.FileName;
            }
            catch (Exception ex)
            {
                AppLogService.Error(ex, "Remove wallpaper failed.");
                MessageBox.Show(ex.Message, "移除失败", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenWallpaperFolder()
        {
            try
            {
                WallpaperService.EnsureLibrary();
                System.Diagnostics.Process.Start("explorer.exe", "\"" + WallpaperService.LibraryDirectory + "\"");
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Open wallpaper folder failed: {Msg}", ex.Message);
            }
        }

        private static string StyleLabel(WallpaperService.WallpaperStyle s)
        {
            switch (s)
            {
                case WallpaperService.WallpaperStyle.Fit: return "适应";
                case WallpaperService.WallpaperStyle.Stretch: return "拉伸";
                case WallpaperService.WallpaperStyle.Tile: return "平铺";
                case WallpaperService.WallpaperStyle.Center: return "居中";
                default: return "填充";
            }
        }
    }

    /// <summary>图库列表项。带懒加载缩略图。</summary>
    public class WallpaperItem
    {
        public string Path { get; }
        public string FileName { get; }

        private BitmapImage _thumbnail;
        public BitmapImage Thumbnail
        {
            get
            {
                if (_thumbnail == null) _thumbnail = LoadThumbnail(Path);
                return _thumbnail;
            }
        }

        public WallpaperItem(string path)
        {
            Path = path;
            FileName = System.IO.Path.GetFileName(path);
        }

        private static BitmapImage LoadThumbnail(string path)
        {
            try
            {
                var img = new BitmapImage();
                img.BeginInit();
                img.CacheOption = BitmapCacheOption.OnLoad;
                img.DecodePixelWidth = 300; // 缩略图宽 300px，节省内存
                img.UriSource = new Uri(path, UriKind.Absolute);
                img.EndInit();
                img.Freeze();
                return img;
            }
            catch { return null; }
        }
    }
}
