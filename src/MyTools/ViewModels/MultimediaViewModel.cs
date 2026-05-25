using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using MyTools.Services;
using MyTools.Views;

namespace MyTools.ViewModels
{
    public enum MultimediaPreferredFilter
    {
        All,
        Image,
        AudioVideo
    }

    public sealed class MultimediaViewModel : INotifyPropertyChanged, IDisposable
    {
        private readonly MainViewModel _owner;
        private bool _initialized;
        private CancellationTokenSource _loadCancellationTokenSource;
        private MediaFileItem _selectedMediaFile;
        private string _selectedFolderPath = string.Empty;
        private string _statusMessage = "选择左侧文件夹浏览图片、音频、视频和 PDF。";
        private MediaKind _previewKind = MediaKind.Other;
        private MediaKind _immersiveKind = MediaKind.Other;
        private bool _isImmersive;
        private bool _isPreviewMuted;
        private bool _isPreviewPlaying;
        private bool _isImmersiveMuted;
        private bool _isImmersivePlaying;
        private Uri _previewMediaSource;
        private Uri _immersiveMediaSource;
        private BitmapImage _previewImageSource;
        private BitmapImage _immersiveImageSource;
        private string _previewImageInfo = string.Empty;
        private MultimediaPreferredFilter _preferredFilter = MultimediaPreferredFilter.All;

        public MultimediaViewModel(MainViewModel owner)
        {
            _owner = owner;
            FolderTreeNodes = new ObservableCollection<MediaFolderNode>();
            MediaFiles = new ObservableCollection<MediaFileItem>();
            ExitMultimediaCommand = new RelayCommand(ExitMultimedia);
            ExitImmersiveCommand = new RelayCommand(ExitImmersive, () => IsImmersive);
            EnterImmersiveCommand = new RelayCommand(EnterImmersive, CanEnterImmersive);
            PreviousImageCommand = new RelayCommand(ShowPreviousImage, CanShowPreviousImage);
            NextImageCommand = new RelayCommand(ShowNextImage, CanShowNextImage);
            OpenConvertDialogCommand = new RelayParameterCommand(OpenConvertDialog, CanOpenConvertDialog);
            ConvertImageToPdfCommand = new AsyncRelayParameterCommand(ConvertImageToPdfAsync, CanConvertImageToPdf);
            ConvertPdfToImagesCommand = new AsyncRelayParameterCommand(ConvertPdfToImagesAsync, CanConvertPdfToImages);
            OpenExternalCommand = new RelayCommand(OpenSelectedExternally, () => SelectedMediaFile != null && File.Exists(SelectedMediaFile.Path));
        }

        public ObservableCollection<MediaFolderNode> FolderTreeNodes { get; }
        public ObservableCollection<MediaFileItem> MediaFiles { get; }
        public ICommand ExitMultimediaCommand { get; }
        public ICommand ExitImmersiveCommand { get; }
        public ICommand EnterImmersiveCommand { get; }
        public ICommand PreviousImageCommand { get; }
        public ICommand NextImageCommand { get; }
        public ICommand OpenConvertDialogCommand { get; }
        public ICommand ConvertImageToPdfCommand { get; }
        public ICommand ConvertPdfToImagesCommand { get; }
        public ICommand OpenExternalCommand { get; }

        public MultimediaPreferredFilter PreferredFilter
        {
            get => _preferredFilter;
            set
            {
                if (_preferredFilter == value) return;
                _preferredFilter = value;
                OnPropertyChanged();
            }
        }

        public MediaFileItem SelectedMediaFile
        {
            get => _selectedMediaFile;
            set
            {
                if (ReferenceEquals(_selectedMediaFile, value)) return;
                _selectedMediaFile = value;
                OnPropertyChanged();
                ApplySelectedMediaFile(value);
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public string SelectedFolderPath
        {
            get => _selectedFolderPath;
            private set { _selectedFolderPath = value ?? string.Empty; OnPropertyChanged(); }
        }

        public string StatusMessage
        {
            get => _statusMessage;
            set { _statusMessage = value ?? string.Empty; OnPropertyChanged(); }
        }

        public MediaKind PreviewKind
        {
            get => _previewKind;
            private set
            {
                if (_previewKind == value) return;
                _previewKind = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsPreviewImage));
                OnPropertyChanged(nameof(IsPreviewAudio));
                OnPropertyChanged(nameof(IsPreviewVideo));
                OnPropertyChanged(nameof(IsPreviewEmpty));
            }
        }

        public MediaKind ImmersiveKind
        {
            get => _immersiveKind;
            private set
            {
                if (_immersiveKind == value) return;
                _immersiveKind = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsImmersiveImage));
                OnPropertyChanged(nameof(IsImmersiveAudio));
                OnPropertyChanged(nameof(IsImmersiveVideo));
            }
        }

        public bool IsImmersive
        {
            get => _isImmersive;
            private set
            {
                if (_isImmersive == value) return;
                _isImmersive = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsExplorerMode));
                OnPropertyChanged(nameof(IsImmersiveImage));
                OnPropertyChanged(nameof(IsImmersiveAudio));
                OnPropertyChanged(nameof(IsImmersiveVideo));
                CommandManager.InvalidateRequerySuggested();
            }
        }

        public bool IsExplorerMode => !IsImmersive;
        public bool IsPreviewImage => PreviewKind == MediaKind.Image;
        public bool IsPreviewAudio => PreviewKind == MediaKind.Audio;
        public bool IsPreviewVideo => PreviewKind == MediaKind.Video;
        public bool IsPreviewEmpty => PreviewKind == MediaKind.Other || PreviewKind == MediaKind.Pdf;
        public bool IsImmersiveImage => IsImmersive && ImmersiveKind == MediaKind.Image;
        public bool IsImmersiveAudio => IsImmersive && ImmersiveKind == MediaKind.Audio;
        public bool IsImmersiveVideo => IsImmersive && ImmersiveKind == MediaKind.Video;

        public bool IsPreviewMuted
        {
            get => _isPreviewMuted;
            set { _isPreviewMuted = value; OnPropertyChanged(); }
        }

        public bool IsPreviewPlaying
        {
            get => _isPreviewPlaying;
            set { _isPreviewPlaying = value; OnPropertyChanged(); }
        }

        public bool IsImmersiveMuted
        {
            get => _isImmersiveMuted;
            set { _isImmersiveMuted = value; OnPropertyChanged(); }
        }

        public bool IsImmersivePlaying
        {
            get => _isImmersivePlaying;
            set { _isImmersivePlaying = value; OnPropertyChanged(); }
        }

        public Uri PreviewMediaSource
        {
            get => _previewMediaSource;
            private set { _previewMediaSource = value; OnPropertyChanged(); }
        }

        public Uri ImmersiveMediaSource
        {
            get => _immersiveMediaSource;
            private set { _immersiveMediaSource = value; OnPropertyChanged(); }
        }

        public BitmapImage PreviewImageSource
        {
            get => _previewImageSource;
            private set { _previewImageSource = value; OnPropertyChanged(); }
        }

        public BitmapImage ImmersiveImageSource
        {
            get => _immersiveImageSource;
            private set { _immersiveImageSource = value; OnPropertyChanged(); }
        }

        public string PreviewImageInfo
        {
            get => _previewImageInfo;
            private set { _previewImageInfo = value ?? string.Empty; OnPropertyChanged(); }
        }

        public async Task InitializeOnEnterAsync()
        {
            if (_initialized) return;
            _initialized = true;
            StatusMessage = "正在加载磁盘根节点...";
            var roots = await Task.Run(() =>
            {
                var nodes = new List<MediaFolderNode>();
                foreach (var drive in DriveInfo.GetDrives())
                {
                    try
                    {
                        if (!drive.IsReady) continue;
                        var label = string.IsNullOrEmpty(drive.VolumeLabel)
                            ? drive.Name.TrimEnd('\\')
                            : string.Format("{0} ({1})", drive.Name.TrimEnd('\\'), drive.VolumeLabel);
                        var node = new MediaFolderNode { Name = label, FullPath = drive.RootDirectory.FullName };
                        if (HasVisibleChildFolders(drive.RootDirectory.FullName))
                            node.AddDummyChild();
                        nodes.Add(node);
                    }
                    catch
                    {
                    }
                }
                return nodes;
            });

            FolderTreeNodes.Clear();
            foreach (var root in roots)
                FolderTreeNodes.Add(root);
            StatusMessage = roots.Count == 0 ? "未发现可用磁盘。" : "选择左侧文件夹浏览图片、音频、视频和 PDF。";
        }

        public async Task ExpandFolderAsync(MediaFolderNode node)
        {
            if (node == null || node.IsDummy || node.ChildrenLoaded || string.IsNullOrWhiteSpace(node.FullPath)) return;
            node.ChildrenLoaded = true;
            var children = await Task.Run(() => LoadChildFolders(node.FullPath));
            node.Children.Clear();
            foreach (var child in children)
                node.Children.Add(child);
        }

        public async Task SelectFolderAsync(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath) || !Directory.Exists(folderPath)) return;
            _loadCancellationTokenSource?.Cancel();
            _loadCancellationTokenSource?.Dispose();
            _loadCancellationTokenSource = new CancellationTokenSource();
            var token = _loadCancellationTokenSource.Token;
            SelectedFolderPath = folderPath;
            SelectedMediaFile = null;
            StopPreview();
            MediaFiles.Clear();
            StatusMessage = "正在加载媒体文件...";

            try
            {
                var files = await MediaFileTypeHelper.EnumerateMediaFilesAsync(folderPath, token);
                if (token.IsCancellationRequested) return;
                foreach (var descriptor in ApplyPreferredFilter(files))
                    MediaFiles.Add(new MediaFileItem(descriptor));
                StatusMessage = MediaFiles.Count == 0 ? "当前文件夹没有可显示的媒体文件。" : string.Format("已加载 {0} 个媒体文件。", MediaFiles.Count);
            }
            catch (OperationCanceledException)
            {
            }
            catch (UnauthorizedAccessException)
            {
                StatusMessage = "无权限访问该文件夹";
            }
            catch (Exception ex)
            {
                StatusMessage = "加载媒体文件失败：" + ex.Message;
            }
        }

        public async Task SelectFileAsync(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath)) return;
            var folder = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(folder))
                await SelectFolderAsync(folder);
            var item = MediaFiles.FirstOrDefault(file => string.Equals(file.Path, filePath, StringComparison.OrdinalIgnoreCase));
            if (item == null)
            {
                var info = new FileInfo(filePath);
                item = new MediaFileItem(new MediaFileDescriptor
                {
                    Path = info.FullName,
                    Name = info.Name,
                    Kind = MediaFileTypeHelper.Classify(info.Extension),
                    SizeBytes = info.Length,
                    ModifiedAt = info.LastWriteTime
                });
                MediaFiles.Add(item);
            }
            SelectedMediaFile = item;
        }

        public void LoadExternalPlaylist(IEnumerable<string> filePaths)
        {
            StopPreview();
            MediaFiles.Clear();
            foreach (var path in (filePaths ?? Enumerable.Empty<string>()).Where(File.Exists).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var info = new FileInfo(path);
                var kind = MediaFileTypeHelper.Classify(info.Extension);
                if (kind == MediaKind.Other) continue;
                MediaFiles.Add(new MediaFileItem(new MediaFileDescriptor
                {
                    Path = info.FullName,
                    Name = info.Name,
                    Kind = kind,
                    SizeBytes = info.Length,
                    ModifiedAt = info.LastWriteTime
                }));
            }
            SelectedFolderPath = string.Empty;
            SelectedMediaFile = MediaFiles.FirstOrDefault();
            StatusMessage = MediaFiles.Count == 0 ? "没有可播放的媒体文件。" : string.Format("已加载临时播放队列 {0} 个文件。", MediaFiles.Count);
        }

        public void ExitImmersive()
        {
            StopImmersive();
            IsImmersive = false;
            ImmersiveKind = MediaKind.Other;
        }

        public void ExitMultimedia()
        {
            StopPreview();
            StopImmersive();
            IsImmersive = false;
            SelectedMediaFile = null;
            _owner.SwitchToHomeFromMultimedia();
        }

        public bool CopyCurrentViewedImageToClipboard()
        {
            var image = IsImmersive && ImmersiveKind == MediaKind.Image
                ? ImmersiveImageSource
                : PreviewKind == MediaKind.Image
                    ? PreviewImageSource
                    : null;
            if (image == null || SelectedMediaFile == null || SelectedMediaFile.Kind != MediaKind.Image)
            {
                StatusMessage = "当前没有可复制的图片。";
                return false;
            }

            try
            {
                ScreenshotService.SetClipboardCompatible(image);
                StatusMessage = "已复制图片到剪贴板：" + SelectedMediaFile.Name;
                return true;
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Multimedia image copy failed: {Msg}", ex.Message);
                StatusMessage = "复制图片失败：" + ex.Message;
                MessageBox.Show("复制图片失败：" + ex.Message, "图片查看", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private bool CanShowPreviousImage()
        {
            return GetAdjacentImage(-1) != null;
        }

        private bool CanShowNextImage()
        {
            return GetAdjacentImage(1) != null;
        }

        private void ShowPreviousImage()
        {
            ShowAdjacentImage(-1);
        }

        private void ShowNextImage()
        {
            ShowAdjacentImage(1);
        }

        private void ShowAdjacentImage(int direction)
        {
            var target = GetAdjacentImage(direction);
            if (target == null) return;

            var keepImmersiveImage = IsImmersive && ImmersiveKind == MediaKind.Image;
            SelectedMediaFile = target;
            if (keepImmersiveImage)
            {
                ImmersiveKind = MediaKind.Image;
                _ = LoadPreviewImageAsync(target.Path, true);
            }
        }

        private MediaFileItem GetAdjacentImage(int direction)
        {
            if (SelectedMediaFile == null || SelectedMediaFile.Kind != MediaKind.Image || direction == 0)
            {
                return null;
            }

            var currentIndex = MediaFiles.IndexOf(SelectedMediaFile);
            if (currentIndex < 0)
            {
                return null;
            }

            var nextIndex = currentIndex + Math.Sign(direction);
            while (nextIndex >= 0 && nextIndex < MediaFiles.Count)
            {
                var item = MediaFiles[nextIndex];
                if (item != null && item.Kind == MediaKind.Image && File.Exists(item.Path))
                {
                    return item;
                }

                nextIndex += Math.Sign(direction);
            }

            return null;
        }
        public void StopPreview()
        {
            IsPreviewPlaying = false;
            PreviewMediaSource = null;
            PreviewImageSource = null;
            PreviewImageInfo = string.Empty;
            PreviewKind = MediaKind.Other;
        }

        public void StopImmersive()
        {
            IsImmersivePlaying = false;
            ImmersiveMediaSource = null;
            ImmersiveImageSource = null;
        }

        private static IList<MediaFolderNode> LoadChildFolders(string folderPath)
        {
            var result = new List<MediaFolderNode>();
            try
            {
                foreach (var folder in Directory.EnumerateDirectories(folderPath))
                {
                    try
                    {
                        var info = new DirectoryInfo(folder);
                        if ((info.Attributes & (FileAttributes.Hidden | FileAttributes.System)) != 0) continue;
                        var node = new MediaFolderNode { Name = info.Name, FullPath = info.FullName };
                        if (HasVisibleChildFolders(info.FullName))
                            node.AddDummyChild();
                        result.Add(node);
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }
            return result.OrderBy(item => item.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
        }

        private static bool HasVisibleChildFolders(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath)) return false;
            try
            {
                foreach (var folder in Directory.EnumerateDirectories(folderPath))
                {
                    try
                    {
                        var info = new DirectoryInfo(folder);
                        if ((info.Attributes & (FileAttributes.Hidden | FileAttributes.System)) == 0)
                            return true;
                    }
                    catch
                    {
                    }
                }
            }
            catch
            {
            }

            return false;
        }

        private IEnumerable<MediaFileDescriptor> ApplyPreferredFilter(IEnumerable<MediaFileDescriptor> files)
        {
            switch (PreferredFilter)
            {
                case MultimediaPreferredFilter.Image:
                    return files.Where(item => item.Kind == MediaKind.Image);
                case MultimediaPreferredFilter.AudioVideo:
                    return files.Where(item => item.Kind == MediaKind.Audio || item.Kind == MediaKind.Video);
                default:
                    return files;
            }
        }

        private void ApplySelectedMediaFile(MediaFileItem item)
        {
            StopPreview();
            if (item == null || !File.Exists(item.Path))
            {
                StatusMessage = "选择媒体文件进行预览。";
                return;
            }

            PreviewKind = item.Kind;
            if (item.Kind == MediaKind.Image)
            {
                _ = LoadPreviewImageAsync(item.Path, false);
                StatusMessage = "已选择图片：" + item.Name;
            }
            else if (item.Kind == MediaKind.Audio || item.Kind == MediaKind.Video)
            {
                PreviewMediaSource = new Uri(item.Path);
                StatusMessage = "已选择媒体：" + item.Name;
            }
            else if (item.Kind == MediaKind.Pdf)
            {
                StatusMessage = "已选择 PDF：" + item.Name;
            }
        }

        private async Task LoadPreviewImageAsync(string filePath, bool immersive)
        {
            try
            {
                var image = await Task.Run(() => LoadBitmap(filePath));
                if (immersive)
                    ImmersiveImageSource = image;
                else
                {
                    PreviewImageSource = image;
                    PreviewImageInfo = image == null ? string.Empty : string.Format("{0} × {1} · {2}", image.PixelWidth, image.PixelHeight, SelectedMediaFile?.SizeText ?? string.Empty);
                }
            }
            catch
            {
                StatusMessage = "无法预览该文件";
            }
        }

        private static BitmapImage LoadBitmap(string filePath)
        {
            var bitmap = new BitmapImage();
            bitmap.BeginInit();
            bitmap.CacheOption = BitmapCacheOption.OnLoad;
            bitmap.UriSource = new Uri(filePath);
            bitmap.EndInit();
            bitmap.Freeze();
            return bitmap;
        }

        private bool CanEnterImmersive()
        {
            var item = SelectedMediaFile;
            return item != null && File.Exists(item.Path) && item.Kind != MediaKind.Pdf && item.Kind != MediaKind.Other;
        }

        private void EnterImmersive()
        {
            var item = SelectedMediaFile;
            if (item == null || !File.Exists(item.Path)) return;
            StopPreview();
            ImmersiveKind = item.Kind;
            IsImmersive = true;
            if (item.Kind == MediaKind.Image)
                _ = LoadPreviewImageAsync(item.Path, true);
            else if (item.Kind == MediaKind.Audio || item.Kind == MediaKind.Video)
                ImmersiveMediaSource = new Uri(item.Path);
            StatusMessage = "已进入沉浸模式，按 Esc 返回列表。";
        }

        private bool IsStandardConvertTarget(MediaFileItem item)
        {
            return item != null
                && File.Exists(item.Path)
                && (item.Kind == MediaKind.Image || item.Kind == MediaKind.Audio || item.Kind == MediaKind.Video);
        }

        private MediaFileItem ResolveContextMediaItem(object parameter)
        {
            return parameter as MediaFileItem ?? SelectedMediaFile;
        }

        private bool CanConvertImageToPdf(object parameter)
        {
            var target = ResolveContextMediaItem(parameter);
            return target != null && target.Kind == MediaKind.Image && File.Exists(target.Path);
        }

        private bool CanConvertPdfToImages(object parameter)
        {
            var target = ResolveContextMediaItem(parameter);
            return target != null && target.Kind == MediaKind.Pdf && File.Exists(target.Path);
        }

        private async Task ConvertImageToPdfAsync(object parameter)
        {
            var target = ResolveContextMediaItem(parameter);
            if (!CanConvertImageToPdf(target))
            {
                StatusMessage = "请选择要转换为 PDF 的图片。";
                return;
            }

            StatusMessage = "正在转换图片为 PDF：" + target.Name;
            var result = await PdfConvertService.ConvertImageToPdfAsync(target.Path, CancellationToken.None);
            await RefreshGeneratedPdfConvertOutputAsync(target, result);
            StatusMessage = result.Message;
            if (!result.Success)
            {
                MessageBox.Show(result.Message, "图片转 PDF", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task ConvertPdfToImagesAsync(object parameter)
        {
            var target = ResolveContextMediaItem(parameter);
            if (!CanConvertPdfToImages(target))
            {
                StatusMessage = "请选择要转换为图片的 PDF。";
                return;
            }

            StatusMessage = "正在转换 PDF 为图片：" + target.Name;
            var result = await PdfConvertService.ConvertPdfToImagesAsync(target.Path, CancellationToken.None);
            await RefreshGeneratedPdfConvertOutputAsync(target, result);
            StatusMessage = result.Message;
            if (!result.Success)
            {
                MessageBox.Show(result.Message, "PDF 转图片", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task RefreshGeneratedPdfConvertOutputAsync(MediaFileItem source, ConvertResult result)
        {
            if (source == null || result == null || !result.Success || string.IsNullOrWhiteSpace(SelectedFolderPath))
            {
                return;
            }

            var sourceFolder = Path.GetDirectoryName(source.Path) ?? string.Empty;
            if (string.Equals(sourceFolder, SelectedFolderPath, StringComparison.OrdinalIgnoreCase))
            {
                await SelectFolderAsync(SelectedFolderPath);
            }
        }

        private bool CanOpenConvertDialog(object parameter)
        {
            return IsStandardConvertTarget(ResolveContextMediaItem(parameter))
                || MediaFiles.Any(item => item.IsChecked && IsStandardConvertTarget(item));
        }

        private async void OpenConvertDialog(object parameter)
        {
            var targets = ResolveConvertTargets(parameter as MediaFileItem);
            if (targets.Count == 0)
            {
                StatusMessage = "没有可转换的媒体文件";
                return;
            }

            var dialog = new MediaConvertDialog(targets, _owner.IsFfmpegAvailable)
            {
                Owner = Application.Current?.MainWindow
            };
            if (dialog.ShowDialog() != true || dialog.Parameters == null) return;
            StatusMessage = "正在启动格式转换...";
            await _owner.ConvertMultimediaFilesAsync(targets, dialog.Parameters);
            StatusMessage = _owner.ConvertStatusMessage;
        }

        private IList<MediaFileItem> ResolveConvertTargets(MediaFileItem contextItem)
        {
            var checkedItems = MediaFiles.Where(item => item.IsChecked && IsStandardConvertTarget(item)).ToList();
            if (checkedItems.Count > 0) return checkedItems;
            if (IsStandardConvertTarget(contextItem)) return new List<MediaFileItem> { contextItem };
            if (IsStandardConvertTarget(SelectedMediaFile)) return new List<MediaFileItem> { SelectedMediaFile };
            return new List<MediaFileItem>();
        }

        private void OpenSelectedExternally()
        {
            var item = SelectedMediaFile;
            if (item == null || !File.Exists(item.Path)) return;
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(item.Path) { UseShellExecute = true });
            }
            catch (Exception ex)
            {
                StatusMessage = "外部打开失败：" + ex.Message;
            }
        }

        public void Dispose()
        {
            _loadCancellationTokenSource?.Cancel();
            _loadCancellationTokenSource?.Dispose();
            StopPreview();
            StopImmersive();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public sealed class MediaFolderNode : INotifyPropertyChanged
    {
        private bool _isExpanded;
        private bool _isSelected;

        public string Name { get; set; }
        public string FullPath { get; set; }
        public ObservableCollection<MediaFolderNode> Children { get; } = new ObservableCollection<MediaFolderNode>();
        public bool ChildrenLoaded { get; set; }
        public bool IsDummy => FullPath == null;

        public bool IsExpanded
        {
            get => _isExpanded;
            set { _isExpanded = value; OnPropertyChanged(); }
        }

        public bool IsSelected
        {
            get => _isSelected;
            set { _isSelected = value; OnPropertyChanged(); }
        }

        public void AddDummyChild()
        {
            if (Children.Count == 0)
                Children.Add(new MediaFolderNode { Name = "...", FullPath = null });
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public sealed class MediaFileItem : INotifyPropertyChanged
    {
        private bool _isChecked;

        public MediaFileItem(MediaFileDescriptor descriptor)
        {
            Path = descriptor.Path ?? string.Empty;
            Name = descriptor.Name ?? string.Empty;
            Kind = descriptor.Kind;
            SizeBytes = descriptor.SizeBytes;
            ModifiedAt = descriptor.ModifiedAt;
        }

        public string Path { get; }
        public string Name { get; }
        public MediaKind Kind { get; }
        public long SizeBytes { get; }
        public DateTime ModifiedAt { get; }
        public string KindText => Kind == MediaKind.Image ? "图片" : Kind == MediaKind.Audio ? "音频" : Kind == MediaKind.Video ? "视频" : Kind == MediaKind.Pdf ? "PDF" : "其他";
        public string SizeText => MediaFileTypeHelper.FormatFileSize(SizeBytes);
        public string ModifiedText => ModifiedAt.ToString("yyyy-MM-dd HH:mm");

        public bool IsChecked
        {
            get => _isChecked;
            set { _isChecked = value; OnPropertyChanged(); }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public sealed class MediaConvertParameters
    {
        public string ImageFormat { get; set; } = "jpg";
        public int ImageMaxWidth { get; set; }
        public int ImageMaxHeight { get; set; }
        public int ImageQuality { get; set; } = 85;
        public string MediaFormat { get; set; } = "mp4";
        public string MediaExtraArgs { get; set; } = string.Empty;
        public string OutputMode { get; set; } = "同目录";
        public string OutputFolder { get; set; } = string.Empty;
    }
}
