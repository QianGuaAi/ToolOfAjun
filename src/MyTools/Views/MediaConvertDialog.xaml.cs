using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using MahApps.Metro.Controls;
using MyTools.Services;
using MyTools.ViewModels;
using WinForms = System.Windows.Forms;

namespace MyTools.Views
{
    public partial class MediaConvertDialog : MetroWindow
    {
        private readonly IList<MediaFileItem> _targets;
        private readonly bool _isFfmpegAvailable;
        private readonly bool _hasImages;
        private readonly bool _hasMedia;

        public MediaConvertDialog(IList<MediaFileItem> targets, bool isFfmpegAvailable)
        {
            InitializeComponent();
            _targets = targets ?? new List<MediaFileItem>();
            _isFfmpegAvailable = isFfmpegAvailable;
            _hasImages = _targets.Any(item => item.Kind == MediaKind.Image);
            _hasMedia = _targets.Any(item => item.Kind == MediaKind.Audio || item.Kind == MediaKind.Video);
            Parameters = new MediaConvertParameters();
            InitializeState();
        }

        public MediaConvertParameters Parameters { get; private set; }

        private void InitializeState()
        {
            TargetCountText.Text = "转换目标：" + _targets.Count + " 个";
            ImageSection.Visibility = _hasImages ? Visibility.Visible : Visibility.Collapsed;
            MediaSection.Visibility = _hasMedia ? Visibility.Visible : Visibility.Collapsed;
            FfmpegWarningText.Visibility = _hasMedia && !_isFfmpegAvailable ? Visibility.Visible : Visibility.Collapsed;
            OkButton.IsEnabled = !_hasMedia || _isFfmpegAvailable;
        }

        private void BrowseOutputFolder_OnClick(object sender, RoutedEventArgs e)
        {
            using (var dialog = new WinForms.FolderBrowserDialog())
            {
                dialog.Description = "选择转换输出目录";
                dialog.ShowNewFolderButton = true;
                if (!string.IsNullOrWhiteSpace(OutputFolderBox.Text) && Directory.Exists(OutputFolderBox.Text))
                    dialog.SelectedPath = OutputFolderBox.Text;
                if (dialog.ShowDialog() == WinForms.DialogResult.OK)
                {
                    OutputFolderBox.Text = dialog.SelectedPath;
                    CustomFolderRadio.IsChecked = true;
                }
            }
        }

        private void Cancel_OnClick(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void Ok_OnClick(object sender, RoutedEventArgs e)
        {
            if (!TryBuildParameters(out var parameters, out var error))
            {
                ValidationText.Text = error;
                ValidationText.Visibility = Visibility.Visible;
                return;
            }

            Parameters = parameters;
            DialogResult = true;
            Close();
        }

        private bool TryBuildParameters(out MediaConvertParameters parameters, out string error)
        {
            parameters = new MediaConvertParameters();
            error = string.Empty;

            if (_hasMedia && !_isFfmpegAvailable)
            {
                error = "未检测到 ffmpeg.exe，音视频转换不可用。";
                return false;
            }

            if (!TryReadNonNegativeInt(ImageMaxWidthBox.Text, out var width))
            {
                error = "最大宽度必须是 0 或正整数。";
                return false;
            }

            if (!TryReadNonNegativeInt(ImageMaxHeightBox.Text, out var height))
            {
                error = "最大高度必须是 0 或正整数。";
                return false;
            }

            var outputMode = CustomFolderRadio.IsChecked == true ? "指定目录" : "同目录";
            var outputFolder = (OutputFolderBox.Text ?? string.Empty).Trim();
            if (string.Equals(outputMode, "指定目录", StringComparison.Ordinal) && !Directory.Exists(outputFolder))
            {
                error = "指定输出目录不存在。";
                return false;
            }

            parameters.ImageFormat = GetComboText(ImageFormatBox, "jpg");
            parameters.ImageMaxWidth = width;
            parameters.ImageMaxHeight = height;
            parameters.ImageQuality = Math.Max(1, Math.Min(100, (int)ImageQualitySlider.Value));
            parameters.MediaFormat = GetComboText(MediaFormatBox, "mp4");
            parameters.MediaExtraArgs = MediaExtraArgsBox.Text ?? string.Empty;
            parameters.OutputMode = outputMode;
            parameters.OutputFolder = outputFolder;
            return true;
        }

        private static bool TryReadNonNegativeInt(string text, out int value)
        {
            if (!int.TryParse((text ?? string.Empty).Trim(), out value)) return false;
            return value >= 0;
        }

        private static string GetComboText(ComboBox comboBox, string fallback)
        {
            if (comboBox?.SelectedItem is ComboBoxItem item && item.Content != null)
                return item.Content.ToString();
            return fallback;
        }
    }
}
