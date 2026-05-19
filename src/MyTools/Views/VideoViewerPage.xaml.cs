using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using MyTools.Services;
using MyTools.ViewModels;

namespace MyTools.Views
{
    public partial class VideoViewerPage : UserControl
    {
        private readonly DispatcherTimer _videoProgressTimer;
        private bool _isDraggingVideoProgress;
        private bool _isDraggingVideoWaveform;
        private bool _isSelectingVideoWaveformRange;
        private bool _isDraggingVideoWaveformRangeHandle;
        private bool _isDraggingVideoWaveformStartHandle;
        private double _videoWaveformRangeStartSeconds;
        private Point _videoPlaylistDragStartPoint;
        private VideoPlaylistItem _videoPlaylistDragItem;

        public VideoViewerPage()
        {
            InitializeComponent();
            _videoProgressTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(350)
            };
            _videoProgressTimer.Tick += VideoProgressTimer_OnTick;
        }

        public bool HandleShortcut(KeyEventArgs e)
        {
            if (!(DataContext is MainViewModel vm)
                || vm.CurrentModule != "VideoViewer"
                || !vm.HasVideoViewerVideo
                || Keyboard.Modifiers != ModifierKeys.None)
            {
                return false;
            }

            switch (e.Key)
            {
                case Key.Space:
                    TogglePlayPause();
                    e.Handled = true;
                    return true;
                case Key.Left:
                    var previousSeconds = Math.Max(0, vm.VideoViewerPositionSeconds - 5);
                    SeekVideoViewer(previousSeconds);
                    vm.VideoViewerPositionSeconds = previousSeconds;
                    vm.UpdateVideoViewerSubtitle(previousSeconds);
                    e.Handled = true;
                    return true;
                case Key.Right:
                    var nextSeconds = vm.VideoViewerDurationSeconds > 0
                        ? Math.Min(vm.VideoViewerDurationSeconds, vm.VideoViewerPositionSeconds + 5)
                        : vm.VideoViewerPositionSeconds + 5;
                    SeekVideoViewer(nextSeconds);
                    vm.VideoViewerPositionSeconds = nextSeconds;
                    vm.UpdateVideoViewerSubtitle(nextSeconds);
                    e.Handled = true;
                    return true;
                case Key.Up:
                    vm.VideoViewerVolume = Math.Min(1.0, vm.VideoViewerVolume + 0.05);
                    e.Handled = true;
                    return true;
                case Key.Down:
                    vm.VideoViewerVolume = Math.Max(0.0, vm.VideoViewerVolume - 0.05);
                    e.Handled = true;
                    return true;
                default:
                    return false;
            }
        }

        public void ResetForSourceChange()
        {
            if (!(DataContext is MainViewModel vm))
            {
                return;
            }

            try
            {
                VideoPlayer.Stop();
            }
            catch
            {
            }

            _videoProgressTimer.Stop();
            vm.IsVideoViewerPlaying = false;
            vm.VideoViewerPositionSeconds = 0;
            vm.UpdateVideoViewerSubtitle(0);
            ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
        }

        public void StopForDroppedMedia()
        {
            if (!(DataContext is MainViewModel vm))
            {
                return;
            }

            try
            {
                VideoPlayer.Stop();
            }
            catch
            {
            }

            vm.IsVideoViewerPlaying = false;
        }

        public void StopVideoViewer()
        {
            _videoProgressTimer.Stop();
            try
            {
                VideoPlayer.Stop();
            }
            catch
            {
            }
        }

        public void PauseVideoViewer()
        {
            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo)
            {
                return;
            }

            try
            {
                VideoPlayer.Pause();
            }
            catch
            {
            }

            _videoProgressTimer.Stop();
            vm.IsVideoViewerPlaying = false;
        }

        public void ApplyVideoSpeed(double speedRatio)
        {
            if (double.IsNaN(speedRatio) || double.IsInfinity(speedRatio) || speedRatio <= 0)
            {
                speedRatio = 1.0;
            }

            try
            {
                VideoPlayer.SpeedRatio = speedRatio;
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Video speed change failed: {Msg}", ex.Message);
            }
        }

        private void VideoPlayer_MediaOpened(object sender, RoutedEventArgs e)
        {
            if (!(DataContext is MainViewModel vm))
            {
                return;
            }

            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                vm.VideoViewerDurationSeconds = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
            }

            vm.VideoViewerPositionSeconds = 0;
            vm.UpdateVideoViewerSubtitle(0);
            ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
            vm.VideoViewerStatusMessage = "媒体已就绪。";
        }

        private void VideoPlayer_MediaEnded(object sender, RoutedEventArgs e)
        {
            if (!(DataContext is MainViewModel vm))
            {
                return;
            }

            var endSeconds = vm.VideoViewerDurationSeconds > 0
                ? vm.VideoViewerDurationSeconds
                : VideoPlayer.Position.TotalSeconds;
            if (vm.ShouldLoopVideoViewerAt(endSeconds, out var loopTargetSeconds))
            {
                SeekVideoViewer(loopTargetSeconds);
                ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
                VideoPlayer.Play();
                _videoProgressTimer.Start();
                vm.IsVideoViewerPlaying = true;
                vm.VideoViewerPositionSeconds = loopTargetSeconds;
                vm.UpdateVideoViewerSubtitle(loopTargetSeconds);
                vm.VideoViewerStatusMessage = "正在 A/B 循环。";
                return;
            }

            VideoPlayer.Stop();
            _videoProgressTimer.Stop();
            vm.IsVideoViewerPlaying = false;
            vm.VideoViewerPositionSeconds = 0;
            vm.UpdateVideoViewerSubtitle(0);
            if (vm.OpenNextVideoViewerPlaylistItem())
            {
                ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
                VideoPlayer.Play();
                _videoProgressTimer.Start();
                vm.IsVideoViewerPlaying = true;
                vm.VideoViewerStatusMessage = vm.VideoViewerAutoAdvanceText;
            }
            else
            {
                vm.VideoViewerStatusMessage = "播放结束。";
            }
        }

        private void VideoPlayer_MediaFailed(object sender, ExceptionRoutedEventArgs e)
        {
            if (!(DataContext is MainViewModel vm))
            {
                return;
            }

            var message = e.ErrorException?.Message ?? "系统缺少对应解码器";
            _videoProgressTimer.Stop();
            vm.IsVideoViewerPlaying = false;
            vm.VideoViewerStatusMessage = "播放失败：" + message;
            MessageBox.Show(
                "播放失败：" + message + "\n\n可点击“外部打开”用系统默认播放器尝试播放。",
                "音视频播放",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
        }

        private void VideoPlayPauseButton_OnClick(object sender, RoutedEventArgs e)
        {
            TogglePlayPause();
        }

        private void TogglePlayPause()
        {
            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo)
            {
                return;
            }

            if (vm.IsVideoViewerPlaying)
            {
                PauseVideoViewer();
                vm.VideoViewerStatusMessage = "已暂停。";
            }
            else
            {
                ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
                VideoPlayer.Play();
                _videoProgressTimer.Start();
                vm.IsVideoViewerPlaying = true;
                vm.VideoViewerStatusMessage = "正在播放。";
            }
        }

        private void VideoStopButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo)
            {
                return;
            }

            VideoPlayer.Stop();
            _videoProgressTimer.Stop();
            vm.IsVideoViewerPlaying = false;
            vm.VideoViewerPositionSeconds = 0;
            vm.UpdateVideoViewerSubtitle(0);
            vm.VideoViewerStatusMessage = "已停止。";
        }

        private void VideoProgressSlider_OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _isDraggingVideoProgress = true;
        }

        private void VideoProgressSlider_OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!(DataContext is MainViewModel vm))
            {
                _isDraggingVideoProgress = false;
                return;
            }

            _isDraggingVideoProgress = false;
            SeekVideoViewer(VideoProgressSlider.Value);
            vm.VideoViewerPositionSeconds = VideoProgressSlider.Value;
            vm.UpdateVideoViewerSubtitle(VideoProgressSlider.Value);
        }

        private void VideoWaveform_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!(sender is UIElement element))
            {
                return;
            }

            element.CaptureMouse();
            if ((Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift)
            {
                BeginVideoWaveformRangeSelection(sender, e.GetPosition(element));
            }
            else
            {
                _isDraggingVideoWaveform = true;
                SeekVideoViewerByWaveformPoint(sender, e.GetPosition(element), "正在按波形定位：");
            }

            e.Handled = true;
        }

        private void VideoWaveform_OnMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (!(sender is UIElement element))
            {
                return;
            }

            element.CaptureMouse();
            BeginVideoWaveformRangeSelection(sender, e.GetPosition(element));
            e.Handled = true;
        }

        private void VideoWaveform_OnMouseMove(object sender, MouseEventArgs e)
        {
            if (_isSelectingVideoWaveformRange)
            {
                if (e.LeftButton != MouseButtonState.Pressed && e.RightButton != MouseButtonState.Pressed)
                {
                    EndVideoWaveformInteraction(sender);
                    return;
                }

                if (sender is UIElement rangeElement)
                {
                    UpdateVideoWaveformRangeSelection(sender, e.GetPosition(rangeElement), false);
                }

                e.Handled = true;
                return;
            }

            if (!_isDraggingVideoWaveform)
            {
                return;
            }

            if (e.LeftButton != MouseButtonState.Pressed)
            {
                EndVideoWaveformInteraction(sender);
                return;
            }

            if (!(sender is UIElement element))
            {
                return;
            }

            SeekVideoViewerByWaveformPoint(sender, e.GetPosition(element), "正在按波形定位：");
            e.Handled = true;
        }

        private void VideoWaveform_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isSelectingVideoWaveformRange)
            {
                if (sender is UIElement rangeElement)
                {
                    UpdateVideoWaveformRangeSelection(sender, e.GetPosition(rangeElement), true);
                }

                EndVideoWaveformInteraction(sender);
                e.Handled = true;
                return;
            }

            if (sender is UIElement element)
            {
                SeekVideoViewerByWaveformPoint(sender, e.GetPosition(element), "已按波形跳转：");
            }

            EndVideoWaveformInteraction(sender);
            e.Handled = true;
        }

        private void VideoWaveform_OnMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isSelectingVideoWaveformRange)
            {
                if (sender is UIElement element)
                {
                    UpdateVideoWaveformRangeSelection(sender, e.GetPosition(element), true);
                }

                EndVideoWaveformInteraction(sender);
            }

            e.Handled = true;
        }

        private void VideoWaveformStartHandle_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            BeginVideoWaveformHandleDrag(sender, true);
            e.Handled = true;
        }

        private void VideoWaveformEndHandle_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            BeginVideoWaveformHandleDrag(sender, false);
            e.Handled = true;
        }

        private void VideoWaveformHandle_OnMouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDraggingVideoWaveformRangeHandle)
            {
                return;
            }

            if (e.LeftButton != MouseButtonState.Pressed)
            {
                EndVideoWaveformInteraction(sender);
                return;
            }

            UpdateVideoWaveformHandleDrag(e.GetPosition(VideoWaveformImage), false);
            e.Handled = true;
        }

        private void VideoWaveformHandle_OnMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_isDraggingVideoWaveformRangeHandle)
            {
                UpdateVideoWaveformHandleDrag(e.GetPosition(VideoWaveformImage), true);
            }

            EndVideoWaveformInteraction(sender);
            e.Handled = true;
        }

        private void SeekVideoViewerByWaveformPoint(object sender, Point point, string statusPrefix)
        {
            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo || vm.VideoViewerDurationSeconds <= 0)
            {
                return;
            }

            if (!(sender is FrameworkElement element) || element.ActualWidth <= 1)
            {
                return;
            }

            var ratio = Math.Max(0, Math.Min(1, point.X / element.ActualWidth));
            var seconds = ratio * vm.VideoViewerDurationSeconds;
            SeekVideoViewer(seconds);
            vm.VideoViewerPositionSeconds = seconds;
            vm.UpdateVideoViewerSubtitle(seconds);
            vm.VideoViewerStatusMessage = (statusPrefix ?? "已按波形跳转：") + vm.VideoViewerPositionText;
        }

        private void BeginVideoWaveformRangeSelection(object sender, Point point)
        {
            _isDraggingVideoWaveform = false;
            if (!TryGetVideoWaveformSeconds(sender, point, out var seconds))
            {
                _isSelectingVideoWaveformRange = false;
                return;
            }

            _isSelectingVideoWaveformRange = true;
            _videoWaveformRangeStartSeconds = seconds;
            if (DataContext is MainViewModel vm)
            {
                var previewEndSeconds = Math.Min(vm.VideoViewerDurationSeconds, seconds + 0.25);
                if (previewEndSeconds <= seconds + 0.2)
                {
                    previewEndSeconds = Math.Max(0, seconds - 0.25);
                }

                vm.SetVideoViewerLoopRange(seconds, previewEndSeconds, false);
                vm.VideoViewerStatusMessage = "正在选择 A/B 区间：" + vm.VideoViewerLoopRangeText;
            }
        }

        private void UpdateVideoWaveformRangeSelection(object sender, Point point, bool enableLoop)
        {
            if (!_isSelectingVideoWaveformRange || !TryGetVideoWaveformSeconds(sender, point, out var seconds))
            {
                return;
            }

            if (DataContext is MainViewModel vm)
            {
                vm.SetVideoViewerLoopRange(_videoWaveformRangeStartSeconds, seconds, enableLoop);
                if (!enableLoop && vm.HasVideoViewerLoopRange)
                {
                    vm.VideoViewerStatusMessage = "正在选择 A/B 区间：" + vm.VideoViewerLoopRangeText;
                }
            }
        }

        private void BeginVideoWaveformHandleDrag(object sender, bool isStartHandle)
        {
            _isDraggingVideoWaveform = false;
            _isSelectingVideoWaveformRange = false;
            _isDraggingVideoWaveformRangeHandle = true;
            _isDraggingVideoWaveformStartHandle = isStartHandle;
            if (sender is UIElement element)
            {
                element.CaptureMouse();
            }

            if (DataContext is MainViewModel vm)
            {
                vm.VideoViewerStatusMessage = isStartHandle
                    ? "正在微调 A 点：" + vm.VideoViewerLoopStartText
                    : "正在微调 B 点：" + vm.VideoViewerLoopEndText;
            }
        }

        private void UpdateVideoWaveformHandleDrag(Point point, bool finalize)
        {
            if (!_isDraggingVideoWaveformRangeHandle || !TryGetVideoWaveformSeconds(VideoWaveformImage, point, out var seconds))
            {
                return;
            }

            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerLoopRange)
            {
                return;
            }

            seconds = _isDraggingVideoWaveformStartHandle
                ? Math.Min(seconds, Math.Max(0, vm.VideoViewerLoopEndSeconds - 0.25))
                : Math.Max(seconds, vm.VideoViewerLoopStartSeconds + 0.25);
            var startSeconds = _isDraggingVideoWaveformStartHandle ? seconds : vm.VideoViewerLoopStartSeconds;
            var endSeconds = _isDraggingVideoWaveformStartHandle ? vm.VideoViewerLoopEndSeconds : seconds;
            if (vm.SetVideoViewerLoopRange(startSeconds, endSeconds, vm.VideoViewerIsLoopEnabled))
            {
                vm.VideoViewerStatusMessage = finalize
                    ? "已微调 A/B 循环：" + vm.VideoViewerLoopRangeText
                    : (_isDraggingVideoWaveformStartHandle
                        ? "正在微调 A 点：" + vm.VideoViewerLoopStartText
                        : "正在微调 B 点：" + vm.VideoViewerLoopEndText);
            }
        }

        private bool TryGetVideoWaveformSeconds(object sender, Point point, out double seconds)
        {
            seconds = 0;
            if (!(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo || vm.VideoViewerDurationSeconds <= 0)
            {
                return false;
            }

            if (!(sender is FrameworkElement element) || element.ActualWidth <= 1)
            {
                return false;
            }

            var ratio = Math.Max(0, Math.Min(1, point.X / element.ActualWidth));
            seconds = ratio * vm.VideoViewerDurationSeconds;
            return true;
        }

        private void EndVideoWaveformInteraction(object sender)
        {
            _isDraggingVideoWaveform = false;
            _isSelectingVideoWaveformRange = false;
            _isDraggingVideoWaveformRangeHandle = false;
            if (sender is UIElement element)
            {
                element.ReleaseMouseCapture();
            }
        }

        private void VideoProgressTimer_OnTick(object sender, EventArgs e)
        {
            if (_isDraggingVideoProgress || _isDraggingVideoWaveform || _isSelectingVideoWaveformRange || _isDraggingVideoWaveformRangeHandle || !(DataContext is MainViewModel vm) || !vm.HasVideoViewerVideo)
            {
                return;
            }

            var positionSeconds = VideoPlayer.Position.TotalSeconds;
            vm.VideoViewerPositionSeconds = positionSeconds;
            vm.UpdateVideoViewerSubtitle(positionSeconds);
            if (VideoPlayer.NaturalDuration.HasTimeSpan)
            {
                vm.VideoViewerDurationSeconds = VideoPlayer.NaturalDuration.TimeSpan.TotalSeconds;
            }

            if (vm.ShouldLoopVideoViewerAt(positionSeconds, out var loopTargetSeconds))
            {
                SeekVideoViewer(loopTargetSeconds);
                vm.VideoViewerPositionSeconds = loopTargetSeconds;
                vm.UpdateVideoViewerSubtitle(loopTargetSeconds);
            }
        }

        private void VideoSpeedComboBox_OnSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (DataContext is MainViewModel vm)
            {
                ApplyVideoSpeed(vm.VideoViewerSpeedRatio);
            }
        }

        private void SeekVideoViewer(double seconds)
        {
            if (double.IsNaN(seconds) || double.IsInfinity(seconds) || seconds < 0)
            {
                return;
            }

            try
            {
                VideoPlayer.Position = TimeSpan.FromSeconds(seconds);
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Video seek failed: {Msg}", ex.Message);
            }
        }

        private void VideoPlaylistListBox_OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _videoPlaylistDragStartPoint = e.GetPosition(null);
            _videoPlaylistDragItem = FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject)?.DataContext as VideoPlaylistItem;
        }

        private void VideoPlaylistListBox_OnMouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed || _videoPlaylistDragItem == null)
            {
                return;
            }

            var currentPosition = e.GetPosition(null);
            if (Math.Abs(currentPosition.X - _videoPlaylistDragStartPoint.X) < SystemParameters.MinimumHorizontalDragDistance
                && Math.Abs(currentPosition.Y - _videoPlaylistDragStartPoint.Y) < SystemParameters.MinimumVerticalDragDistance)
            {
                return;
            }

            DragDrop.DoDragDrop(VideoPlaylistListBox, _videoPlaylistDragItem, DragDropEffects.Move);
            _videoPlaylistDragItem = null;
        }

        private void VideoPlaylistListBox_OnDrop(object sender, DragEventArgs e)
        {
            if (!(DataContext is MainViewModel vm))
            {
                return;
            }

            var item = e.Data.GetData(typeof(VideoPlaylistItem)) as VideoPlaylistItem;
            var target = FindAncestor<ListBoxItem>(e.OriginalSource as DependencyObject)?.DataContext as VideoPlaylistItem;
            if (item == null || target == null)
            {
                return;
            }

            vm.MoveVideoViewerPlaylistItem(item, target);
            e.Handled = true;
        }

        private void VideoPlaylistItem_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is MainViewModel vm && sender is ListBoxItem item)
            {
                vm.OpenVideoViewerPlaylistItemCommand.Execute(item.DataContext);
                e.Handled = true;
            }
        }

        private void VideoLoopRangeItem_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is MainViewModel vm && sender is ListBoxItem item && item.DataContext is VideoLoopRangeItem range)
            {
                SeekVideoViewer(range.StartSeconds);
                vm.OpenVideoViewerLoopRangeCommand.Execute(range);
                e.Handled = true;
            }
        }

        private void VideoLoopRangeOpenButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is VideoLoopRangeItem range)
            {
                SeekVideoViewer(range.StartSeconds);
            }
        }

        private void VideoBookmarkItem_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (DataContext is MainViewModel vm && sender is ListBoxItem item && item.DataContext is VideoBookmarkItem bookmark)
            {
                SeekVideoViewer(bookmark.PositionSeconds);
                vm.OpenVideoViewerBookmarkCommand.Execute(bookmark);
                e.Handled = true;
            }
        }

        private void VideoBookmarkJumpButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.DataContext is VideoBookmarkItem bookmark)
            {
                SeekVideoViewer(bookmark.PositionSeconds);
            }
        }

        private static T FindAncestor<T>(DependencyObject current) where T : DependencyObject
        {
            while (current != null)
            {
                if (current is T match)
                {
                    return match;
                }

                current = System.Windows.Media.VisualTreeHelper.GetParent(current);
            }

            return null;
        }
    }
}
