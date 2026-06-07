using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using MyTools.ViewModels;

namespace MyTools.Views
{
    public partial class MultimediaPage : UserControl
    {
        private readonly DispatcherTimer _positionTimer;
        private readonly HashSet<Slider> _draggingPositionSliders = new HashSet<Slider>();

        public MultimediaPage()
        {
            InitializeComponent();
            _positionTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            _positionTimer.Tick += PositionTimer_OnTick;
        }

        private MultimediaViewModel ViewModel => DataContext as MultimediaViewModel;

        private void MultimediaPage_OnLoaded(object sender, RoutedEventArgs e)
        {
            Focus();
            ApplyAllVolumeSliders();
            _positionTimer.Start();
        }

        private void MultimediaPage_OnUnloaded(object sender, RoutedEventArgs e)
        {
            StopAllMediaElements();
            _positionTimer.Stop();
        }

        private async void FolderTreeItem_OnExpanded(object sender, RoutedEventArgs e)
        {
            if (ViewModel != null && sender is TreeViewItem item && item.DataContext is MediaFolderNode node)
                await ViewModel.ExpandFolderAsync(node);
        }

        private async void FolderTree_OnSelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (ViewModel != null && e.NewValue is MediaFolderNode node && !node.IsDummy)
                await ViewModel.SelectFolderAsync(node.FullPath);
        }

        private async void FolderNodeHeader_OnMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var item = FindVisualParent<TreeViewItem>(sender as DependencyObject);
            if (item == null || !(item.DataContext is MediaFolderNode node) || node.IsDummy)
                return;

            item.IsSelected = true;
            if (item.HasItems || node.Children.Count > 0)
            {
                item.IsExpanded = !item.IsExpanded;
                if (item.IsExpanded && ViewModel != null)
                    await ViewModel.ExpandFolderAsync(node);
            }

            e.Handled = true;
        }

        private static T FindVisualParent<T>(DependencyObject source) where T : DependencyObject
        {
            while (source != null)
            {
                source = VisualTreeHelper.GetParent(source);
                if (source is T target)
                    return target;
            }

            return null;
        }

        private void MediaFileList_OnPreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            var item = FindVisualParent<ListViewItem>(e.OriginalSource as DependencyObject);
            if (item != null)
            {
                item.IsSelected = true;
                item.Focus();
            }
        }

        private void MediaFileList_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var vm = ViewModel;
            if (vm?.EnterImmersiveCommand.CanExecute(null) == true)
                vm.EnterImmersiveCommand.Execute(null);
        }

        private void ViewModeButton_OnClick(object sender, RoutedEventArgs e)
        {
            if (sender is Button button && button.ContextMenu != null)
            {
                button.ContextMenu.PlacementTarget = button;
                button.ContextMenu.IsOpen = true;
            }
        }

        private void MultimediaPage_OnPreviewKeyDown(object sender, KeyEventArgs e)
        {
            var vm = ViewModel;
            if (vm == null)
            {
                return;
            }

            if (e.Key == Key.C && (Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                e.Handled = vm.CopyCurrentViewedImageToClipboard();
                return;
            }

            if (vm.IsImmersive && e.Key == Key.Escape)
            {
                StopImmersiveElements();
                vm.ExitImmersive();
                e.Handled = true;
            }
        }

        private void PreviewPlayPause_OnClick(object sender, RoutedEventArgs e)
        {
            ToggleElement(PreviewMediaElement, true);
        }

        private void PreviewAudioPlayPause_OnClick(object sender, RoutedEventArgs e)
        {
            ToggleElement(PreviewAudioElement, true);
        }

        private void PreviewMute_OnClick(object sender, RoutedEventArgs e)
        {
            ToggleMute(PreviewMediaElement);
        }

        private void PreviewAudioMute_OnClick(object sender, RoutedEventArgs e)
        {
            ToggleMute(PreviewAudioElement);
        }

        private void ImmersivePlayPause_OnClick(object sender, RoutedEventArgs e)
        {
            ToggleElement(ImmersiveMediaElement, false);
        }

        private void ImmersiveAudioPlayPause_OnClick(object sender, RoutedEventArgs e)
        {
            ToggleElement(ImmersiveAudioElement, false);
        }

        private void ImmersiveMute_OnClick(object sender, RoutedEventArgs e)
        {
            ToggleMute(ImmersiveMediaElement);
        }

        private void ImmersiveAudioMute_OnClick(object sender, RoutedEventArgs e)
        {
            ToggleMute(ImmersiveAudioElement);
        }

        private void MediaElement_OnMediaOpened(object sender, RoutedEventArgs e)
        {
            if (sender is MediaElement element)
            {
                ApplyVolumeForElement(element);
                var slider = ResolveSliderForElement(element);
                UpdateSlider(element, slider);
            }
        }

        private void PreviewMediaElement_OnMediaEnded(object sender, RoutedEventArgs e)
        {
            if (sender is MediaElement element)
            {
                element.Stop();
                ResetSlider(ResolveSliderForElement(element));
                if (ViewModel != null) ViewModel.IsPreviewPlaying = false;
            }
        }

        private void ImmersiveMediaElement_OnMediaEnded(object sender, RoutedEventArgs e)
        {
            if (sender is MediaElement element)
            {
                element.Stop();
                ResetSlider(ResolveSliderForElement(element));
                if (ViewModel != null) ViewModel.IsImmersivePlaying = false;
            }
        }

        private void ToggleElement(MediaElement element, bool preview)
        {
            if (element == null || element.Source == null) return;
            ApplyVolumeForElement(element);
            var vm = ViewModel;
            if (preview)
            {
                if (vm != null && vm.IsPreviewPlaying)
                {
                    element.Pause();
                    vm.IsPreviewPlaying = false;
                }
                else
                {
                    StopOtherPreviewElement(element);
                    element.Play();
                    if (vm != null) vm.IsPreviewPlaying = true;
                }
            }
            else
            {
                if (vm != null && vm.IsImmersivePlaying)
                {
                    element.Pause();
                    vm.IsImmersivePlaying = false;
                }
                else
                {
                    StopOtherImmersiveElement(element);
                    element.Play();
                    if (vm != null) vm.IsImmersivePlaying = true;
                }
            }
        }

        private static void ToggleMute(MediaElement element)
        {
            if (element != null) element.IsMuted = !element.IsMuted;
        }

        private void PositionSlider_OnPreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (sender is Slider slider)
            {
                _draggingPositionSliders.Add(slider);
            }
        }

        private void PositionSlider_OnPreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (sender is Slider slider)
            {
                CommitPositionSliderDeferred(slider);
            }
        }

        private void PositionSlider_OnLostMouseCapture(object sender, MouseEventArgs e)
        {
            if (sender is Slider slider && _draggingPositionSliders.Contains(slider))
            {
                CommitPositionSliderDeferred(slider);
            }
        }

        private void PositionSlider_OnKeyUp(object sender, KeyEventArgs e)
        {
            if (sender is Slider slider && IsPositionCommitKey(e.Key))
            {
                CommitPositionSlider(slider, true);
                e.Handled = true;
            }
        }

        private void CommitPositionSliderDeferred(Slider slider)
        {
            Dispatcher.BeginInvoke(new Action(() => CommitPositionSlider(slider, true)), DispatcherPriority.Background);
        }

        private static bool IsPositionCommitKey(Key key)
        {
            return key == Key.Left
                || key == Key.Right
                || key == Key.Up
                || key == Key.Down
                || key == Key.Home
                || key == Key.End
                || key == Key.PageUp
                || key == Key.PageDown;
        }

        private void CommitPositionSlider(Slider slider, bool resumePlayback)
        {
            _draggingPositionSliders.Remove(slider);
            var element = ResolveElementForPositionSlider(slider);
            if (element == null || element.Source == null || !element.NaturalDuration.HasTimeSpan)
            {
                return;
            }

            var total = element.NaturalDuration.TimeSpan.TotalSeconds;
            if (total <= 0)
            {
                return;
            }

            var targetSeconds = Math.Min(total, Math.Max(0, slider.Value));
            element.Position = TimeSpan.FromSeconds(targetSeconds);
            if (resumePlayback)
            {
                if (ReferenceEquals(element, PreviewMediaElement) || ReferenceEquals(element, PreviewAudioElement))
                {
                    StopOtherPreviewElement(element);
                    element.Play();
                    if (ViewModel != null) ViewModel.IsPreviewPlaying = true;
                }
                else
                {
                    StopOtherImmersiveElement(element);
                    element.Play();
                    if (ViewModel != null) ViewModel.IsImmersivePlaying = true;
                }
            }
        }

        private MediaElement ResolveElementForPositionSlider(Slider slider)
        {
            if (ReferenceEquals(slider, PreviewPositionSlider)) return PreviewMediaElement;
            if (ReferenceEquals(slider, PreviewAudioPositionSlider)) return PreviewAudioElement;
            if (ReferenceEquals(slider, ImmersivePositionSlider)) return ImmersiveMediaElement;
            if (ReferenceEquals(slider, ImmersiveAudioPositionSlider)) return ImmersiveAudioElement;
            return null;
        }

        private Slider ResolveSliderForElement(MediaElement element)
        {
            if (ReferenceEquals(element, PreviewMediaElement)) return PreviewPositionSlider;
            if (ReferenceEquals(element, PreviewAudioElement)) return PreviewAudioPositionSlider;
            if (ReferenceEquals(element, ImmersiveMediaElement)) return ImmersivePositionSlider;
            if (ReferenceEquals(element, ImmersiveAudioElement)) return ImmersiveAudioPositionSlider;
            return null;
        }

        private static void ResetSlider(Slider slider)
        {
            if (slider == null) return;
            slider.Maximum = 1;
            slider.Value = 0;
        }

        private void VolumeSlider_OnValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (sender is Slider slider)
            {
                ApplyVolumeForSlider(slider, true);
            }
        }

        private void ApplyAllVolumeSliders()
        {
            ApplyVolumeForSlider(PreviewVolumeSlider, false);
            ApplyVolumeForSlider(PreviewAudioVolumeSlider, false);
            ApplyVolumeForSlider(ImmersiveVolumeSlider, false);
            ApplyVolumeForSlider(ImmersiveAudioVolumeSlider, false);
        }

        private void ApplyVolumeForElement(MediaElement element)
        {
            if (ReferenceEquals(element, PreviewMediaElement)) ApplyVolumeForSlider(PreviewVolumeSlider, false);
            else if (ReferenceEquals(element, PreviewAudioElement)) ApplyVolumeForSlider(PreviewAudioVolumeSlider, false);
            else if (ReferenceEquals(element, ImmersiveMediaElement)) ApplyVolumeForSlider(ImmersiveVolumeSlider, false);
            else if (ReferenceEquals(element, ImmersiveAudioElement)) ApplyVolumeForSlider(ImmersiveAudioVolumeSlider, false);
        }

        private void ApplyVolumeForSlider(Slider slider, bool updateMuteState)
        {
            var element = ResolveElementForVolumeSlider(slider);
            if (element == null || slider == null)
            {
                return;
            }

            var volume = Math.Min(1, Math.Max(0, slider.Value));
            element.Volume = volume;
            if (updateMuteState)
            {
                element.IsMuted = volume <= 0.001;
            }
        }

        private MediaElement ResolveElementForVolumeSlider(Slider slider)
        {
            if (ReferenceEquals(slider, PreviewVolumeSlider)) return PreviewMediaElement;
            if (ReferenceEquals(slider, PreviewAudioVolumeSlider)) return PreviewAudioElement;
            if (ReferenceEquals(slider, ImmersiveVolumeSlider)) return ImmersiveMediaElement;
            if (ReferenceEquals(slider, ImmersiveAudioVolumeSlider)) return ImmersiveAudioElement;
            return null;
        }

        private void StopOtherPreviewElement(MediaElement current)
        {
            StopElementIfOther(current, PreviewMediaElement, PreviewPositionSlider);
            StopElementIfOther(current, PreviewAudioElement, PreviewAudioPositionSlider);
        }

        private void StopOtherImmersiveElement(MediaElement current)
        {
            StopElementIfOther(current, ImmersiveMediaElement, ImmersivePositionSlider);
            StopElementIfOther(current, ImmersiveAudioElement, ImmersiveAudioPositionSlider);
        }

        private static void StopElementIfOther(MediaElement current, MediaElement target, Slider slider)
        {
            if (ReferenceEquals(current, target)) return;
            target.Stop();
            ResetSlider(slider);
        }

        private void StopAllMediaElements()
        {
            PreviewMediaElement.Stop();
            PreviewAudioElement.Stop();
            ResetSlider(PreviewPositionSlider);
            ResetSlider(PreviewAudioPositionSlider);
            StopImmersiveElements();
        }

        private void StopImmersiveElements()
        {
            ImmersiveMediaElement.Stop();
            ImmersiveAudioElement.Stop();
            ResetSlider(ImmersivePositionSlider);
            ResetSlider(ImmersiveAudioPositionSlider);
        }

        private void PositionTimer_OnTick(object sender, EventArgs e)
        {
            UpdateSlider(PreviewMediaElement, PreviewPositionSlider);
            UpdateSlider(PreviewAudioElement, PreviewAudioPositionSlider);
            UpdateSlider(ImmersiveMediaElement, ImmersivePositionSlider);
            UpdateSlider(ImmersiveAudioElement, ImmersiveAudioPositionSlider);
        }

        private void UpdateSlider(MediaElement element, Slider slider)
        {
            if (element == null || slider == null || !element.NaturalDuration.HasTimeSpan) return;
            if (_draggingPositionSliders.Contains(slider) || slider.IsMouseCaptureWithin) return;
            var total = element.NaturalDuration.TimeSpan.TotalSeconds;
            if (total <= 0) return;
            slider.Maximum = total;
            slider.Value = Math.Min(total, Math.Max(0, element.Position.TotalSeconds));
        }
    }
}
