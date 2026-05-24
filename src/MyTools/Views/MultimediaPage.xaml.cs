using System;
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

        private void MediaFileList_OnMouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            var vm = ViewModel;
            if (vm?.EnterImmersiveCommand.CanExecute(null) == true)
                vm.EnterImmersiveCommand.Execute(null);
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

        private void PreviewMediaElement_OnMediaEnded(object sender, RoutedEventArgs e)
        {
            if (sender is MediaElement element)
            {
                element.Stop();
                if (ViewModel != null) ViewModel.IsPreviewPlaying = false;
            }
        }

        private void ImmersiveMediaElement_OnMediaEnded(object sender, RoutedEventArgs e)
        {
            if (sender is MediaElement element)
            {
                element.Stop();
                if (ViewModel != null) ViewModel.IsImmersivePlaying = false;
            }
        }

        private void ToggleElement(MediaElement element, bool preview)
        {
            if (element == null || element.Source == null) return;
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

        private void StopOtherPreviewElement(MediaElement current)
        {
            if (!ReferenceEquals(current, PreviewMediaElement)) PreviewMediaElement.Stop();
            if (!ReferenceEquals(current, PreviewAudioElement)) PreviewAudioElement.Stop();
        }

        private void StopOtherImmersiveElement(MediaElement current)
        {
            if (!ReferenceEquals(current, ImmersiveMediaElement)) ImmersiveMediaElement.Stop();
            if (!ReferenceEquals(current, ImmersiveAudioElement)) ImmersiveAudioElement.Stop();
        }

        private void StopAllMediaElements()
        {
            PreviewMediaElement.Stop();
            PreviewAudioElement.Stop();
            StopImmersiveElements();
        }

        private void StopImmersiveElements()
        {
            ImmersiveMediaElement.Stop();
            ImmersiveAudioElement.Stop();
        }

        private void PositionTimer_OnTick(object sender, EventArgs e)
        {
            UpdateSlider(PreviewMediaElement, PreviewPositionSlider);
            UpdateSlider(PreviewAudioElement, PreviewAudioPositionSlider);
            UpdateSlider(ImmersiveMediaElement, ImmersivePositionSlider);
            UpdateSlider(ImmersiveAudioElement, ImmersiveAudioPositionSlider);
        }

        private static void UpdateSlider(MediaElement element, Slider slider)
        {
            if (element == null || slider == null || !element.NaturalDuration.HasTimeSpan) return;
            var total = element.NaturalDuration.TimeSpan.TotalSeconds;
            if (total <= 0) return;
            slider.Maximum = total;
            slider.Value = Math.Min(total, Math.Max(0, element.Position.TotalSeconds));
        }
    }
}
