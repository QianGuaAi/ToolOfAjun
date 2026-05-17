using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using MaterialDesignThemes.Wpf;
using MyTools.Services;

namespace MyTools
{
    public partial class RecordRegionWindow : Window
    {
        public event EventHandler ToggleRecordingRequested;

        private readonly DispatcherTimer _recordingTimer;
        private readonly CancellationTokenSource _lifetimeCancellation = new CancellationTokenSource();
        private DateTime _recordingStartedAt;
        private bool _isCountingDown;

        public RecordRegionWindow()
        {
            InitializeComponent();
            _recordingTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            _recordingTimer.Tick += RecordingTimer_OnTick;
        }

        public RecordingRegion GetCaptureRegion()
        {
            var dpiScale = VisualTreeHelper.GetDpi(this);
            return new RecordingRegion
            {
                X = (int)Math.Round(Left * dpiScale.DpiScaleX),
                Y = (int)Math.Round(Top * dpiScale.DpiScaleY),
                Width = (int)Math.Round(ActualWidth * dpiScale.DpiScaleX),
                Height = (int)Math.Round(ActualHeight * dpiScale.DpiScaleY)
            };
        }

        public async Task<bool> RunStartCountdownAsync(CancellationToken cancellationToken)
        {
            if (_isCountingDown)
            {
                return false;
            }

            using (var linked = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, _lifetimeCancellation.Token))
            {
                _isCountingDown = true;
                ToggleRecordButton.IsEnabled = false;
                CountdownOverlay.Visibility = Visibility.Visible;

                try
                {
                    for (var value = 3; value >= 1; value--)
                    {
                        CountdownText.Text = value.ToString();
                        await Task.Delay(1000, linked.Token);
                    }

                    CountdownText.Text = "录";
                    await Task.Delay(160, linked.Token);
                    return true;
                }
                catch (OperationCanceledException)
                {
                    return false;
                }
                finally
                {
                    CountdownOverlay.Visibility = Visibility.Collapsed;
                    ToggleRecordButton.IsEnabled = true;
                    _isCountingDown = false;
                }
            }
        }

        public void SetRecordingState(bool isRecording)
        {
            if (RecordIcon == null)
            {
                return;
            }

            if (isRecording)
            {
                RecordIcon.Kind = PackIconKind.StopCircle;
                RecordIcon.Foreground = Brushes.Gray;
                ToggleRecordButton.ToolTip = "停止录像";
                RecordingStatusPanel.Visibility = Visibility.Visible;
                _recordingStartedAt = DateTime.Now;
                UpdateRecordingStatus();
                _recordingTimer.Start();
            }
            else
            {
                _recordingTimer.Stop();
                RecordingStatusPanel.Visibility = Visibility.Collapsed;
                CountdownOverlay.Visibility = Visibility.Collapsed;
                _isCountingDown = false;
                RecordIcon.Kind = PackIconKind.RecordRec;
                RecordIcon.Foreground = Brushes.White;
                ToggleRecordButton.ToolTip = "开始录像";
            }
        }

        private void ToggleRecordButton_Click(object sender, RoutedEventArgs e)
        {
            if (_isCountingDown)
            {
                return;
            }

            ToggleRecordingRequested?.Invoke(this, EventArgs.Empty);
        }

        private void RecordingTimer_OnTick(object sender, EventArgs e)
        {
            UpdateRecordingStatus();
        }

        private void UpdateRecordingStatus()
        {
            var elapsed = DateTime.Now - _recordingStartedAt;
            if (elapsed < TimeSpan.Zero)
            {
                elapsed = TimeSpan.Zero;
            }

            RecordingStatusText.Text = elapsed.TotalHours >= 1
                ? "录制中 " + elapsed.ToString(@"h\:mm\:ss")
                : "录制中 " + elapsed.ToString(@"mm\:ss");
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton != MouseButtonState.Pressed)
            {
                return;
            }

            try
            {
                DragMove();
            }
            catch
            {
                // Ignore drag interruptions.
            }
        }

        protected override void OnClosed(EventArgs e)
        {
            _recordingTimer.Stop();
            _recordingTimer.Tick -= RecordingTimer_OnTick;
            _lifetimeCancellation.Cancel();
            _lifetimeCancellation.Dispose();
            base.OnClosed(e);
        }
    }
}
