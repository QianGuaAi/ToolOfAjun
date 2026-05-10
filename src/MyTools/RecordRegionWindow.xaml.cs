using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using MaterialDesignThemes.Wpf;
using MyTools.Services;

namespace MyTools
{
    public partial class RecordRegionWindow : Window
    {
        public event EventHandler ToggleRecordingRequested;

        public RecordRegionWindow()
        {
            InitializeComponent();
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
            }
            else
            {
                RecordIcon.Kind = PackIconKind.RecordRec;
                RecordIcon.Foreground = Brushes.White;
                ToggleRecordButton.ToolTip = "开始录像";
            }
        }

        private void ToggleRecordButton_Click(object sender, RoutedEventArgs e)
        {
            ToggleRecordingRequested?.Invoke(this, EventArgs.Empty);
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
    }
}
