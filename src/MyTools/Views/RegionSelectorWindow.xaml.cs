using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace MyTools.Views
{
    /// <summary>
    /// 全屏覆盖式区域选择窗口。
    /// 以传入的截图作为背景，半透明蒙层；用户拖动产生矩形，确认后通过 <see cref="SelectedRectPx"/> 暴露物理像素矩形（基于原 BitmapSource 的像素坐标）。
    /// </summary>
    public partial class RegionSelectorWindow : Window
    {
        [DllImport("user32.dll")] private static extern int GetSystemMetrics(int nIndex);
        [DllImport("user32.dll", SetLastError = true)] private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);
        private const int SM_XVIRTUALSCREEN = 76;
        private const int SM_YVIRTUALSCREEN = 77;
        private const int SM_CXVIRTUALSCREEN = 78;
        private const int SM_CYVIRTUALSCREEN = 79;
        private const uint SWP_NOZORDER = 0x0004;

        private readonly BitmapSource _snapshot;
        private readonly int _virtualWidthPx;
        private readonly int _virtualHeightPx;

        private bool _dragging;
        private Point _startPoint;

        /// <summary>
        /// 选区在 <paramref name="snapshot"/> 像素坐标系下的矩形（左上为 0,0）。
        /// 关闭后若 DialogResult=true 则有效。
        /// </summary>
        public Int32Rect SelectedRectPx { get; private set; }

        public RegionSelectorWindow(BitmapSource snapshot)
        {
            InitializeComponent();
            _snapshot = snapshot ?? throw new ArgumentNullException(nameof(snapshot));

            var virtualLeftPx = GetSystemMetrics(SM_XVIRTUALSCREEN);
            var virtualTopPx = GetSystemMetrics(SM_YVIRTUALSCREEN);
            _virtualWidthPx = GetSystemMetrics(SM_CXVIRTUALSCREEN);
            _virtualHeightPx = GetSystemMetrics(SM_CYVIRTUALSCREEN);

            BackgroundImage.Source = _snapshot;
            SelectionImage.Source = _snapshot;

            PlaceWindowByFallbackDpi(virtualLeftPx, virtualTopPx);
            SourceInitialized += (s, e) => PlaceWindowByHwndDpi(virtualLeftPx, virtualTopPx);

            Loaded += (s, e) =>
            {
                Activate();
                Focus();
                Keyboard.Focus(this);
            };
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                DialogResult = false;
                Close();
            }
            else if (e.Key == Key.Enter)
            {
                if (TryFinalize()) { DialogResult = true; Close(); }
            }
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _dragging = true;
            _startPoint = e.GetPosition(RootGrid);
            SelectionBorder.Visibility = Visibility.Visible;
            SelectionImage.Visibility = Visibility.Visible;
            SizeHint.Visibility = Visibility.Visible;
            CaptureMouse();
            UpdateSelection(_startPoint, _startPoint);
        }

        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_dragging) return;
            var p = e.GetPosition(RootGrid);
            UpdateSelection(_startPoint, p);
        }

        private void Window_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_dragging) return;
            _dragging = false;
            ReleaseMouseCapture();
            if (TryFinalize()) { DialogResult = true; Close(); }
        }

        private void UpdateSelection(Point a, Point b)
        {
            double x = Math.Min(a.X, b.X), y = Math.Min(a.Y, b.Y);
            double w = Math.Abs(a.X - b.X), h = Math.Abs(a.Y - b.Y);

            Canvas.SetLeft(SelectionBorder, x);
            Canvas.SetTop(SelectionBorder, y);
            SelectionBorder.Width = w;
            SelectionBorder.Height = h;

            // Keep the preview image aligned with the background, then clip out only the selected area.
            Canvas.SetLeft(SelectionImage, 0);
            Canvas.SetTop(SelectionImage, 0);
            SelectionImage.Width = ActualWidth;
            SelectionImage.Height = ActualHeight;
            SelectionImage.Clip = new RectangleGeometry(new Rect(x, y, w, h));
            SelectionImage.RenderTransform = null;

            // 尺寸提示
            var scale = GetSnapshotScale();
            int wPx = (int)Math.Round(w * scale.X);
            int hPx = (int)Math.Round(h * scale.Y);
            SizeHintText.Text = $"{wPx} × {hPx}";
            // 提示框放在选区上方；若空间不够则放下方
            double hintX = x;
            double hintY = y - 28;
            if (hintY < 4) hintY = y + h + 6;
            Canvas.SetLeft(SizeHint, hintX);
            Canvas.SetTop(SizeHint, hintY);
        }

        /// <summary>把当前 DIP 选区换算成 _snapshot 像素坐标。</summary>
        private bool TryFinalize()
        {
            if (SelectionBorder.Visibility != Visibility.Visible) return false;
            double w = SelectionBorder.Width, h = SelectionBorder.Height;
            if (w < 4 || h < 4) return false;
            double x = Canvas.GetLeft(SelectionBorder);
            double y = Canvas.GetTop(SelectionBorder);

            var scale = GetSnapshotScale();
            int xPx = (int)Math.Floor(x * scale.X);
            int yPx = (int)Math.Floor(y * scale.Y);
            int rightPx = (int)Math.Ceiling((x + w) * scale.X);
            int bottomPx = (int)Math.Ceiling((y + h) * scale.Y);
            xPx = Clamp(xPx, 0, _snapshot.PixelWidth - 1);
            yPx = Clamp(yPx, 0, _snapshot.PixelHeight - 1);
            rightPx = Clamp(rightPx, xPx + 1, _snapshot.PixelWidth);
            bottomPx = Clamp(bottomPx, yPx + 1, _snapshot.PixelHeight);

            int wPx = rightPx - xPx;
            int hPx = bottomPx - yPx;
            if (wPx < 1 || hPx < 1) return false;

            SelectedRectPx = new Int32Rect(xPx, yPx, wPx, hPx);
            return true;
        }

        private Point GetSnapshotScale()
        {
            double width = ActualWidth > 0 ? ActualWidth : Width;
            double height = ActualHeight > 0 ? ActualHeight : Height;
            if (width <= 0 || height <= 0)
            {
                return new Point(1.0, 1.0);
            }

            return new Point(_snapshot.PixelWidth / width, _snapshot.PixelHeight / height);
        }

        private void PlaceWindowByFallbackDpi(int virtualLeftPx, int virtualTopPx)
        {
            var dpi = VisualTreeHelper.GetDpi(this);
            Left = virtualLeftPx / dpi.DpiScaleX;
            Top = virtualTopPx / dpi.DpiScaleY;
            Width = _virtualWidthPx / dpi.DpiScaleX;
            Height = _virtualHeightPx / dpi.DpiScaleY;
        }

        private void PlaceWindowByHwndDpi(int virtualLeftPx, int virtualTopPx)
        {
            try
            {
                var source = PresentationSource.FromVisual(this);
                var transform = source?.CompositionTarget?.TransformFromDevice ?? Matrix.Identity;
                var topLeftDip = transform.Transform(new Point(virtualLeftPx, virtualTopPx));
                var bottomRightDip = transform.Transform(new Point(virtualLeftPx + _virtualWidthPx, virtualTopPx + _virtualHeightPx));

                Left = topLeftDip.X;
                Top = topLeftDip.Y;
                Width = Math.Max(1, bottomRightDip.X - topLeftDip.X);
                Height = Math.Max(1, bottomRightDip.Y - topLeftDip.Y);

                var hwnd = new WindowInteropHelper(this).Handle;
                if (hwnd != IntPtr.Zero)
                {
                    SetWindowPos(hwnd, IntPtr.Zero, virtualLeftPx, virtualTopPx, _virtualWidthPx, _virtualHeightPx, SWP_NOZORDER);
                }
            }
            catch (Exception ex)
            {
                Services.AppLogService.Warning("Place region selector window failed: {Msg}", ex.Message);
            }
        }

        private static int Clamp(int value, int min, int max)
        {
            if (max < min) return min;
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }
    }
}
