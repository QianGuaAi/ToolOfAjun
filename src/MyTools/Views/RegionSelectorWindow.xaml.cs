using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
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
        private const int SM_XVIRTUALSCREEN = 76;
        private const int SM_YVIRTUALSCREEN = 77;
        private const int SM_CXVIRTUALSCREEN = 78;
        private const int SM_CYVIRTUALSCREEN = 79;

        private readonly BitmapSource _snapshot;
        private readonly int _virtualLeftPx;
        private readonly int _virtualTopPx;
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

            _virtualLeftPx = GetSystemMetrics(SM_XVIRTUALSCREEN);
            _virtualTopPx = GetSystemMetrics(SM_YVIRTUALSCREEN);
            _virtualWidthPx = GetSystemMetrics(SM_CXVIRTUALSCREEN);
            _virtualHeightPx = GetSystemMetrics(SM_CYVIRTUALSCREEN);

            BackgroundImage.Source = _snapshot;
            SelectionImage.Source = _snapshot;

            // 把窗口铺满整个虚拟屏幕（DIP）。
            // 因为本进程是 Per-Monitor V2 DPI-aware：
            // - WPF 的 Top/Left/Width/Height 单位是 "WPF DIP（按主显示器 DPI）"；
            //   而 GetSystemMetrics 返回物理像素。这里把进程上下文 DPI 用 PresentationSource 推算出来。
            // 简化处理：用主显示器 DPI。
            var dpi = VisualTreeHelper.GetDpi(this);
            Left = _virtualLeftPx / dpi.DpiScaleX;
            Top = _virtualTopPx / dpi.DpiScaleY;
            Width = _virtualWidthPx / dpi.DpiScaleX;
            Height = _virtualHeightPx / dpi.DpiScaleY;

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

            // 高亮选区：把同一张图按相对偏移显示
            Canvas.SetLeft(SelectionImage, x);
            Canvas.SetTop(SelectionImage, y);
            SelectionImage.Width = w;
            SelectionImage.Height = h;
            // 用 ImageBrush 也行，这里用 Image + Clip 的 RectangleGeometry
            SelectionImage.Clip = new RectangleGeometry(new Rect(0, 0, w, h));
            // 通过 RenderTransform 把图像的视区平移到选中那一块
            SelectionImage.RenderTransform = new TranslateTransform(-x, -y);
            // 让 SelectionImage 自身仍铺满整个虚拟屏幕大小
            SelectionImage.Width = ActualWidth;
            SelectionImage.Height = ActualHeight;

            // 尺寸提示
            var dpi = VisualTreeHelper.GetDpi(this);
            int wPx = (int)Math.Round(w * dpi.DpiScaleX);
            int hPx = (int)Math.Round(h * dpi.DpiScaleY);
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

            var dpi = VisualTreeHelper.GetDpi(this);
            int xPx = Math.Max(0, (int)Math.Round(x * dpi.DpiScaleX));
            int yPx = Math.Max(0, (int)Math.Round(y * dpi.DpiScaleY));
            int wPx = (int)Math.Round(w * dpi.DpiScaleX);
            int hPx = (int)Math.Round(h * dpi.DpiScaleY);
            // 防溢出
            wPx = Math.Min(wPx, _snapshot.PixelWidth - xPx);
            hPx = Math.Min(hPx, _snapshot.PixelHeight - yPx);
            if (wPx < 1 || hPx < 1) return false;

            SelectedRectPx = new Int32Rect(xPx, yPx, wPx, hPx);
            return true;
        }
    }
}
