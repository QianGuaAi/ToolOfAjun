using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace MyTools
{
    public partial class ScreenshotEditorWindow : Window
    {
        private Color _currentColor = Colors.Red;

        public ScreenshotEditorWindow()
        {
            InitializeComponent();
            SetBrushColor(Colors.Red);
            MarkActiveColorButton(BtnColorRed);
        }

        public void LoadScreenshot(BitmapSource screenshot)
        {
            ScreenshotImage.Source = screenshot;
            DrawingCanvas.Width  = screenshot.PixelWidth;
            DrawingCanvas.Height = screenshot.PixelHeight;
            CanvasContainer.Width  = screenshot.PixelWidth;
            CanvasContainer.Height = screenshot.PixelHeight;
        }

        private void SetBrushColor(Color color)
        {
            _currentColor = color;
            DrawingCanvas.DefaultDrawingAttributes.Color = color;
        }

        private void MarkActiveColorButton(Button active)
        {
            Button[] all = {
                BtnColorRed, BtnColorOrange, BtnColorYellow, BtnColorGreen,
                BtnColorCyan, BtnColorBlue, BtnColorPurple, BtnColorWhite, BtnColorBlack
            };
            foreach (var btn in all)
                btn.BorderBrush = Brushes.Transparent;
            if (active != null)
                active.BorderBrush = Brushes.White;
        }

        private RenderTargetBitmap FlattenToImage()
        {
            var w = (int)DrawingCanvas.ActualWidth;
            var h = (int)DrawingCanvas.ActualHeight;
            if (w <= 0 || h <= 0) return null;

            var dpi = 96.0;
            var rtb = new RenderTargetBitmap(w, h, dpi, dpi, PixelFormats.Pbgra32);
            rtb.Render(CanvasContainer);
            return rtb;
        }

        private void ColorBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button btn)) return;
            TglEraser.IsChecked = false;
            DrawingCanvas.EditingMode = InkCanvasEditingMode.Ink;
            var hex = btn.Tag as string;
            if (!string.IsNullOrEmpty(hex))
            {
                try
                {
                    var color = (Color)ColorConverter.ConvertFromString(hex);
                    SetBrushColor(color);
                    MarkActiveColorButton(btn);
                }
                catch { }
            }
        }

        private void SldBrushSize_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (DrawingCanvas == null || TxtBrushSize == null) return;
            var size = (int)e.NewValue;
            TxtBrushSize.Text = size.ToString();
            DrawingCanvas.DefaultDrawingAttributes.Width  = size;
            DrawingCanvas.DefaultDrawingAttributes.Height = size;
        }

        private void TglEraser_Checked(object sender, RoutedEventArgs e)
        {
            DrawingCanvas.EditingMode = InkCanvasEditingMode.EraseByPoint;
            DrawingCanvas.EraserShape = new EllipseStylusShape(20, 20);
        }

        private void TglEraser_Unchecked(object sender, RoutedEventArgs e)
        {
            DrawingCanvas.EditingMode = InkCanvasEditingMode.Ink;
        }

        private void BtnUndo_Click(object sender, RoutedEventArgs e)
        {
            var strokes = DrawingCanvas.Strokes;
            if (strokes.Count > 0)
                strokes.RemoveAt(strokes.Count - 1);
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            if (DrawingCanvas.Strokes.Count == 0) return;
            var result = MessageBox.Show("清除所有标注？", "确认", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result == MessageBoxResult.Yes)
                DrawingCanvas.Strokes.Clear();
        }

        private void BtnCopyClipboard_Click(object sender, RoutedEventArgs e)
        {
            var img = FlattenToImage();
            if (img == null) return;
            Clipboard.SetImage(img);
            MessageBox.Show("已复制到剪贴板", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new SaveFileDialog
            {
                Title = "另存为图片",
                Filter = "PNG 图片|*.png|JPEG 图片|*.jpg|BMP 图片|*.bmp",
                DefaultExt = ".png",
                FileName = $"截图_{DateTime.Now:yyyyMMdd_HHmmss}"
            };
            if (dlg.ShowDialog() != true) return;

            var img = FlattenToImage();
            if (img == null) return;

            BitmapEncoder encoder;
            var ext = Path.GetExtension(dlg.FileName).ToLowerInvariant();
            if (ext == ".jpg" || ext == ".jpeg")
                encoder = new JpegBitmapEncoder { QualityLevel = 95 };
            else if (ext == ".bmp")
                encoder = new BmpBitmapEncoder();
            else
                encoder = new PngBitmapEncoder();

            encoder.Frames.Add(BitmapFrame.Create(img));
            using (var stream = File.OpenWrite(dlg.FileName))
                encoder.Save(stream);

            MessageBox.Show($"已保存：{dlg.FileName}", "完成", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Close();
            }
            else if (e.Key == Key.Z && Keyboard.Modifiers == ModifierKeys.Control)
            {
                BtnUndo_Click(null, null);
            }
            else if (e.Key == Key.C && Keyboard.Modifiers == ModifierKeys.Control)
            {
                BtnCopyClipboard_Click(null, null);
            }
            else if (e.Key == Key.S && Keyboard.Modifiers == ModifierKeys.Control)
            {
                BtnSaveAs_Click(null, null);
            }
            else if (e.Key == Key.E)
            {
                TglEraser.IsChecked = !TglEraser.IsChecked;
            }
        }
    }
}
