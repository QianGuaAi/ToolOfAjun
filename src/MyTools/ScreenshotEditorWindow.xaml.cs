using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Microsoft.Win32;

namespace MyTools
{
    public partial class ScreenshotEditorWindow : Window
    {
        private enum DrawingMode { Pen, Rect, Line, Eraser }
        private enum UndoType { Stroke, Shape }

        private struct UndoItem
        {
            public UndoType Type;
            public Stroke Stroke;
            public UIElement Shape;
        }

        private DrawingMode _mode = DrawingMode.Pen;
        private Color _currentColor = Colors.Red;
        private readonly Stack<UndoItem> _undoStack = new Stack<UndoItem>();

        private Point _shapeStart;
        private Shape _previewShape;

        public ScreenshotEditorWindow()
        {
            InitializeComponent();
            SetBrushColor(Colors.Red);
            if (BtnColorRed != null)
            {
                MarkActiveColorButton(BtnColorRed);
            }
            DrawingCanvas.Strokes.StrokesChanged += Strokes_StrokesChanged;
        }

        public void LoadScreenshot(BitmapSource screenshot)
        {
            ScreenshotImage.Source = screenshot;
            var w = screenshot.PixelWidth;
            var h = screenshot.PixelHeight;
            DrawingCanvas.Width  = w;
            DrawingCanvas.Height = h;
            ShapesCanvas.Width   = w;
            ShapesCanvas.Height  = h;
            CanvasContainer.Width  = w;
            CanvasContainer.Height = h;
        }

        private void SetBrushColor(Color color)
        {
            _currentColor = color;
            DrawingCanvas.DefaultDrawingAttributes.Color = color;
        }

        private void MarkActiveColorButton(Button active)
        {
            if (active == null)
            {
                return;
            }

            Button[] all = {
                BtnColorRed, BtnColorOrange, BtnColorYellow, BtnColorGreen,
                BtnColorCyan, BtnColorBlue, BtnColorPurple, BtnColorWhite, BtnColorBlack
            };
            foreach (var btn in all)
            {
                if (btn != null)
                {
                    btn.BorderBrush = Brushes.Transparent;
                }
            }

            active.BorderBrush = Brushes.White;
        }

        private void SetMode(DrawingMode mode)
        {
            _mode = mode;
            switch (mode)
            {
                case DrawingMode.Pen:
                    DrawingCanvas.EditingMode = InkCanvasEditingMode.Ink;
                    DrawingCanvas.IsHitTestVisible = true;
                    ShapesCanvas.IsHitTestVisible  = false;
                    TglPen.IsChecked    = true;
                    TglRect.IsChecked   = false;
                    TglLine.IsChecked   = false;
                    TglEraser.IsChecked = false;
                    break;
                case DrawingMode.Rect:
                case DrawingMode.Line:
                    DrawingCanvas.EditingMode = InkCanvasEditingMode.None;
                    DrawingCanvas.IsHitTestVisible = false;
                    ShapesCanvas.IsHitTestVisible  = true;
                    TglPen.IsChecked    = false;
                    TglRect.IsChecked   = (mode == DrawingMode.Rect);
                    TglLine.IsChecked   = (mode == DrawingMode.Line);
                    TglEraser.IsChecked = false;
                    break;
                case DrawingMode.Eraser:
                    DrawingCanvas.EditingMode = InkCanvasEditingMode.EraseByPoint;
                    DrawingCanvas.EraserShape = new EllipseStylusShape(20, 20);
                    DrawingCanvas.IsHitTestVisible = true;
                    ShapesCanvas.IsHitTestVisible  = false;
                    TglPen.IsChecked    = false;
                    TglRect.IsChecked   = false;
                    TglLine.IsChecked   = false;
                    TglEraser.IsChecked = true;
                    break;
            }
        }

        private RenderTargetBitmap FlattenToImage()
        {
            var w = (int)CanvasContainer.ActualWidth;
            var h = (int)CanvasContainer.ActualHeight;
            if (w <= 0 || h <= 0) return null;

            var rtb = new RenderTargetBitmap(w, h, 96.0, 96.0, PixelFormats.Pbgra32);
            rtb.Render(CanvasContainer);
            return rtb;
        }

        private void Strokes_StrokesChanged(object sender, StrokeCollectionChangedEventArgs e)
        {
            foreach (var s in e.Added)
                _undoStack.Push(new UndoItem { Type = UndoType.Stroke, Stroke = s });
        }

        private void ColorBtn_Click(object sender, RoutedEventArgs e)
        {
            if (!(sender is Button btn)) return;
            var hex = btn.Tag as string;
            if (string.IsNullOrEmpty(hex)) return;
            try
            {
                var color = (Color)ColorConverter.ConvertFromString(hex);
                SetBrushColor(color);
                MarkActiveColorButton(btn);
            }
            catch { }
            if (_mode == DrawingMode.Eraser) SetMode(DrawingMode.Pen);
        }

        private void SldBrushSize_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (DrawingCanvas == null || TxtBrushSize == null) return;
            var size = (int)e.NewValue;
            TxtBrushSize.Text = size.ToString();
            DrawingCanvas.DefaultDrawingAttributes.Width  = size;
            DrawingCanvas.DefaultDrawingAttributes.Height = size;
        }

        private void TglPen_Checked(object sender, RoutedEventArgs e)    => SetMode(DrawingMode.Pen);
        private void TglRect_Checked(object sender, RoutedEventArgs e)   => SetMode(DrawingMode.Rect);
        private void TglLine_Checked(object sender, RoutedEventArgs e)   => SetMode(DrawingMode.Line);
        private void TglEraser_Checked(object sender, RoutedEventArgs e) => SetMode(DrawingMode.Eraser);
        private void TglEraser_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_mode == DrawingMode.Eraser) SetMode(DrawingMode.Pen);
        }

        private void ShapesCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (_mode != DrawingMode.Rect && _mode != DrawingMode.Line) return;
            _shapeStart = e.GetPosition(ShapesCanvas);
            ShapesCanvas.CaptureMouse();

            var thickness = (int)SldBrushSize.Value;
            var stroke = new SolidColorBrush(_currentColor);
            var dash = new DoubleCollection(new double[] { 6, 3 });

            if (_mode == DrawingMode.Rect)
            {
                _previewShape = new Rectangle
                {
                    Stroke = stroke,
                    StrokeThickness = thickness,
                    Fill = Brushes.Transparent,
                    StrokeDashArray = dash
                };
                Canvas.SetLeft(_previewShape, _shapeStart.X);
                Canvas.SetTop(_previewShape, _shapeStart.Y);
            }
            else
            {
                _previewShape = new Line
                {
                    Stroke = stroke,
                    StrokeThickness = thickness,
                    StrokeDashArray = dash,
                    StrokeStartLineCap = PenLineCap.Round,
                    StrokeEndLineCap = PenLineCap.Round,
                    X1 = _shapeStart.X, Y1 = _shapeStart.Y,
                    X2 = _shapeStart.X, Y2 = _shapeStart.Y
                };
            }
            ShapesCanvas.Children.Add(_previewShape);
        }

        private void ShapesCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (_previewShape == null || e.LeftButton != MouseButtonState.Pressed) return;
            var pos = e.GetPosition(ShapesCanvas);

            if (_previewShape is Rectangle rect)
            {
                var x = Math.Min(pos.X, _shapeStart.X);
                var y = Math.Min(pos.Y, _shapeStart.Y);
                rect.Width  = Math.Abs(pos.X - _shapeStart.X);
                rect.Height = Math.Abs(pos.Y - _shapeStart.Y);
                Canvas.SetLeft(rect, x);
                Canvas.SetTop(rect, y);
            }
            else if (_previewShape is Line line)
            {
                line.X2 = pos.X;
                line.Y2 = pos.Y;
            }
        }

        private void ShapesCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (_previewShape == null) return;
            ShapesCanvas.ReleaseMouseCapture();

            _previewShape.StrokeDashArray = null;

            var valid = false;
            if (_previewShape is Rectangle r)
                valid = r.Width > 2 && r.Height > 2;
            else if (_previewShape is Line l)
                valid = Math.Abs(l.X2 - l.X1) > 2 || Math.Abs(l.Y2 - l.Y1) > 2;

            if (!valid)
                ShapesCanvas.Children.Remove(_previewShape);
            else
                _undoStack.Push(new UndoItem { Type = UndoType.Shape, Shape = _previewShape });

            _previewShape = null;
        }

        private void BtnUndo_Click(object sender, RoutedEventArgs e)
        {
            if (_undoStack.Count == 0) return;
            var item = _undoStack.Pop();
            if (item.Type == UndoType.Stroke)
                DrawingCanvas.Strokes.Remove(item.Stroke);
            else
                ShapesCanvas.Children.Remove(item.Shape);
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            if (DrawingCanvas.Strokes.Count == 0 && ShapesCanvas.Children.Count == 0) return;
            var result = MessageBox.Show("清除所有标注？", "确认", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes) return;
            DrawingCanvas.Strokes.Clear();
            ShapesCanvas.Children.Clear();
            _undoStack.Clear();
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
            var ext = System.IO.Path.GetExtension(dlg.FileName).ToLowerInvariant();
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

        private void BtnClose_Click(object sender, RoutedEventArgs e) => Close();

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            var ctrl = Keyboard.Modifiers == ModifierKeys.Control;
            if (e.Key == Key.Escape) { Close(); return; }
            if (ctrl && e.Key == Key.Z) { BtnUndo_Click(null, null); e.Handled = true; return; }
            if (ctrl && e.Key == Key.C) { BtnCopyClipboard_Click(null, null); e.Handled = true; return; }
            if (ctrl && e.Key == Key.S) { BtnSaveAs_Click(null, null); e.Handled = true; return; }

            if (Keyboard.Modifiers != ModifierKeys.None) return;
            if (e.Key == Key.P) SetMode(DrawingMode.Pen);
            else if (e.Key == Key.R) SetMode(DrawingMode.Rect);
            else if (e.Key == Key.L) SetMode(DrawingMode.Line);
            else if (e.Key == Key.E) SetMode(_mode == DrawingMode.Eraser ? DrawingMode.Pen : DrawingMode.Eraser);
        }

        protected override void OnClosed(EventArgs e)
        {
            if (DrawingCanvas != null)
            {
                DrawingCanvas.Strokes.StrokesChanged -= Strokes_StrokesChanged;
                DrawingCanvas.Strokes.Clear();
            }

            if (ScreenshotImage != null)
            {
                ScreenshotImage.Source = null;
            }

            if (ShapesCanvas != null)
            {
                ShapesCanvas.Children.Clear();
            }

            _undoStack.Clear();

            base.OnClosed(e);
        }
    }
}
