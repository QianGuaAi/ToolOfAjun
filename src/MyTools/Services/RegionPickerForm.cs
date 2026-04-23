using System;
using System.Drawing;
using System.Windows.Forms;

namespace MyTools.Services
{
    /// <summary>
    /// Full-screen overlay to pick a screen region in physical pixels (multi-monitor safe).
    /// </summary>
    internal sealed class RegionPickerForm : Form
    {
        private Point _start;
        private Point _end;
        private bool _dragging;
        private readonly Pen _rubberBandPen;

        public Rectangle Selection { get; private set; }

        public RegionPickerForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.Black;
            Opacity = 0.35;
            Cursor = Cursors.Cross;
            DoubleBuffered = true;
            KeyPreview = true;

            var bounds = SystemInformation.VirtualScreen;
            SetBounds(bounds.X, bounds.Y, bounds.Width, bounds.Height);

            _rubberBandPen = new Pen(Color.DeepSkyBlue, 2) { DashStyle = System.Drawing.Drawing2D.DashStyle.Dash };
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (e.Button != MouseButtons.Left)
            {
                return;
            }

            _dragging = true;
            _start = e.Location;
            _end = e.Location;
            Selection = Rectangle.Empty;
            Invalidate();
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (!_dragging)
            {
                return;
            }

            _end = e.Location;
            Selection = NormalizeRect(_start, _end);
            Invalidate();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (!_dragging || e.Button != MouseButtons.Left)
            {
                return;
            }

            _dragging = false;
            _end = e.Location;
            Selection = NormalizeRect(_start, _end);

            if (Selection.Width < 3 || Selection.Height < 3)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return;
            }

            // Selection is in client coordinates; map to screen (virtual) coordinates.
            var screenRect = new Rectangle(
                Left + Selection.X,
                Top + Selection.Y,
                Selection.Width,
                Selection.Height);

            Selection = screenRect;
            DialogResult = DialogResult.OK;
            Close();
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Escape)
            {
                _dragging = false;
                DialogResult = DialogResult.Cancel;
                Close();
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (!_dragging && Selection.IsEmpty)
            {
                return;
            }

            var rect = _dragging ? NormalizeRect(_start, _end) : Rectangle.Empty;
            if (!rect.IsEmpty && rect.Width > 0 && rect.Height > 0)
            {
                e.Graphics.DrawRectangle(_rubberBandPen, rect);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _rubberBandPen.Dispose();
            }

            base.Dispose(disposing);
        }

        private static Rectangle NormalizeRect(Point a, Point b)
        {
            int x1 = Math.Min(a.X, b.X);
            int y1 = Math.Min(a.Y, b.Y);
            int x2 = Math.Max(a.X, b.X);
            int y2 = Math.Max(a.Y, b.Y);
            return new Rectangle(x1, y1, Math.Max(0, x2 - x1), Math.Max(0, y2 - y1));
        }
    }
}
