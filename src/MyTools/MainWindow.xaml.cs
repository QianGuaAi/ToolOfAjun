using System.ComponentModel;
using System.Drawing;
using System.Windows;
using MahApps.Metro.Controls;

namespace MyTools
{
    public partial class MainWindow : MetroWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            InitializeTrayIcon();
        }

        private void InitializeTrayIcon()
        {
            var executablePath = System.Reflection.Assembly.GetEntryAssembly()?.Location;
            if (string.IsNullOrWhiteSpace(executablePath))
            {
                return;
            }

            var icon = System.Drawing.Icon.ExtractAssociatedIcon(executablePath);
            if (icon != null)
            {
                TrayIcon.Icon = icon;
            }
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            try { System.IO.File.AppendAllText("debug.log", "OnClosing hit. IsExiting: " + App.IsExiting + "\n"); } catch {}
            
            // 如果程序不是正在从托盘退出，则隐藏窗口并取消关闭
            if (!App.IsExiting)
            {
                e.Cancel = true;
                this.Hide();
                try { System.IO.File.AppendAllText("debug.log", "Window hidden.\n"); } catch {}
            }
            base.OnClosing(e);
        }
    }
}
