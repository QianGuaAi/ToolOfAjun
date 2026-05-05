using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using MyTools.Services;

namespace MyTools
{
    public partial class App : Application
    {
        private static readonly string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MyTools.startup.log");

        public static bool IsExiting { get; set; }

        public App()
        {
            RegisterGlobalExceptionHandlers();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                AppLogService.Initialize();
                AppLogService.Information("Application starting");

                base.OnStartup(e);

                var mainWindow = new MainWindow();
                MainWindow = mainWindow;
                mainWindow.Show();
            }
            catch (Exception ex)
            {
                HandleFatalException("应用启动失败", ex);
            }
        }

        protected override void OnExit(ExitEventArgs e)
        {
            AppLogService.Information("Application exiting with code {ExitCode}", e.ApplicationExitCode);
            AppLogService.CloseAndFlush();
            base.OnExit(e);
        }

        private void RegisterGlobalExceptionHandlers()
        {
            DispatcherUnhandledException += OnDispatcherUnhandledException;
            AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
            TaskScheduler.UnobservedTaskException += OnUnobservedTaskException;
        }

        private void OnDispatcherUnhandledException(object sender, DispatcherUnhandledExceptionEventArgs e)
        {
            LogException("UI 线程未处理异常", e.Exception);
            MessageBox.Show(BuildUserMessage(e.Exception), "MyTools 运行错误", MessageBoxButton.OK, MessageBoxImage.Error);
            e.Handled = true;
            Shutdown(-1);
        }

        private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            var exception = e.ExceptionObject as Exception ?? new Exception("未知未处理异常");
            LogException("AppDomain 未处理异常", exception);
        }

        private void OnUnobservedTaskException(object sender, UnobservedTaskExceptionEventArgs e)
        {
            LogException("Task 未观察到的异常", e.Exception);
            e.SetObserved();
        }

        private void HandleFatalException(string title, Exception ex)
        {
            LogException(title, ex);
            MessageBox.Show(BuildUserMessage(ex), title, MessageBoxButton.OK, MessageBoxImage.Error);
            Shutdown(-1);
        }

        private static string BuildUserMessage(Exception ex)
        {
            return "程序运行时发生异常，详细信息已写入同目录日志文件：\n"
                + LogPath
                + "\n\n异常类型："
                + ex.GetType().FullName
                + "\n异常消息："
                + ex.Message;
        }

        private static void LogException(string title, Exception ex)
        {
            try
            {
                AppLogService.Error(ex, "{Title}", title);
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath) ?? AppDomain.CurrentDomain.BaseDirectory);
                File.AppendAllText(
                    LogPath,
                    $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {title}{Environment.NewLine}{ex}{Environment.NewLine}{Environment.NewLine}",
                    Encoding.UTF8);
            }
            catch
            {
            }
        }
    }
}
