using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using MyTools.Services;

namespace MyTools
{
    public partial class App : Application
    {
        private static readonly string LogPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MyTools.startup.log");

        private static Mutex _singleInstanceMutex;
        private static EventWaitHandle _activationEvent;

        public static bool IsExiting { get; set; }

        public App()
        {
            RegisterGlobalExceptionHandlers();
        }

        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                bool isNewInstance;
                _singleInstanceMutex = new Mutex(true, @"Local\MyTools_SingleInstance", out isNewInstance);
                if (!isNewInstance)
                {
                    try
                    {
                        using (var ev = EventWaitHandle.OpenExisting(@"Local\MyTools_Activate"))
                            ev.Set();
                    }
                    catch { }
                    Shutdown(0);
                    return;
                }

                _activationEvent = new EventWaitHandle(false, EventResetMode.AutoReset, @"Local\MyTools_Activate");
                var listenerThread = new Thread(() =>
                {
                    while (!IsExiting)
                    {
                        try
                        {
                            _activationEvent.WaitOne();
                            if (IsExiting)
                            {
                                break;
                            }

                            Dispatcher.InvokeAsync(() =>
                            {
                                var win = MainWindow;
                                if (win == null) return;
                                win.Show();
                                if (win.WindowState == WindowState.Minimized)
                                    win.WindowState = WindowState.Normal;
                                win.Activate();
                                win.Focus();
                            });
                        }
                        catch { }
                    }
                }) { IsBackground = true, Name = "ActivationListener" };
                listenerThread.Start();

                AppLogService.Initialize();
                AppLogService.Information("Application starting on {Os}, 64bit={Is64}, Framework={Fx}",
                    OsVersionService.DisplayName,
                    Environment.Is64BitOperatingSystem,
                    System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription);

                try
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(LogPath) ?? AppDomain.CurrentDomain.BaseDirectory);
                    File.AppendAllText(
                        LogPath,
                        $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 启动 — OS={OsVersionService.DisplayName}, 64bit={Environment.Is64BitOperatingSystem}, .NET={System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription}{Environment.NewLine}",
                        Encoding.UTF8);
                }
                catch
                {
                }

                if (!OsVersionService.IsWindows10OrGreater)
                {
                    AppLogService.Warning("Running on legacy Windows ({Os}). Some modules (Lock Win10 22H2 / Defender / Auto Update / DXGI capture) will be hidden or unavailable.",
                        OsVersionService.DisplayName);
                }

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
            IsExiting = true;
            _activationEvent?.Set();
            _activationEvent?.Dispose();
            try { _singleInstanceMutex?.ReleaseMutex(); } catch { }
            _singleInstanceMutex?.Dispose();
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
            MessageBox.Show(BuildUserMessage(e.Exception), "阿君的工具运行错误", MessageBoxButton.OK, MessageBoxImage.Error);
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
