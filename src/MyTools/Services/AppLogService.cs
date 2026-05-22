using System;
using System.IO;
using System.Runtime.CompilerServices;
using Serilog;

namespace MyTools.Services
{
    public static class AppLogService
    {
        private static readonly object SyncRoot = new object();
        private static bool _initialized;

        public static bool IsInitialized => _initialized;

        public static void Initialize()
        {
            if (_initialized)
            {
                return;
            }

            lock (SyncRoot)
            {
                if (_initialized)
                {
                    return;
                }

                var logDirectory = AppDomain.CurrentDomain.BaseDirectory;
                Directory.CreateDirectory(logDirectory);
                CosturaBootstrap.EnsureInitialized();

                ConfigureLogger(logDirectory);

                _initialized = true;
            }
        }

        public static void Information(string messageTemplate, params object[] propertyValues)
        {
            Initialize();
            WriteInformation(messageTemplate, propertyValues);
        }

        public static void InformationIfInitialized(string messageTemplate, params object[] propertyValues)
        {
            if (!_initialized)
            {
                return;
            }

            WriteInformation(messageTemplate, propertyValues);
        }

        public static void Warning(string messageTemplate, params object[] propertyValues)
        {
            Initialize();
            WriteWarning(messageTemplate, propertyValues);
        }

        public static void Error(Exception exception, string messageTemplate, params object[] propertyValues)
        {
            Initialize();
            WriteError(exception, messageTemplate, propertyValues);
        }

        public static void Error(string messageTemplate, params object[] propertyValues)
        {
            Initialize();
            WriteError(messageTemplate, propertyValues);
        }

        public static void CloseAndFlush()
        {
            if (!_initialized)
            {
                return;
            }

            FlushLogger();
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void ConfigureLogger(string logDirectory)
        {
            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Information()
                .WriteTo.File(
                    Path.Combine(logDirectory, "MyTools.log"),
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 14,
                    fileSizeLimitBytes: 5 * 1024 * 1024,
                    rollOnFileSizeLimit: true,
                    shared: false,
                    buffered: true,
                    flushToDiskInterval: TimeSpan.FromSeconds(2),
                    encoding: System.Text.Encoding.UTF8)
                .CreateLogger();
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void WriteInformation(string messageTemplate, params object[] propertyValues)
        {
            Log.Information(messageTemplate, propertyValues);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void WriteWarning(string messageTemplate, params object[] propertyValues)
        {
            Log.Warning(messageTemplate, propertyValues);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void WriteError(Exception exception, string messageTemplate, params object[] propertyValues)
        {
            Log.Error(exception, messageTemplate, propertyValues);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void WriteError(string messageTemplate, params object[] propertyValues)
        {
            Log.Error(messageTemplate, propertyValues);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void FlushLogger()
        {
            Log.CloseAndFlush();
        }
    }
}
