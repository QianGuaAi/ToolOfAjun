using System;
using System.IO;
using Serilog;

namespace MyTools.Services
{
    public static class AppLogService
    {
        private static readonly object SyncRoot = new object();
        private static bool _initialized;

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

                Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Information()
                    .WriteTo.File(
                        Path.Combine(logDirectory, "MyTools.log"),
                        rollingInterval: RollingInterval.Day,
                        retainedFileCountLimit: 14,
                        shared: true,
                        encoding: System.Text.Encoding.UTF8)
                    .CreateLogger();

                _initialized = true;
            }
        }

        public static void Information(string messageTemplate, params object[] propertyValues)
        {
            Initialize();
            Log.Information(messageTemplate, propertyValues);
        }

        public static void Error(Exception exception, string messageTemplate, params object[] propertyValues)
        {
            Initialize();
            Log.Error(exception, messageTemplate, propertyValues);
        }

        public static void CloseAndFlush()
        {
            if (!_initialized)
            {
                return;
            }

            Log.CloseAndFlush();
        }
    }
}
