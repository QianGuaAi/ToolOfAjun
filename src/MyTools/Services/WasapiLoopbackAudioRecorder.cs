using System;
using System.IO;
using System.Threading;
using NAudio.Wave;

namespace MyTools.Services
{
    internal sealed class WasapiLoopbackAudioRecorder : IDisposable
    {
        private readonly object _syncRoot = new object();
        private readonly ManualResetEventSlim _stoppedEvent = new ManualResetEventSlim(false);
        private WasapiLoopbackCapture _capture;
        private WaveFileWriter _writer;
        private bool _disposed;
        private bool _stopRequested;

        private WasapiLoopbackAudioRecorder(string outputPath, WasapiLoopbackCapture capture, WaveFileWriter writer)
        {
            OutputPath = outputPath;
            DeviceName = capture.GetType().Name;
            _capture = capture;
            _writer = writer;
        }

        public string OutputPath { get; }

        public string DeviceName { get; private set; }

        public long BytesRecorded { get; private set; }

        public bool HasAudioData => BytesRecorded > 0 && File.Exists(OutputPath) && new FileInfo(OutputPath).Length > 44;

        public static WasapiLoopbackAudioRecorder Start(string outputPath)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? AppDomain.CurrentDomain.BaseDirectory);
            TryDeleteFile(outputPath);

            var capture = new WasapiLoopbackCapture();
            var writer = new WaveFileWriter(outputPath, capture.WaveFormat);
            var recorder = new WasapiLoopbackAudioRecorder(outputPath, capture, writer)
            {
                DeviceName = capture.WaveFormat.ToString()
            };

            capture.DataAvailable += recorder.HandleDataAvailable;
            capture.RecordingStopped += recorder.HandleRecordingStopped;

            try
            {
                capture.StartRecording();
                return recorder;
            }
            catch
            {
                recorder.Dispose();
                throw;
            }
        }

        public void Stop()
        {
            WasapiLoopbackCapture capture;
            lock (_syncRoot)
            {
                if (_disposed || _stopRequested)
                {
                    return;
                }

                _stopRequested = true;
                capture = _capture;
            }

            try
            {
                capture?.StopRecording();
                _stoppedEvent.Wait(TimeSpan.FromSeconds(2));
            }
            catch (Exception ex)
            {
                AppLogService.Warning("WASAPI loopback stop failed: {Msg}", ex.Message);
            }
            finally
            {
                Dispose();
            }
        }

        public void Dispose()
        {
            WasapiLoopbackCapture capture;
            WaveFileWriter writer;
            lock (_syncRoot)
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                capture = _capture;
                writer = _writer;
                _capture = null;
                _writer = null;
            }

            if (capture != null)
            {
                capture.DataAvailable -= HandleDataAvailable;
                capture.RecordingStopped -= HandleRecordingStopped;
                capture.Dispose();
            }

            writer?.Dispose();
            _stoppedEvent.Dispose();
        }

        private void HandleDataAvailable(object sender, WaveInEventArgs e)
        {
            lock (_syncRoot)
            {
                if (_disposed || _writer == null || e.BytesRecorded <= 0)
                {
                    return;
                }

                _writer.Write(e.Buffer, 0, e.BytesRecorded);
                BytesRecorded += e.BytesRecorded;
            }
        }

        private void HandleRecordingStopped(object sender, StoppedEventArgs e)
        {
            if (e.Exception != null)
            {
                AppLogService.Warning("WASAPI loopback capture stopped with error: {Msg}", e.Exception.Message);
            }

            _stoppedEvent.Set();
        }

        private static void TryDeleteFile(string filePath)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath))
                {
                    File.Delete(filePath);
                }
            }
            catch
            {
            }
        }
    }
}
