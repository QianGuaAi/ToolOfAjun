using System;
using System.IO;
using System.Threading;
using NAudio.Wave;

namespace MyTools.Services
{
    internal sealed class WasapiLoopbackAudioRecorder : IDisposable
    {
        internal const double AudibleSampleThreshold = 0.0001d;
        internal const long MinAudibleSampleCount = 128;
        private static readonly Guid PcmSubFormat = new Guid("00000001-0000-0010-8000-00aa00389b71");
        private static readonly Guid IeeeFloatSubFormat = new Guid("00000003-0000-0010-8000-00aa00389b71");
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

        public long AudibleSampleCount { get; private set; }

        public double MaxSampleLevel { get; private set; }

        private double CurrentSampleLevel { get; set; }

        public bool HasAudibleAudio => GetLiveAudioLevelStats().HasAudibleAudio;

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
                TrackAudioLevel(e.Buffer, e.BytesRecorded, _writer.WaveFormat);
            }
        }

        public AudioLevelStats GetLiveAudioLevelStats()
        {
            lock (_syncRoot)
            {
                return new AudioLevelStats(BytesRecorded, MaxSampleLevel, AudibleSampleCount, CurrentSampleLevel);
            }
        }

        public AudioLevelStats ConsumeLiveAudioLevelStats()
        {
            lock (_syncRoot)
            {
                var stats = new AudioLevelStats(BytesRecorded, MaxSampleLevel, AudibleSampleCount, CurrentSampleLevel);
                CurrentSampleLevel = 0d;
                return stats;
            }
        }

        public static AudioLevelStats ScanWaveFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath) || new FileInfo(filePath).Length <= 44)
            {
                return new AudioLevelStats(0, 0d, 0);
            }

            try
            {
                using (var reader = new WaveFileReader(filePath))
                {
                    var sampleProvider = reader.ToSampleProvider();
                    var buffer = new float[Math.Max(1024, sampleProvider.WaveFormat.SampleRate / 10)];
                    long audibleSampleCount = 0;
                    double maxSampleLevel = 0d;
                    int samplesRead;
                    while ((samplesRead = sampleProvider.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        for (var index = 0; index < samplesRead; index++)
                        {
                            var sample = buffer[index];
                            if (float.IsNaN(sample) || float.IsInfinity(sample))
                            {
                                continue;
                            }

                            var level = Math.Abs(sample);
                            if (level > maxSampleLevel)
                            {
                                maxSampleLevel = level;
                            }

                            if (level >= AudibleSampleThreshold)
                            {
                                audibleSampleCount++;
                            }

                            if (audibleSampleCount >= MinAudibleSampleCount && maxSampleLevel >= AudibleSampleThreshold)
                            {
                                return new AudioLevelStats(GetFileSize(filePath), maxSampleLevel, audibleSampleCount);
                            }
                        }
                    }

                    return new AudioLevelStats(GetFileSize(filePath), maxSampleLevel, audibleSampleCount);
                }
            }
            catch (Exception ex)
            {
                AppLogService.Warning("Scanning loopback audio file failed: {Msg}", ex.Message);
                return new AudioLevelStats(GetFileSize(filePath), 0d, 0);
            }
        }

        private void TrackAudioLevel(byte[] buffer, int bytesRecorded, WaveFormat waveFormat)
        {
            if (buffer == null || waveFormat == null || bytesRecorded <= 0)
            {
                return;
            }

            var bytesPerSample = waveFormat.BitsPerSample / 8;
            if (bytesPerSample <= 0)
            {
                bytesPerSample = waveFormat.Encoding == WaveFormatEncoding.IeeeFloat ? 4 : 2;
            }

            var limit = Math.Min(bytesRecorded, buffer.Length);
            var currentPeak = 0d;
            for (var offset = 0; offset + bytesPerSample <= limit; offset += bytesPerSample)
            {
                var level = ReadSampleLevel(buffer, offset, bytesPerSample, waveFormat);
                if (level <= 0)
                {
                    continue;
                }

                if (level > currentPeak)
                {
                    currentPeak = level;
                }

                if (level > MaxSampleLevel)
                {
                    MaxSampleLevel = level;
                }

                if (level >= AudibleSampleThreshold)
                {
                    AudibleSampleCount++;
                }
            }

            if (currentPeak > CurrentSampleLevel)
            {
                CurrentSampleLevel = currentPeak;
            }
        }

        private static double ReadSampleLevel(byte[] buffer, int offset, int bytesPerSample, WaveFormat waveFormat)
        {
            var encoding = waveFormat.Encoding;
            if (waveFormat is WaveFormatExtensible extensible)
            {
                if (extensible.SubFormat == IeeeFloatSubFormat)
                {
                    return bytesPerSample == 4 ? ReadFloatLevel(buffer, offset) : 0d;
                }

                if (extensible.SubFormat == PcmSubFormat)
                {
                    return ReadPcmLevel(buffer, offset, bytesPerSample);
                }
            }

            if (encoding == WaveFormatEncoding.IeeeFloat)
            {
                return bytesPerSample == 4 ? ReadFloatLevel(buffer, offset) : 0d;
            }

            return ReadPcmLevel(buffer, offset, bytesPerSample);
        }

        private static double ReadFloatLevel(byte[] buffer, int offset)
        {
            var sample = BitConverter.ToSingle(buffer, offset);
            return !double.IsNaN(sample) && !double.IsInfinity(sample)
                ? Math.Abs(sample)
                : 0d;
        }

        private static double ReadPcmLevel(byte[] buffer, int offset, int bytesPerSample)
        {
            switch (bytesPerSample)
            {
                case 1:
                    return Math.Abs((buffer[offset] - 128) / 128d);
                case 2:
                    return Math.Abs(BitConverter.ToInt16(buffer, offset) / 32768d);
                case 3:
                    var sample24 = buffer[offset] | (buffer[offset + 1] << 8) | (buffer[offset + 2] << 16);
                    if ((sample24 & 0x800000) != 0)
                    {
                        sample24 |= unchecked((int)0xFF000000);
                    }

                    return Math.Abs(sample24 / 8388608d);
                case 4:
                    return Math.Abs(BitConverter.ToInt32(buffer, offset) / 2147483648d);
                default:
                    return 0d;
            }
        }

        private static long GetFileSize(string filePath)
        {
            try
            {
                return !string.IsNullOrWhiteSpace(filePath) && File.Exists(filePath) ? new FileInfo(filePath).Length : 0L;
            }
            catch
            {
                return 0L;
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

    internal struct AudioLevelStats
    {
        public AudioLevelStats(long bytesRecorded, double maxSampleLevel, long audibleSampleCount, double currentSampleLevel = 0d)
        {
            BytesRecorded = bytesRecorded;
            MaxSampleLevel = maxSampleLevel;
            AudibleSampleCount = audibleSampleCount;
            CurrentSampleLevel = currentSampleLevel;
        }

        public long BytesRecorded { get; }

        public double MaxSampleLevel { get; }

        public long AudibleSampleCount { get; }

        public double CurrentSampleLevel { get; }

        public bool HasAudibleAudio => BytesRecorded > 44
            && AudibleSampleCount >= WasapiLoopbackAudioRecorder.MinAudibleSampleCount
            && MaxSampleLevel >= WasapiLoopbackAudioRecorder.AudibleSampleThreshold;
    }
}
