using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace MyTools.Services
{
    public class BenchmarkResult
    {
        public string Name { get; set; }
        public string Score { get; set; }
        public string Detail { get; set; }
        public TimeSpan Elapsed { get; set; }
    }

    public static class BenchmarkService
    {
        // ======================== CPU Benchmark ========================
        public static async Task<BenchmarkResult> RunCpuSingleThreadAsync(IProgress<string> progress, CancellationToken ct)
        {
            progress?.Report("CPU 单线程：运行中…");
            return await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                var sw = Stopwatch.StartNew();
                int iterations = 200_000;
                var data = new byte[1024];
                new Random(42).NextBytes(data);

                using (var sha = new SHA256CryptoServiceProvider())
                {
                    for (int i = 0; i < iterations; i++)
                    {
                        ct.ThrowIfCancellationRequested();
                        data = sha.ComputeHash(data);
                    }
                }
                sw.Stop();

                double opsPerSec = iterations / sw.Elapsed.TotalSeconds;
                return new BenchmarkResult
                {
                    Name = "CPU 单线程 (SHA-256)",
                    Score = $"{opsPerSec:N0} ops/s",
                    Detail = $"{iterations:N0} 次 SHA-256 · 耗时 {sw.Elapsed.TotalSeconds:0.##} 秒",
                    Elapsed = sw.Elapsed
                };
            }, ct).ConfigureAwait(false);
        }

        public static async Task<BenchmarkResult> RunCpuMultiThreadAsync(IProgress<string> progress, CancellationToken ct)
        {
            int threadCount = Environment.ProcessorCount;
            progress?.Report($"CPU 多线程 ({threadCount}T)：运行中…");

            return await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                int iterationsPerThread = 200_000;
                var sw = Stopwatch.StartNew();
                long totalOps = 0;

                var tasks = new Task[threadCount];
                for (int t = 0; t < threadCount; t++)
                {
                    int seed = t;
                    tasks[t] = Task.Run(() =>
                    {
                        var data = new byte[1024];
                        new Random(seed).NextBytes(data);
                        using (var sha = new SHA256CryptoServiceProvider())
                        {
                            for (int i = 0; i < iterationsPerThread; i++)
                            {
                                ct.ThrowIfCancellationRequested();
                                data = sha.ComputeHash(data);
                            }
                        }
                        Interlocked.Add(ref totalOps, iterationsPerThread);
                    }, ct);
                }
                Task.WaitAll(tasks);
                sw.Stop();

                double opsPerSec = totalOps / sw.Elapsed.TotalSeconds;
                return new BenchmarkResult
                {
                    Name = $"CPU 多线程 ({threadCount}T SHA-256)",
                    Score = $"{opsPerSec:N0} ops/s",
                    Detail = $"{totalOps:N0} 次 SHA-256 · {threadCount} 线程 · 耗时 {sw.Elapsed.TotalSeconds:0.##} 秒",
                    Elapsed = sw.Elapsed
                };
            }, ct).ConfigureAwait(false);
        }

        // ======================== Memory Benchmark ========================
        public static async Task<BenchmarkResult> RunMemoryBandwidthAsync(IProgress<string> progress, CancellationToken ct)
        {
            progress?.Report("内存带宽：运行中…");
            return await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                // Sequential read/write over 256 MB buffer
                int sizeMb = 256;
                int sizeBytes = sizeMb * 1024 * 1024;
                var buffer = new long[sizeBytes / 8];
                var sw = Stopwatch.StartNew();

                // Write pass
                for (int i = 0; i < buffer.Length; i++)
                {
                    buffer[i] = (long)i * 7 + 13;
                }
                ct.ThrowIfCancellationRequested();

                // Read pass (sum to prevent optimization)
                long sum = 0;
                for (int i = 0; i < buffer.Length; i++)
                {
                    sum += buffer[i];
                }
                sw.Stop();

                // Prevent dead code elimination
                GC.KeepAlive(sum);

                // 2 passes (write + read) each touching sizeMb
                double totalGb = (2.0 * sizeMb) / 1024.0;
                double bandwidthGBs = totalGb / sw.Elapsed.TotalSeconds;
                var result = new BenchmarkResult
                {
                    Name = "内存带宽 (顺序读写)",
                    Score = $"{bandwidthGBs:0.##} GB/s",
                    Detail = $"{sizeMb * 2} MB 读写 · 耗时 {sw.Elapsed.TotalMilliseconds:0} ms",
                    Elapsed = sw.Elapsed
                };

                // Eagerly release the 256 MB buffer
                buffer = null;
                GC.Collect(2, GCCollectionMode.Optimized, false);
                return result;
            }, ct).ConfigureAwait(false);
        }

        public static async Task<BenchmarkResult> RunMemoryLatencyAsync(IProgress<string> progress, CancellationToken ct)
        {
            progress?.Report("内存延迟：运行中…");
            return await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                // Pointer-chasing latency test (16 MB working set)
                int count = 4 * 1024 * 1024; // 4M entries × 4 bytes = 16 MB
                var chain = new int[count];
                // Build random permutation for pointer chasing
                var rng = new Random(12345);
                for (int i = count - 1; i > 0; i--)
                {
                    int j = rng.Next(i + 1);
                    chain[i] = chain[j];
                    chain[j] = i;
                }

                ct.ThrowIfCancellationRequested();
                int steps = 10_000_000;
                int idx = 0;
                var sw = Stopwatch.StartNew();
                for (int i = 0; i < steps; i++)
                {
                    idx = chain[idx];
                }
                sw.Stop();
                GC.KeepAlive(idx);

                double nsPerAccess = sw.Elapsed.TotalMilliseconds * 1_000_000.0 / steps;
                return new BenchmarkResult
                {
                    Name = "内存延迟 (随机访问)",
                    Score = $"{nsPerAccess:0.#} ns",
                    Detail = $"{steps / 1_000_000}M 次随机跳转 · 16 MB 工作集 · 耗时 {sw.Elapsed.TotalMilliseconds:0} ms",
                    Elapsed = sw.Elapsed
                };
            }, ct).ConfigureAwait(false);
        }

        // ======================== GPU Info (lightweight) ========================
        public static async Task<BenchmarkResult> RunGpuInfoAsync(IProgress<string> progress, CancellationToken ct)
        {
            progress?.Report("GPU 信息：读取中…");
            return await Task.Run(() =>
            {
                ct.ThrowIfCancellationRequested();
                var sw = Stopwatch.StartNew();
                var gpus = new List<string>();
                try
                {
                    using (var searcher = new System.Management.ManagementObjectSearcher("SELECT Name, AdapterRAM, DriverVersion FROM Win32_VideoController"))
                    using (var results = searcher.Get())
                    foreach (System.Management.ManagementObject mo in results)
                    {
                        try
                        {
                            var name = mo["Name"]?.ToString() ?? "Unknown GPU";
                            var ramRaw = mo["AdapterRAM"];
                            string ram = "";
                            if (ramRaw != null && ulong.TryParse(ramRaw.ToString(), out var bytes))
                            {
                                ram = $" · {bytes / 1024.0 / 1024.0 / 1024.0:0.##} GB";
                            }
                            var driver = mo["DriverVersion"]?.ToString();
                            gpus.Add($"{name}{ram} (驱动 {driver})");
                        }
                        finally { mo.Dispose(); }
                    }
                }
                catch (Exception ex)
                {
                    gpus.Add($"读取失败：{ex.Message}");
                }
                sw.Stop();

                return new BenchmarkResult
                {
                    Name = "GPU 信息",
                    Score = gpus.FirstOrDefault() ?? "无 GPU",
                    Detail = string.Join("\n", gpus),
                    Elapsed = sw.Elapsed
                };
            }, ct).ConfigureAwait(false);
        }

        // ======================== Run All ========================
        public static async Task<List<BenchmarkResult>> RunAllAsync(IProgress<string> progress, CancellationToken ct)
        {
            var results = new List<BenchmarkResult>();

            results.Add(await RunCpuSingleThreadAsync(progress, ct));
            results.Add(await RunCpuMultiThreadAsync(progress, ct));
            results.Add(await RunMemoryBandwidthAsync(progress, ct));
            results.Add(await RunMemoryLatencyAsync(progress, ct));
            results.Add(await RunGpuInfoAsync(progress, ct));

            progress?.Report("全部测试完成。");
            return results;
        }
    }
}
