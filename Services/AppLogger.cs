using System;
using System.IO;
using System.Text;

namespace GlobeMapper.Services
{
    /// <summary>
    /// 단일 변환 실행 단위로 log/log_yyyyMMdd_HHmmss.txt 파일에 기록.
    /// </summary>
    public sealed class AppLogger : IDisposable
    {
        private readonly StreamWriter _writer;
        private bool _disposed;

        public string LogPath { get; }

        public AppLogger()
        {
            var logDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "log");
            Directory.CreateDirectory(logDir);

            LogPath = Path.Combine(logDir, $"log_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
            _writer = new StreamWriter(LogPath, append: false, encoding: new UTF8Encoding(false))
            {
                AutoFlush = true,
            };
            WriteLine($"=== GlobeMapper 변환 시작 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
        }

        public void WriteLine(string message)
        {
            if (_disposed) return;
            _writer.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] {message}");
        }

        public void WriteLines(System.Collections.Generic.IEnumerable<string> lines, string header = null)
        {
            if (header != null) WriteLine(header);
            foreach (var line in lines)
                _writer.WriteLine($"  {line}");
        }

        public void WriteException(Exception ex, string context = null)
        {
            if (context != null) WriteLine($"[ERROR] {context}");
            var e = ex;
            while (e != null)
            {
                _writer.WriteLine($"  {e.GetType().FullName}: {e.Message}");
                if (e.StackTrace != null)
                {
                    foreach (var line in e.StackTrace.Split('\n'))
                        _writer.WriteLine($"    {line.TrimEnd()}");
                }
                e = e.InnerException;
                if (e != null) _writer.WriteLine("  → 원인:");
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            WriteLine($"=== 변환 종료 {DateTime.Now:yyyy-MM-dd HH:mm:ss} ===");
            _writer.Dispose();
            _disposed = true;
        }
    }
}
