using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace WizGrep.Helpers
{
    public class LoggerHelper
    {
        private static readonly Lazy<LoggerHelper> _instance = new(() => new LoggerHelper());

        // get temp file path for logging (e.g., C:\Users\Username\AppData\Local\Temp\WizGrep.log)
        private static readonly string
            logFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "WizGrep.log");

        private readonly Lock _lock = new Lock();

        private LoggerHelper()
        {
        }

        public static LoggerHelper Instance => _instance.Value;

        public void LogInfo(string message)
        {
            WriteLog("INFO", message);
        }

        public void LogError(string message)
        {
            WriteLog("ERROR", message);
        }

        private void WriteLog(string level, string message)
        {
            if(string.IsNullOrEmpty(message))
            {
                return;
            }

            var logEntry = $"{DateTime.Now} [{level}] {message}";
            lock (_lock)
            {
                System.IO.File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
            }
        }
    }
}
