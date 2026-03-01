using System;
using System.Collections.Generic;
using System.Text;

namespace WizGrep.Helpers
{
    public class LoggerHelper
    {
        private static LoggerHelper? _instance;

        // get temp file path for logging (e.g., C:\Users\Username\AppData\Local\Temp\WizGrep.log)
        private static readonly string
            logFilePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "WizGrep.log");

        private LoggerHelper()
        {
        }

        public static LoggerHelper Instance
        {
            get
            {
                _instance ??= new LoggerHelper();

                return _instance;
            }
        }

        public void LogInfo(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return;
            }

            string logEntry = $"{DateTime.Now} [INFO] {message}";
            System.IO.File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
        }

        public void LogError(string message)
        {
            if (string.IsNullOrEmpty(message))
            {
                return;
            }

            string logEntry = $"{DateTime.Now} [ERROR] {message}";
            System.IO.File.AppendAllText(logFilePath, logEntry + Environment.NewLine);
        }
    }
}
