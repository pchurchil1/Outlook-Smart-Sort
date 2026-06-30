using System;
using System.IO;

namespace OutlookClassifierAddIn5.Services
{
    public static class AppLogger
    {
        private static readonly object Gate = new object();

        public static string DataDirectory
        {
            get
            {
                return Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "OutlookClassifier");
            }
        }

        public static string LogDirectory
        {
            get { return Path.Combine(DataDirectory, "logs"); }
        }

        public static void Info(string message)
        {
            Write("INFO", message, null);
        }

        public static void Warn(string message)
        {
            Write("WARN", message, null);
        }

        public static void Error(string message)
        {
            Write("ERROR", message, null);
        }

        public static void Error(Exception ex, string message)
        {
            Write("ERROR", message, ex);
        }

        private static void Write(string level, string message, Exception ex)
        {
            try
            {
                Directory.CreateDirectory(LogDirectory);
                var path = Path.Combine(LogDirectory, "smartsort-" + DateTime.Now.ToString("yyyy-MM-dd") + ".log");
                var line = DateTime.UtcNow.ToString("o") + " [" + level + "] " + (message ?? string.Empty);
                if (ex != null)
                {
                    line += Environment.NewLine + ex;
                }

                lock (Gate)
                {
                    File.AppendAllText(path, line + Environment.NewLine);
                }
            }
            catch
            {
                // Logging must never interrupt Outlook.
            }
        }
    }
}
