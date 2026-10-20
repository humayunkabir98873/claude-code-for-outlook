using System;
using System.IO;

namespace ClaudeMailAssistant
{
    // Appends to %LOCALAPPDATA%\ClaudeMailAssistant\log.txt. Never throws: logging must not take Outlook down.
    public static class Log
    {
        // CLAUDE_MAIL_DATA_DIR lets the test harness keep its own browser profile next to a running Outlook.
        public static readonly string DataDir = Environment.GetEnvironmentVariable("CLAUDE_MAIL_DATA_DIR")
            ?? Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ClaudeMailAssistant");

        static readonly object Gate = new object();

        public static void Write(string message)
        {
            try
            {
                lock (Gate)
                {
                    Directory.CreateDirectory(DataDir);
                    var file = Path.Combine(DataDir, "log.txt");
                    if (File.Exists(file) && new FileInfo(file).Length > 1_000_000) File.Delete(file);
                    File.AppendAllText(file, $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} {message}{Environment.NewLine}");
                }
            }
            catch
            {
                // nowhere left to report it
            }
        }

        public static void Error(string where, Exception ex) => Write($"ERROR in {where}: {ex}");
    }
}
