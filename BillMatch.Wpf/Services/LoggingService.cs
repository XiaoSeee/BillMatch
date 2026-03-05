using System;
using System.IO;
using NLog;
using NLog.Config;
using NLog.Targets;

namespace BillMatch.Wpf.Services
{
    /// <summary>
    /// 日志服务 - 使用NLog实现
    /// </summary>
    public interface ILoggingService
    {
        void Debug(string message);
        void Debug(string message, Exception exception);
        void Info(string message);
        void Info(string message, Exception exception);
        void Warn(string message);
        void Warn(string message, Exception exception);
        void Error(string message);
        void Error(string message, Exception exception);
        void Fatal(string message);
        void Fatal(string message, Exception exception);
        string GetLogFilePath();
    }

    public class LoggingService : ILoggingService
    {
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();
        private static readonly string LogFilePath;

        static LoggingService()
        {
            // 初始化NLog配置
            var config = new LoggingConfiguration();

            // 日志文件路径: %LocalAppData%\BillMatch\logs\BillMatch_{日期}.log
            var logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "BillMatch", "logs");
            
            if (!Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }

            LogFilePath = Path.Combine(logDir, "BillMatch_${date:format=yyyyMMdd}.log");

            // 文件目标
            var fileTarget = new FileTarget("logfile")
            {
                FileName = LogFilePath,
                Layout = "${longdate} | ${level:uppercase=true} | ${message} ${exception:format=ToString}",
                ArchiveFileName = Path.Combine(logDir, "BillMatch_{#}.log"),
                ArchiveNumbering = ArchiveNumberingMode.Date,
                ArchiveDateFormat = "yyyyMMdd",
                MaxArchiveFiles = 30,
                KeepFileOpen = false,
                ConcurrentWrites = true
            };

            // 控制台目标 (用于调试)
            var consoleTarget = new ConsoleTarget("logconsole")
            {
                Layout = "${longdate} | ${level:uppercase=true} | ${message} ${exception:format=ToString}"
            };

            // 添加规则
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, fileTarget);
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, consoleTarget);

            // 应用配置
            LogManager.Configuration = config;
        }

        public void Debug(string message)
        {
            Logger.Debug(message);
        }

        public void Debug(string message, Exception exception)
        {
            Logger.Debug(exception, message);
        }

        public void Info(string message)
        {
            Logger.Info(message);
        }

        public void Info(string message, Exception exception)
        {
            Logger.Info(exception, message);
        }

        public void Warn(string message)
        {
            Logger.Warn(message);
        }

        public void Warn(string message, Exception exception)
        {
            Logger.Warn(exception, message);
        }

        public void Error(string message)
        {
            Logger.Error(message);
        }

        public void Error(string message, Exception exception)
        {
            Logger.Error(exception, message);
        }

        public void Fatal(string message)
        {
            Logger.Fatal(message);
        }

        public void Fatal(string message, Exception exception)
        {
            Logger.Fatal(exception, message);
        }

        public string GetLogFilePath()
        {
            return LogFilePath?.Replace("${date:format=yyyyMMdd}", DateTime.Now.ToString("yyyyMMdd")) ?? string.Empty;
        }
    }
}
