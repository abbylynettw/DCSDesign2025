using log4net.Appender;
using log4net.Core;
using log4net.Layout;
using log4net.Repository.Hierarchy;
using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace MyOffice.LogHelper
{
    public static class LogManager
    {
        private static bool _isInitialized = false;
        private static readonly object _syncLock = new object();

        // 初始化log4net
        public static void Initialize()
        {
            if (_isInitialized)
                return;

            lock (_syncLock)
            {
                if (_isInitialized)
                    return;

                try
                {
                    // 代码配置log4net
                    ConfigureLog4Net();

                    _isInitialized = true;
                    GetLogger(typeof(LogManager)).Info("Log4net 初始化成功");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"初始化日志系统失败: {ex.Message}");
                }
            }
        }

        // 使用代码配置log4net
        private static void ConfigureLog4Net()
        {
            // 创建一个新的log4net配置
            Hierarchy hierarchy = (Hierarchy)log4net.LogManager.GetRepository();
            hierarchy.Root.RemoveAllAppenders(); // 清除所有现有appender

            // 创建PatternLayout
            PatternLayout patternLayout = new PatternLayout();
            patternLayout.ConversionPattern = "%date [%thread] %-5level %logger - %message%newline";
            patternLayout.ActivateOptions();

            // 创建标准日志文件输出
            string logDirectory = EnsureLogDirectory();
            RollingFileAppender rollingAppender = new RollingFileAppender();
            rollingAppender.File = Path.Combine(logDirectory, "application.log");
            rollingAppender.AppendToFile = true;
            rollingAppender.RollingStyle = RollingFileAppender.RollingMode.Composite;
            rollingAppender.MaxSizeRollBackups = 10;
            rollingAppender.MaximumFileSize = "5MB";
            rollingAppender.DatePattern = "yyyyMMdd";
            rollingAppender.PreserveLogFileNameExtension = true;
            rollingAppender.StaticLogFileName = false;
            rollingAppender.Layout = patternLayout;
            rollingAppender.ActivateOptions();

            // 创建错误日志文件输出
            PatternLayout errorLayout = new PatternLayout();
            errorLayout.ConversionPattern = "%date [%thread] %-5level %logger - %message%newline%exception";
            errorLayout.ActivateOptions();

            RollingFileAppender errorAppender = new RollingFileAppender();
            errorAppender.File = Path.Combine(logDirectory, "error.log");
            errorAppender.AppendToFile = true;
            errorAppender.RollingStyle = RollingFileAppender.RollingMode.Date;
            errorAppender.DatePattern = "yyyyMMdd";
            errorAppender.PreserveLogFileNameExtension = true;
            errorAppender.StaticLogFileName = false;
            errorAppender.Layout = errorLayout;

            // 添加过滤器，只记录ERROR和FATAL级别
            log4net.Filter.LevelRangeFilter filter = new log4net.Filter.LevelRangeFilter();
            filter.LevelMin = Level.Error;
            filter.LevelMax = Level.Fatal;
            filter.ActivateOptions();
            errorAppender.AddFilter(filter);

            errorAppender.ActivateOptions();

            // 将appender添加到root logger
            hierarchy.Root.AddAppender(rollingAppender);
            hierarchy.Root.AddAppender(errorAppender);

            // 设置默认日志级别
            hierarchy.Root.Level = Level.Info;

            // 激活配置
            hierarchy.Configured = true;
        }

        // 获取指定类型的日志记录器
        public static ILog GetLogger(Type type)
        {
            if (!_isInitialized)
                Initialize();

            return log4net.LogManager.GetLogger(type);
        }

        // 获取泛型类型的日志记录器
        public static ILog GetLogger<T>()
        {
            return GetLogger(typeof(T));
        }

        // 获取指定名称的日志记录器
        public static ILog GetLogger(string name)
        {
            if (!_isInitialized)
                Initialize();

            return log4net.LogManager.GetLogger(name);
        }

        // 确保日志目录存在
        public static string EnsureLogDirectory()
        {
            string logDirectory = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                "logs");

            if (!Directory.Exists(logDirectory))
                Directory.CreateDirectory(logDirectory);

            return logDirectory;
        }
    }
}
