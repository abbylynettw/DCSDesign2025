using log4net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyOffice.LogHelper
{
    public static class LogExtensions
    {
        public static void LogDebug(this ILog log, string message)
        {
            if (log.IsDebugEnabled)
                log.Debug(message);
        }

        public static void LogDebug(this ILog log, string message, Exception exception)
        {
            if (log.IsDebugEnabled)
                log.Debug(message, exception);
        }

        public static void LogInfo(this ILog log, string message)
        {
            if (log.IsInfoEnabled)
                log.Info(message);
        }

        public static void LogInfo(this ILog log, string message, Exception exception)
        {
            if (log.IsInfoEnabled)
                log.Info(message, exception);
        }

        public static void LogWarn(this ILog log, string message)
        {
            if (log.IsWarnEnabled)
                log.Warn(message);
        }

        public static void LogWarn(this ILog log, string message, Exception exception)
        {
            if (log.IsWarnEnabled)
                log.Warn(message, exception);
        }

        public static void LogError(this ILog log, string message)
        {
            if (log.IsErrorEnabled)
                log.Error(message);
        }

        public static void LogError(this ILog log, string message, Exception exception)
        {
            if (log.IsErrorEnabled)
                log.Error(message, exception);
        }

        public static void LogFatal(this ILog log, string message)
        {
            if (log.IsFatalEnabled)
                log.Fatal(message);
        }

        public static void LogFatal(this ILog log, string message, Exception exception)
        {
            if (log.IsFatalEnabled)
                log.Fatal(message, exception);
        }
    }
}
