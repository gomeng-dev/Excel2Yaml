using Microsoft.Extensions.Logging;

namespace ExcelToJsonCSharp.Extensions
{
    public static class LoggingExtensions
    {
        public static void Debug(this ILogger logger, string message, params object[] args)
        {
            logger.LogDebug(message, args);
        }

        public static void Trace(this ILogger logger, string message)
        {
            logger.LogTrace(message);
        }
    }
}
