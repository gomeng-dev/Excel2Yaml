using ExcelToJsonAddin.Logging;

namespace ExcelToJsonCSharp.Extensions
{
    public static class LoggingExtensions
    {
        public static void Debug(this ISimpleLogger logger, string message, params object[] args)
        {
            logger.Debug(message, args);
        }

        public static void Trace(this ISimpleLogger logger, string message)
        {
            logger.Debug(message);
        }
    }
}
