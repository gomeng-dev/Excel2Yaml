using System;

namespace ExcelToJsonAddin.Logging
{
    public interface ISimpleLogger
    {
        void Debug(string message);
        void Debug(string messageTemplate, params object[] args);
        void Information(string message);
        void Information(string messageTemplate, params object[] args);
        void Warning(string message);
        void Warning(string messageTemplate, params object[] args);
        void Error(string message);
        void Error(string messageTemplate, params object[] args);
        void Error(Exception exception, string message);
        void Error(Exception exception, string messageTemplate, params object[] args);
    }
}
