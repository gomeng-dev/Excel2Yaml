using System;
using System.Diagnostics;
using System.IO;

namespace ExcelToJsonAddin.Logging
{
    public class SimpleLogger : ISimpleLogger
    {
        private readonly string _categoryName;
        private static readonly string LogDirectory = 
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), 
                       "ExcelToJsonAddin", "Logs");
        
        public SimpleLogger(string categoryName)
        {
            _categoryName = categoryName;
            
            try
            {
                if (!Directory.Exists(LogDirectory))
                {
                    Directory.CreateDirectory(LogDirectory);
                }
            }
            catch
            {
                // 디렉토리 생성 실패 시 무시
            }
        }
        
        public void Debug(string message)
        {
            Log("DEBUG", message);
        }
        
        public void Debug(string messageTemplate, params object[] args)
        {
            try
            {
                if (args != null && args.Length > 0)
                {
                    Log("DEBUG", string.Format(messageTemplate, args));
                }
                else
                {
                    Log("DEBUG", messageTemplate);
                }
            }
            catch (Exception ex)
            {
                // 포맷팅 오류가 발생하면 원본 문자열과 에러 메시지를 로깅
                Log("DEBUG", "로깅 오류: " + messageTemplate + " - 예외: " + ex.Message);
            }
        }
        
        public void Information(string message)
        {
            Log("INFO", message);
        }
        
        public void Information(string messageTemplate, params object[] args)
        {
            try
            {
                if (args != null && args.Length > 0)
                {
                    Log("INFO", string.Format(messageTemplate, args));
                }
                else
                {
                    Log("INFO", messageTemplate);
                }
            }
            catch (Exception ex)
            {
                // 포맷팅 오류가 발생하면 원본 문자열과 에러 메시지를 로깅
                Log("INFO", "로깅 오류: " + messageTemplate + " - 예외: " + ex.Message);
            }
        }
        
        public void Warning(string message)
        {
            Log("WARN", message);
        }
        
        public void Warning(string messageTemplate, params object[] args)
        {
            try
            {
                if (args != null && args.Length > 0)
                {
                    Log("WARN", string.Format(messageTemplate, args));
                }
                else
                {
                    Log("WARN", messageTemplate);
                }
            }
            catch (Exception ex)
            {
                // 포맷팅 오류가 발생하면 원본 문자열과 에러 메시지를 로깅
                Log("WARN", "로깅 오류: " + messageTemplate + " - 예외: " + ex.Message);
            }
        }
        
        public void Error(string message)
        {
            Log("ERROR", message);
        }
        
        public void Error(string messageTemplate, params object[] args)
        {
            try
            {
                if (args != null && args.Length > 0)
                {
                    Log("ERROR", string.Format(messageTemplate, args));
                }
                else
                {
                    Log("ERROR", messageTemplate);
                }
            }
            catch (Exception ex)
            {
                // 포맷팅 오류가 발생하면 원본 문자열과 에러 메시지를 로깅
                Log("ERROR", "로깅 오류: " + messageTemplate + " - 예외: " + ex.Message);
            }
        }
        
        public void Error(Exception exception, string message)
        {
            Log("ERROR", $"{message} Exception: {exception}");
        }
        
        public void Error(Exception exception, string messageTemplate, params object[] args)
        {
            Log("ERROR", $"{string.Format(messageTemplate, args)} Exception: {exception}");
        }
        
        private void Log(string level, string message)
        {
            var logMessage = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} [{level}] [{_categoryName}] {message}";
            
            // 디버그 콘솔에 출력
            System.Diagnostics.Debug.WriteLine(logMessage);
            
            try
            {
                // 파일에도 로깅 (하루 단위로 파일 생성)
                var logFile = Path.Combine(LogDirectory, $"log_{DateTime.Now:yyyyMMdd}.txt");
                File.AppendAllText(logFile, logMessage + Environment.NewLine);
            }
            catch
            {
                // 파일 로깅 실패 시 무시
            }
        }
    }
    
    public static class SimpleLoggerFactory
    {
        public static ISimpleLogger CreateLogger<T>()
        {
            return new SimpleLogger(typeof(T).Name);
        }
        
        public static ISimpleLogger CreateLogger(string categoryName)
        {
            return new SimpleLogger(categoryName);
        }
    }
}
