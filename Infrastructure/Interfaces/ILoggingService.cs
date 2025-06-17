using System;

namespace ExcelToYamlAddin.Infrastructure.Interfaces
{
    /// <summary>
    /// 로깅 서비스 인터페이스
    /// </summary>
    public interface ILoggingService
    {
        /// <summary>
        /// 디버그 메시지를 기록합니다.
        /// </summary>
        /// <param name="message">메시지</param>
        /// <param name="args">메시지 인자</param>
        void Debug(string message, params object[] args);

        /// <summary>
        /// 정보 메시지를 기록합니다.
        /// </summary>
        /// <param name="message">메시지</param>
        /// <param name="args">메시지 인자</param>
        void Information(string message, params object[] args);

        /// <summary>
        /// 경고 메시지를 기록합니다.
        /// </summary>
        /// <param name="message">메시지</param>
        /// <param name="args">메시지 인자</param>
        void Warning(string message, params object[] args);

        /// <summary>
        /// 오류 메시지를 기록합니다.
        /// </summary>
        /// <param name="message">메시지</param>
        /// <param name="args">메시지 인자</param>
        void Error(string message, params object[] args);

        /// <summary>
        /// 예외와 함께 오류 메시지를 기록합니다.
        /// </summary>
        /// <param name="exception">예외</param>
        /// <param name="message">메시지</param>
        /// <param name="args">메시지 인자</param>
        void Error(Exception exception, string message, params object[] args);

        /// <summary>
        /// 치명적 오류 메시지를 기록합니다.
        /// </summary>
        /// <param name="message">메시지</param>
        /// <param name="args">메시지 인자</param>
        void Fatal(string message, params object[] args);

        /// <summary>
        /// 예외와 함께 치명적 오류 메시지를 기록합니다.
        /// </summary>
        /// <param name="exception">예외</param>
        /// <param name="message">메시지</param>
        /// <param name="args">메시지 인자</param>
        void Fatal(Exception exception, string message, params object[] args);

        /// <summary>
        /// 구조화된 속성과 함께 로그를 기록합니다.
        /// </summary>
        /// <param name="propertyName">속성 이름</param>
        /// <param name="value">속성 값</param>
        /// <param name="destructureObjects">객체 구조 분해 여부</param>
        /// <returns>로거 인스턴스</returns>
        ILoggingService ForContext(string propertyName, object value, bool destructureObjects = false);
    }
}