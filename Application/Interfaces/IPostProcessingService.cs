using System.Collections.Generic;

namespace ExcelToYamlAddin.Application.Interfaces
{
    /// <summary>
    /// 후처리 서비스 인터페이스
    /// </summary>
    public interface IPostProcessingService
    {
        /// <summary>
        /// 서비스 이름
        /// </summary>
        string ServiceName { get; }

        /// <summary>
        /// 처리 우선순위 (낮은 값이 먼저 실행됨)
        /// </summary>
        int Priority { get; }

        /// <summary>
        /// 후처리를 수행합니다.
        /// </summary>
        /// <param name="content">처리할 컨텐츠</param>
        /// <param name="options">처리 옵션</param>
        /// <returns>처리된 컨텐츠</returns>
        string Process(string content, Dictionary<string, object> options);

        /// <summary>
        /// 주어진 컨텐츠 타입을 처리할 수 있는지 확인합니다.
        /// </summary>
        /// <param name="contentType">컨텐츠 타입</param>
        /// <returns>처리 가능 여부</returns>
        bool CanProcess(string contentType);
    }
}