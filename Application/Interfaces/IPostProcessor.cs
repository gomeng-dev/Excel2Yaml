using ExcelToYamlAddin.Application.PostProcessing;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToYamlAddin.Application.Interfaces
{
    /// <summary>
    /// 후처리기의 기본 인터페이스입니다.
    /// </summary>
    public interface IPostProcessor
    {
        /// <summary>
        /// 처리 우선순위를 가져옵니다. 낮을수록 먼저 실행됩니다.
        /// </summary>
        int Priority { get; }

        /// <summary>
        /// 이 프로세서가 주어진 컨텍스트를 처리할 수 있는지 확인합니다.
        /// </summary>
        /// <param name="context">처리 컨텍스트</param>
        /// <returns>처리 가능 여부</returns>
        bool CanProcess(ProcessingContext context);

        /// <summary>
        /// 후처리를 비동기적으로 수행합니다.
        /// </summary>
        /// <param name="input">입력 문자열</param>
        /// <param name="context">처리 컨텍스트</param>
        /// <param name="cancellationToken">취소 토큰</param>
        /// <returns>처리 결과</returns>
        Task<ProcessingResult> ProcessAsync(string input, ProcessingContext context, CancellationToken cancellationToken = default);
    }
}