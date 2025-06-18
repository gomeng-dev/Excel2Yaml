using ExcelToYamlAddin.Application.PostProcessing;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToYamlAddin.Application.Interfaces
{
    /// <summary>
    /// 후처리 파이프라인 인터페이스입니다.
    /// </summary>
    public interface IProcessingPipeline
    {
        /// <summary>
        /// 파이프라인을 통해 입력을 처리합니다.
        /// </summary>
        /// <param name="input">입력 문자열</param>
        /// <param name="context">처리 컨텍스트</param>
        /// <param name="progress">진행률 보고</param>
        /// <param name="cancellationToken">취소 토큰</param>
        /// <returns>처리 결과</returns>
        Task<ProcessingResult> ProcessAsync(
            string input, 
            ProcessingContext context,
            IProgress<ProcessingProgress> progress = null,
            CancellationToken cancellationToken = default);
    }
}