using ExcelToYamlAddin.Application.Interfaces;
using System;
using System.Diagnostics;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToYamlAddin.Application.PostProcessing
{
    /// <summary>
    /// 모든 후처리기의 기본 추상 클래스입니다.
    /// </summary>
    public abstract class PostProcessorBase : IPostProcessor
    {
        /// <summary>
        /// 처리 우선순위를 가져옵니다.
        /// </summary>
        public abstract int Priority { get; }

        /// <summary>
        /// 이 프로세서가 주어진 컨텍스트를 처리할 수 있는지 확인합니다.
        /// </summary>
        public abstract bool CanProcess(ProcessingContext context);

        /// <summary>
        /// 후처리를 비동기적으로 수행합니다.
        /// </summary>
        public async Task<ProcessingResult> ProcessAsync(string input, ProcessingContext context, CancellationToken cancellationToken = default)
        {
            var stopwatch = Stopwatch.StartNew();
            var processorName = GetType().Name;

            try
            {
                cancellationToken.ThrowIfCancellationRequested();

                // 전처리 검증
                if (string.IsNullOrEmpty(input))
                {
                    return ProcessingResult.CreateSuccess(input, processorName);
                }

                // 실제 처리 수행
                var result = await ProcessCoreAsync(input, context, cancellationToken);

                stopwatch.Stop();
                return new ProcessingResult
                {
                    Output = result,
                    Success = true,
                    ProcessingTime = stopwatch.Elapsed,
                    ProcessorName = processorName
                };
            }
            catch (OperationCanceledException)
            {
                throw; // 취소는 그대로 전파
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                return new ProcessingResult
                {
                    Success = false,
                    ErrorMessage = $"{processorName} 처리 중 오류 발생: {ex.Message}",
                    ProcessingTime = stopwatch.Elapsed,
                    ProcessorName = processorName
                };
            }
        }

        /// <summary>
        /// 파생 클래스에서 실제 처리 로직을 구현합니다.
        /// </summary>
        /// <param name="input">입력 문자열</param>
        /// <param name="context">처리 컨텍스트</param>
        /// <param name="cancellationToken">취소 토큰</param>
        /// <returns>처리된 문자열</returns>
        protected abstract Task<string> ProcessCoreAsync(string input, ProcessingContext context, CancellationToken cancellationToken);
    }
}