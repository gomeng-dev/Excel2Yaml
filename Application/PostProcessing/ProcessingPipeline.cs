using ExcelToYamlAddin.Application.Interfaces;
using ExcelToYamlAddin.Infrastructure.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToYamlAddin.Application.PostProcessing
{
    /// <summary>
    /// 후처리 파이프라인을 구현하는 클래스입니다.
    /// 여러 후처리기를 순서대로 실행합니다.
    /// </summary>
    public class ProcessingPipeline : IProcessingPipeline
    {
        private readonly IEnumerable<IPostProcessor> _processors;
        private readonly ISimpleLogger _logger;

        /// <summary>
        /// ProcessingPipeline의 새 인스턴스를 초기화합니다.
        /// </summary>
        /// <param name="processors">실행할 후처리기들</param>
        public ProcessingPipeline(IEnumerable<IPostProcessor> processors)
        {
            _processors = processors?.OrderBy(p => p.Priority) ?? throw new ArgumentNullException(nameof(processors));
            _logger = SimpleLoggerFactory.CreateLogger<ProcessingPipeline>();
        }

        /// <summary>
        /// 파이프라인을 통해 입력을 처리합니다.
        /// </summary>
        public async Task<ProcessingResult> ProcessAsync(
            string input, 
            ProcessingContext context,
            IProgress<ProcessingProgress> progress = null,
            CancellationToken cancellationToken = default)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));

            var currentOutput = input;
            var overallStopwatch = Stopwatch.StartNew();
            var processorResults = new List<ProcessingResult>();

            _logger.Information($"후처리 파이프라인 시작 - 파일: {context.FilePath}");

            var processorList = _processors.ToList();
            var enabledProcessors = processorList.Where(p => p.CanProcess(context)).ToList();

            if (enabledProcessors.Count == 0)
            {
                _logger.Information("활성화된 후처리기가 없습니다.");
                return ProcessingResult.CreateSuccess(input, "Pipeline");
            }

            _logger.Information($"활성화된 후처리기: {enabledProcessors.Count}개");

            for (int i = 0; i < enabledProcessors.Count; i++)
            {
                var processor = enabledProcessors[i];
                cancellationToken.ThrowIfCancellationRequested();

                try
                {
                    _logger.Information($"[{i + 1}/{enabledProcessors.Count}] {processor.GetType().Name} 실행 중...");

                    // 진행률 보고
                    progress?.Report(new ProcessingProgress
                    {
                        CurrentProcessor = processor.GetType().Name,
                        ProcessorIndex = i,
                        TotalProcessors = enabledProcessors.Count,
                        Message = $"{processor.GetType().Name} 처리 중..."
                    });

                    // 처리 실행
                    var result = await processor.ProcessAsync(currentOutput, context, cancellationToken);
                    processorResults.Add(result);

                    if (!result.Success)
                    {
                        _logger.Warning($"{processor.GetType().Name} 처리 실패: {result.ErrorMessage}");
                        
                        // 실패한 경우 파이프라인 중단
                        overallStopwatch.Stop();
                        return new ProcessingResult
                        {
                            Success = false,
                            ErrorMessage = result.ErrorMessage,
                            ProcessingTime = overallStopwatch.Elapsed,
                            ProcessorName = "Pipeline",
                            Output = currentOutput // 실패 전까지의 출력 유지
                        };
                    }

                    currentOutput = result.Output;
                    _logger.Information($"{processor.GetType().Name} 완료 (소요시간: {result.ProcessingTime.TotalMilliseconds:F2}ms)");
                }
                catch (Exception ex)
                {
                    _logger.Error($"{processor.GetType().Name} 처리 중 예외 발생: {ex.Message}", ex);
                    overallStopwatch.Stop();
                    
                    return new ProcessingResult
                    {
                        Success = false,
                        ErrorMessage = $"{processor.GetType().Name} 처리 중 오류: {ex.Message}",
                        ProcessingTime = overallStopwatch.Elapsed,
                        ProcessorName = "Pipeline",
                        Output = currentOutput
                    };
                }
            }

            overallStopwatch.Stop();
            
            // 최종 진행률 보고
            progress?.Report(new ProcessingProgress
            {
                CurrentProcessor = "완료",
                ProcessorIndex = enabledProcessors.Count,
                TotalProcessors = enabledProcessors.Count,
                Message = "모든 후처리 완료"
            });

            _logger.Information($"후처리 파이프라인 완료 (총 소요시간: {overallStopwatch.Elapsed.TotalMilliseconds:F2}ms)");

            return new ProcessingResult
            {
                Success = true,
                Output = currentOutput,
                ProcessingTime = overallStopwatch.Elapsed,
                ProcessorName = "Pipeline"
            };
        }
    }

    /// <summary>
    /// 처리 진행률 정보
    /// </summary>
    public class ProcessingProgress
    {
        public string CurrentProcessor { get; set; }
        public int ProcessorIndex { get; set; }
        public int TotalProcessors { get; set; }
        public string Message { get; set; }
        
        public double PercentComplete => TotalProcessors > 0 
            ? (double)ProcessorIndex / TotalProcessors * 100 
            : 0;
    }
}