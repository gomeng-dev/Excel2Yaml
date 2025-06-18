using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.Logging;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToYamlAddin.Application.PostProcessing.Processors
{
    /// <summary>
    /// YAML 파일의 항목들을 지정된 키 경로를 기반으로 병합하는 후처리기입니다.
    /// </summary>
    public class YamlMergeProcessor : PostProcessorBase
    {
        private readonly ISimpleLogger _logger;

        public YamlMergeProcessor()
        {
            _logger = SimpleLoggerFactory.CreateLogger<YamlMergeProcessor>();
        }

        /// <summary>
        /// 처리 우선순위 (병합은 다른 처리보다 먼저 수행되어야 함)
        /// </summary>
        public override int Priority => 10;

        /// <summary>
        /// 이 프로세서가 처리할 수 있는지 확인합니다.
        /// </summary>
        public override bool CanProcess(ProcessingContext context)
        {
            return context.OutputFormat == OutputFormat.Yaml &&
                   context.Options.EnableMerge &&
                   !string.IsNullOrWhiteSpace(context.Options.MergeKeyPaths);
        }

        /// <summary>
        /// YAML 병합 처리를 수행합니다.
        /// </summary>
        protected override async Task<string> ProcessCoreAsync(string input, ProcessingContext context, CancellationToken cancellationToken)
        {
            _logger.Information($"YAML 병합 처리 시작: {context.Options.MergeKeyPaths}");

            try
            {
                // 임시 파일로 저장하여 처리
                var tempPath = Path.GetTempFileName();
                await Task.Run(() => File.WriteAllText(tempPath, input), cancellationToken);

                // 기존 YamlMergeKeyPathsProcessor 사용
                bool success = YamlMergeKeyPathsProcessor.ProcessYamlFileFromConfig(
                    tempPath,
                    context.Options.MergeKeyPaths,
                    context.Options.IncludeEmptyFields);

                if (success)
                {
                    _logger.Information("YAML 병합 처리 완료");
                    // 처리된 파일 읽기
                    var result = await Task.Run(() => File.ReadAllText(tempPath), cancellationToken);
                    
                    // 임시 파일 삭제
                    try { File.Delete(tempPath); } catch { }
                    
                    return result;
                }
                else
                {
                    // 임시 파일 삭제
                    try { File.Delete(tempPath); } catch { }
                    _logger.Warning("YAML 병합 처리 실패");
                    // 실패 시 원본 반환
                    return input;
                }
            }
            catch (Exception ex)
            {
                _logger.Error($"YAML 병합 중 오류: {ex.Message}", ex);
                throw;
            }
        }
    }
}