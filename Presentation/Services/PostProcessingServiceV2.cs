using ExcelToYamlAddin.Application.Interfaces;
using ExcelToYamlAddin.Application.PostProcessing;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.Configuration;
using ExcelToYamlAddin.Infrastructure.Logging;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
// 네임스페이스 충돌을 피하기 위해 별칭 사용
using YamlMergeProcessorV2 = ExcelToYamlAddin.Application.PostProcessing.Processors.YamlMergeProcessor;
using YamlFlowStyleProcessorV2 = ExcelToYamlAddin.Application.PostProcessing.Processors.YamlFlowStyleProcessor;
using JsonFormatterProcessorV2 = ExcelToYamlAddin.Application.PostProcessing.Processors.JsonFormatterProcessor;
using XmlFormatterProcessorV2 = ExcelToYamlAddin.Application.PostProcessing.Processors.XmlFormatterProcessor;

namespace ExcelToYamlAddin.Presentation.Services
{
    /// <summary>
    /// 파이프라인 기반의 향상된 YAML 후처리 서비스
    /// </summary>
    public class PostProcessingServiceV2
    {
        private readonly IProcessingPipeline _pipeline;
        private readonly ISimpleLogger _logger;

        public PostProcessingServiceV2()
        {
            _logger = SimpleLoggerFactory.CreateLogger<PostProcessingServiceV2>();
            
            // 프로세서들 초기화
            var processors = new List<IPostProcessor>
            {
                new YamlMergeProcessorV2(),
                new YamlFlowStyleProcessorV2(),
                new JsonFormatterProcessorV2(),
                new XmlFormatterProcessorV2()
            };

            _pipeline = new ProcessingPipeline(processors);
        }

        /// <summary>
        /// 지정된 파일 목록에 대해 파이프라인 기반 후처리를 적용합니다.
        /// </summary>
        public async Task<(int successCount, Dictionary<string, ProcessingResult> results)> ApplyPostProcessingAsync(
            List<string> filePaths,
            List<Excel.Worksheet> convertibleSheets,
            IProgress<Forms.ProgressForm.ProgressInfo> progress,
            CancellationToken cancellationToken,
            int initialProgressPercentage,
            int progressRange,
            bool isForJsonConversion,
            bool addEmptyYamlFields,
            OutputFormat outputFormat)
        {
            var results = new Dictionary<string, ProcessingResult>();
            int successCount = 0;
            int filesProcessed = 0;
            int totalFiles = filePaths.Count;

            _logger.Information($"파이프라인 기반 후처리 시작: {totalFiles}개 파일");

            foreach (var filePath in filePaths)
            {
                cancellationToken.ThrowIfCancellationRequested();

                string fileName = Path.GetFileNameWithoutExtension(filePath);
                
                // 진행률 계산
                int fileStartProgress = initialProgressPercentage + (int)((double)filesProcessed / totalFiles * progressRange);
                int fileEndProgress = initialProgressPercentage + (int)((double)(filesProcessed + 1) / totalFiles * progressRange);

                ReportProgress(progress, fileStartProgress,
                    $"[{filesProcessed + 1}/{totalFiles}] '{fileName}' 후처리 중...");

                try
                {
                    // 매칭되는 시트 찾기
                    string matchedSheetName = FindMatchingSheet(fileName, convertibleSheets);
                    
                    if (matchedSheetName != null)
                    {
                        // 처리 컨텍스트 생성
                        var context = CreateProcessingContext(filePath, matchedSheetName, addEmptyYamlFields, isForJsonConversion, outputFormat);
                        
                        // 파일 읽기 (.NET Framework에서는 동기 메서드 사용)
                        string content = await Task.Run(() => File.ReadAllText(filePath), cancellationToken);
                        
                        // 파이프라인 진행률 래퍼
                        var pipelineProgress = new Progress<ProcessingProgress>(p =>
                        {
                            int subProgress = fileStartProgress + (int)((double)p.PercentComplete / 100 * (fileEndProgress - fileStartProgress));
                            ReportProgress(progress, subProgress, 
                                $"[{filesProcessed + 1}/{totalFiles}] '{fileName}' - {p.Message}");
                        });

                        // 파이프라인 실행
                        var result = await _pipeline.ProcessAsync(content, context, pipelineProgress, cancellationToken);
                        results[filePath] = result;

                        if (result.Success)
                        {
                            // 결과 저장 (.NET Framework에서는 동기 메서드 사용)
                            await Task.Run(() => File.WriteAllText(filePath, result.Output), cancellationToken);
                            successCount++;
                            _logger.Information($"'{fileName}' 후처리 성공");
                        }
                        else
                        {
                            _logger.Warning($"'{fileName}' 후처리 실패: {result.ErrorMessage}");
                        }
                    }
                    else
                    {
                        _logger.Information($"'{fileName}' - 매칭되는 시트 없음, 건너뜀");
                        results[filePath] = ProcessingResult.CreateSuccess(string.Empty, "Skipped");
                    }
                }
                catch (Exception ex)
                {
                    _logger.Error($"'{fileName}' 처리 중 오류: {ex.Message}", ex);
                    results[filePath] = ProcessingResult.CreateFailure(ex.Message);
                }

                filesProcessed++;
                ReportProgress(progress, fileEndProgress,
                    $"[{filesProcessed}/{totalFiles}] '{fileName}' 처리 완료");
            }

            _logger.Information($"파이프라인 기반 후처리 완료: 성공 {successCount}/{totalFiles}");
            return (successCount, results);
        }

        /// <summary>
        /// 처리 컨텍스트를 생성합니다.
        /// </summary>
        private ProcessingContext CreateProcessingContext(
            string filePath, 
            string sheetName, 
            bool addEmptyYamlFields,
            bool isForJsonConversion,
            OutputFormat outputFormat)
        {
            // 설정 읽기
            bool yamlEmptyFieldsOption = GetYamlEmptyFieldsOption(sheetName, addEmptyYamlFields);
            string mergeKeyPaths = GetMergeKeyPaths(sheetName);
            string flowStyleConfig = GetFlowStyleConfig(sheetName, filePath);
            
            // 기존 YamlFlowStyleProcessor의 정적 메서드 사용
            bool isFlowStyleEnabled = !Application.PostProcessing.YamlFlowStyleProcessor.IsConfigEffectivelyEmpty(flowStyleConfig);

            return new ProcessingContext
            {
                FilePath = filePath,
                SheetName = sheetName,
                OutputFormat = outputFormat,
                Options = new ProcessingOptions
                {
                    EnableMerge = !string.IsNullOrEmpty(mergeKeyPaths) && outputFormat == OutputFormat.Yaml,
                    ApplyFlowStyle = isFlowStyleEnabled && 
                                     outputFormat == OutputFormat.Yaml && 
                                     !isForJsonConversion,
                    IncludeEmptyFields = yamlEmptyFieldsOption,
                    MergeKeyPaths = mergeKeyPaths,
                    FlowStyleConfig = flowStyleConfig
                }
            };
        }

        /// <summary>
        /// 매칭되는 시트 이름을 찾습니다.
        /// </summary>
        private string FindMatchingSheet(string fileName, List<Excel.Worksheet> convertibleSheets)
        {
            foreach (var sheet in convertibleSheets)
            {
                string currentSheetNameForMatch = sheet.Name;
                if (currentSheetNameForMatch.StartsWith("!"))
                    currentSheetNameForMatch = currentSheetNameForMatch.Substring(1);

                if (string.Compare(currentSheetNameForMatch, fileName, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    return sheet.Name;
                }
            }
            return null;
        }

        /// <summary>
        /// YAML 빈 필드 옵션을 가져옵니다.
        /// </summary>
        private bool GetYamlEmptyFieldsOption(string sheetName, bool addEmptyYamlFields)
        {
            bool yamlEmptyFieldsOption = ExcelConfigManager.Instance.GetConfigBool(sheetName, "YamlEmptyFields", false);
            if (!yamlEmptyFieldsOption)
                yamlEmptyFieldsOption = SheetPathManager.Instance.GetYamlEmptyFieldsOption(sheetName);
            if (!yamlEmptyFieldsOption && addEmptyYamlFields)
                yamlEmptyFieldsOption = addEmptyYamlFields;

            return yamlEmptyFieldsOption;
        }

        /// <summary>
        /// 병합 키 경로 설정을 가져옵니다.
        /// </summary>
        private string GetMergeKeyPaths(string sheetName)
        {
            string sheetMergeKeyPaths = ExcelConfigManager.Instance.GetConfigValue(sheetName, "MergeKeyPaths", "");
            if (string.IsNullOrEmpty(sheetMergeKeyPaths))
                sheetMergeKeyPaths = SheetPathManager.Instance.GetMergeKeyPaths(sheetName);
            
            return sheetMergeKeyPaths;
        }

        /// <summary>
        /// Flow 스타일 설정을 가져옵니다.
        /// </summary>
        private string GetFlowStyleConfig(string sheetName, string filePath)
        {
            string sheetFlowStyle = ExcelConfigManager.Instance.GetConfigValue(sheetName, "FlowStyle", "");
            if (string.IsNullOrWhiteSpace(sheetFlowStyle))
                sheetFlowStyle = SheetPathManager.Instance.GetFlowStyleConfig(sheetName ?? Path.GetFileNameWithoutExtension(filePath));
            
            return sheetFlowStyle;
        }

        /// <summary>
        /// 진행 상태를 보고합니다.
        /// </summary>
        private void ReportProgress(IProgress<Forms.ProgressForm.ProgressInfo> progress, int percentage, string message)
        {
            progress?.Report(new Forms.ProgressForm.ProgressInfo
            {
                Percentage = percentage,
                StatusMessage = message
            });
        }
    }
}