using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using ExcelToYamlAddin.Application.PostProcessing;
using ExcelToYamlAddin.Infrastructure.Configuration;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToYamlAddin.Presentation.Services
{
    /// <summary>
    /// YAML 후처리 관련 기능을 처리하는 서비스
    /// </summary>
    public class PostProcessingService
    {
        /// <summary>
        /// 지정된 YAML 파일 목록에 대해 공통 후처리 작업을 적용합니다.
        /// </summary>
        /// <param name="yamlFilePaths">후처리를 적용할 YAML 파일 경로 목록입니다.</param>
        /// <param name="convertibleSheets">변환 가능한 원본 Excel 시트 목록입니다.</param>
        /// <param name="progress">진행 상태를 보고할 IProgress 객체입니다.</param>
        /// <param name="cancellationToken">작업 취소를 위한 CancellationToken입니다.</param>
        /// <param name="initialProgressPercentage">이 후처리 단계가 시작될 때의 전체 진행률입니다.</param>
        /// <param name="progressRange">이 후처리 단계가 전체 진행률에서 차지하는 범위입니다.</param>
        /// <param name="isForJsonConversion">JSON으로 변환하기 위한 중간 단계인 경우 true</param>
        /// <param name="addEmptyYamlFields">전역 빈 필드 설정</param>
        /// <returns>키 경로 병합 및 Flow 스타일 처리 성공 횟수를 포함하는 튜플을 반환합니다.</returns>
        public (int mergeKeyPathsSuccessCount, int flowStyleSuccessCount) ApplyYamlPostProcessing(
            List<string> yamlFilePaths,
            List<Excel.Worksheet> convertibleSheets,
            IProgress<Forms.ProgressForm.ProgressInfo> progress,
            CancellationToken cancellationToken,
            int initialProgressPercentage,
            int progressRange,
            bool isForJsonConversion,
            bool addEmptyYamlFields)
        {
            int mergeKeyPathsSuccessCount = 0;
            int flowStyleSuccessCount = 0;
            int filesProcessedInThisStep = 0;
            int totalFilesToProcessInThisStep = yamlFilePaths.Count;

            ReportProgress(progress, initialProgressPercentage, 
                $"YAML 후처리 시작: {totalFilesToProcessInThisStep}개 파일 처리 예정");

            foreach (var yamlFilePath in yamlFilePaths)
            {
                cancellationToken.ThrowIfCancellationRequested();

                string fileName = Path.GetFileNameWithoutExtension(yamlFilePath);
                bool mergeKeyPathsProcessingAttemptedThisFile = false;
                bool flowStyleProcessingAttemptedThisFile = false;

                // 파일별 진행률 계산
                int fileStartProgress = initialProgressPercentage + (int)((double)filesProcessedInThisStep / totalFilesToProcessInThisStep * progressRange);
                int fileEndProgress = initialProgressPercentage + (int)((double)(filesProcessedInThisStep + 1) / totalFilesToProcessInThisStep * progressRange);
                int fileProgressRange = fileEndProgress - fileStartProgress;

                ReportProgress(progress, fileStartProgress, 
                    $"[{filesProcessedInThisStep + 1}/{totalFilesToProcessInThisStep}] '{fileName}' 후처리 준비 중...");

                // 매칭되는 시트 찾기
                string matchedSheetName = FindMatchingSheet(fileName, convertibleSheets);

                if (matchedSheetName != null)
                {
                    // YAML 빈 필드 옵션 확인
                    bool yamlEmptyFieldsOption = GetYamlEmptyFieldsOption(matchedSheetName, addEmptyYamlFields);

                    // 키 경로 병합 처리 (25% of file progress)
                    ReportProgress(progress, fileStartProgress + fileProgressRange / 8, 
                        $"[{filesProcessedInThisStep + 1}/{totalFilesToProcessInThisStep}] '{fileName}' - 키 경로 병합 확인 중...");
                    
                    if (ProcessMergeKeyPaths(yamlFilePath, matchedSheetName, yamlEmptyFieldsOption))
                    {
                        mergeKeyPathsProcessingAttemptedThisFile = true;
                        mergeKeyPathsSuccessCount++;
                        ReportProgress(progress, fileStartProgress + fileProgressRange / 4, 
                            $"[{filesProcessedInThisStep + 1}/{totalFilesToProcessInThisStep}] '{fileName}' - 키 경로 병합 완료");
                    }
                    else
                    {
                        ReportProgress(progress, fileStartProgress + fileProgressRange / 4, 
                            $"[{filesProcessedInThisStep + 1}/{totalFilesToProcessInThisStep}] '{fileName}' - 키 경로 병합 건너뜀");
                    }

                    if (!isForJsonConversion)
                    {
                        // Flow Style 처리 (25% of file progress)
                        ReportProgress(progress, fileStartProgress + fileProgressRange / 2, 
                            $"[{filesProcessedInThisStep + 1}/{totalFilesToProcessInThisStep}] '{fileName}' - Flow Style 확인 중...");
                        
                        if (ProcessFlowStyle(yamlFilePath, matchedSheetName))
                        {
                            flowStyleProcessingAttemptedThisFile = true;
                            flowStyleSuccessCount++;
                            ReportProgress(progress, fileStartProgress + (fileProgressRange * 5) / 8, 
                                $"[{filesProcessedInThisStep + 1}/{totalFilesToProcessInThisStep}] '{fileName}' - Flow Style 적용 완료");
                        }
                        else
                        {
                            ReportProgress(progress, fileStartProgress + (fileProgressRange * 5) / 8, 
                                $"[{filesProcessedInThisStep + 1}/{totalFilesToProcessInThisStep}] '{fileName}' - Flow Style 건너뜀");
                        }

                        // 빈 배열 처리 (12.5% of file progress)
                        ReportProgress(progress, fileStartProgress + (fileProgressRange * 3) / 4, 
                            $"[{filesProcessedInThisStep + 1}/{totalFilesToProcessInThisStep}] '{fileName}' - 빈 배열 처리 중...");
                        ProcessEmptyArrays(yamlFilePath, yamlEmptyFieldsOption, addEmptyYamlFields);

                        // 최종 문자열 정리 (12.5% of file progress)
                        if (!mergeKeyPathsProcessingAttemptedThisFile && !flowStyleProcessingAttemptedThisFile)
                        {
                            ReportProgress(progress, fileStartProgress + (fileProgressRange * 7) / 8, 
                                $"[{filesProcessedInThisStep + 1}/{totalFilesToProcessInThisStep}] '{fileName}' - 최종 문자열 정리 중...");
                            Debug.WriteLine($"[ApplyYamlPostProcessing] 최종 Raw 문자열 변환 후처리 실행: {yamlFilePath}");
                            new FinalRawStringConverter().ProcessYamlFile(yamlFilePath);
                        }
                    }
                }
                else
                {
                    ReportProgress(progress, fileStartProgress + fileProgressRange / 2, 
                        $"[{filesProcessedInThisStep + 1}/{totalFilesToProcessInThisStep}] '{fileName}' - 매칭되는 시트 없음, 건너뜀");
                }

                filesProcessedInThisStep++;
                ReportProgress(progress, fileEndProgress, 
                    $"[{filesProcessedInThisStep}/{totalFilesToProcessInThisStep}] '{fileName}' 후처리 완료 (병합: {mergeKeyPathsSuccessCount}, 스타일: {flowStyleSuccessCount})");
            }
            return (mergeKeyPathsSuccessCount, flowStyleSuccessCount);
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
        /// 키 경로 병합을 처리합니다.
        /// </summary>
        private bool ProcessMergeKeyPaths(string yamlFilePath, string sheetName, bool yamlEmptyFieldsOption)
        {
            string sheetMergeKeyPaths = ExcelConfigManager.Instance.GetConfigValue(sheetName, "MergeKeyPaths", "");
            if (string.IsNullOrEmpty(sheetMergeKeyPaths)) 
                sheetMergeKeyPaths = SheetPathManager.Instance.GetMergeKeyPaths(sheetName);

            if (!string.IsNullOrEmpty(sheetMergeKeyPaths))
            {
                Debug.WriteLine($"[ProcessMergeKeyPaths] YAML 병합 후처리 실행: {yamlFilePath}, 설정: {sheetMergeKeyPaths}");
                bool success = YamlMergeKeyPathsProcessor.ProcessYamlFileFromConfig(yamlFilePath, sheetMergeKeyPaths, yamlEmptyFieldsOption);
                if (success) 
                {
                    Debug.WriteLine($"[ProcessMergeKeyPaths] YAML 병합 후처리 완료: {yamlFilePath}");
                }
                else 
                {
                    Debug.WriteLine($"[ProcessMergeKeyPaths] YAML 병합 후처리 실패: {yamlFilePath}");
                }
                return success;
            }
            return false;
        }

        /// <summary>
        /// Flow Style을 처리합니다.
        /// </summary>
        private bool ProcessFlowStyle(string yamlFilePath, string sheetName)
        {
            string sheetFlowStyle = ExcelConfigManager.Instance.GetConfigValue(sheetName, "FlowStyle", "");
            if (string.IsNullOrWhiteSpace(sheetFlowStyle)) 
                sheetFlowStyle = SheetPathManager.Instance.GetFlowStyleConfig(sheetName ?? Path.GetFileNameWithoutExtension(yamlFilePath));

            if (!YamlFlowStyleProcessor.IsConfigEffectivelyEmpty(sheetFlowStyle))
            {
                Debug.WriteLine($"[ProcessFlowStyle] YAML 흐름 스타일 후처리 실행: {yamlFilePath}, 설정: {sheetFlowStyle}");
                bool success = YamlFlowStyleProcessor.ProcessYamlFileFromConfig(yamlFilePath, sheetFlowStyle);
                if (success) 
                {
                    Debug.WriteLine($"[ProcessFlowStyle] YAML 흐름 스타일 후처리 완료: {yamlFilePath}");
                }
                else 
                {
                    Debug.WriteLine($"[ProcessFlowStyle] YAML 흐름 스타일 후처리 실패: {yamlFilePath}");
                }
                return success;
            }
            else 
            {
                Debug.WriteLine($"[ProcessFlowStyle] YAML 흐름 스타일 후처리 건너뜀: {yamlFilePath}, 설정: '{sheetFlowStyle}'");
            }
            return false;
        }

        /// <summary>
        /// 빈 배열을 처리합니다.
        /// </summary>
        private void ProcessEmptyArrays(string yamlFilePath, bool yamlEmptyFieldsOption, bool addEmptyYamlFields)
        {
            bool processEmptyArrays = yamlEmptyFieldsOption || addEmptyYamlFields;
            if (processEmptyArrays) 
            {
                Debug.WriteLine($"[ProcessEmptyArrays] YAML 빈 배열 처리: OrderedYamlFactory에서 처리 (파일: {yamlFilePath}), 시트별: {yamlEmptyFieldsOption}, 전역: {addEmptyYamlFields}");
            }
            else 
            {
                Debug.WriteLine($"[ProcessEmptyArrays] YAML 빈 배열 처리 건너뜀: 관련 옵션 비활성화. 시트별: {yamlEmptyFieldsOption}, 전역: {addEmptyYamlFields}");
            }
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