using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using ExcelToYamlAddin.Application.Services;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.Configuration;
using ExcelToYamlAddin.Infrastructure.Excel;
using ExcelToYamlAddin.Infrastructure.Logging;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToYamlAddin.Presentation.Services
{
    /// <summary>
    /// Excel 파일 변환 관련 기능을 처리하는 서비스
    /// </summary>
    public class ConversionService
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<ConversionService>();
        
        /// <summary>
        /// Excel 파일을 지정된 형식으로 변환합니다.
        /// </summary>
        public List<string> ConvertExcelFile(
            ExcelToYamlConfig config,
            Excel.Workbook activeWorkbook,
            string tempFile,
            List<Excel.Worksheet> convertibleSheets,
            IProgress<Forms.ProgressForm.ProgressInfo> progress,
            CancellationToken cancellationToken)
        {
            var convertedFiles = new List<string>();
            
            if (activeWorkbook == null || convertibleSheets == null || convertibleSheets.Count == 0)
            {
                return convertedFiles;
            }

            // 워크북 경로 설정
            string workbookPath = activeWorkbook.FullName;
            SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);

            // 처리할 시트 목록 계산
            var sheetsToProcess = FilterSheetsToProcess(convertibleSheets);
            
            if (sheetsToProcess.Count == 0)
            {
                ReportProgress(progress, 100, "처리할 시트가 없습니다.", isCompleted: true);
                return convertedFiles;
            }

            string outputFormat = config.OutputFormat == OutputFormat.Json ? "JSON" : "YAML";
            int processedSheets = 0;
            int successCount = 0;
            int skipCount = 0;

            // 초기 분석 단계 (5%)
            ReportProgress(progress, 5, "Excel 워크북 분석 중...");
            cancellationToken.ThrowIfCancellationRequested();

            ReportProgress(progress, 10, $"{sheetsToProcess.Count}개 시트 검증 중...");
            cancellationToken.ThrowIfCancellationRequested();

            // 각 시트별 변환 진행률 계산 (10% ~ 90% 사용)
            int baseProgress = 15;
            int conversionProgressRange = 75; // 15% ~ 90%

            // 모든 변환 가능한 시트에 대해 처리
            foreach (var sheet in sheetsToProcess)
            {
                cancellationToken.ThrowIfCancellationRequested();

                string sheetName = sheet.Name;
                string cleanSheetName = sheetName.StartsWith("!") ? sheetName.Substring(1) : sheetName;
                
                // 시트 시작 진행률
                int sheetStartProgress = baseProgress + (int)((double)processedSheets / sheetsToProcess.Count * conversionProgressRange);
                int sheetEndProgress = baseProgress + (int)((double)(processedSheets + 1) / sheetsToProcess.Count * conversionProgressRange);
                int sheetProgressRange = sheetEndProgress - sheetStartProgress;

                // 시트별 세부 단계
                ReportProgress(progress, sheetStartProgress, 
                    $"[{processedSheets + 1}/{sheetsToProcess.Count}] '{cleanSheetName}' 시트 데이터 읽기 중...");
                cancellationToken.ThrowIfCancellationRequested();

                ReportProgress(progress, sheetStartProgress + sheetProgressRange / 4, 
                    $"[{processedSheets + 1}/{sheetsToProcess.Count}] '{cleanSheetName}' 스키마 분석 중...");
                cancellationToken.ThrowIfCancellationRequested();

                ReportProgress(progress, sheetStartProgress + sheetProgressRange / 2, 
                    $"[{processedSheets + 1}/{sheetsToProcess.Count}] '{cleanSheetName}' {outputFormat} 변환 중...");

                // 시트 변환 처리
                var resultFile = ProcessSheet(sheet, tempFile, config);
                
                if (!string.IsNullOrEmpty(resultFile))
                {
                    successCount++;
                    convertedFiles.Add(resultFile);
                    ReportProgress(progress, sheetStartProgress + (sheetProgressRange * 3) / 4, 
                        $"[{processedSheets + 1}/{sheetsToProcess.Count}] '{cleanSheetName}' 파일 저장 중...");
                }
                else
                {
                    skipCount++;
                    ReportProgress(progress, sheetStartProgress + (sheetProgressRange * 3) / 4, 
                        $"[{processedSheets + 1}/{sheetsToProcess.Count}] '{cleanSheetName}' 변환 실패");
                }

                processedSheets++;
                ReportProgress(progress, sheetEndProgress, 
                    $"[{processedSheets}/{sheetsToProcess.Count}] '{cleanSheetName}' 완료 ({successCount}개 성공, {skipCount}개 실패)");
                cancellationToken.ThrowIfCancellationRequested();
            }

            // 최종 정리 단계 (90% ~ 100%)
            ReportProgress(progress, 95, "변환 결과 정리 중...");
            cancellationToken.ThrowIfCancellationRequested();

            // 변환 결과 로그 작성
            Debug.WriteLine($"변환 완료: {successCount}개 성공, {skipCount}개 실패");

            // 작업 완료 알림
            ReportProgress(progress, 100, 
                $"전체 변환 완료: {successCount}개 시트 성공" + (skipCount > 0 ? $", {skipCount}개 실패" : ""), 
                isCompleted: true);

            return convertedFiles;
        }

        /// <summary>
        /// Excel 파일을 임시 디렉토리에 YAML로 변환합니다.
        /// </summary>
        public List<string> ConvertExcelFileToTemp(
            ExcelToYamlConfig config,
            string tempDir,
            List<Excel.Worksheet> sheetsToProcess,
            string tempFile,
            IProgress<Forms.ProgressForm.ProgressInfo> progress = null,
            CancellationToken cancellationToken = default)
        {
            var convertedFiles = new List<string>();

            try
            {
                if (sheetsToProcess == null || sheetsToProcess.Count == 0)
                {
                    return convertedFiles;
                }

                int processedSheets = 0;
                int successCount = 0;
                int skipCount = 0;

                // 활성화된 시트만 필터링
                var enabledSheets = sheetsToProcess.Where(sheet => 
                    SheetPathManager.Instance.IsSheetEnabled(sheet.Name)).ToList();

                ReportProgress(progress, 5, $"{enabledSheets.Count}개 시트 임시 변환 준비 중...");

                // 모든 변환 가능한 시트에 대해 처리
                foreach (var sheet in enabledSheets)
                {
                    cancellationToken.ThrowIfCancellationRequested();

                    string sheetName = sheet.Name;
                    string fileName = sheetName.StartsWith("!") ? sheetName.Substring(1) : sheetName;
                    string resultFile = Path.Combine(tempDir, $"{fileName}.yaml");

                    // 시트별 진행률 계산 (10% ~ 90%)
                    int baseProgress = 10;
                    int conversionRange = 80;
                    int currentProgress = baseProgress + (int)((double)processedSheets / enabledSheets.Count * conversionRange);

                    ReportProgress(progress, currentProgress, 
                        $"[{processedSheets + 1}/{enabledSheets.Count}] '{fileName}' 임시 변환 중...");

                    try
                    {
                        // 변환 처리 - 시트 이름 지정
                        var excelReader = new ExcelReader(config);
                        excelReader.ProcessExcelFile(tempFile, resultFile, sheetName);

                        convertedFiles.Add(resultFile);
                        successCount++;
                        
                        ReportProgress(progress, currentProgress + (conversionRange / enabledSheets.Count / 2), 
                            $"[{processedSheets + 1}/{enabledSheets.Count}] '{fileName}' 임시 파일 저장 완료");
                    }
                    catch (Exception ex)
                    {
                        skipCount++;
                        Debug.WriteLine($"[ConvertExcelFileToTemp] 시트 '{sheetName}' 변환 중 오류 발생: {ex.Message}");
                        ReportProgress(progress, currentProgress + (conversionRange / enabledSheets.Count / 2), 
                            $"[{processedSheets + 1}/{enabledSheets.Count}] '{fileName}' 변환 실패");
                    }

                    processedSheets++;
                }

                ReportProgress(progress, 95, "임시 변환 결과 정리 중...");
                Debug.WriteLine($"[ConvertExcelFileToTemp] 임시 변환 완료: {successCount}개 성공, {skipCount}개 실패");

                return convertedFiles;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ConvertExcelFileToTemp] 변환 중 오류 발생: {ex.Message}");
                return convertedFiles;
            }
        }

        /// <summary>
        /// 처리할 시트 목록을 필터링합니다.
        /// </summary>
        private List<Excel.Worksheet> FilterSheetsToProcess(List<Excel.Worksheet> convertibleSheets)
        {
            var sheetsToProcess = new List<Excel.Worksheet>();
            
            foreach (var sheet in convertibleSheets)
            {
                string sheetName = sheet.Name;
                bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);
                string savePath = SheetPathManager.Instance.GetSheetPath(sheetName);

                if (isEnabled && !string.IsNullOrEmpty(savePath))
                {
                    sheetsToProcess.Add(sheet);
                }
            }

            return sheetsToProcess;
        }

        /// <summary>
        /// 개별 시트를 처리합니다.
        /// </summary>
        private string ProcessSheet(Excel.Worksheet sheet, string tempFile, ExcelToYamlConfig config)
        {
            string sheetName = sheet.Name;
            
            // 앞의 '!' 문자 제거 (표시용)
            string fileName = sheetName.StartsWith("!") ? sheetName.Substring(1) : sheetName;

            // 시트별 저장 경로 가져오기 - 원래 이름 유지
            string savePath = SheetPathManager.Instance.GetSheetPath(sheetName);

            // 디버깅을 위한 로그 추가
            Debug.WriteLine($"시트 '{sheetName}'의 저장 경로: {savePath ?? "설정되지 않음"}");

            // 저장 경로가 없으면 '!'가 없는 이름으로도 시도
            if (string.IsNullOrEmpty(savePath) && sheetName.StartsWith("!"))
            {
                string altSheetName = sheetName.Substring(1);
                savePath = SheetPathManager.Instance.GetSheetPath(altSheetName);
                Debug.WriteLine($"대체 시트명 '{altSheetName}'으로 경로 검색 결과: {savePath ?? "설정되지 않음"}");
            }

            // 저장 경로가 유효하지 않으면 건너뛰기
            if (string.IsNullOrEmpty(savePath))
            {
                return null;
            }

            // 활성화 상태 확인 - 비활성화된 시트는 건너뛰기
            bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);
            if (!isEnabled)
            {
                return null;
            }

            // 경로 존재 확인 및 생성
            if (!Directory.Exists(savePath))
            {
                try
                {
                    Debug.WriteLine($"경로가 존재하지 않아 생성합니다: {savePath}");
                    Directory.CreateDirectory(savePath);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"경로 생성 실패: {ex.Message}");
                    return null;
                }
            }

            // 파일 확장자 결정
            string ext = config.OutputFormat == OutputFormat.Json ? ".json" : ".yaml";

            // 결과 파일 경로
            string resultFile = Path.Combine(savePath, $"{fileName}{ext}");

            try
            {
                // 변환 처리 - 시트 이름 지정
                var excelReader = new ExcelReader(config);
                excelReader.ProcessExcelFile(tempFile, resultFile, sheetName);

                return resultFile;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ProcessSheet] 시트 '{sheetName}' 변환 중 오류 발생: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 진행 상태를 보고합니다.
        /// </summary>
        private void ReportProgress(
            IProgress<Forms.ProgressForm.ProgressInfo> progress,
            int percentage,
            string message,
            bool isCompleted = false,
            bool hasError = false,
            string errorMessage = null)
        {
            progress?.Report(new Forms.ProgressForm.ProgressInfo
            {
                Percentage = percentage,
                StatusMessage = message,
                IsCompleted = isCompleted,
                HasError = hasError,
                ErrorMessage = errorMessage
            });
        }
    }
}