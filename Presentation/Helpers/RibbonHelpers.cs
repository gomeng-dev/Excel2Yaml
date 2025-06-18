using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ExcelToYamlAddin.Infrastructure.Configuration;
using ExcelToYamlAddin.Infrastructure.Excel;

namespace ExcelToYamlAddin.Presentation.Helpers
{
    /// <summary>
    /// Ribbon UI 관련 유틸리티 기능을 제공하는 헬퍼 클래스
    /// </summary>
    public static class RibbonHelpers
    {
        /// <summary>
        /// 시트 이름에서 '!' 접두사를 제거합니다.
        /// </summary>
        /// <param name="sheetName">원본 시트 이름</param>
        /// <returns>'!' 접두사가 제거된 시트 이름</returns>
        public static string GetCleanSheetName(string sheetName)
        {
            return sheetName?.StartsWith("!") == true ? sheetName.Substring(1) : sheetName ?? string.Empty;
        }

        /// <summary>
        /// 활성 워크북의 유효성을 검사하고 초기화합니다.
        /// </summary>
        /// <returns>워크북이 유효하면 true, 그렇지 않으면 false</returns>
        public static bool ValidateAndInitializeWorkbook()
        {
            var app = Globals.ThisAddIn.Application;
            var activeWorkbook = app.ActiveWorkbook;

            if (activeWorkbook == null)
            {
                MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            string workbookPath = activeWorkbook.FullName;
            SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);
            ExcelConfigManager.Instance.SetCurrentWorkbook(workbookPath);
            
            return true;
        }

        /// <summary>
        /// Import 작업을 위한 워크북 유효성을 검사합니다.
        /// </summary>
        /// <returns>유효한 워크북 객체, 없으면 null</returns>
        public static dynamic ValidateWorkbookForImport()
        {
            var app = Globals.ThisAddIn.Application;
            var currentWorkbook = app.ActiveWorkbook;

            if (currentWorkbook == null)
            {
                MessageBox.Show("활성 워크북이 없습니다. Excel 파일을 먼저 열어주세요.",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return null;
            }

            return currentWorkbook;
        }

        /// <summary>
        /// 임시 디렉토리를 안전하게 삭제합니다.
        /// </summary>
        /// <param name="tempDir">삭제할 임시 디렉토리 경로</param>
        /// <param name="methodName">호출한 메소드 이름 (로깅용)</param>
        public static void CleanupTempDirectory(string tempDir, string methodName)
        {
            try
            {
                if (Directory.Exists(tempDir))
                {
                    Directory.Delete(tempDir, true);
                    Debug.WriteLine($"[{methodName}] 임시 디렉토리 정리 완료: {tempDir}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[{methodName}] 임시 디렉토리 정리 중 오류: {ex.Message}");
            }
        }

        /// <summary>
        /// 진행 상태를 보고합니다.
        /// </summary>
        /// <param name="progress">진행 상태 보고 객체</param>
        /// <param name="percentage">진행률 (0-100)</param>
        /// <param name="message">상태 메시지</param>
        /// <param name="isCompleted">완료 여부</param>
        /// <param name="hasError">오류 발생 여부</param>
        /// <param name="errorMessage">오류 메시지</param>
        public static void ReportProgress(
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

        /// <summary>
        /// 현재 활성 시트의 설정과 전역 설정을 조합하여 최종 빈 필드 포함 설정을 반환합니다.
        /// </summary>
        /// <param name="configKey">Excel 설정에서 조회할 키 이름</param>
        /// <param name="globalAddEmptyFields">전역 빈 필드 설정</param>
        /// <returns>시트별 설정 또는 전역 설정 중 하나라도 true이면 true, 그렇지 않으면 false</returns>
        public static bool GetEffectiveEmptyFieldsSetting(string configKey, bool globalAddEmptyFields)
        {
            bool sheetSpecificSetting = false;
            string currentSheetName = string.Empty;

            try
            {
                // 현재 활성 시트의 설정 확인
                if (Globals.ThisAddIn.Application.ActiveSheet != null)
                {
                    currentSheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
                    sheetSpecificSetting = ExcelConfigManager.Instance.GetConfigBool(currentSheetName, configKey, false);
                    Debug.WriteLine($"[GetEffectiveEmptyFieldsSetting] 시트 '{currentSheetName}'의 {configKey} 설정: {sheetSpecificSetting}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetEffectiveEmptyFieldsSetting] 시트별 설정 확인 중 오류: {ex.Message}");
            }

            // 시트별 설정이나 전역 설정 중 하나라도 true이면 빈 필드 포함
            bool result = sheetSpecificSetting || globalAddEmptyFields;
            Debug.WriteLine($"[GetEffectiveEmptyFieldsSetting] 최종 설정 (시트별: {sheetSpecificSetting}, 전역: {globalAddEmptyFields}, 결과: {result})");
            
            return result;
        }

        /// <summary>
        /// 변환을 위한 시트 준비 및 유효성 검사를 수행합니다.
        /// </summary>
        /// <param name="outConvertibleSheets">변환 가능한 시트 목록입니다.</param>
        /// <returns>변환을 계속 진행할 수 있으면 true, 그렇지 않으면 false를 반환합니다.</returns>
        public static bool PrepareAndValidateSheets(out List<Microsoft.Office.Interop.Excel.Worksheet> outConvertibleSheets)
        {
            outConvertibleSheets = null;

            SheetPathManager.Instance.Initialize();
            Debug.WriteLine("[PrepareAndValidateSheets] SheetPathManager 초기화 완료");

            if (!ValidateAndInitializeWorkbook())
            {
                return false;
            }

            var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(Globals.ThisAddIn.Application.ActiveWorkbook);

            if (convertibleSheets.Count == 0)
            {
                MessageBox.Show("변환 가능한 시트가 없습니다. 변환하려는 시트 이름 앞에 '!'를 추가하세요.",
                    "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            ExcelConfigManager.Instance.EnsureConfigSheetExists();
            if (Properties.Settings.Default.FirstConfigMigration)
            {
                ExcelConfigManager.Instance.MigrateFromXmlSettings(SheetPathManager.Instance);
                Properties.Settings.Default.FirstConfigMigration = false;
                Properties.Settings.Default.Save();
                Debug.WriteLine("[PrepareAndValidateSheets] XML 설정을 Excel 설정으로 마이그레이션 완료");
            }

            int enabledSheetsCount = 0;
            Debug.WriteLine($"[PrepareAndValidateSheets] 변환 가능한 시트 수: {convertibleSheets.Count}");
            foreach (var sheet in convertibleSheets)
            {
                string currentSheetName = sheet.Name;
                bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(currentSheetName);
                Debug.WriteLine($"[PrepareAndValidateSheets] 시트 '{currentSheetName}' 활성화 상태: {isEnabled}");
                if (isEnabled)
                {
                    enabledSheetsCount++;
                }
            }
            Debug.WriteLine($"[PrepareAndValidateSheets] 활성화된 시트 수: {enabledSheetsCount}, 비활성화된 시트 수: {convertibleSheets.Count - enabledSheetsCount}");

            if (enabledSheetsCount == 0)
            {
                Debug.WriteLine("[PrepareAndValidateSheets] 경고: 활성화된 시트가 없습니다. 시트 활성화 상태 상세 정보 출력:");
                Dictionary<string, string> allEnabledPaths = SheetPathManager.Instance.GetAllEnabledSheetPaths();
                Debug.WriteLine($"[PrepareAndValidateSheets] GetAllEnabledSheetPaths 결과: {allEnabledPaths.Count}개 시트");
                foreach (var kvp in allEnabledPaths) Debug.WriteLine($"[PrepareAndValidateSheets] 활성화된 시트: '{kvp.Key}', 경로: '{kvp.Value}'");

                foreach (var sheet in convertibleSheets)
                {
                    string sheetName = sheet.Name;
                    bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);
                    if (isEnabled && !allEnabledPaths.ContainsKey(sheetName))
                    {
                        Debug.WriteLine($"[PrepareAndValidateSheets] 활성화 상태 불일치 감지. 시트 '{sheetName}'를 강제로 활성화합니다.");
                        SheetPathManager.Instance.SetSheetEnabled(sheetName, true);
                        enabledSheetsCount++;
                    }
                }

                if (enabledSheetsCount == 0)
                {
                    MessageBox.Show("활성화된 시트가 없어 변환을 취소합니다.\n\n시트 설정 창에서 시트를 활성화하십시오.", "변환 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    using (var form = new Forms.SheetPathSettingsForm(convertibleSheets)) 
                    { 
                        form.StartPosition = FormStartPosition.CenterScreen; 
                        form.ShowDialog(); 
                    }
                    return false;
                }
            }

            if (enabledSheetsCount < convertibleSheets.Count)
            {
                int disabledCount = convertibleSheets.Count - enabledSheetsCount;
                string message = $"{convertibleSheets.Count}개의 변환 가능한 시트 중 {disabledCount}개의 시트가 비활성화되어 있습니다.\n\n활성화된 {enabledSheetsCount}개의 시트만 변환하시겠습니까?\n\n아니오를 선택하면 시트별 경로 설정 창이 열립니다.";
                DialogResult result = MessageBox.Show(message, "시트 활성화 확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.No)
                {
                    var form = new Forms.SheetPathSettingsForm(convertibleSheets);
                    form.StartPosition = FormStartPosition.CenterScreen; 
                    form.ShowDialog();
                    return false;
                }
            }

            outConvertibleSheets = convertibleSheets;
            return true;
        }
    }
}