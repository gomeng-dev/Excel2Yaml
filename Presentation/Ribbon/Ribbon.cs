using ExcelToYamlAddin.Infrastructure.Configuration;
using ExcelToYamlAddin.Core;
using ExcelToYamlAddin.Application.PostProcessing;
using ExcelToYamlAddin.Forms;
using ExcelToYamlAddin.Infrastructure.Logging;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Threading; // CancellationToken을 위해 추가
using Excel = Microsoft.Office.Interop.Excel;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Application.Services;
using ExcelToYamlAddin.Infrastructure.Excel;

namespace ExcelToYamlAddin
{
    public partial class Ribbon : RibbonBase
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<Ribbon>();
        
        // 옵션 설정
        private bool includeEmptyFields = false;
        private bool enableHashGen = false;
        private bool addEmptyYamlFields = false;

        private readonly ExcelToYamlConfig config = new ExcelToYamlConfig();

        // 설정 폼 인스턴스 저장을 위한 필드
        private Forms.SheetPathSettingsForm settingsForm = null;

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();

            // 설정 불러오기
            try
            {
                // 초기값은 false로 설정(Ribbon_Load에서 시트별 설정을 로드)
                addEmptyYamlFields = false;

                // Properties.Settings.Default에서 기본 설정 로드
                addEmptyYamlFields = Properties.Settings.Default.AddEmptyYamlFields;
                Debug.WriteLine($"[Ribbon] 기본 설정값 로드: addEmptyYamlFields={addEmptyYamlFields}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Ribbon] 설정 로드 중 오류 발생: {ex.Message}");
            }
        }

        // 리본 로드 시 호출
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                Debug.WriteLine("리본 로드 시작");

                // SheetPathManager 인스턴스 초기화
                var pathManager = ExcelToYamlAddin.Infrastructure.Configuration.SheetPathManager.Instance;
                if (pathManager == null)
                {
                    Debug.WriteLine("[Ribbon_Load] SheetPathManager 인스턴스를 가져올 수 없습니다.");
                }

                // ExcelConfigManager 인스턴스 초기화
                var excelConfigManager = ExcelConfigManager.Instance;
                if (excelConfigManager == null)
                {
                    Debug.WriteLine("[Ribbon_Load] ExcelConfigManager 인스턴스를 가져올 수 없습니다.");
                }

                Debug.WriteLine("리본 UI가 로드되었습니다.");

                // 디버그 모드에서만 HTML 내보내기 버튼 표시
#if DEBUG
                btnExportToHtml.Visible = true;
                Debug.WriteLine("디버그 모드: HTML 내보내기 버튼 표시");
#else
                btnExportToHtml.Visible = false;
                Debug.WriteLine("릴리즈 모드: HTML 내보내기 버튼 숨김");
#endif

                // 활성 시트의 설정 로드
                try
                {
                    var app = Globals.ThisAddIn.Application;
                    if (app != null && app.ActiveSheet != null && excelConfigManager != null)
                    {
                        Excel.Worksheet activeSheet = app.ActiveSheet as Excel.Worksheet;
                        if (activeSheet != null)
                        {
                            string sheetName = activeSheet.Name;
                            addEmptyYamlFields = excelConfigManager.GetConfigBool(sheetName, "YamlEmptyFields", addEmptyYamlFields);
                            Debug.WriteLine($"[Ribbon_Load] 시트 '{sheetName}'에서 YAML 빈 배열 필드 설정 로드: {addEmptyYamlFields}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[Ribbon_Load] 활성 시트 설정 로드 중 오류: {ex.Message}");
                }

                // 설정 로드 확인
                if (pathManager != null)
                {
                    pathManager.Initialize(); // 설정 다시 로드

                    // 현재 워크북 설정
                    if (Globals.ThisAddIn.Application.ActiveWorkbook != null)
                    {
                        string workbookPath = Globals.ThisAddIn.Application.ActiveWorkbook.FullName;
                        pathManager.SetCurrentWorkbook(workbookPath);
                        Debug.WriteLine($"[Ribbon_Load] 현재 워크북 설정: {workbookPath}");

                        // 워크북의 모든 시트에 대한 설정 확인
                        foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets)
                        {
                            bool yamlOption = pathManager.GetYamlEmptyFieldsOption(sheet.Name);
                            Debug.WriteLine($"[Ribbon_Load] 시트 '{sheet.Name}' YAML 설정: {yamlOption}");
                        }
                    }
                    else
                    {
                        Debug.WriteLine("[Ribbon_Load] 활성화된 워크북이 없습니다.");
                    }
                }
                else
                {
                    Debug.WriteLine("[Ribbon_Load] SheetPathManager 인스턴스를 가져올 수 없습니다.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"리본 로드 중 오류: {ex.Message}");
            }
        }

        /// <summary>
        /// 변환을 위한 시트 준비 및 유효성 검사를 수행합니다.
        /// 공통 초기화, 시트 분석, 사용자 확인 등의 로직을 포함합니다.
        /// </summary>
        /// <param name="outConvertibleSheets">변환 가능한 시트 목록입니다.</param>
        /// <returns>변환을 계속 진행할 수 있으면 true, 그렇지 않으면 false를 반환합니다.</returns>
        private bool PrepareAndValidateSheets(out List<Excel.Worksheet> outConvertibleSheets)
        {
            outConvertibleSheets = null;

            SheetPathManager.Instance.Initialize();
            Debug.WriteLine("[PrepareAndValidateSheets] SheetPathManager 초기화 완료");

            var addIn = Globals.ThisAddIn;
            var app = addIn.Application;
            var activeWorkbook = app.ActiveWorkbook;

            if (activeWorkbook == null)
            {
                MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            string workbookPath = activeWorkbook.FullName;
            SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);
            ExcelConfigManager.Instance.SetCurrentWorkbook(workbookPath);

            var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(activeWorkbook);

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
                    using (var form = new Forms.SheetPathSettingsForm(convertibleSheets)) { form.StartPosition = FormStartPosition.CenterScreen; form.ShowDialog(); }
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
                    if (settingsForm != null && !settingsForm.IsDisposed) { settingsForm.Activate(); return false; }
                    
                    // 워크시트 이름들을 미리 추출 (COM 객체 무효화 방지)
                    var sheetNames = new List<string>();
                    foreach (var sheet in convertibleSheets)
                    {
                        try
                        {
                            sheetNames.Add(sheet.Name);
                        }
                        catch (System.Runtime.InteropServices.COMException ex)
                        {
                            Debug.WriteLine($"워크시트 이름 추출 실패 (COM Exception: {ex.ErrorCode:X})");
                        }
                    }
                    
                    settingsForm = new Forms.SheetPathSettingsForm(convertibleSheets);
                    settingsForm.FormClosed += (s, args) => { HandleSettingsFormClosedSafe(sheetNames); settingsForm = null; };
                    settingsForm.StartPosition = FormStartPosition.CenterScreen; settingsForm.Show();
                    return false;
                }
            }

            outConvertibleSheets = convertibleSheets;
            return true;
        }

        /// <summary>
        /// 지정된 YAML 파일 목록에 대해 공통 후처리 작업을 적용합니다.
        /// </summary>
        /// <param name="yamlFilePaths">후처리를 적용할 YAML 파일 경로 목록입니다.</param>
        /// <param name="convertibleSheets">변환 가능한 원본 Excel 시트 목록입니다. 파일 이름과 시트 설정을 매칭하는 데 사용됩니다.</param>
        /// <param name="progress">진행 상태를 보고할 IProgress 객체입니다.</param>
        /// <param name="cancellationToken">작업 취소를 위한 CancellationToken입니다.</param>
        /// <param name="initialProgressPercentage">이 후처리 단계가 시작될 때의 전체 진행률입니다.</param>
        /// <param name="progressRange">이 후처리 단계가 전체 진행률에서 차지하는 범위입니다.</param>
        /// <param name="isForJsonConversion">JSON으로 변환하기 위한 중간 단계의 YAML 파일에 대한 후처리인 경우 true로 설정합니다. 이 경우 일부 후처리 단계가 생략될 수 있습니다.</param>
        /// <returns>키 경로 병합 및 Flow 스타일 처리 성공 횟수를 포함하는 튜플을 반환합니다.</returns>
        private (int mergeKeyPathsSuccessCount, int flowStyleSuccessCount) ApplyYamlPostProcessing(
            List<string> yamlFilePaths,
            List<Excel.Worksheet> convertibleSheets,
            IProgress<Forms.ProgressForm.ProgressInfo> progress,
            CancellationToken cancellationToken,
            int initialProgressPercentage,
            int progressRange,
            bool isForJsonConversion = false)
        {
            int mergeKeyPathsSuccessCount = 0;
            int flowStyleSuccessCount = 0;
            int filesProcessedInThisStep = 0;
            int totalFilesToProcessInThisStep = yamlFilePaths.Count;

            foreach (var yamlFilePath in yamlFilePaths)
            {
                cancellationToken.ThrowIfCancellationRequested();

                string fileName = Path.GetFileNameWithoutExtension(yamlFilePath);
                bool mergeKeyPathsProcessingAttemptedThisFile = false;
                bool flowStyleProcessingAttemptedThisFile = false;

                progress.Report(new Forms.ProgressForm.ProgressInfo
                {
                    Percentage = initialProgressPercentage + (int)((double)filesProcessedInThisStep / totalFilesToProcessInThisStep * progressRange),
                    StatusMessage = $"'{fileName}' YAML 후처리 중..."
                });

                string matchedSheetName = null;
                foreach (var sheet in convertibleSheets)
                {
                    string currentSheetNameForMatch = sheet.Name;
                    if (currentSheetNameForMatch.StartsWith("!"))
                        currentSheetNameForMatch = currentSheetNameForMatch.Substring(1);

                    if (string.Compare(currentSheetNameForMatch, fileName, StringComparison.OrdinalIgnoreCase) == 0)
                    {
                        matchedSheetName = sheet.Name;
                        break;
                    }
                }

                if (matchedSheetName != null)
                {
                    bool yamlEmptyFieldsOption = ExcelConfigManager.Instance.GetConfigBool(matchedSheetName, "YamlEmptyFields", false);
                    if (!yamlEmptyFieldsOption) yamlEmptyFieldsOption = SheetPathManager.Instance.GetYamlEmptyFieldsOption(matchedSheetName);
                    if (!yamlEmptyFieldsOption && addEmptyYamlFields) yamlEmptyFieldsOption = addEmptyYamlFields;

                    progress.Report(new Forms.ProgressForm.ProgressInfo { StatusMessage = $"'{fileName}' - 키 경로 병합 처리 중..." });
                    string sheetMergeKeyPaths = ExcelConfigManager.Instance.GetConfigValue(matchedSheetName, "MergeKeyPaths", "");
                    if (string.IsNullOrEmpty(sheetMergeKeyPaths)) sheetMergeKeyPaths = SheetPathManager.Instance.GetMergeKeyPaths(matchedSheetName);

                    if (!string.IsNullOrEmpty(sheetMergeKeyPaths))
                    {
                        mergeKeyPathsProcessingAttemptedThisFile = true;
                        Debug.WriteLine($"[ApplyYamlPostProcessing] YAML 병합 후처리 실행: {yamlFilePath}, 설정: {sheetMergeKeyPaths}");
                        bool success = YamlMergeKeyPathsProcessor.ProcessYamlFileFromConfig(yamlFilePath, sheetMergeKeyPaths, yamlEmptyFieldsOption);
                        if (success) { Debug.WriteLine($"[ApplyYamlPostProcessing] YAML 병합 후처리 완료: {yamlFilePath}"); mergeKeyPathsSuccessCount++; }
                        else { Debug.WriteLine($"[ApplyYamlPostProcessing] YAML 병합 후처리 실패: {yamlFilePath}"); }
                    }

                    if (!isForJsonConversion)
                    {
                        progress.Report(new Forms.ProgressForm.ProgressInfo { StatusMessage = $"'{fileName}' - Flow Style 처리 중..." });
                        string sheetFlowStyle = ExcelConfigManager.Instance.GetConfigValue(matchedSheetName, "FlowStyle", "");
                        if (string.IsNullOrWhiteSpace(sheetFlowStyle)) sheetFlowStyle = SheetPathManager.Instance.GetFlowStyleConfig(matchedSheetName ?? fileName);

                        if (!YamlFlowStyleProcessor.IsConfigEffectivelyEmpty(sheetFlowStyle))
                        {
                            flowStyleProcessingAttemptedThisFile = true;
                            Debug.WriteLine($"[ApplyYamlPostProcessing] YAML 흐름 스타일 후처리 실행: {yamlFilePath}, 설정: {sheetFlowStyle}");
                            bool success = YamlFlowStyleProcessor.ProcessYamlFileFromConfig(yamlFilePath, sheetFlowStyle);
                            if (success) { Debug.WriteLine($"[ApplyYamlPostProcessing] YAML 흐름 스타일 후처리 완료: {yamlFilePath}"); flowStyleSuccessCount++; }
                            else { Debug.WriteLine($"[ApplyYamlPostProcessing] YAML 흐름 스타일 후처리 실패: {yamlFilePath}"); }
                        }
                        else { Debug.WriteLine($"[ApplyYamlPostProcessing] YAML 흐름 스타일 후처리 건너뜀: {yamlFilePath}, 설정: '{sheetFlowStyle}'"); }

                        progress.Report(new Forms.ProgressForm.ProgressInfo { StatusMessage = $"'{fileName}' - 빈 배열 처리 중..." });
                        bool processEmptyArrays = yamlEmptyFieldsOption || addEmptyYamlFields;
                        if (processEmptyArrays) { Debug.WriteLine($"[ApplyYamlPostProcessing] YAML 빈 배열 처리: OrderedYamlFactory에서 처리 (파일: {yamlFilePath}), 시트별: {yamlEmptyFieldsOption}, 전역: {addEmptyYamlFields}"); }
                        else { Debug.WriteLine($"[ApplyYamlPostProcessing] YAML 빈 배열 처리 건너뜀: 관련 옵션 비활성화. 시트별: {yamlEmptyFieldsOption}, 전역: {addEmptyYamlFields}");}

                        if (!mergeKeyPathsProcessingAttemptedThisFile && !flowStyleProcessingAttemptedThisFile)
                        {
                            progress.Report(new Forms.ProgressForm.ProgressInfo { StatusMessage = $"'{fileName}' - 최종 문자열 정리 중..." });
                            Debug.WriteLine($"[ApplyYamlPostProcessing] 최종 Raw 문자열 변환 후처리 실행: {yamlFilePath}");
                            new Application.PostProcessing.FinalRawStringConverter().ProcessYamlFile(yamlFilePath);
                        }
                        else { Debug.WriteLine($"[ApplyYamlPostProcessing] 최종 Raw 문자열 변환 건너뜀 (Merge: {mergeKeyPathsProcessingAttemptedThisFile}, Flow: {flowStyleProcessingAttemptedThisFile}): {yamlFilePath}"); }
                    }
                }
                filesProcessedInThisStep++;
            }
            return (mergeKeyPathsSuccessCount, flowStyleSuccessCount);
        }

        // YAML으로 변환 버튼 클릭
        public void OnConvertToYamlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 설정 초기화 및 다시 로드
                List<Excel.Worksheet> convertibleSheets;
                if (!PrepareAndValidateSheets(out convertibleSheets))
                {
                    return;
                }

                // YAML 변환 설정
                config.OutputFormat = OutputFormat.Yaml;

                // 설정 적용 (포맷에 따라 다른 옵션 적용)
                if (config.OutputFormat == OutputFormat.Yaml)
                {
                    // YAML 변환 시 시트별 설정과 전역 설정을 모두 고려합니다
                    // 현재 활성화된 시트의 시트별 설정 확인
                    bool sheetSpecificSetting = false;

                    try
                    {
                        // 현재 활성 시트의 설정 확인
                        if (Globals.ThisAddIn.Application.ActiveSheet != null)
                        {
                            string currentSheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
                            sheetSpecificSetting = ExcelConfigManager.Instance.GetConfigBool(currentSheetName, "YamlEmptyFields", false);
                            Debug.WriteLine($"[OnConvertToYamlClick] 현재 시트 '{currentSheetName}'의 YamlEmptyFields 설정: {sheetSpecificSetting}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[OnConvertToYamlClick] 시트별 설정 확인 중 오류: {ex.Message}");
                    }

                    // 시트별 설정이나 전역 설정 중 하나라도 true이면 빈 필드 포함
                    bool includeEmptySettings = sheetSpecificSetting || addEmptyYamlFields;
                    config.IncludeEmptyFields = includeEmptySettings;

                    Debug.WriteLine($"[OnConvertToYamlClick] YAML 변환 시 빈 필드 포함 설정: {config.IncludeEmptyFields} (시트별 설정: {sheetSpecificSetting}, 전역 설정: {addEmptyYamlFields})");
                    Debug.WriteLine($"[OnConvertToYamlClick] 최종 설정값 추적: 리본 변수값={addEmptyYamlFields}, 시트별 설정값={sheetSpecificSetting}, Config 최종 설정값={config.IncludeEmptyFields}");
                }
                else
                {
                    // JSON 변환 시에는 includeEmptyFields 옵션만 사용
                    config.IncludeEmptyFields = includeEmptyFields;
                    Debug.WriteLine($"[OnConvertToJsonClick] JSON 변환 시 빈 필드 포함 설정: {config.IncludeEmptyFields} (includeEmptyFields: {includeEmptyFields})");
                }

                config.EnableHashGen = enableHashGen;

                // 변환 전 설정 다시 로드 및 동기화
                SheetPathManager.Instance.Initialize();
                Debug.WriteLine("[OnConvertToYamlClick] 변환 전 SheetPathManager 재초기화 완료");

                // 변환 처리
                List<string> convertedFiles = ConvertExcelFile(config);

                // 변환 결과 추적
                int successCount = 0;
                int currentPostProcessMergeSuccessCount = 0; // 스코프 문제 해결을 위해 이름 변경 및 외부 선언
                int currentPostProcessFlowSuccessCount = 0;  // 스코프 문제 해결을 위해 이름 변경 및 외부 선언

                // YAML 후처리 기능 적용
                if (convertedFiles != null && convertedFiles.Count > 0)
                {
                    try
                    {
                        Debug.WriteLine($"[Ribbon] YAML 후처리 확인: {convertedFiles.Count}개 파일");
                        successCount = convertedFiles.Count;


                        // 후처리를 위한 프로그레스 바 표시
                        using (var progressForm = new Forms.ProgressForm())
                        {
                            progressForm.RunOperation((progress, cancellationToken) =>
                            {
                                try
                                {
                                    (currentPostProcessMergeSuccessCount, currentPostProcessFlowSuccessCount) = ApplyYamlPostProcessing(
                                        convertedFiles, convertibleSheets, progress, cancellationToken,
                                        initialProgressPercentage: 0, progressRange: 100, isForJsonConversion: false);

                                    // 모든 작업 완료
                                    progress.Report(new Forms.ProgressForm.ProgressInfo
                                    {
                                        Percentage = 100,
                                        StatusMessage = "모든 파일 처리 완료",
                                        IsCompleted = true
                                    });
                                }
                                catch (OperationCanceledException)
                                {
                                    // 취소 처리
                                    progress.Report(new Forms.ProgressForm.ProgressInfo
                                    {
                                        Percentage = 100,
                                        StatusMessage = "후처리가 취소되었습니다.",
                                        IsCompleted = true
                                    });
                                }
                            }, "YAML 후처리 중...");

                            progressForm.ShowDialog();

                            // 취소된 경우
                            if (progressForm.DialogResult == DialogResult.Cancel)
                            {
                                MessageBox.Show("후처리 작업이 취소되었습니다.", "작업 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[Ribbon] YAML 후처리 중 오류 발생: {ex.Message}");
                    }

                    // 모든 후처리 작업이 완료된 후 메시지 표시
                    string message = $"{successCount}개의 시트가 성공적으로 변환되었습니다.";

                    if (currentPostProcessMergeSuccessCount > 0)
                        message += $"\n키 경로 병합 처리: {currentPostProcessMergeSuccessCount}개 파일";

                    if (currentPostProcessFlowSuccessCount > 0)
                        message += $"\nFlow 스타일 처리: {currentPostProcessFlowSuccessCount}개 파일";

                    if (convertedFiles.Count > 0)
                    {
                        message += "\n\n변환된 파일:";
                        foreach (var file in convertedFiles.Take(5))  // 첫 5개만 표시
                        {
                            message += $"\n{file}";
                        }

                        if (convertedFiles.Count > 5)
                        {
                            message += $"\n... 외 {convertedFiles.Count - 5}개 파일";
                        }
                    }

                    MessageBox.Show(message, "변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("변환된 시트가 없습니다. 시트별 저장 경로를 설정했는지 확인하세요.",
                        "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"YAML 변환 중 오류가 발생했습니다: {ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"[Ribbon] YAML 변환 오류: {ex.Message}");
                Debug.WriteLine($"[Ribbon] 스택 트레이스: {ex.StackTrace}");
            }
        }

        // XML로 변환 버튼 클릭 (새로 추가된 메소드)
        public void OnConvertToXmlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                List<Excel.Worksheet> convertibleSheets;
                if (!PrepareAndValidateSheets(out convertibleSheets))
                {
                    return;
                }

                // XML 변환 설정
                // 1단계: YAML로 먼저 변환하므로 OutputFormat.Yaml 설정
                config.OutputFormat = OutputFormat.Yaml;

                // 현재 활성 시트 또는 전역 설정에 따라 빈 필드 포함 여부 결정
                bool sheetSpecificEmptyFieldSetting = false;
                try
                {
                    if (Globals.ThisAddIn.Application.ActiveSheet != null)
                    {
                        string currentSheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
                        sheetSpecificEmptyFieldSetting = ExcelConfigManager.Instance.GetConfigBool(currentSheetName, "YamlEmptyFields", false);
                        Debug.WriteLine($"[OnConvertToXmlClick] 현재 시트 '{currentSheetName}'의 YamlEmptyFields 설정 (YAML 단계용): {sheetSpecificEmptyFieldSetting}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[OnConvertToXmlClick] 시트별 YamlEmptyFields 설정 확인 중 오류 (YAML 단계용): {ex.Message}");
                }
                config.IncludeEmptyFields = sheetSpecificEmptyFieldSetting || addEmptyYamlFields;
                Debug.WriteLine($"[OnConvertToXmlClick] YAML 변환 시 빈 필드 포함 설정: {config.IncludeEmptyFields} (시트별: {sheetSpecificEmptyFieldSetting}, 전역: {addEmptyYamlFields})");

                config.EnableHashGen = enableHashGen; // 해시 생성 옵션 (YAML 단계용)

                SheetPathManager.Instance.Initialize();
                Debug.WriteLine("[OnConvertToXmlClick] 변환 전 SheetPathManager 재초기화 완료");

                // 프로그레스 바 표시 및 변환 작업 수행
                using (var progressForm = new Forms.ProgressForm())
                {
                    List<string> finalXmlFiles = new List<string>();
                    progressForm.RunOperation((progress, cancellationToken) =>
                    {
                        progress.Report(new Forms.ProgressForm.ProgressInfo
                        {
                            Percentage = 0,
                            StatusMessage = "Excel → YAML 변환 준비 중..."
                        });

                        // 1단계: Excel을 YAML로 변환 (임시 파일에 저장)
                        string tempDir = Path.Combine(Path.GetTempPath(), "ExcelToXmlTemp_" + Guid.NewGuid().ToString().Substring(0, 8));
                        Directory.CreateDirectory(tempDir);

                        ExcelToYamlConfig yamlConfig = new ExcelToYamlConfig
                        {
                            IncludeEmptyFields = config.IncludeEmptyFields,
                            EnableHashGen = config.EnableHashGen, // YAML 단계에서 해시 생성은 선택사항
                            OutputFormat = OutputFormat.Yaml,
                            WorkingDirectory = tempDir // 임시 YAML 파일 저장 경로
                        };

                        List<string> tempYamlFiles = new List<string>();

                        try
                        {
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 10,
                                StatusMessage = "Excel → YAML 파일 생성 중..."
                            });

                            // Excel을 임시 YAML 파일로 변환
                            tempYamlFiles = ConvertExcelFileToTemp(yamlConfig, tempDir, convertibleSheets);

                            if (tempYamlFiles.Count == 0)
                            {
                                progress.Report(new Forms.ProgressForm.ProgressInfo { Percentage = 100, StatusMessage = "변환할 YAML 파일이 없습니다.", IsCompleted = true });
                                return;
                            }

                            // 2단계: YAML 후처리 (MergeKeyPaths, FlowStyle 등)
                            progress.Report(new Forms.ProgressForm.ProgressInfo { Percentage = 30, StatusMessage = "YAML 후처리 진행 중..." });
                            ApplyYamlPostProcessing(tempYamlFiles, convertibleSheets, progress, cancellationToken, 30, 30, isForJsonConversion: false); // isForJsonConversion = false로 모든 후처리 적용

                            // 3단계: 후처리된 YAML을 XML로 변환
                            progress.Report(new Forms.ProgressForm.ProgressInfo { Percentage = 60, StatusMessage = "YAML → XML 변환 중..." });

                            var yamlParser = new YamlDotNet.Serialization.DeserializerBuilder().Build();
                            int processedXmlCount = 0;

                            foreach (var yamlFilePath in tempYamlFiles)
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                                string sheetFileName = Path.GetFileNameWithoutExtension(yamlFilePath); // 예: "Sheet1"
                                string originalSheetName = convertibleSheets.FirstOrDefault(s => 
                                    (s.Name.StartsWith("!") ? s.Name.Substring(1) : s.Name).Equals(sheetFileName, StringComparison.OrdinalIgnoreCase))?.Name ?? sheetFileName;

                                progress.Report(new Forms.ProgressForm.ProgressInfo
                                {
                                    Percentage = 60 + (int)((double)processedXmlCount / tempYamlFiles.Count * 35),
                                    StatusMessage = $"'{sheetFileName}' YAML → XML 변환 중..."
                                });

                                string yamlContent = File.ReadAllText(yamlFilePath);
                                
                                // YAML이 배열인지 객체인지 동적으로 판단하여 역직렬화
                                object deserializedYaml = yamlParser.Deserialize<object>(yamlContent);
                                
                                IDictionary<string, object> dataForXml;
                                string xmlRootElementName = sheetFileName; // 기본값

                                if (deserializedYaml is IDictionary<string, object> yamlDict)
                                {
                                    // YAML이 객체인 경우
                                    if (yamlDict.Count == 1 && yamlDict.Values.First() is IDictionary<string, object> innerData)
                                    {
                                        xmlRootElementName = yamlDict.Keys.First();
                                        dataForXml = innerData;
                                    }
                                    else
                                    {
                                        dataForXml = yamlDict;
                                    }
                                }
                                else if (deserializedYaml is IList<object> yamlList)
                                {
                                    // YAML이 배열인 경우 (루트 요소 제거 후 배열이 된 경우)
                                    dataForXml = new Dictionary<string, object>
                                    {
                                        { "Items", yamlList }  // 배열을 Items로 감싸서 XML 구조 생성
                                    };
                                    xmlRootElementName = sheetFileName;
                                }
                                else
                                {
                                    // 기타 경우 (스칼라 값 등)
                                    dataForXml = new Dictionary<string, object>
                                    {
                                        { "Value", deserializedYaml }
                                    };
                                    xmlRootElementName = sheetFileName;
                                }
                                
                                string xmlString = Application.PostProcessing.YamlToXmlConverter.ConvertToXmlString(dataForXml, xmlRootElementName);

                                string savePath = SheetPathManager.Instance.GetSheetPath(originalSheetName);
                                if (string.IsNullOrEmpty(savePath)) savePath = Path.GetDirectoryName(Globals.ThisAddIn.Application.ActiveWorkbook.FullName); // 기본 경로

                                string xmlFilePath = Path.Combine(savePath, $"{sheetFileName}.xml");
                                Directory.CreateDirectory(Path.GetDirectoryName(xmlFilePath)); // 경로 생성
                                File.WriteAllText(xmlFilePath, xmlString);
                                finalXmlFiles.Add(xmlFilePath);
                                processedXmlCount++;
                            }

                            progress.Report(new Forms.ProgressForm.ProgressInfo { Percentage = 100, StatusMessage = $"{finalXmlFiles.Count}개 파일 XML 변환 완료", IsCompleted = true });
                        }
                        catch (OperationCanceledException)
                        {
                            progress.Report(new Forms.ProgressForm.ProgressInfo { Percentage = 100, StatusMessage = "XML 변환 작업이 취소되었습니다.", IsCompleted = true });
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[OnConvertToXmlClick] XML 변환 중 오류 발생: {ex.Message}\n{ex.StackTrace}");
                            progress.Report(new Forms.ProgressForm.ProgressInfo { Percentage = 100, StatusMessage = $"오류 발생: {ex.Message}", IsCompleted = true, HasError = true, ErrorMessage = ex.Message });
                        }
                        finally
                        {
                            try { if (Directory.Exists(tempDir)) Directory.Delete(tempDir, true); }
                            catch (Exception ex) { Debug.WriteLine($"[OnConvertToXmlClick] 임시 디렉토리 정리 중 오류: {ex.Message}"); }
                        }
                    }, "Excel → YAML → XML 변환 중...");

                    progressForm.ShowDialog();

                    if (progressForm.DialogResult != DialogResult.Cancel && finalXmlFiles.Count > 0)
                    {
                        MessageBox.Show($"{finalXmlFiles.Count}개의 시트가 성공적으로 XML로 변환되었습니다.", "XML 변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else if (progressForm.DialogResult != DialogResult.Cancel)
                    {
                        MessageBox.Show("변환된 XML 파일이 없습니다.", "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"XML 변환 중 오류가 발생했습니다: {ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"[Ribbon] XML 변환 오류: {ex.Message}");
                Debug.WriteLine($"[Ribbon] 스택 트레이스: {ex.StackTrace}");
            }
        }


        // YAML을 JSON으로 변환 버튼 클릭
        public void OnConvertYamlToJsonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                List<Excel.Worksheet> convertibleSheets;
                if (!PrepareAndValidateSheets(out convertibleSheets))
                {
                    return;
                }

                // 설정 적용 - YAML 변환으로 설정
                config.IncludeEmptyFields = addEmptyYamlFields;
                config.EnableHashGen = enableHashGen;
                config.OutputFormat = OutputFormat.Yaml;

                // YAML 변환 시 시트별 설정과 전역 설정을 모두 고려
                bool sheetSpecificSetting = false;

                try
                {
                    // 현재 활성 시트의 설정 확인
                    if (Globals.ThisAddIn.Application.ActiveSheet != null)
                    {
                        string currentSheetName = Globals.ThisAddIn.Application.ActiveSheet.Name;
                        sheetSpecificSetting = ExcelConfigManager.Instance.GetConfigBool(currentSheetName, "YamlEmptyFields", false);
                        Debug.WriteLine($"[OnConvertYamlToJsonClick] 현재 시트 '{currentSheetName}'의 YamlEmptyFields 설정: {sheetSpecificSetting}");
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[OnConvertYamlToJsonClick] 시트별 설정 확인 중 오류: {ex.Message}");
                }

                // 시트별 설정이나 전역 설정 중 하나라도 true이면 빈 필드 포함
                bool includeEmptySettings = sheetSpecificSetting || addEmptyYamlFields;
                config.IncludeEmptyFields = includeEmptySettings;

                // 변환 전 설정 다시 로드 및 동기화
                SheetPathManager.Instance.Initialize();
                Debug.WriteLine("[OnConvertYamlToJsonClick] 변환 전 SheetPathManager 재초기화 완료");

                // 프로그레스 바 표시 및 변환 작업 수행
                using (var progressForm = new Forms.ProgressForm())
                {
                    progressForm.RunOperation((progress, cancellationToken) =>
                    {
                        progress.Report(new Forms.ProgressForm.ProgressInfo
                        {
                            Percentage = 0,
                            StatusMessage = "YAML 파일 변환 준비 중..."
                        });

                        // 1단계: YAML로 변환 (임시 파일에 저장)
                        // 임시 디렉토리 생성
                        string tempDir = Path.Combine(Path.GetTempPath(), "Excel2YamlTemp_" + Guid.NewGuid().ToString().Substring(0, 8));
                        Directory.CreateDirectory(tempDir);

                        // 설정 복제
                        ExcelToYamlConfig yamlConfig = new ExcelToYamlConfig
                        {
                            IncludeEmptyFields = config.IncludeEmptyFields,
                            EnableHashGen = config.EnableHashGen,
                            OutputFormat = OutputFormat.Yaml,
                            WorkingDirectory = tempDir
                        };

                        List<string> yamlFiles = new List<string>();
                        List<Tuple<string, string>> convertPairs = new List<Tuple<string, string>>();

                        try
                        {
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 10,
                                StatusMessage = "YAML 파일 생성 중..."
                            });

                            // YAML 파일 생성
                            yamlFiles = ConvertExcelFileToTemp(yamlConfig, tempDir, convertibleSheets); // convertibleSheets 전달

                            if (yamlFiles.Count == 0)
                            {
                                progress.Report(new Forms.ProgressForm.ProgressInfo
                                {
                                    Percentage = 100,
                                    StatusMessage = "변환할 YAML 파일이 없습니다.",
                                    IsCompleted = true
                                });
                                return;
                            }

                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 30,
                                StatusMessage = "YAML 후처리 진행 중..."
                            });
                            ApplyYamlPostProcessing(yamlFiles, convertibleSheets, progress, cancellationToken, 30, 20, isForJsonConversion: true);

                            // 3단계: JSON 변환 준비
                            progress.Report(new ProgressForm.ProgressInfo
                            {
                                Percentage = 50,
                                StatusMessage = $"{yamlFiles.Count}개 YAML 파일 JSON 변환 준비 중..."
                            });

                            // YAML 파일을 JSON으로 변환할 쌍 생성
                            foreach (var yamlFile in yamlFiles)
                            {
                                cancellationToken.ThrowIfCancellationRequested();

                                string fileName = Path.GetFileNameWithoutExtension(yamlFile);
                                string sheetName = fileName.StartsWith("!") ? fileName : "!" + fileName;

                                // 실제 저장 경로 가져오기
                                string savePath = SheetPathManager.Instance.GetSheetPath(sheetName);
                                if (string.IsNullOrEmpty(savePath))
                                {
                                    // !가 없는 이름으로 시도
                                    savePath = SheetPathManager.Instance.GetSheetPath(fileName);
                                }

                                if (!string.IsNullOrEmpty(savePath))
                                {
                                    string jsonFilePath = Path.Combine(savePath, fileName + ".json");
                                    convertPairs.Add(new Tuple<string, string>(yamlFile, jsonFilePath));
                                }
                            }

                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 70,
                                StatusMessage = "YAML에서 JSON으로 변환 중..."
                            });

                            // YAML을 JSON으로 변환
                            List<string> convertedJsonFiles = Application.PostProcessing.YamlToJsonProcessor.BatchConvertYamlToJson(convertPairs);

                            // 결과 보고
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 100,
                                StatusMessage = $"{convertedJsonFiles.Count}개 파일 변환 완료",
                                IsCompleted = true
                            });
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[OnConvertYamlToJsonClick] 변환 중 오류 발생: {ex.Message}");
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 100,
                                StatusMessage = $"오류 발생: {ex.Message}",
                                IsCompleted = true
                            });
                        }
                        finally
                        {
                            // 임시 디렉토리 정리
                            try
                            {
                                if (Directory.Exists(tempDir))
                                {
                                    Directory.Delete(tempDir, true);
                                }
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[OnConvertYamlToJsonClick] 임시 디렉토리 정리 중 오류: {ex.Message}");
                            }
                        }
                    }, "YAML → JSON 변환 중...");

                    progressForm.ShowDialog();

                    // 변환 완료 메시지
                    if (progressForm.DialogResult != DialogResult.Cancel)
                    {
                        MessageBox.Show("YAML → JSON 변환이 완료되었습니다.", "변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"YAML → JSON 변환 중 오류가 발생했습니다: {ex.Message}",
                    "변환 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"[OnConvertYamlToJsonClick] 오류: {ex.Message}\n{ex.StackTrace}");
            }
        }


        // 빈 필드 포함 옵션 체크박스 상태 가져오기
        public bool GetEmptyFieldsState(IRibbonControl control)
        {
            return includeEmptyFields;
        }

        // 빈 필드 포함 옵션 체크박스 클릭
        public void OnEmptyFieldsClicked(IRibbonControl control, bool pressed)
        {
            includeEmptyFields = pressed;
        }

        // MD5 해시 생성 옵션 체크박스 상태 가져오기
        public bool GetHashGenState(IRibbonControl control)
        {
            return enableHashGen;
        }

        // MD5 해시 생성 옵션 체크박스 클릭
        public void OnHashGenClicked(IRibbonControl control, bool pressed)
        {
            enableHashGen = pressed;
        }

        // YAML 선택적 필드 처리 옵션 체크박스 상태 가져오기
        public bool GetAddEmptyYamlState(IRibbonControl control)
        {
            return addEmptyYamlFields;
        }

        // YAML 선택적 필드 처리 옵션 체크박스 클릭
        public void OnAddEmptyYamlClicked(IRibbonControl control, bool pressed)
        {
            addEmptyYamlFields = pressed;
            Debug.WriteLine($"[Ribbon] YAML 빈 배열 필드 추가 옵션 변경: {addEmptyYamlFields}");

            try
            {
                // Properties.Settings.Default에 설정 저장
                Properties.Settings.Default.AddEmptyYamlFields = pressed;
                Properties.Settings.Default.Save();
                Debug.WriteLine($"[Ribbon] YAML 기본 설정 저장: {pressed}");

                // Globals.ThisAddIn이 초기화되었는지 확인 후 시트별 설정 저장
                if (Globals.ThisAddIn != null && Globals.ThisAddIn.Application != null &&
                    Globals.ThisAddIn.Application.ActiveSheet != null)
                {
                    // 체크박스 상태가 변경되면 Excel Config에도 저장 (현재 선택된 시트에만 적용)
                    var activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
                    if (activeWorksheet != null)
                    {
                        string sheetName = activeWorksheet.Name;
                        Debug.WriteLine($"[Ribbon] 현재 시트 '{sheetName}'에 YAML 빈 배열 필드 설정 저장: {addEmptyYamlFields}");
                        ExcelConfigManager.Instance.SetConfigValue(sheetName, "YamlEmptyFields", addEmptyYamlFields.ToString().ToLower());
                    }
                }
                else
                {
                    Debug.WriteLine("[Ribbon] ThisAddIn이 초기화되지 않았거나 활성 시트가 없어 시트별 설정을 저장하지 않습니다.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Ribbon] 설정 저장 중 오류 발생: {ex.Message}");
            }
        }

        // 고급 설정 버튼 클릭
        public void OnSettingsClick(IRibbonControl control)
        {
            OnSheetPathSettingsClick(null, null);
        }

        // 도움말 버튼 클릭 이벤트 핸들러
        private void OnHelpButtonClick(object sender, RibbonControlEventArgs e)
        {
            OnHelpClick(null);
        }

        public void OnHelpClick(IRibbonControl control)
        {
            try
            {
                // 임베디드 리소스에서 HTML 내용 로드
                string htmlContent = null;
                using (Stream stream = System.Reflection.Assembly.GetExecutingAssembly()
                    .GetManifestResourceStream("ExcelToYamlAddin.Docs.Readme.html"))
                {
                    if (stream == null)
                    {
                        // 임베디드 리소스를 찾을 수 없는 경우, 물리적 파일을 시도해봅니다.
                        string addinPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                        string readmePath = Path.Combine(addinPath, "Docs", "Readme.html");

                        if (File.Exists(readmePath))
                        {
                            // 물리적 파일이 존재하면 내용을 읽어옵니다.
                            htmlContent = File.ReadAllText(readmePath, Encoding.UTF8);
                        }
                        else
                        {
                            MessageBox.Show("도움말 리소스를 찾을 수 없습니다.",
                                           "리소스 없음",
                                           MessageBoxButtons.OK,
                                           MessageBoxIcon.Warning);
                            Debug.WriteLine("[OnHelpClick] 도움말 리소스를 찾을 수 없습니다.");
                            return;
                        }
                    }
                    else
                    {
                        // 임베디드 리소스가 있으면 내용을 읽어옵니다.
                        using (StreamReader reader = new StreamReader(stream))
                        {
                            htmlContent = reader.ReadToEnd();
                        }
                    }
                }

                // 도움말 폼 생성 (비모달)
                Form helpForm = new Form()
                {
                    Text = "Excel2YAML 사용 설명서",
                    Size = new Size(1000, 700),
                    StartPosition = FormStartPosition.CenterScreen,
                    FormBorderStyle = FormBorderStyle.SizableToolWindow,
                    ShowInTaskbar = false,
                    TopMost = true
                };

                var browser = new WebBrowser();
                browser.Dock = DockStyle.Fill;
                browser.ScriptErrorsSuppressed = true;
                browser.DocumentText = htmlContent;

                helpForm.Controls.Add(browser);
                helpForm.Show(); // 비모달로 표시

                Debug.WriteLine("[OnHelpClick] 도움말 내용을 플로팅 윈도우에 표시했습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"도움말 파일을 열 수 없습니다: {ex.Message}",
                               "오류",
                               MessageBoxButtons.OK,
                               MessageBoxIcon.Error);
                Debug.WriteLine($"[OnHelpClick] 오류: {ex.Message}");
            }
        }

        // 시트별 경로 설정 버튼 클릭
        public void OnSheetPathSettingsClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 현재 워크북 가져오기
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;

                if (app.ActiveWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 워크북 경로 설정
                string workbookPath = app.ActiveWorkbook.FullName;
                SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);

                // 변환 가능한 시트 찾기
                var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(app.ActiveWorkbook);

                if (convertibleSheets.Count == 0)
                {
                    MessageBox.Show("변환 가능한 시트가 없습니다. 변환하려는 시트 이름 앞에 '!'를 추가하세요.",
                        "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 시트별 경로 설정 폼 열기
                using (var form = new SheetPathSettingsForm(convertibleSheets))
                {
                    form.StartPosition = FormStartPosition.CenterScreen;
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"시트별 경로 설정 중 오류가 발생했습니다: {ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                Debug.WriteLine($"시트별 경로 설정 오류: {ex}");
            }
        }

        // 시트별 경로 설정 대화상자 표시
        private void ManageSheetPathsButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 현재 워크북 가져오기
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;

                if (app.ActiveWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 변환 가능한 시트 가져오기
                var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(app.ActiveWorkbook);

                if (convertibleSheets.Count == 0)
                {
                    if (MessageBox.Show("변환 가능한 시트(이름이 !로 시작하는 시트)가 없습니다. 시트 설정 화면을 여시겠습니까?",
                        "알림", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        // 빈 목록으로 폼 열기
                        convertibleSheets = new List<Worksheet>();
                        foreach (Worksheet sheet in app.ActiveWorkbook.Sheets)
                        {
                            convertibleSheets.Add(sheet);
                        }
                    }
                    else
                    {
                        return;
                    }
                }

                // 워크북 경로 설정
                SheetPathManager.Instance.SetCurrentWorkbook(app.ActiveWorkbook.FullName);

                // 이미 열려있는 설정 폼이 있으면 활성화
                if (settingsForm != null && !settingsForm.IsDisposed)
                {
                    settingsForm.Activate();
                    return;
                }

                // 시트별 경로 설정 대화상자 표시 (비모달)
                settingsForm = new Forms.SheetPathSettingsForm(convertibleSheets);
                settingsForm.FormClosed += (s, args) => { HandleSettingsFormClosed(convertibleSheets); settingsForm = null; };
                settingsForm.StartPosition = FormStartPosition.CenterScreen;
                settingsForm.Show();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"도움말을 표시하는 중 오류가 발생했습니다: {ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// SheetPathSettingsForm이 닫힐 때 호출되는 안전한 핸들러입니다.
        /// </summary>
        /// <param name="sheetNames">시트 이름 목록입니다.</param>
        private void HandleSettingsFormClosedSafe(List<string> sheetNames)
        {
            // 설정 후 다시 활성화된 시트 수 확인
            int enabledSheetsCount = 0;
            foreach (var sheetName in sheetNames)
            {
                if (SheetPathManager.Instance.IsSheetEnabled(sheetName))
                {
                    enabledSheetsCount++;
                }
            }

            // 활성화된 시트가 없으면 변환 취소 메시지 (선택적)
            if (enabledSheetsCount == 0) { /* MessageBox.Show("활성화된 시트가 없어 변환을 취소합니다.", "변환 취소", MessageBoxButtons.OK, MessageBoxIcon.Information); */ }
        }

        /// <summary>
        /// SheetPathSettingsForm이 닫힐 때 호출되는 공통 핸들러입니다.
        /// </summary>
        /// <param name="convertibleSheets">설정 폼에 전달되었던 시트 목록입니다.</param>
        private void HandleSettingsFormClosed(List<Excel.Worksheet> convertibleSheets)
        {
            // 설정 후 다시 활성화된 시트 수 확인
            int enabledSheetsCount = 0;
            foreach (var sheet in convertibleSheets)
            {
                try
                {
                    // Excel COM 객체에 안전하게 접근
                    string sheetName = sheet.Name;
                    if (SheetPathManager.Instance.IsSheetEnabled(sheetName))
                    {
                        enabledSheetsCount++;
                    }
                }
                catch (System.Runtime.InteropServices.COMException ex)
                {
                    // COM 객체가 무효한 경우 로그만 남기고 건너뛰기
                    Debug.WriteLine($"워크시트 접근 실패 (COM Exception: {ex.ErrorCode:X}): 시트가 삭제되었거나 무효한 상태입니다.");
                }
                catch (Exception ex)
                {
                    // 기타 예외 처리
                    Debug.WriteLine($"워크시트 접근 중 예외 발생: {ex.Message}");
                }
            }

            // 활성화된 시트가 없으면 변환 취소 메시지 (선택적)
            if (enabledSheetsCount == 0) { /* MessageBox.Show("활성화된 시트가 없어 변환을 취소합니다.", "변환 취소", MessageBoxButtons.OK, MessageBoxIcon.Information); */ }
        }

        // Excel 파일 변환 처리 (수정: 변환된 파일 목록 반환)
        private List<string> ConvertExcelFile(ExcelToYamlConfig config)
        {
            string tempFile = null;
            List<string> convertedFiles = new List<string>();

            try
            {
                // 현재 워크북 가져오기
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;
                var activeWorkbook = app.ActiveWorkbook;

                if (activeWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return convertedFiles;
                }

                // 워크북 경로 설정
                string workbookPath = activeWorkbook.FullName;
                SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);

                // 디버깅을 위한 로그 추가
                Debug.WriteLine($"현재 워크북 경로: {workbookPath}");

                // 변환 가능한 시트 찾기
                var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(activeWorkbook);

                Debug.WriteLine($"변환 가능한 시트 수: {convertibleSheets.Count}");
                foreach (var sheet in convertibleSheets)
                {
                    Debug.WriteLine($"시트 이름: {sheet.Name}");
                }

                if (convertibleSheets.Count == 0)
                {
                    MessageBox.Show("변환 가능한 시트가 없습니다. 변환하려는 시트 이름 앞에 '!'를 추가하세요.",
                        "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return convertedFiles;
                }

                // 임시 파일로 저장
                tempFile = addIn.SaveToTempFile();
                if (string.IsNullOrEmpty(tempFile))
                {
                    MessageBox.Show("임시 파일을 생성할 수 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return convertedFiles;
                }

                // 프로그레스 바 표시 및 변환 작업 수행
                string outputFormat = config.OutputFormat == OutputFormat.Json ? "JSON" : "YAML";
                using (var progressForm = new Forms.ProgressForm())
                {
                    progressForm.RunOperation((progress, cancellationToken) =>
                    {
                        int totalSheets = convertibleSheets.Count;
                        int processedSheets = 0;
                        int successCount = 0;
                        int skipCount = 0;

                        // 초기 프로그레스 업데이트
                        progress.Report(new Forms.ProgressForm.ProgressInfo
                        {
                            Percentage = 0,
                            StatusMessage = "Excel 데이터 분석 중..."
                        });

                        // 처리할 시트 목록 계산
                        var sheetsToProcess = new List<Worksheet>();
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

                        // 처리할 시트가 없으면 종료
                        if (sheetsToProcess.Count == 0)
                        {
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 100,
                                StatusMessage = "처리할 시트가 없습니다.",
                                IsCompleted = true
                            });
                            return;
                        }

                        // 모든 변환 가능한 시트에 대해 처리
                        foreach (var sheet in sheetsToProcess)
                        {
                            // 작업 취소 확인
                            if (cancellationToken.IsCancellationRequested)
                            {
                                progress.Report(new Forms.ProgressForm.ProgressInfo
                                {
                                    Percentage = 100,
                                    StatusMessage = "작업이 취소되었습니다.",
                                    IsCompleted = true
                                });
                                return;
                            }

                            string sheetName = sheet.Name;
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = (int)((double)processedSheets / sheetsToProcess.Count * 100),
                                StatusMessage = $"'{sheetName}' 시트 변환 중..."
                            });

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
                                processedSheets++;
                                skipCount++;
                                continue;
                            }

                            // 활성화 상태 확인 - 비활성화된 시트는 건너뛰기
                            bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);
                            if (!isEnabled)
                            {
                                processedSheets++;
                                skipCount++;
                                continue;
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
                                    processedSheets++;
                                    skipCount++;
                                    continue;
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

                                successCount++;
                                convertedFiles.Add(resultFile);
                            }
                            catch (Exception ex)
                            {
                                Debug.WriteLine($"[ConvertExcelFile] 시트 '{sheetName}' 변환 중 오류 발생: {ex.Message}");
                                skipCount++;
                            }

                            processedSheets++;
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = (int)((double)processedSheets / sheetsToProcess.Count * 100),
                                StatusMessage = $"'{sheetName}' 시트 변환 완료"
                            });
                        }

                        // 변환 결과 로그 작성
                        Debug.WriteLine($"변환 완료: {successCount}개 성공, {skipCount}개 실패");

                        // 작업 완료 알림
                        progress.Report(new Forms.ProgressForm.ProgressInfo
                        {
                            Percentage = 100,
                            StatusMessage = $"변환 완료: {successCount}개 시트 변환 성공",
                            IsCompleted = true
                        });
                    }, $"Excel → {outputFormat} 변환 중");

                    progressForm.ShowDialog();

                    // 취소된 경우
                    if (progressForm.DialogResult == DialogResult.Cancel)
                    {
                        MessageBox.Show("변환 작업이 취소되었습니다.", "작업 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }

                return convertedFiles;
            }
            catch (IOException ex)
            {
                MessageBox.Show($"파일 처리 중 오류가 발생했습니다: {ex.Message}",
                    "파일 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (UnauthorizedAccessException ex)
            {
                MessageBox.Show($"파일 접근 권한이 없습니다: {ex.Message}",
                    "권한 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"변환 중 오류가 발생했습니다: {ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 임시 파일 정리
                if (!string.IsNullOrEmpty(tempFile))
                {
                    try
                    {
                        File.Delete(tempFile);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[ConvertExcelFile] 임시 파일 삭제 중 오류 발생: {ex.Message}");
                    }
                }
            }

            // 예외 발생 시 이 지점에 도달하므로 빈 목록 반환
            return convertedFiles;
        }

        // HTML 내보내기 버튼 클릭
        public void OnExportToHtmlClick(object sender, RibbonControlEventArgs e)
        {
            ExportCurrentSheetToHtml();
        }

        // 디버깅: 현재 시트를 HTML로 내보내기
        public void ExportCurrentSheetToHtml()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var activeSheet = app.ActiveSheet as Excel.Worksheet;
                
                if (activeSheet == null)
                {
                    MessageBox.Show("활성 시트가 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 저장 경로 선택
                var saveDialog = new SaveFileDialog
                {
                    Title = "HTML 파일 저장",
                    Filter = "HTML 파일 (*.html)|*.html",
                    FileName = $"{activeSheet.Name}_debug.html",
                    RestoreDirectory = true
                };

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    ExcelToHtmlExporter.ExportToHtml(activeSheet, saveDialog.FileName);
                    
                    // 브라우저에서 열기
                    var result = MessageBox.Show($"HTML 파일이 생성되었습니다.\n\n{saveDialog.FileName}\n\n브라우저에서 열어보시겠습니까?",
                        "내보내기 완료", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                    
                    if (result == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start(saveDialog.FileName);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"HTML 내보내기 중 오류가 발생했습니다:\n\n{ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // XML 가져오기 버튼 클릭
        public void OnImportXmlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 파일 선택 대화상자
                var openFileDialog = new OpenFileDialog
                {
                    Title = "XML 파일 선택",
                    Filter = "XML 파일 (*.xml)|*.xml|모든 파일 (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true
                };

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                string xmlFilePath = openFileDialog.FileName;
                string xmlContent = File.ReadAllText(xmlFilePath);
                string fileName = Path.GetFileNameWithoutExtension(xmlFilePath);

                // XML을 Excel로 변환 (새로운 방식: XML → YAML → Excel)
                var converter = new Core.XmlToExcelViaYamlConverter();
                var workbook = converter.ConvertToExcel(xmlContent, fileName);

                // 현재 워크북 가져오기
                var app = Globals.ThisAddIn.Application;
                var currentWorkbook = app.ActiveWorkbook;
                
                if (currentWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다. Excel 파일을 먼저 열어주세요.",
                        "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 변환된 시트 복사
                var convertedSheet = workbook.Worksheets.First();
                
                // ClosedXML 워크북을 임시 파일로 저장
                string tempFile = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.xlsx");
                workbook.SaveAs(tempFile);
                
                // 임시 파일을 Excel에서 열기
                var tempWorkbook = app.Workbooks.Open(tempFile);
                var sourceSheet = tempWorkbook.Worksheets[1];
                
                // 현재 워크북의 마지막에 시트 복사
                sourceSheet.Copy(After: currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);
                
                // 새로 추가된 시트 선택
                var newSheet = currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count];
                newSheet.Activate();
                
                // 임시 워크북 닫기
                tempWorkbook.Close(false);
                
                // 임시 파일 삭제
                try
                {
                    File.Delete(tempFile);
                }
                catch { }

                MessageBox.Show($"XML 파일이 성공적으로 Excel로 변환되었습니다.\n\n파일: {fileName}.xml",
                    "변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"XML 변환 중 오류가 발생했습니다:\n\n{ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // YAML 가져오기 버튼 클릭
        public void OnImportYamlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 파일 선택 대화상자
                var openFileDialog = new OpenFileDialog
                {
                    Title = "YAML 파일 선택",
                    Filter = "YAML 파일 (*.yaml;*.yml)|*.yaml;*.yml|모든 파일 (*.*)|*.*",
                    FilterIndex = 1,
                    RestoreDirectory = true
                };

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                string yamlFilePath = openFileDialog.FileName;
                string fileName = Path.GetFileNameWithoutExtension(yamlFilePath);

                // 현재 워크북 가져오기
                var app = Globals.ThisAddIn.Application;
                var currentWorkbook = app.ActiveWorkbook;
                
                if (currentWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다. Excel 파일을 먼저 열어주세요.",
                        "오류", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // YAML 파일 읽기
                string yamlContent;
                try
                {
                    yamlContent = File.ReadAllText(yamlFilePath);
                    Debug.WriteLine($"[OnImportYamlClick] YAML 파일 읽기 완료: {yamlFilePath}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"YAML 파일을 읽는 중 오류가 발생했습니다:\n\n{ex.Message}",
                        "파일 읽기 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                // 루트 요소가 있는 경우 제거 (XML에서 변환된 경우 처리)
                try
                {
                    Debug.WriteLine("[OnImportYamlClick] YAML 구조 분석 시작");
                    var yaml = new YamlDotNet.RepresentationModel.YamlStream();
                    yaml.Load(new StringReader(yamlContent));
                    
                    if (yaml.Documents.Count > 0 && yaml.Documents[0].RootNode is YamlDotNet.RepresentationModel.YamlMappingNode rootMapping)
                    {
                        // 루트가 단일 키를 가진 매핑이면, 그 값을 직접 사용
                        if (rootMapping.Children.Count == 1)
                        {
                            var rootKey = rootMapping.Children.Keys.First();
                            var rootValue = rootMapping.Children[rootKey];
                            
                            Debug.WriteLine($"[OnImportYamlClick] 루트 요소 '{rootKey}' 감지, 제거 중...");
                            
                            // 루트 값을 새로운 YAML 문서로 변환
                            var newYamlDoc = new YamlDotNet.RepresentationModel.YamlDocument(rootValue);
                            var newYamlStream = new YamlDotNet.RepresentationModel.YamlStream(newYamlDoc);
                            var writer = new StringWriter();
                            newYamlStream.Save(writer, false);
                            yamlContent = writer.ToString();
                            
                            Debug.WriteLine($"[OnImportYamlClick] 루트 요소 제거 완료");
                        }
                    }
                    Debug.WriteLine("[OnImportYamlClick] YAML 구조 분석 완료");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[OnImportYamlClick] 루트 요소 처리 중 오류 (무시하고 계속): {ex.Message}");
                }
                
                // YAML을 Excel로 변환
                Core.YamlToExcel.YamlToExcelConverter converter;
                try
                {
                    Debug.WriteLine("[OnImportYamlClick] YamlToExcelConverter 생성 시작");
                    converter = new Core.YamlToExcel.YamlToExcelConverter();
                    Debug.WriteLine("[OnImportYamlClick] YamlToExcelConverter 생성 완료");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"YAML 변환기 생성 중 오류가 발생했습니다:\n\n{ex.Message}",
                        "변환기 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                // 임시 파일로 변환
                string tempFile = Path.Combine(Path.GetTempPath(), $"temp_{Guid.NewGuid()}.xlsx");
                try
                {
                    Debug.WriteLine($"[OnImportYamlClick] YAML to Excel 변환 시작: {tempFile}");
                    converter.ConvertFromContent(yamlContent, tempFile);
                    Debug.WriteLine($"[OnImportYamlClick] YAML to Excel 변환 완료: {tempFile}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"YAML을 Excel로 변환하는 중 오류가 발생했습니다:\n\n{ex.Message}",
                        "변환 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                // 임시 파일을 Excel에서 열기
                dynamic tempWorkbook;
                dynamic sourceSheet;
                try
                {
                    Debug.WriteLine($"[OnImportYamlClick] 임시 파일 열기 시작: {tempFile}");
                    tempWorkbook = app.Workbooks.Open(tempFile);
                    sourceSheet = tempWorkbook.Worksheets[1];
                    Debug.WriteLine($"[OnImportYamlClick] 임시 파일 열기 완료");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"변환된 파일을 여는 중 오류가 발생했습니다:\n\n{ex.Message}",
                        "파일 열기 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    try { File.Delete(tempFile); } catch { }
                    return;
                }
                
                // 새 시트 이름 생성 (중복 방지)
                string newSheetName = $"!{fileName}";
                int suffix = 1;
                while (WorksheetExists(currentWorkbook, newSheetName))
                {
                    newSheetName = $"!{fileName}_{suffix++}";
                }
                
                // 현재 워크북에 시트 복사
                try
                {
                    Debug.WriteLine($"[OnImportYamlClick] 시트 복사 시작: {newSheetName}");
                    sourceSheet.Copy(After: currentWorkbook.Worksheets[currentWorkbook.Worksheets.Count]);
                    var newSheet = currentWorkbook.ActiveSheet;
                    newSheet.Name = newSheetName;
                    Debug.WriteLine($"[OnImportYamlClick] 시트 복사 완료: {newSheetName}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"시트를 복사하는 중 오류가 발생했습니다:\n\n{ex.Message}",
                        "시트 복사 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    tempWorkbook.Close(false);
                    try { File.Delete(tempFile); } catch { }
                    return;
                }
                
                // 임시 워크북 닫기
                try
                {
                    Debug.WriteLine("[OnImportYamlClick] 임시 워크북 닫기");
                    tempWorkbook.Close(false);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[OnImportYamlClick] 임시 워크북 닫기 중 오류 (무시): {ex.Message}");
                }
                
                // 임시 파일 삭제
                try 
                { 
                    File.Delete(tempFile); 
                    Debug.WriteLine($"[OnImportYamlClick] 임시 파일 삭제 완료: {tempFile}");
                } 
                catch (Exception ex) 
                { 
                    Debug.WriteLine($"[OnImportYamlClick] 임시 파일 삭제 실패 (무시): {ex.Message}");
                }
                
                MessageBox.Show($"YAML 파일이 성공적으로 가져와졌습니다.\n\n시트 이름: {newSheetName}",
                    "가져오기 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"YAML 변환 중 오류가 발생했습니다:\n\n{ex.Message}",
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 워크시트 존재 여부 확인 헬퍼 메서드
        private bool WorksheetExists(dynamic workbook, string sheetName)
        {
            foreach (dynamic sheet in workbook.Worksheets)
            {
                if (sheet.Name == sheetName)
                    return true;
            }
            return false;
        }



        // 엑셀 파일을 임시 디렉토리에 YAML로 변환하는 메서드
        private List<string> ConvertExcelFileToTemp(ExcelToYamlConfig config, string tempDir, List<Excel.Worksheet> sheetsToProcess)
        {
            string tempFile = null;
            List<string> convertedFiles = new List<string>();

            try
            {
                // 현재 워크북 가져오기
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;

                if (app.ActiveWorkbook == null || sheetsToProcess == null || sheetsToProcess.Count == 0)
                {
                    return convertedFiles;
                }

                // sheetsToProcess는 PrepareAndValidateSheets에서 이미 필터링된 목록을 사용

                // 임시 파일로 저장
                tempFile = addIn.SaveToTempFile();
                if (string.IsNullOrEmpty(tempFile))
                {
                    return convertedFiles;
                }

                // 모든 변환 가능한 시트에 대해 처리
                foreach (var sheet in sheetsToProcess)
                {
                    string sheetName = sheet.Name;
                    bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);

                    // 비활성화된 시트는 건너뛰기
                    if (!isEnabled)
                    {
                        continue;
                    }

                    // 앞의 '!' 문자 제거 (표시용)
                    string fileName = sheetName.StartsWith("!") ? sheetName.Substring(1) : sheetName;

                    // 결과 파일 경로
                    string resultFile = Path.Combine(tempDir, $"{fileName}.yaml");

                    try
                    {
                        // 변환 처리 - 시트 이름 지정
                        var excelReader = new ExcelReader(config);
                        excelReader.ProcessExcelFile(tempFile, resultFile, sheetName);

                        convertedFiles.Add(resultFile);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[ConvertExcelFileToTemp] 시트 '{sheetName}' 변환 중 오류 발생: {ex.Message}");
                    }
                }

                return convertedFiles;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ConvertExcelFileToTemp] 변환 중 오류 발생: {ex.Message}");
                return convertedFiles;
            }
        }

        // JSON 가져오기 버튼 클릭
        private void OnImportJsonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Logger.Information("JSON 가져오기 버튼 클릭");

                // JSON 파일 선택
                using (var openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Title = "JSON 파일 선택";
                    openFileDialog.Filter = "JSON 파일 (*.json)|*.json|모든 파일 (*.*)|*.*";
                    openFileDialog.FilterIndex = 1;

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        string jsonPath = openFileDialog.FileName;
                        Logger.Information($"선택된 JSON 파일: {jsonPath}");

                        // JSON을 YAML로 변환 후 Excel로 변환
                        string jsonContent = File.ReadAllText(jsonPath);
                        
                        // JSON을 YAML로 변환 (Newtonsoft.Json 사용)
                        var jsonObject = Newtonsoft.Json.JsonConvert.DeserializeObject(jsonContent);
                        var serializer = new YamlDotNet.Serialization.SerializerBuilder().Build();
                        string yamlContent = serializer.Serialize(jsonObject);

                        // YAML을 Excel로 변환
                        var yamlToExcel = new Core.YamlToExcel.YamlToExcelConverter();
                        string fileName = Path.GetFileNameWithoutExtension(jsonPath);
                        string excelPath = Path.Combine(Path.GetDirectoryName(jsonPath), $"{fileName}.xlsx");

                        yamlToExcel.ConvertFromContent(yamlContent, excelPath);

                        Logger.Information($"JSON → Excel 변환 완료: {excelPath}");

                        // 결과 알림
                        var result = MessageBox.Show(
                            $"JSON 파일을 Excel로 변환했습니다.\n\n생성된 파일: {excelPath}\n\n파일을 열어보시겠습니까?",
                            "변환 완료",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Information);

                        if (result == DialogResult.Yes)
                        {
                            System.Diagnostics.Process.Start(excelPath);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"JSON 가져오기 중 오류 발생: {ex.Message}", ex);
                MessageBox.Show($"JSON 가져오기 중 오류가 발생했습니다:\n{ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        // 리소스 텍스트 가져오기
        private static string GetResourceText(string resourceName)
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();

            foreach (string name in assembly.GetManifestResourceNames())
            {
                if (string.Compare(resourceName, name, StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (var stream = assembly.GetManifestResourceStream(name))
                    {
                        if (stream != null)
                        {
                            using (var reader = new StreamReader(stream))
                            {
                                return reader.ReadToEnd();
                            }
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// 테스트 실행 버튼 클릭 이벤트 핸들러
        /// </summary>
        private void OnRunTestsClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                Logger.Information("테스트 실행 시작");
                
                var runner = new Tests.TestRunner();
                
                // 테스트 시트 생성 및 실행 옵션
                var result = MessageBox.Show(
                    "테스트 시트를 생성하고 테스트를 실행하시겠습니까?\n" +
                    "'예'를 클릭하면 테스트 시트를 생성합니다.\n" +
                    "'아니오'를 클릭하면 현재 시트로만 테스트합니다.",
                    "테스트 실행",
                    MessageBoxButtons.YesNoCancel,
                    MessageBoxIcon.Question);
                
                if (result == DialogResult.Cancel)
                {
                    return;
                }
                
                if (result == DialogResult.Yes)
                {
                    // 테스트 시트 생성 및 테스트 실행
                    runner.GenerateTestSheetsAndRun();
                }
                else
                {
                    // 현재 시트로만 테스트 실행
                    runner.RunIntegrationTests();
                }
                
                var summary = runner.GetResultSummary();
                MessageBox.Show(
                    summary,
                    "테스트 완료",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "테스트 실행 중 오류");
                MessageBox.Show(
                    $"테스트 실행 중 오류가 발생했습니다:\n{ex.Message}",
                    "오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}