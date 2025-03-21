using ExcelToYamlAddin.Config;
using ExcelToYamlAddin.Core;
using ExcelToYamlAddin.Core.YamlPostProcessors;
using ExcelToYamlAddin.Forms;
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
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelToYamlAddin
{
    public partial class Ribbon : RibbonBase
    {
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
                var pathManager = ExcelToYamlAddin.Config.SheetPathManager.Instance;
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

        // YAML으로 변환 버튼 클릭
        public void OnConvertToYamlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 설정 초기화 및 다시 로드
                SheetPathManager.Instance.Initialize();
                Debug.WriteLine("[OnConvertToYamlClick] SheetPathManager 초기화 완료");

                // 현재 워크북 가져오기
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;
                var activeWorkbook = app.ActiveWorkbook;

                if (activeWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 워크북 경로 설정
                string workbookPath = activeWorkbook.FullName;
                SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);

                // Excel 설정 관리자 초기화
                ExcelConfigManager.Instance.SetCurrentWorkbook(workbookPath);

                // 변환 가능한 시트 찾기
                var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(activeWorkbook);

                if (convertibleSheets.Count == 0)
                {
                    MessageBox.Show("변환 가능한 시트가 없습니다. 변환하려는 시트 이름 앞에 '!'를 추가하세요.",
                        "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // '!'로 시작하는 시트가 있으면 excel2yamlconfig 시트 확인/생성
                ExcelConfigManager.Instance.EnsureConfigSheetExists();

                // XML 설정에서 Excel 설정으로 마이그레이션 (최초 1회)
                if (Properties.Settings.Default.FirstConfigMigration)
                {
                    ExcelConfigManager.Instance.MigrateFromXmlSettings(ExcelToYamlAddin.Config.SheetPathManager.Instance);
                    Properties.Settings.Default.FirstConfigMigration = false;
                    Properties.Settings.Default.Save();
                    Debug.WriteLine("[Ribbon] XML 설정을 Excel 설정으로 마이그레이션 완료");
                }

                // 활성화된 시트 수 확인
                int enabledSheetsCount = 0;
                Debug.WriteLine($"[OnConvertToYamlClick] 변환 가능한 시트 수: {convertibleSheets.Count}");
                foreach (var sheet in convertibleSheets)
                {
                    string currentSheetName = sheet.Name;
                    bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(currentSheetName);
                    Debug.WriteLine($"[OnConvertToYamlClick] 시트 '{currentSheetName}' 활성화 상태: {isEnabled}");

                    if (isEnabled)
                    {
                        enabledSheetsCount++;
                    }
                }

                Debug.WriteLine($"[OnConvertToYamlClick] 활성화된 시트 수: {enabledSheetsCount}, 비활성화된 시트 수: {convertibleSheets.Count - enabledSheetsCount}");

                // 활성화된 시트가 없는 경우 처리
                if (enabledSheetsCount == 0)
                {
                    Debug.WriteLine("[OnConvertToYamlClick] 경고: 활성화된 시트가 없습니다. 시트 활성화 상태 상세 정보 출력:");

                    // 활성화 상태 자세히 확인 (디버그용)
                    Dictionary<string, string> allEnabledPaths = SheetPathManager.Instance.GetAllEnabledSheetPaths();
                    Debug.WriteLine($"[OnConvertToYamlClick] GetAllEnabledSheetPaths 결과: {allEnabledPaths.Count}개 시트");
                    foreach (var kvp in allEnabledPaths)
                    {
                        Debug.WriteLine($"[OnConvertToYamlClick] 활성화된 시트: '{kvp.Key}', 경로: '{kvp.Value}'");
                    }

                    foreach (var sheet in convertibleSheets)
                    {
                        string sheetName = sheet.Name;
                        bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);
                        string sheetPath = SheetPathManager.Instance.GetSheetPath(sheetName);

                        Debug.WriteLine($"[OnConvertToYamlClick] 시트 '{sheetName}' - 활성화: {isEnabled}, 경로: '{sheetPath}'");

                        // 활성화 상태가 믿을 수 없는 경우 해당 시트를 강제로 활성화
                        if (isEnabled && !allEnabledPaths.ContainsKey(sheetName))
                        {
                            Debug.WriteLine($"[OnConvertToYamlClick] 활성화 상태 불일치 감지. 시트 '{sheetName}'를 강제로 활성화합니다.");
                            SheetPathManager.Instance.SetSheetEnabled(sheetName, true);
                            enabledSheetsCount++;
                        }
                    }

                    // 다시 활성화된 시트 수 확인
                    if (enabledSheetsCount == 0)
                    {
                        MessageBox.Show("활성화된 시트가 없어 변환을 취소합니다.\n\n시트 설정 창에서 시트를 활성화하십시오.",
                            "변환 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // 시트별 경로 설정 창 열기
                        using (var form = new Forms.SheetPathSettingsForm(convertibleSheets))
                        {
                            form.StartPosition = FormStartPosition.CenterScreen;
                            form.ShowDialog();
                        }
                        return;
                    }
                }

                // 활성화된 시트 수가 탐지된 시트 수보다 적은 경우 확인창 표시
                if (enabledSheetsCount < convertibleSheets.Count)
                {
                    int disabledCount = convertibleSheets.Count - enabledSheetsCount;
                    string message = $"{convertibleSheets.Count}개의 변환 가능한 시트 중 {disabledCount}개의 시트가 비활성화되어 있습니다.\n\n" +
                                    $"활성화된 {enabledSheetsCount}개의 시트만 변환하시겠습니까?\n\n" +
                                    "아니오를 선택하면 시트별 경로 설정 창이 열립니다.";

                    DialogResult result = MessageBox.Show(message, "시트 활성화 확인",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.No)
                    {
                        // 시트별 경로 설정 창 열기
                        if (settingsForm != null && !settingsForm.IsDisposed)
                        {
                            settingsForm.Activate();
                            return;
                        }

                        settingsForm = new Forms.SheetPathSettingsForm(convertibleSheets);
                        settingsForm.FormClosed += (s, args) =>
                        {
                            // 설정 후 다시 활성화된 시트 수 확인
                            enabledSheetsCount = 0;
                            foreach (var sheet in convertibleSheets)
                            {
                                if (SheetPathManager.Instance.IsSheetEnabled(sheet.Name))
                                {
                                    enabledSheetsCount++;
                                }
                            }

                            // 활성화된 시트가 없으면 변환 취소
                            if (enabledSheetsCount == 0)
                            {
                                MessageBox.Show("활성화된 시트가 없어 변환을 취소합니다.",
                                    "변환 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }

                            settingsForm = null;
                        };
                        settingsForm.StartPosition = FormStartPosition.CenterScreen;
                        settingsForm.Show();
                        return;
                    }
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
                int mergeKeyPathsSuccessCount = 0;
                int flowStyleSuccessCount = 0;

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
                                int totalFiles = convertedFiles.Count;
                                int processedFiles = 0;

                                // 초기 프로그레스 업데이트
                                progress.Report(new Forms.ProgressForm.ProgressInfo
                                {
                                    Percentage = 0,
                                    StatusMessage = "YAML 후처리 준비 중..."
                                });

                                // 취소 여부 확인을 위한 헬퍼 메서드
                                void CheckCancellation()
                                {
                                    // 취소 토큰 확인
                                    cancellationToken.ThrowIfCancellationRequested();

                                    // 더 이상 이전 방식의 취소 확인은 사용하지 않음
                                    // cancellationToken으로 충분함
                                }

                                try
                                {
                                    foreach (var filePath in convertedFiles)
                                    {
                                        // 각 파일 처리 전 취소 여부 확인
                                        CheckCancellation();

                                        if (File.Exists(filePath) && Path.GetExtension(filePath).ToLower() == ".yaml")
                                        {
                                            // 파일 이름 정보 추출
                                            string fileName = Path.GetFileNameWithoutExtension(filePath);

                                            progress.Report(new Forms.ProgressForm.ProgressInfo
                                            {
                                                Percentage = (int)((double)processedFiles / totalFiles * 100),
                                                StatusMessage = $"'{fileName}' 파일 처리 중..."
                                            });

                                            // 작업 중간에도 취소 여부 확인
                                            CheckCancellation();

                                            // 파일 경로에서 시트 이름 추출
                                            string workbookName = Path.GetFileName(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);

                                            // 가능한 시트 이름 형식
                                            List<string> possibleSheetNames = new List<string>
                                            {
                                                fileName,                  // 파일명 그대로
                                                "!" + fileName,            // !접두사 추가
                                                fileName.StartsWith("!") ? fileName.Substring(1) : fileName   // !접두사 제거
                                            };

                                            Debug.WriteLine($"[Ribbon] YAML 파일 처리: {filePath}");

                                            // 찾은 실제 시트 이름
                                            string matchedSheetName = null;

                                            // YAML 선택적 필드 후처리
                                            // 워크북 내 시트 이름 매칭
                                            foreach (var sheet in convertibleSheets)
                                            {
                                                string currentSheetName = sheet.Name;
                                                if (currentSheetName.StartsWith("!"))
                                                    currentSheetName = currentSheetName.Substring(1);

                                                if (string.Compare(currentSheetName, fileName, true) == 0)
                                                {
                                                    matchedSheetName = sheet.Name;
                                                    break;
                                                }
                                            }

                                            // 1단계: YAML 선택적 필드 처리
                                            progress.Report(new Forms.ProgressForm.ProgressInfo
                                            {
                                                StatusMessage = $"'{fileName}' - 선택적 필드 처리 중..."
                                            });

                                            // YAML 선택적 필드 처리
                                            if (matchedSheetName != null)
                                            {
                                                // Excel excel2yamlconfig 시트에서 먼저 확인 (우선순위 변경: Excel이 더 높은 우선순위)
                                                bool option = ExcelConfigManager.Instance.GetConfigBool(matchedSheetName, "YamlEmptyFields", false);

                                                // Excel에 설정이 없으면 SheetPathManager에서 확인
                                                if (!option)
                                                {
                                                    option = SheetPathManager.Instance.GetYamlEmptyFieldsOption(matchedSheetName);
                                                }

                                                // 둘 다 없으면 기본 설정 사용
                                                if (!option && addEmptyYamlFields)
                                                {
                                                    option = addEmptyYamlFields;
                                                }

                                                // 2단계: 키 경로 병합 처리
                                                progress.Report(new Forms.ProgressForm.ProgressInfo
                                                {
                                                    StatusMessage = $"'{fileName}' - 키 경로 병합 처리 중..."
                                                });

                                                // 키 경로 병합 후처리
                                                if (matchedSheetName != null)
                                                {
                                                    // Excel excel2yamlconfig 시트에서 먼저 확인 (우선순위 변경: Excel이 더 높은 우선순위)
                                                    string sheetMergeKeyPaths = ExcelConfigManager.Instance.GetConfigValue(matchedSheetName, "MergeKeyPaths", "");

                                                    // Excel에 설정이 없으면 SheetPathManager에서 확인
                                                    if (string.IsNullOrEmpty(sheetMergeKeyPaths))
                                                    {
                                                        sheetMergeKeyPaths = SheetPathManager.Instance.GetMergeKeyPaths(matchedSheetName);
                                                    }

                                                    // YAML 병합 키 경로 처리 실행 (YamlMergeKeyPathsProcessor 사용)
                                                    if (!string.IsNullOrEmpty(sheetMergeKeyPaths))
                                                    {
                                                        Debug.WriteLine($"[Ribbon] YAML 병합 후처리 실행: {filePath}, 설정: {sheetMergeKeyPaths}");

                                                        // 개선된 YamlMergeKeyPathsProcessor를 사용하여 병합 처리
                                                        // YamlEmptyFields 옵션이 true이면 빈 필드를 유지하도록 includeEmptyFields 매개변수를 전달
                                                        bool success = YamlMergeKeyPathsProcessor.ProcessYamlFileFromConfig(
                                                            filePath,
                                                            sheetMergeKeyPaths,
                                                            option // YamlEmptyFields 옵션 전달
                                                        );

                                                        if (success)
                                                        {
                                                            Debug.WriteLine($"[Ribbon] YAML 병합 후처리 완료: {filePath}");
                                                            mergeKeyPathsSuccessCount++;
                                                        }
                                                        else
                                                        {
                                                            Debug.WriteLine($"[Ribbon] YAML 병합 후처리 실패: {filePath}");
                                                        }
                                                    }
                                                }
                                            }

                                            // 3단계: Flow Style 처리
                                            progress.Report(new Forms.ProgressForm.ProgressInfo
                                            {
                                                StatusMessage = $"'{fileName}' - Flow Style 처리 중..."
                                            });

                                            // Flow Style 처리
                                            if (matchedSheetName != null)
                                            {
                                                // 시트 이름을 로깅하여 디버깅을 돕습니다
                                                Debug.WriteLine($"[Ribbon] Flow Style 처리 검사 중: 파일 경로={filePath}, 파일 이름={fileName}, 매칭된 시트 이름={matchedSheetName}");

                                                // Excel excel2yamlconfig 시트에서 먼저 확인 (우선순위 변경: Excel이 더 높은 우선순위)
                                                string sheetFlowStyle = ExcelConfigManager.Instance.GetConfigValue(matchedSheetName, "FlowStyle", "");
                                                Debug.WriteLine($"[Ribbon] ExcelConfigManager Flow Style 설정 값: '{sheetFlowStyle}'");

                                                // Excel에 설정이 없으면 SheetPathManager에서 확인
                                                if (string.IsNullOrEmpty(sheetFlowStyle))
                                                {
                                                    sheetFlowStyle = SheetPathManager.Instance.GetFlowStyleConfig(matchedSheetName ?? fileName);
                                                    Debug.WriteLine($"[Ribbon] SheetPathManager Flow Style 설정 값: '{sheetFlowStyle}'");
                                                }

                                                // YAML 흐름 스타일 처리 실행
                                                if (!string.IsNullOrEmpty(sheetFlowStyle))
                                                {
                                                    Debug.WriteLine($"[Ribbon] YAML 흐름 스타일 후처리 실행: {filePath}, 설정: {sheetFlowStyle}");
                                                    bool success = YamlFlowStyleProcessor.ProcessYamlFileFromConfig(filePath, sheetFlowStyle);
                                                    if (success)
                                                    {
                                                        Debug.WriteLine($"[Ribbon] YAML 흐름 스타일 후처리 완료: {filePath}");
                                                        flowStyleSuccessCount++;
                                                    }
                                                }
                                                else
                                                {
                                                    Debug.WriteLine($"[Ribbon] YAML 흐름 스타일 후처리 건너뜀: {filePath}");
                                                    Debug.WriteLine($"[Ribbon] 설정 값: '{sheetFlowStyle}'");
                                                }
                                            }

                                            // 빈 배열 필드 처리
                                            progress.Report(new Forms.ProgressForm.ProgressInfo
                                            {
                                                StatusMessage = $"'{fileName}' - 빈 배열 처리 중..."
                                            });

                                            if (matchedSheetName != null)
                                            {
                                                // YamlEmptyFields 옵션 확인
                                                bool yamlEmptyFieldsOption = ExcelConfigManager.Instance.GetConfigBool(matchedSheetName, "YamlEmptyFields", false);

                                                // 전역 설정이나 시트별 설정이 활성화된 경우에만 처리
                                                bool processEmptyArrays = yamlEmptyFieldsOption || addEmptyYamlFields;

                                                if (processEmptyArrays)
                                                {
                                                    Debug.WriteLine($"[Ribbon] YAML 빈 배열 처리: OrderedYamlFactory에서 처리 (파일: {filePath})");
                                                    Debug.WriteLine($"[Ribbon] - 시트별 설정: {yamlEmptyFieldsOption}, 전역 설정: {addEmptyYamlFields}");
                                                }
                                                else
                                                {
                                                    Debug.WriteLine($"[Ribbon] YAML 빈 배열 처리 건너뜀: 관련 옵션이 비활성화되어 있습니다.");
                                                    Debug.WriteLine($"[Ribbon] - 시트별 설정: {yamlEmptyFieldsOption}, 전역 설정: {addEmptyYamlFields}");
                                                }
                                            }

                                            processedFiles++;
                                        }
                                    }

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

                    if (mergeKeyPathsSuccessCount > 0)
                        message += $"\n키 경로 병합 처리: {mergeKeyPathsSuccessCount}개 파일";

                    if (flowStyleSuccessCount > 0)
                        message += $"\nFlow 스타일 처리: {flowStyleSuccessCount}개 파일";

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

        // JSON으로 변환 버튼 클릭
        public void OnConvertToJsonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 설정 초기화 및 다시 로드
                SheetPathManager.Instance.Initialize();
                Debug.WriteLine("[OnConvertToJsonClick] SheetPathManager 초기화 완료");

                // 현재 워크북 가져오기
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;
                var activeWorkbook = app.ActiveWorkbook;

                if (activeWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 워크북 경로 설정
                string workbookPath = activeWorkbook.FullName;
                SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);

                // Excel 설정 관리자 초기화
                ExcelConfigManager.Instance.SetCurrentWorkbook(workbookPath);

                // 변환 가능한 시트 찾기
                var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(activeWorkbook);

                if (convertibleSheets.Count == 0)
                {
                    MessageBox.Show("변환 가능한 시트가 없습니다. 변환하려는 시트 이름 앞에 '!'를 추가하세요.",
                        "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // '!'로 시작하는 시트가 있으면 excel2yamlconfig 시트 확인/생성
                ExcelConfigManager.Instance.EnsureConfigSheetExists();

                // XML 설정에서 Excel 설정으로 마이그레이션 (최초 1회)
                if (Properties.Settings.Default.FirstConfigMigration)
                {
                    ExcelConfigManager.Instance.MigrateFromXmlSettings(ExcelToYamlAddin.Config.SheetPathManager.Instance);
                    Properties.Settings.Default.FirstConfigMigration = false;
                    Properties.Settings.Default.Save();
                    Debug.WriteLine("[Ribbon] XML 설정을 Excel 설정으로 마이그레이션 완료");
                }

                // 활성화된 시트 수 확인
                int enabledSheetsCount = 0;
                Debug.WriteLine($"[OnConvertToJsonClick] 변환 가능한 시트 수: {convertibleSheets.Count}");
                foreach (var sheet in convertibleSheets)
                {
                    string currentSheetName = sheet.Name;
                    bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(currentSheetName);
                    Debug.WriteLine($"[OnConvertToJsonClick] 시트 '{currentSheetName}' 활성화 상태: {isEnabled}");

                    if (isEnabled)
                    {
                        enabledSheetsCount++;
                    }
                }

                Debug.WriteLine($"[OnConvertToJsonClick] 활성화된 시트 수: {enabledSheetsCount}, 비활성화된 시트 수: {convertibleSheets.Count - enabledSheetsCount}");

                // 활성화된 시트가 없는 경우 처리
                if (enabledSheetsCount == 0)
                {
                    Debug.WriteLine("[OnConvertToJsonClick] 경고: 활성화된 시트가 없습니다. 시트 활성화 상태 상세 정보 출력:");

                    // 활성화 상태 자세히 확인 (디버그용)
                    Dictionary<string, string> allEnabledPaths = SheetPathManager.Instance.GetAllEnabledSheetPaths();
                    Debug.WriteLine($"[OnConvertToJsonClick] GetAllEnabledSheetPaths 결과: {allEnabledPaths.Count}개 시트");
                    foreach (var kvp in allEnabledPaths)
                    {
                        Debug.WriteLine($"[OnConvertToJsonClick] 활성화된 시트: '{kvp.Key}', 경로: '{kvp.Value}'");
                    }

                    foreach (var sheet in convertibleSheets)
                    {
                        string sheetName = sheet.Name;
                        bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);
                        string sheetPath = SheetPathManager.Instance.GetSheetPath(sheetName);

                        Debug.WriteLine($"[OnConvertToJsonClick] 시트 '{sheetName}' - 활성화: {isEnabled}, 경로: '{sheetPath}'");

                        // 활성화 상태가 믿을 수 없는 경우 해당 시트를 강제로 활성화
                        if (isEnabled && !allEnabledPaths.ContainsKey(sheetName))
                        {
                            Debug.WriteLine($"[OnConvertToJsonClick] 활성화 상태 불일치 감지. 시트 '{sheetName}'를 강제로 활성화합니다.");
                            SheetPathManager.Instance.SetSheetEnabled(sheetName, true);
                            enabledSheetsCount++;
                        }
                    }

                    // 다시 활성화된 시트 수 확인
                    if (enabledSheetsCount == 0)
                    {
                        MessageBox.Show("활성화된 시트가 없어 변환을 취소합니다.\n\n시트 설정 창에서 시트를 활성화하십시오.",
                            "변환 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // 시트별 경로 설정 창 열기
                        using (var form = new Forms.SheetPathSettingsForm(convertibleSheets))
                        {
                            form.StartPosition = FormStartPosition.CenterScreen;
                            form.ShowDialog();
                        }
                        return;
                    }
                }

                // 활성화된 시트 수가 탐지된 시트 수보다 적은 경우 확인창 표시
                if (enabledSheetsCount < convertibleSheets.Count)
                {
                    int disabledCount = convertibleSheets.Count - enabledSheetsCount;
                    string message = $"{convertibleSheets.Count}개의 변환 가능한 시트 중 {disabledCount}개의 시트가 비활성화되어 있습니다.\n\n" +
                                    $"활성화된 {enabledSheetsCount}개의 시트만 변환하시겠습니까?\n\n" +
                                    "아니오를 선택하면 시트별 경로 설정 창이 열립니다.";

                    DialogResult result = MessageBox.Show(message, "시트 활성화 확인",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.No)
                    {
                        // 시트별 경로 설정 창 열기
                        if (settingsForm != null && !settingsForm.IsDisposed)
                        {
                            settingsForm.Activate();
                            return;
                        }

                        settingsForm = new Forms.SheetPathSettingsForm(convertibleSheets);
                        settingsForm.FormClosed += (s, args) =>
                        {
                            // 설정 후 다시 활성화된 시트 수 확인
                            enabledSheetsCount = 0;
                            foreach (var sheet in convertibleSheets)
                            {
                                if (SheetPathManager.Instance.IsSheetEnabled(sheet.Name))
                                {
                                    enabledSheetsCount++;
                                }
                            }

                            // 활성화된 시트가 없으면 변환 취소
                            if (enabledSheetsCount == 0)
                            {
                                MessageBox.Show("활성화된 시트가 없어 변환을 취소합니다.",
                                    "변환 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }

                            settingsForm = null;
                        };
                        settingsForm.StartPosition = FormStartPosition.CenterScreen;
                        settingsForm.Show();
                        return;
                    }
                }

                // 설정 적용 - JSON 변환으로 설정
                config.IncludeEmptyFields = includeEmptyFields;
                config.EnableHashGen = enableHashGen;
                config.OutputFormat = OutputFormat.Json;

                // 변환 전 설정 다시 로드 및 동기화
                SheetPathManager.Instance.Initialize();
                Debug.WriteLine("[OnConvertToJsonClick] 변환 전 SheetPathManager 재초기화 완료");

                // 변환 처리
                List<string> convertedFiles = ConvertExcelFile(config);

                // 변환 결과 메시지 표시
                if (convertedFiles != null && convertedFiles.Count > 0)
                {
                    MessageBox.Show($"{convertedFiles.Count}개의 파일이 성공적으로 변환되었습니다.",
                        "변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("변환된 파일이 없습니다.",
                        "변환 완료", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"변환 중 오류가 발생했습니다: {ex.Message}",
                    "변환 오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"[OnConvertToJsonClick] 오류: {ex.Message}\n{ex.StackTrace}");
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
                var convertibleSheets = Core.SheetAnalyzer.GetConvertibleSheets(app.ActiveWorkbook);

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
                settingsForm.FormClosed += (s, args) => { settingsForm = null; };
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

        // YAML을 JSON으로 변환 버튼 클릭
        public void OnConvertYamlToJsonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 설정 초기화 및 다시 로드
                SheetPathManager.Instance.Initialize();
                Debug.WriteLine("[OnConvertYamlToJsonClick] SheetPathManager 초기화 완료");

                // 현재 워크북 가져오기
                var addIn = Globals.ThisAddIn;
                var app = addIn.Application;
                var activeWorkbook = app.ActiveWorkbook;

                if (activeWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 워크북 경로 설정
                string workbookPath = activeWorkbook.FullName;
                SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);

                // Excel 설정 관리자 초기화
                ExcelConfigManager.Instance.SetCurrentWorkbook(workbookPath);

                // 변환 가능한 시트 찾기
                var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(activeWorkbook);

                if (convertibleSheets.Count == 0)
                {
                    MessageBox.Show("변환 가능한 시트가 없습니다. 변환하려는 시트 이름 앞에 '!'를 추가하세요.",
                        "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // '!'로 시작하는 시트가 있으면 excel2yamlconfig 시트 확인/생성
                ExcelConfigManager.Instance.EnsureConfigSheetExists();

                // 활성화된 시트 수 확인
                int enabledSheetsCount = 0;
                Debug.WriteLine($"[OnConvertYamlToJsonClick] 변환 가능한 시트 수: {convertibleSheets.Count}");
                foreach (var sheet in convertibleSheets)
                {
                    string currentSheetName = sheet.Name;
                    bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(currentSheetName);
                    Debug.WriteLine($"[OnConvertYamlToJsonClick] 시트 '{currentSheetName}' 활성화 상태: {isEnabled}");

                    if (isEnabled)
                    {
                        enabledSheetsCount++;
                    }
                }

                Debug.WriteLine($"[OnConvertYamlToJsonClick] 활성화된 시트 수: {enabledSheetsCount}, 비활성화된 시트 수: {convertibleSheets.Count - enabledSheetsCount}");

                // 활성화된 시트가 없는 경우 처리
                if (enabledSheetsCount == 0)
                {
                    MessageBox.Show("활성화된 시트가 없어 변환을 취소합니다.\n\n시트 설정 창에서 시트를 활성화하십시오.",
                        "변환 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // 시트별 경로 설정 창 열기
                        using (var form = new Forms.SheetPathSettingsForm(convertibleSheets))
                        {
                            form.StartPosition = FormStartPosition.CenterScreen;
                            form.ShowDialog();
                        }
                        return;
                }

                // 활성화된 시트 수가 탐지된 시트 수보다 적은 경우 확인창 표시
                if (enabledSheetsCount < convertibleSheets.Count)
                {
                    int disabledCount = convertibleSheets.Count - enabledSheetsCount;
                    string message = $"{convertibleSheets.Count}개의 변환 가능한 시트 중 {disabledCount}개의 시트가 비활성화되어 있습니다.\n\n" +
                                    $"활성화된 {enabledSheetsCount}개의 시트만 변환하시겠습니까?\n\n" +
                                    "아니오를 선택하면 시트별 경로 설정 창이 열립니다.";

                    DialogResult result = MessageBox.Show(message, "시트 활성화 확인",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (result == DialogResult.No)
                    {
                        // 시트별 경로 설정 창 열기
                        if (settingsForm != null && !settingsForm.IsDisposed)
                        {
                            settingsForm.Activate();
                            return;
                        }

                        settingsForm = new Forms.SheetPathSettingsForm(convertibleSheets);
                        settingsForm.FormClosed += (s, args) =>
                        {
                            // 설정 후 다시 활성화된 시트 수 확인
                            enabledSheetsCount = 0;
                            foreach (var sheet in convertibleSheets)
                            {
                                if (SheetPathManager.Instance.IsSheetEnabled(sheet.Name))
                                {
                                    enabledSheetsCount++;
                                }
                            }

                            // 활성화된 시트가 없으면 변환 취소
                            if (enabledSheetsCount == 0)
                            {
                                MessageBox.Show("활성화된 시트가 없어 변환을 취소합니다.",
                                    "변환 취소", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }

                            settingsForm = null;
                        };
                        settingsForm.StartPosition = FormStartPosition.CenterScreen;
                        settingsForm.Show();
                        return;
                    }
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
                            yamlFiles = ConvertExcelFileToTemp(yamlConfig, tempDir);

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

                            // 2단계: YAML 후처리 (MergeKeyPaths, FlowStyle 등)
                            progress.Report(new Forms.ProgressForm.ProgressInfo
                            {
                                Percentage = 30,
                                StatusMessage = "YAML 후처리 진행 중..."
                            });

                            int processedFiles = 0;
                            int totalFiles = yamlFiles.Count;

                            // 후처리를 위한 헬퍼 메서드
                            void CheckCancellation()
                            {
                                cancellationToken.ThrowIfCancellationRequested();
                            }

                            // 각 YAML 파일에 대해 후처리 적용
                            foreach (var yamlFilePath in yamlFiles)
                            {
                                CheckCancellation();

                                if (File.Exists(yamlFilePath))
                                {
                                    // 파일 이름 정보 추출
                                    string fileName = Path.GetFileNameWithoutExtension(yamlFilePath);

                                    progress.Report(new ProgressForm.ProgressInfo
                                    {
                                        Percentage = 30 + (int)((double)processedFiles / totalFiles * 20),
                                        StatusMessage = $"'{fileName}' YAML 후처리 중..."
                                    });

                                    // 가능한 시트 이름 형식
                                    string sheetName = null;

                                    // 워크북 내 시트 이름 매칭
                                    foreach (var sheet in convertibleSheets)
                                    {
                                        string currentSheetName = sheet.Name;
                                        if (currentSheetName.StartsWith("!"))
                                            currentSheetName = currentSheetName.Substring(1);

                                        if (string.Compare(currentSheetName, fileName, true) == 0)
                                        {
                                            sheetName = sheet.Name;
                                            break;
                                        }
                                    }

                                    if (sheetName != null)
                                    {
                                        // 1. YAML 선택적 필드 처리
                                        bool yamlEmptyFieldsOption = ExcelConfigManager.Instance.GetConfigBool(sheetName, "YamlEmptyFields", false);

                                        // Excel에 설정이 없으면 SheetPathManager에서 확인
                                        if (!yamlEmptyFieldsOption)
                                        {
                                            yamlEmptyFieldsOption = SheetPathManager.Instance.GetYamlEmptyFieldsOption(sheetName);
                                        }

                                        // 둘 다 없으면 기본 설정 사용
                                        if (!yamlEmptyFieldsOption && addEmptyYamlFields)
                                        {
                                            yamlEmptyFieldsOption = addEmptyYamlFields;
                                        }

                                        // 2. 키 경로 병합 처리
                                        progress.Report(new ProgressForm.ProgressInfo
                                        {
                                            StatusMessage = $"'{fileName}' - 키 경로 병합 처리 중..."
                                        });

                                        // 키 경로 병합 후처리
                                        if (sheetName != null)
                                        {
                                            // Excel excel2yamlconfig 시트에서 먼저 확인 (우선순위 변경: Excel이 더 높은 우선순위)
                                            string sheetMergeKeyPaths = ExcelConfigManager.Instance.GetConfigValue(sheetName, "MergeKeyPaths", "");

                                            // Excel에 설정이 없으면 SheetPathManager에서 확인
                                            if (string.IsNullOrEmpty(sheetMergeKeyPaths))
                                            {
                                                sheetMergeKeyPaths = SheetPathManager.Instance.GetMergeKeyPaths(sheetName);
                                            }

                                            // YAML 병합 키 경로 처리 실행
                                            if (!string.IsNullOrEmpty(sheetMergeKeyPaths))
                                            {
                                                Debug.WriteLine($"[OnConvertYamlToJsonClick] YAML 병합 후처리 실행: {yamlFilePath}, 설정: {sheetMergeKeyPaths}");

                                                bool success = YamlMergeKeyPathsProcessor.ProcessYamlFileFromConfig(
                                                    yamlFilePath,
                                                    sheetMergeKeyPaths,
                                                    yamlEmptyFieldsOption
                                                );

                                                if (success)
                                                {
                                                    Debug.WriteLine($"[OnConvertYamlToJsonClick] YAML 병합 후처리 완료: {yamlFilePath}");
                                                }
                                                else
                                                {
                                                    Debug.WriteLine($"[OnConvertYamlToJsonClick] YAML 병합 후처리 실패: {yamlFilePath}");
                                                }
                                            }
                                        }
                                    }

                                    processedFiles++;
                                }
                            }

                            // 3단계: JSON 변환 준비
                            progress.Report(new ProgressForm.ProgressInfo
                            {
                                Percentage = 50,
                                StatusMessage = $"{yamlFiles.Count}개 YAML 파일 JSON 변환 준비 중..."
                            });

                            // YAML 파일을 JSON으로 변환할 쌍 생성
                            foreach (var yamlFile in yamlFiles)
                            {
                                if (cancellationToken.IsCancellationRequested)
                                {
                                    break;
                                }

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
                            List<string> convertedJsonFiles = Core.YamlPostProcessors.YamlToJsonProcessor.BatchConvertYamlToJson(convertPairs);

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

        // 엑셀 파일을 임시 디렉토리에 YAML로 변환하는 메서드
        private List<string> ConvertExcelFileToTemp(ExcelToYamlConfig config, string tempDir)
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
                    return convertedFiles;
                }

                // 변환 가능한 시트 찾기
                var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(activeWorkbook);

                if (convertibleSheets.Count == 0)
                {
                    return convertedFiles;
                }

                // 임시 파일로 저장
                tempFile = addIn.SaveToTempFile();
                if (string.IsNullOrEmpty(tempFile))
                {
                    return convertedFiles;
                }

                // 모든 변환 가능한 시트에 대해 처리
                foreach (var sheet in convertibleSheets)
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
    }
}