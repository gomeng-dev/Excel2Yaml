using System;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using ExcelToJsonAddin.Config;
using ExcelToJsonAddin.Core;
using ExcelToJsonAddin.Core.YamlPostProcessors;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Microsoft.Office.Core;
using System.Reflection;
using ExcelToJsonAddin.Properties;
using ExcelToJsonAddin.Forms;
using Microsoft.Office.Interop.Excel;

namespace ExcelToJsonAddin
{
    public partial class Ribbon : RibbonBase
    {
        // 옵션 설정
        private bool includeEmptyFields = false;
        private bool enableHashGen = false;
        private bool addEmptyYamlFields = false;
        
        private readonly ExcelToJsonConfig config = new ExcelToJsonConfig();

        public Ribbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
            
            // 설정 불러오기
            try
            {
                addEmptyYamlFields = Properties.Settings.Default.AddEmptyYamlFields;
                Debug.WriteLine($"[Ribbon] 설정에서 YAML 선택적 필드 처리 상태 로드: {addEmptyYamlFields}");
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
                var pathManager = ExcelToJsonAddin.Config.SheetPathManager.Instance;
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
                
                // 설정에서 기본 YAML 옵션 로드
                addEmptyYamlFields = Properties.Settings.Default.AddEmptyYamlFields;
                Debug.WriteLine($"[Ribbon_Load] 기본 YAML 선택적 필드 처리 상태: {addEmptyYamlFields}");
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
                
                // '!'로 시작하는 시트가 있으면 !Config 시트 확인/생성
                ExcelConfigManager.Instance.EnsureConfigSheetExists();
                
                // XML 설정에서 Excel 설정으로 마이그레이션 (최초 1회)
                if (Properties.Settings.Default.FirstConfigMigration)
                {
                    ExcelConfigManager.Instance.MigrateFromXmlSettings(ExcelToJsonAddin.Config.SheetPathManager.Instance);
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
                        using (var form = new Forms.SheetPathSettingsForm(convertibleSheets))
                        {
                            form.ShowDialog();
                            
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
                                return;
                            }
                        }
                    }
                }
                
                // 설정 적용
                config.IncludeEmptyFields = includeEmptyFields || addEmptyYamlFields;
                config.EnableHashGen = enableHashGen;
                config.OutputFormat = OutputFormat.Yaml;
                
                // 변환 전 설정 다시 로드 및 동기화
                SheetPathManager.Instance.Initialize();
                Debug.WriteLine("[OnConvertToYamlClick] 변환 전 SheetPathManager 재초기화 완료");
                
                // 변환 처리
                List<string> convertedFiles = ConvertExcelFile(config);
                
                // 변환 결과 추적
                int successCount = 0;
                int mergeKeyPathsSuccessCount = 0;
                int flowStyleSuccessCount = 0;
                // 선택적 필드 처리가 SerializeToYaml에 통합되어 제거
                //bool processYamlEmptyFields = false;
                
                // YAML 후처리 기능 적용
                if (convertedFiles != null && convertedFiles.Count > 0)
                {
                    try
                    {
                        Debug.WriteLine($"[Ribbon] YAML 후처리 확인: {convertedFiles.Count}개 파일");
                        successCount = convertedFiles.Count;
                        
                        foreach (var filePath in convertedFiles)
                        {
                            if (File.Exists(filePath) && Path.GetExtension(filePath).ToLower() == ".yaml")
                            {
                                // 파일 경로에서 시트 이름 추출
                                string fileName = Path.GetFileNameWithoutExtension(filePath);
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
                                
                                // YAML 선택적 필드 처리
                                if (matchedSheetName != null)
                                {
                                    // Excel !Config 시트에서 먼저 확인 (우선순위 변경: Excel이 더 높은 우선순위)
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
                                    
                                    // 빈 필드 처리는 이제 불필요함 - SerializeToYaml 메서드가 직접 처리
                                    // config.IncludeEmptyFields 옵션을 통해 빈 필드를 유지하거나 제거하도록 설정됨
                                    
                                    // 키 경로 병합 후처리
                                    if (matchedSheetName != null)
                                    {
                                        // Excel !Config 시트에서 먼저 확인 (우선순위 변경: Excel이 더 높은 우선순위)
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
                                
                                // Flow Style 처리
                                if (matchedSheetName != null)
                                {
                                    // 시트 이름을 로깅하여 디버깅을 돕습니다
                                    Debug.WriteLine($"[Ribbon] Flow Style 처리 검사 중: 파일 경로={filePath}, 파일 이름={fileName}, 매칭된 시트 이름={matchedSheetName}");
                                    
                                    // Excel !Config 시트에서 먼저 확인 (우선순위 변경: Excel이 더 높은 우선순위)
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
                            }
                        }
                        
                        Debug.WriteLine($"[Ribbon] YAML 선택적 필드 후처리 완료: {successCount}/{convertedFiles.Count} 파일 처리됨");
                        Debug.WriteLine($"[Ribbon] YAML 키 경로 병합 후처리 완료: {mergeKeyPathsSuccessCount}/{convertedFiles.Count} 파일 처리됨");
                        Debug.WriteLine($"[Ribbon] YAML Flow Style 후처리 완료: {flowStyleSuccessCount}/{convertedFiles.Count} 파일 처리됨");
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
            
            // 설정 저장
            try
            {
                Properties.Settings.Default.AddEmptyYamlFields = pressed;
                Properties.Settings.Default.Save();
                Debug.WriteLine($"[Ribbon] YAML 선택적 필드 처리 상태 저장: {pressed}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Ribbon] 설정 저장 중 오류 발생: {ex.Message}");
            }
        }

        // 고급 설정 버튼 클릭
        public void OnSettingsClick(IRibbonControl control)
        {
            MessageBox.Show("고급 설정 기능은 개발 중입니다.", "정보", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"시트별 경로 설정 중 오류가 발생했습니다: {ex.Message}", "오류", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                
                // 시트별 경로 설정 대화상자 표시
                using (var form = new Forms.SheetPathSettingsForm(convertibleSheets))
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"시트별 경로 설정 중 오류가 발생했습니다: {ex.Message}", 
                    "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Excel 파일 변환 처리 (수정: 변환된 파일 목록 반환)
        private List<string> ConvertExcelFile(ExcelToJsonConfig config)
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

                int successCount = 0;
                int skipCount = 0;
                // 선택적 필드 처리가 SerializeToYaml에 통합되어 제거
                //bool processYamlEmptyFields = false;
                
                // 모든 변환 가능한 시트에 대해 처리
                foreach (var sheet in convertibleSheets)
                {
                    string sheetName = sheet.Name;
                    Debug.WriteLine($"처리 중인 시트: {sheetName}");
                    
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
                        Debug.WriteLine($"시트 '{sheetName}'의 저장 경로가 설정되지 않았습니다. 건너뛰기");
                        
                        // 활성화 상태 확인 - 활성화된 시트인데 경로가 없는 경우 사용자에게 알림
                        bool sheetIsEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);
                        if (sheetIsEnabled)
                        {
                            // 활성화된 모든 시트 경로 가져오기
                            var allEnabledPaths = SheetPathManager.Instance.GetAllEnabledSheetPaths();
                            Debug.WriteLine($"GetAllEnabledSheetPaths 결과: {allEnabledPaths.Count}개의 활성화된 시트");
                            
                            // GetAllEnabledSheetPaths에 포함되어 있지만 저장 경로가 비어있는지 확인
                            if (allEnabledPaths.ContainsKey(sheetName) && string.IsNullOrEmpty(allEnabledPaths[sheetName]))
                            {
                                Debug.WriteLine($"활성화된 시트 '{sheetName}'은 저장 경로가 비어있습니다.");
                            }
                            
                            MessageBox.Show($"시트 '{sheetName}'이 활성화되어 있지만 저장 경로가 설정되지 않았습니다.\n" +
                                           "경로를 설정하시겠습니까?", "경로 설정 필요", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            
                            // 경로 설정 창 열기
                            using (var form = new Forms.SheetPathSettingsForm(convertibleSheets))
                            {
                                form.ShowDialog();
                                
                                // 경로 설정 후 다시 확인
                                savePath = SheetPathManager.Instance.GetSheetPath(sheetName);
                                Debug.WriteLine($"경로 설정 후 시트 '{sheetName}'의 저장 경로: {savePath ?? "설정되지 않음"}");
                                
                                if (string.IsNullOrEmpty(savePath))
                                {
                                    Debug.WriteLine($"시트 '{sheetName}'의 저장 경로가 여전히 설정되지 않았습니다. 건너뛰기");
                                    skipCount++;
                                    continue;
                                }
                            }
                        }
                        else
                        {
                            skipCount++;
                            continue;
                        }
                    }
                    
                    // 활성화 상태 확인 - 비활성화된 시트는 건너뛰기
                    bool isEnabled = SheetPathManager.Instance.IsSheetEnabled(sheetName);
                    if (!isEnabled)
                    {
                        Debug.WriteLine($"시트 '{sheetName}'은 비활성화 상태입니다. 건너뛰기");
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
                        Debug.WriteLine($"시트 '{sheetName}' 변환 중 오류 발생: {ex.Message}");
                        skipCount++;
                    }
                }
                
                // 변환 결과 로그 작성
                Debug.WriteLine($"변환 완료: {successCount}개 성공, {skipCount}개 실패");
                
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
                        Debug.WriteLine($"임시 파일 삭제 중 오류 발생: {ex.Message}");
                    }
                }
            }
            
            return convertedFiles;
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