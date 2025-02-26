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
                Debug.WriteLine("리본 UI가 로드되었습니다.");
                
                // 설정 로드 확인
                var pathManager = SheetPathManager.Instance;
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
        
        // JSON으로 변환 버튼 클릭
        public void OnConvertToJsonClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 설정 적용
                config.IncludeEmptyFields = includeEmptyFields;
                config.EnableHashGen = enableHashGen;
                config.OutputFormat = OutputFormat.Json;
                
                // 변환 처리
                ConvertExcelFile(config);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"JSON 변환 중 오류가 발생했습니다: {ex.Message}", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"JSON 변환 오류: {ex}");
            }
        }
        
        // YAML으로 변환 버튼 클릭
        public void OnConvertToYamlClick(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 설정 적용
                config.IncludeEmptyFields = includeEmptyFields;
                config.EnableHashGen = enableHashGen;
                config.OutputFormat = OutputFormat.Yaml;
                
                // 변환 처리
                List<string> convertedFiles = ConvertExcelFile(config);
                
                // YAML 후처리 기능은 이제 시트별로 적용됨
                // 전역 addEmptyYamlFields 변수 대신 시트별 설정 사용
                if (convertedFiles != null && convertedFiles.Count > 0)
                {
                    try
                    {
                        Debug.WriteLine($"[Ribbon] YAML 후처리 확인: {convertedFiles.Count}개 파일");
                        int successCount = 0;
                        
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
                                Debug.WriteLine($"[Ribbon] 추출한 파일명: {fileName}");
                                Debug.WriteLine($"[Ribbon] 워크북 이름: {workbookName}");
                                
                                // SheetPathManager 초기화 확인
                                if (SheetPathManager.Instance != null)
                                {
                                    SheetPathManager.Instance.SetCurrentWorkbook(Globals.ThisAddIn.Application.ActiveWorkbook.FullName);
                                }
                                else
                                {
                                    Debug.WriteLine($"[Ribbon] 오류: SheetPathManager 인스턴스가 null입니다.");
                                }
                                
                                bool useEmptyOptionals = false;
                                string matchedSheetName = null;
                                
                                // 여러 가능한 시트 이름으로 설정 확인
                                foreach (var sheetName in possibleSheetNames)
                                {
                                    Debug.WriteLine($"[Ribbon] 시트 이름으로 설정 확인: {sheetName}");
                                    bool option = SheetPathManager.Instance.GetYamlEmptyFieldsOption(sheetName);
                                    Debug.WriteLine($"[Ribbon] '{sheetName}' YAML 선택적 필드 설정: {option}");
                                    
                                    if (option)
                                    {
                                        useEmptyOptionals = true;
                                        matchedSheetName = sheetName;
                                        break;
                                    }
                                }
                                
                                Debug.WriteLine($"[Ribbon] 최종 YAML 선택적 필드 처리 여부: {useEmptyOptionals}" + 
                                    (matchedSheetName != null ? $" (시트: {matchedSheetName})" : " (매칭된 시트 없음)"));
                                
                                // useEmptyOptionals가 true인 경우에만 처리
                                if (useEmptyOptionals)
                                {
                                    Debug.WriteLine($"[Ribbon] YAML 파일에 선택적 필드 처리 시작: {filePath}");
                                    
                                    // 새로 만든 프로세서 사용
                                    var processor = new YamlOptionalFieldsProcessor();
                                    if (processor.ProcessYamlFile(filePath))
                                    {
                                        successCount++;
                                        Debug.WriteLine($"[Ribbon] YAML 파일 선택적 필드 처리 완료: {filePath}");
                                    }
                                    else
                                    {
                                        Debug.WriteLine($"[Ribbon] YAML 파일 선택적 필드 처리 실패: {filePath}");
                                    }
                                }
                                else
                                {
                                    Debug.WriteLine($"[Ribbon] YAML 선택적 필드 처리 건너뜀 (설정 꺼짐): {filePath}");
                                }
                            }
                        }
                        
                        Debug.WriteLine($"[Ribbon] YAML 후처리 완료: {successCount}/{convertedFiles.Count} 파일 처리됨");
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"[Ribbon] YAML 후처리 중 오류 발생: {ex.Message}");
                    }
                }
                
                // 변환 완료 전역 메시지는 ConvertExcelFile 메서드에서 처리됨
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
                var activeWorkbook = app.ActiveWorkbook;
                
                if (activeWorkbook == null)
                {
                    MessageBox.Show("활성 워크북이 없습니다.", "오류", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                
                // 워크북 경로 설정
                string workbookPath = activeWorkbook.FullName;
                SheetPathManager.Instance.SetCurrentWorkbook(workbookPath);
                
                // 변환 가능한 시트 찾기
                var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(activeWorkbook);
                
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
                    
                    // 모든 시트 경로 확인
                    var allPaths = SheetPathManager.Instance.GetAllSheetPaths();
                    Debug.WriteLine($"저장된 시트 경로 항목 수: {allPaths.Count}");
                    foreach(var path in allPaths)
                    {
                        Debug.WriteLine($"시트: {path.Key}, 경로: {path.Value}");
                    }
                    
                    // 저장 경로가 유효하지 않으면 건너뛰기
                    if (string.IsNullOrEmpty(savePath))
                    {
                        Debug.WriteLine($"시트 '{sheetName}'의 저장 경로가 설정되지 않았습니다. 건너뛰기");
                        skipCount++;
                        continue;
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
                
                // 변환 결과 메시지
                if (successCount > 0)
                {
                    string message = $"{successCount}개의 시트가 성공적으로 변환되었습니다.";
                    if (skipCount > 0)
                    {
                        message += $"\n{skipCount}개의 시트는 변환되지 않았습니다.";
                    }
                    
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
                    
                    // 첫 번째 파일이 있는 폴더 열기
                    if (convertedFiles.Count > 0 && MessageBox.Show("변환된 파일이 있는 폴더를 열까요?", 
                        "확인", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        Process.Start("explorer.exe", $"/select,\"{convertedFiles[0]}\"");
                    }
                }
                else
                {
                    MessageBox.Show("변환된 시트가 없습니다. 시트별 저장 경로를 설정했는지 확인하세요.", 
                        "알림", MessageBoxButtons.OK, MessageBoxIcon.Information);
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