using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using ExcelToJsonAddin.Config;

namespace ExcelToJsonAddin.Core
{
    /// <summary>
    /// Excel 파일 내의 !Config 시트에 설정을 저장하고 로드하는 관리자 클래스
    /// </summary>
    public class ExcelConfigManager
    {
        // 싱글톤 인스턴스
        private static ExcelConfigManager _instance;
        
        // 설정 시트 이름 상수
        public const string CONFIG_SHEET_NAME = "excel2yamlconfig";
        
        // 현재 워크북 경로
        private string _currentWorkbookPath;
        
        /// <summary>
        /// 현재 워크북 경로를 가져옵니다.
        /// </summary>
        public string WorkbookPath
        {
            get { return _currentWorkbookPath; }
        }
        
        // 캐시된 설정 값
        private Dictionary<string, Dictionary<string, string>> _sheetConfigCache;
        
        // 마지막으로 로드한 시간
        private DateTime _lastLoadTime;
        
        // 설정 열 인덱스 
        private const int SHEET_NAME_COL = 1;      // A열 - 시트 이름
        private const int CONFIG_KEY_COL = 2;      // B열 - 설정 키
        private const int CONFIG_VALUE_COL = 3;    // C열 - 설정 값
        private const int YAML_EMPTY_FIELDS_COL = 4; // D열 - YAML 선택적 필드 설정
        
        // 헤더 로우 인덱스
        private const int HEADER_ROW = 1;
        
        // 데이터 시작 로우 인덱스
        private const int DATA_START_ROW = 2;
        
        /// <summary>
        /// 싱글톤 인스턴스 가져오기
        /// </summary>
        public static ExcelConfigManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new ExcelConfigManager();
                }
                return _instance;
            }
        }
        
        /// <summary>
        /// 생성자
        /// </summary>
        private ExcelConfigManager()
        {
            _sheetConfigCache = new Dictionary<string, Dictionary<string, string>>();
            _lastLoadTime = DateTime.MinValue;
        }
        
        /// <summary>
        /// 현재 워크북 설정
        /// </summary>
        /// <param name="workbookPath">워크북 경로</param>
        public void SetCurrentWorkbook(string workbookPath)
        {
            if (_currentWorkbookPath != workbookPath)
            {
                _currentWorkbookPath = workbookPath;
                _sheetConfigCache.Clear();
                _lastLoadTime = DateTime.MinValue;
            }
        }
        
        /// <summary>
        /// 설정 시트가 존재하는지 확인하고 없으면 생성 (조건부)
        /// </summary>
        public void EnsureConfigSheetExists()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;
                
                Debug.WriteLine("[ExcelConfigManager] EnsureConfigSheetExists 메서드 시작");
                
                if (workbook == null)
                {
                    Debug.WriteLine("[ExcelConfigManager] 활성 워크북이 없습니다.");
                    return;
                }
                
                Debug.WriteLine($"[ExcelConfigManager] 활성 워크북 이름: {workbook.Name}");
                Debug.WriteLine($"[ExcelConfigManager] 활성 워크북 경로: {workbook.FullName}");
                
                // 모든 워크시트 이름 로깅
                Debug.WriteLine("[ExcelConfigManager] 워크북 내 모든 시트 목록:");
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    Debug.WriteLine($"[ExcelConfigManager] - 시트 이름: {sheet.Name}");
                }
                
                // 설정 시트 존재 여부 확인
                bool configSheetExists = false;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name == CONFIG_SHEET_NAME)
                    {
                        configSheetExists = true;
                        Debug.WriteLine("[ExcelConfigManager] !Config 시트가 이미 존재합니다.");
                        break;
                    }
                }
                
                // 이미 존재하면 더 진행하지 않음
                if (configSheetExists)
                {
                    Debug.WriteLine("[ExcelConfigManager] !Config 시트가 이미 존재하므로 생성을 건너뜁니다.");
                    return;
                }
                
                Debug.WriteLine("[ExcelConfigManager] !Config 시트가 존재하지 않습니다. 생성 조건 확인 중...");
                
                // '!'로 시작하는 시트가 있는지 확인
                bool hasExclamationSheet = false;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    if (sheet.Name.StartsWith("!") && sheet.Name != CONFIG_SHEET_NAME)
                    {
                        hasExclamationSheet = true;
                        Debug.WriteLine($"[ExcelConfigManager] '!'로 시작하는 시트 발견: {sheet.Name}");
                        break;
                    }
                }
                
                // '!'로 시작하는 시트가 없으면 생성하지 않음
                if (!hasExclamationSheet)
                {
                    Debug.WriteLine("[ExcelConfigManager] '!'로 시작하는 시트가 없어 !Config 시트를 생성하지 않습니다.");
                    return;
                }

                // 설정 시트가 없고 '!'로 시작하는 시트가 있으면 생성
                Debug.WriteLine("[ExcelConfigManager] 설정 시트 생성 조건 만족: !Config 시트 생성 시작");
                try {
                    Worksheet configSheet = workbook.Worksheets.Add();
                    Debug.WriteLine("[ExcelConfigManager] 새 워크시트 추가 성공");
                    
                    configSheet.Name = CONFIG_SHEET_NAME;
                    Debug.WriteLine("[ExcelConfigManager] 새 워크시트 이름을 !Config로 변경 성공");
                    
                    // 헤더 작성
                    configSheet.Cells[HEADER_ROW, SHEET_NAME_COL] = "SheetName";
                    configSheet.Cells[HEADER_ROW, CONFIG_KEY_COL] = "ConfigKey";
                    configSheet.Cells[HEADER_ROW, CONFIG_VALUE_COL] = "ConfigValue";
                    configSheet.Cells[HEADER_ROW, YAML_EMPTY_FIELDS_COL] = "YamlEmptyFields";
                    Debug.WriteLine("[ExcelConfigManager] 헤더 작성 완료");
                    
                    // 시트 숨기기
                    //configSheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
                    configSheet.Visible = XlSheetVisibility.xlSheetHidden;
                    Debug.WriteLine("[ExcelConfigManager] 시트 숨기기 완료");
                    
                    // 헤더 스타일 설정
                    Range headerRange = configSheet.Range[
                        configSheet.Cells[HEADER_ROW, SHEET_NAME_COL],
                        configSheet.Cells[HEADER_ROW, YAML_EMPTY_FIELDS_COL]
                    ];
                    headerRange.Font.Bold = true;
                    Debug.WriteLine("[ExcelConfigManager] 헤더 스타일 설정 완료");
                    
                    Debug.WriteLine("[ExcelConfigManager] !Config 시트 생성 완료");
                }
                catch (Exception innerEx) {
                    Debug.WriteLine($"[ExcelConfigManager] !Config 시트 생성 중 오류 발생: {innerEx.Message}");
                    Debug.WriteLine($"[ExcelConfigManager] 스택 트레이스: {innerEx.StackTrace}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExcelConfigManager] 설정 시트 확인/생성 중 오류: {ex.Message}");
                Debug.WriteLine($"[ExcelConfigManager] 스택 트레이스: {ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// !Config 시트에서 모든 설정 로드
        /// </summary>
        public void LoadAllSettings()
        {
            try
            {
                _sheetConfigCache.Clear();
                
                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;
                
                if (workbook == null)
                {
                    Debug.WriteLine("[ExcelConfigManager] 설정 로드 실패: 활성 워크북이 없습니다.");
                    return;
                }
                
                // 설정 시트 존재 여부 확인 및 생성
                EnsureConfigSheetExists();
                
                // 설정 시트 가져오기
                Worksheet configSheet = null;
                try
                {
                    configSheet = workbook.Worksheets[CONFIG_SHEET_NAME];
                }
                catch
                {
                    Debug.WriteLine("[ExcelConfigManager] 설정 시트를 찾을 수 없습니다.");
                    return;
                }
                
                // 마지막 사용 셀 찾기
                Range usedRange = configSheet.UsedRange;
                int lastRow = usedRange.Row + usedRange.Rows.Count - 1;
                
                // 헤더 이후의 데이터만 처리
                for (int row = DATA_START_ROW; row <= lastRow; row++)
                {
                    string sheetName = Convert.ToString(configSheet.Cells[row, SHEET_NAME_COL].Value);
                    string configKey = Convert.ToString(configSheet.Cells[row, CONFIG_KEY_COL].Value);
                    string configValue = Convert.ToString(configSheet.Cells[row, CONFIG_VALUE_COL].Value);
                    
                    if (!string.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(configKey))
                    {
                        // 캐시에 시트 설정 추가
                        if (!_sheetConfigCache.ContainsKey(sheetName))
                        {
                            _sheetConfigCache[sheetName] = new Dictionary<string, string>();
                        }
                        
                        _sheetConfigCache[sheetName][configKey] = configValue ?? "";
                    }
                }
                
                _lastLoadTime = DateTime.Now;
                Debug.WriteLine($"[ExcelConfigManager] 설정 로드 완료: {_sheetConfigCache.Count}개 시트의 설정을 로드했습니다.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ExcelConfigManager] 설정 로드 중 오류: " + ex.Message);
            }
        }
        
        /// <summary>
        /// 특정 시트의 모든 설정 값 가져오기
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>설정 키-값 딕셔너리</returns>
        public Dictionary<string, string> GetSheetConfig(string sheetName)
        {
            EnsureConfigLoaded();
            
            if (_sheetConfigCache.ContainsKey(sheetName))
            {
                return new Dictionary<string, string>(_sheetConfigCache[sheetName]);
            }
            
            return new Dictionary<string, string>();
        }
        
        /// <summary>
        /// 특정 시트의 특정 설정 값 가져오기
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="configKey">설정 키</param>
        /// <param name="defaultValue">기본값</param>
        /// <returns>설정 값 또는 기본값</returns>
        public string GetConfigValue(string sheetName, string configKey, string defaultValue = "")
        {
            EnsureConfigLoaded();
            
            if (_sheetConfigCache.ContainsKey(sheetName) && 
                _sheetConfigCache[sheetName].ContainsKey(configKey))
            {
                return _sheetConfigCache[sheetName][configKey];
            }
            
            return defaultValue;
        }
        
        /// <summary>
        /// 특정 시트의 특정 설정 값을 불리언으로 가져오기
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="configKey">설정 키</param>
        /// <param name="defaultValue">기본값</param>
        /// <returns>설정 값 또는 기본값</returns>
        public bool GetConfigBool(string sheetName, string configKey, bool defaultValue = false)
        {
            string value = GetConfigValue(sheetName, configKey, defaultValue.ToString());
            
            if (bool.TryParse(value, out bool result))
            {
                return result;
            }
            
            return defaultValue;
        }
        
        /// <summary>
        /// 특정 시트의 설정 값 저장
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="configKey">설정 키</param>
        /// <param name="configValue">설정 값</param>
        public void SetConfigValue(string sheetName, string configKey, string configValue)
        {
            try
            {
                // 캐시에 먼저 저장
                EnsureConfigLoaded();
                
                if (!_sheetConfigCache.ContainsKey(sheetName))
                {
                    _sheetConfigCache[sheetName] = new Dictionary<string, string>();
                }
                
                _sheetConfigCache[sheetName][configKey] = configValue;
                
                // 엑셀 시트에 저장
                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;
                
                if (workbook == null)
                {
                    Debug.WriteLine("[ExcelConfigManager] 설정 저장 실패: 활성 워크북이 없습니다.");
                    return;
                }
                
                // 설정 시트 존재 여부 확인 및 생성
                EnsureConfigSheetExists();
                
                // 설정 시트 가져오기
                Worksheet configSheet = workbook.Worksheets[CONFIG_SHEET_NAME];
                
                // 기존 설정이 있는지 확인
                Range usedRange = configSheet.UsedRange;
                int lastRow = usedRange.Row + usedRange.Rows.Count - 1;
                int targetRow = -1;
                
                for (int row = DATA_START_ROW; row <= lastRow; row++)
                {
                    string currentSheetName = Convert.ToString(configSheet.Cells[row, SHEET_NAME_COL].Value);
                    string currentKey = Convert.ToString(configSheet.Cells[row, CONFIG_KEY_COL].Value);
                    
                    if (currentSheetName == sheetName && currentKey == configKey)
                    {
                        targetRow = row;
                        break;
                    }
                }
                
                // 기존 설정이 없으면 새로운 행 추가
                if (targetRow == -1)
                {
                    targetRow = lastRow + 1;
                }
                
                // 설정 값 저장
                configSheet.Cells[targetRow, SHEET_NAME_COL] = sheetName;
                configSheet.Cells[targetRow, CONFIG_KEY_COL] = configKey;
                configSheet.Cells[targetRow, CONFIG_VALUE_COL] = configValue;
                
                Debug.WriteLine($"[ExcelConfigManager] 설정 저장 완료: {sheetName}.{configKey} = {configValue}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ExcelConfigManager] 설정 저장 중 오류: " + ex.Message);
            }
        }
        
        /// <summary>
        /// 설정이 로드되었는지 확인하고 필요시 로드
        /// </summary>
        private void EnsureConfigLoaded()
        {
            // 마지막 로드 후 5초 이상 지났거나 캐시가 비어있으면 다시 로드
            if (_lastLoadTime.AddSeconds(5) < DateTime.Now || _sheetConfigCache.Count == 0)
            {
                LoadAllSettings();
            }
        }
        
        /// <summary>
        /// 모든 시트의 특정 설정 키에 대한 값 가져오기
        /// </summary>
        /// <param name="configKey">설정 키</param>
        /// <returns>시트별 설정 값 딕셔너리</returns>
        public Dictionary<string, string> GetAllSheetsConfigValue(string configKey)
        {
            EnsureConfigLoaded();
            
            Dictionary<string, string> result = new Dictionary<string, string>();
            
            foreach (var sheetEntry in _sheetConfigCache)
            {
                if (sheetEntry.Value.ContainsKey(configKey))
                {
                    result[sheetEntry.Key] = sheetEntry.Value[configKey];
                }
            }
            
            return result;
        }
        
        /// <summary>
        /// XML 설정을 Excel 설정으로 마이그레이션
        /// </summary>
        /// <param name="sheetPathManager">시트 경로 매니저</param>
        public void MigrateFromXmlSettings(ExcelToJsonAddin.Config.SheetPathManager sheetPathManager)
        {
            try
            {
                if (sheetPathManager == null)
                {
                    Debug.WriteLine("[ExcelConfigManager] 마이그레이션 실패: SheetPathManager가 null입니다.");
                    return;
                }
                
                Debug.WriteLine("[ExcelConfigManager] XML 설정에서 Excel 설정으로 마이그레이션 시작");
                
                // 활성화된 워크북의 모든 시트에 대한 XML 설정 가져오기
                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;
                
                if (workbook == null)
                {
                    Debug.WriteLine("[ExcelConfigManager] 마이그레이션 실패: 활성 워크북이 없습니다.");
                    return;
                }
                
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    string sheetName = sheet.Name;
                    
                    // 경로 설정은 제외 (시트별 경로는 여전히 XML에서 관리)
                    
                    // YAML 선택적 필드 설정 마이그레이션
                    bool yamlOption = sheetPathManager.GetYamlEmptyFieldsOption(sheetName);
                    SetConfigValue(sheetName, "YamlEmptyFields", yamlOption.ToString());
                    
                    // 병합 키 경로 설정 마이그레이션
                    string mergeKeyPaths = sheetPathManager.GetMergeKeyPaths(sheetName);
                    if (!string.IsNullOrEmpty(mergeKeyPaths))
                    {
                        SetConfigValue(sheetName, "MergeKeyPaths", mergeKeyPaths);
                    }
                    
                    // Flow 스타일 설정 마이그레이션
                    string flowStyleConfig = sheetPathManager.GetFlowStyleConfig(sheetName);
                    if (!string.IsNullOrEmpty(flowStyleConfig))
                    {
                        SetConfigValue(sheetName, "FlowStyle", flowStyleConfig);
                    }
                }
                
                Debug.WriteLine("[ExcelConfigManager] 마이그레이션 완료");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ExcelConfigManager] 마이그레이션 중 오류: " + ex.Message);
            }
        }
        
        /// <summary>
        /// Excel 설정을 XML 설정으로 내보내기
        /// </summary>
        /// <param name="sheetPathManager">시트 경로 매니저</param>
        public void ExportToXmlSettings(ExcelToJsonAddin.Config.SheetPathManager sheetPathManager)
        {
            try
            {
                if (sheetPathManager == null)
                {
                    Debug.WriteLine("[ExcelConfigManager] XML 내보내기 실패: SheetPathManager가 null입니다.");
                    return;
                }
                
                Debug.WriteLine("[ExcelConfigManager] Excel 설정을 XML 설정으로 내보내기 시작");
                
                EnsureConfigLoaded();
                
                foreach (var sheetEntry in _sheetConfigCache)
                {
                    string sheetName = sheetEntry.Key;
                    var configs = sheetEntry.Value;
                    
                    // YAML 선택적 필드 설정 내보내기
                    if (configs.ContainsKey("YamlEmptyFields") && 
                        bool.TryParse(configs["YamlEmptyFields"], out bool yamlOption))
                    {
                        sheetPathManager.SetYamlEmptyFieldsOption(sheetName, yamlOption);
                    }
                    
                    // 병합 키 경로 설정 내보내기
                    if (configs.ContainsKey("MergeKeyPaths"))
                    {
                        sheetPathManager.SetMergeKeyPaths(sheetName, configs["MergeKeyPaths"]);
                    }
                    
                    // Flow 스타일 설정 내보내기
                    if (configs.ContainsKey("FlowStyle"))
                    {
                        sheetPathManager.SetFlowStyleConfig(sheetName, configs["FlowStyle"]);
                    }
                }
                
                Debug.WriteLine("[ExcelConfigManager] XML 내보내기 완료");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[ExcelConfigManager] XML 내보내기 중 오류: " + ex.Message);
            }
        }
        
        /// <summary>
        /// 설정 시트를 보이게 만듭니다.
        /// </summary>
        public void ShowConfigSheet()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var workbook = app.ActiveWorkbook;
                
                if (workbook == null)
                {
                    Debug.WriteLine("[ExcelConfigManager] 설정 시트 표시 실패: 활성 워크북이 없습니다.");
                    return;
                }
                
                // 설정 시트 가져오기
                Worksheet configSheet = null;
                try
                {
                    configSheet = workbook.Worksheets[CONFIG_SHEET_NAME];
                    configSheet.Visible = XlSheetVisibility.xlSheetVisible;
                    Debug.WriteLine("[ExcelConfigManager] 설정 시트를 보이게 변경했습니다.");
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"[ExcelConfigManager] 설정 시트를 찾을 수 없거나 표시하는 중 오류 발생: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ExcelConfigManager] 설정 시트 표시 중 오류: {ex.Message}");
            }
        }
    }
} 