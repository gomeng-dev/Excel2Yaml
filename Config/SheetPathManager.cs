using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using ExcelToJsonAddin.Properties;
using System.Diagnostics;
using System.Xml;

namespace ExcelToJsonAddin.Config
{
    /// <summary>
    /// 시트별 경로 관리를 위한 클래스
    /// </summary>
    public class SheetPathManager
    {
        private static readonly string ConfigFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "ExcelToJsonAddin",
            "SheetPaths.xml");

        // 싱글톤 인스턴스
        private static SheetPathManager _instance;

        // 워크북 파일 경로와 시트 이름을 키로 사용하는 딕셔너리
        // 키: 워크북 경로, 값: 시트 이름과 경로 정보의 딕셔너리
        private Dictionary<string, Dictionary<string, SheetPathInfo>> _sheetPaths;

        // 현재 워크북 경로
        private string _currentWorkbookPath;

        public static SheetPathManager Instance
        {
            get
            {
                return GetInstance();
            }
        }

        private static SheetPathManager GetInstance()
        {
            if (_instance == null)
            {
                _instance = new SheetPathManager();
            }
            return _instance;
        }

        // 생성자
        private SheetPathManager()
        {
            _sheetPaths = new Dictionary<string, Dictionary<string, SheetPathInfo>>();
        }

        // 현재 작업 중인 워크북 설정
        public void SetCurrentWorkbook(string workbookPath)
        {
            if (string.IsNullOrEmpty(workbookPath))
            {
                Debug.WriteLine($"[SetCurrentWorkbook] 오류: 워크북 경로가 비어있습니다.");
                return;
            }

            // 설정이 로드되지 않았으면 로드
            if (_sheetPaths == null)
            {
                Debug.WriteLine($"[SetCurrentWorkbook] 설정이 로드되지 않아 로드합니다.");
                LoadSheetPaths();
            }

            // OneDrive/SharePoint 경로에서 특수 문자 처리
            string normalizedPath = NormalizeWorkbookPath(workbookPath);
            string fileName = Path.GetFileName(workbookPath);
            
            // 디버그용 정보 출력
            Debug.WriteLine($"[SetCurrentWorkbook] 워크북 정보: 원본 경로='{workbookPath}', 정규화된 경로='{normalizedPath}', 파일명='{fileName}'");
            
            // 현재 워크북 정보 설정
            _currentWorkbookPath = workbookPath; // 원본 경로 유지 (참조용)
            
            // 현재 워크북에 대한 시트 정보 출력
            if (LazyLoadSheetPaths().ContainsKey(workbookPath))
            {
                Debug.WriteLine($"[SetCurrentWorkbook] 원본 경로 '{workbookPath}'에 저장된 시트 정보:");
                DumpSheetInfo(LazyLoadSheetPaths()[workbookPath]);
            }
            
            if (normalizedPath != workbookPath && LazyLoadSheetPaths().ContainsKey(normalizedPath))
            {
                Debug.WriteLine($"[SetCurrentWorkbook] 정규화된 경로 '{normalizedPath}'에 저장된 시트 정보:");
                DumpSheetInfo(LazyLoadSheetPaths()[normalizedPath]);
            }
            
            if (LazyLoadSheetPaths().ContainsKey(fileName))
            {
                Debug.WriteLine($"[SetCurrentWorkbook] 파일명 '{fileName}'에 저장된 시트 정보:");
                DumpSheetInfo(LazyLoadSheetPaths()[fileName]);
            }

            // 원본 경로에 대한 딕셔너리 생성
            if (!LazyLoadSheetPaths().ContainsKey(workbookPath))
            {
                LazyLoadSheetPaths()[workbookPath] = new Dictionary<string, SheetPathInfo>();
                Debug.WriteLine($"[SetCurrentWorkbook] 원본 경로 '{workbookPath}'에 대한 새 사전 생성");
            }
            
            // 정규화된 경로에 대한 딕셔너리 생성
            if (normalizedPath != workbookPath && !LazyLoadSheetPaths().ContainsKey(normalizedPath))
            {
                LazyLoadSheetPaths()[normalizedPath] = new Dictionary<string, SheetPathInfo>();
                Debug.WriteLine($"[SetCurrentWorkbook] 정규화된 경로 '{normalizedPath}'에 대한 새 사전 생성");
            }

            // 파일명만으로도 딕셔너리 생성
            if (!string.IsNullOrEmpty(fileName) && !LazyLoadSheetPaths().ContainsKey(fileName))
            {
                LazyLoadSheetPaths()[fileName] = new Dictionary<string, SheetPathInfo>();
                Debug.WriteLine($"[SetCurrentWorkbook] 워크북 '{fileName}'에 대한 새 사전 생성");
            }

            // 현재 로드된 워크북 디버그 정보
            Debug.WriteLine($"[SetCurrentWorkbook] 현재 로드된 워크북 수: {LazyLoadSheetPaths().Count}");
            foreach (var wb in LazyLoadSheetPaths().Keys)
            {
                Debug.WriteLine($"[SetCurrentWorkbook] 로드된 워크북: {wb}, 시트 수: {LazyLoadSheetPaths()[wb].Count}");
            }
            
            // 원본 경로와 정규화된 경로, 파일명 간의 데이터 동기화
            SynchronizeWorkbookData(workbookPath, normalizedPath, fileName);
        }
        
        // 워크북 데이터 동기화 (원본 경로, 정규화된 경로, 파일명 간의 시트 정보 복사)
        private void SynchronizeWorkbookData(string originalPath, string normalizedPath, string fileName)
        {
            // 세 가지 키: 원본 경로, 정규화된 경로, 파일명 
            var paths = new List<string> { originalPath };
            
            if (normalizedPath != originalPath)
                paths.Add(normalizedPath);
                
            if (!string.IsNullOrEmpty(fileName))
                paths.Add(fileName);
                
            // 모든 경로 조합에 대해 서로 데이터 동기화
            for (int i = 0; i < paths.Count; i++)
            {
                for (int j = 0; j < paths.Count; j++)
                {
                    if (i == j) continue; // 같은 경로는 스킵
                    
                    string sourcePath = paths[i];
                    string targetPath = paths[j];
                    
                    // 소스 경로에 시트 정보가 있으면 대상 경로로 복사
                    if (LazyLoadSheetPaths().ContainsKey(sourcePath) && LazyLoadSheetPaths()[sourcePath].Count > 0)
                    {
                        foreach (var sheet in LazyLoadSheetPaths()[sourcePath])
                        {
                            string sheetName = sheet.Key;
                            SheetPathInfo sheetInfo = sheet.Value;
                            
                            // 대상 경로에 시트 정보가 없거나 다른 경우 복사
                            if (!LazyLoadSheetPaths()[targetPath].ContainsKey(sheetName))
                            {
                                LazyLoadSheetPaths()[targetPath][sheetName] = sheetInfo;
                                Debug.WriteLine($"[SynchronizeWorkbookData] '{sourcePath}'에서 '{targetPath}'로 시트 '{sheetName}' 정보 복사");
                            }
                            else
                            {
                                // 활성화 상태 확인 및 업데이트
                                if (LazyLoadSheetPaths()[targetPath][sheetName].Enabled != sheetInfo.Enabled)
                                {
                                    bool sourceEnabled = sheetInfo.Enabled;
                                    bool targetEnabled = LazyLoadSheetPaths()[targetPath][sheetName].Enabled;
                                    
                                    // 둘 중 하나라도 활성화되어 있으면 모두 활성화
                                    if (sourceEnabled || targetEnabled)
                                    {
                                        LazyLoadSheetPaths()[sourcePath][sheetName].Enabled = true;
                                        LazyLoadSheetPaths()[targetPath][sheetName].Enabled = true;
                                        Debug.WriteLine($"[SynchronizeWorkbookData] 시트 '{sheetName}'의 활성화 상태 동기화 (true로 설정)");
                                    }
                                }
                            }
                        }
                    }
                }
            }
            
            // 설정 저장
            SaveSettings();
        }

        // 시트 정보 로깅을 위한 헬퍼 메서드
        private void DumpSheetInfo(Dictionary<string, SheetPathInfo> sheetInfos)
        {
            if (sheetInfos == null || sheetInfos.Count == 0)
            {
                Debug.WriteLine("    (시트 정보 없음)");
                return;
            }
            
            foreach (var sheet in sheetInfos)
            {
                Debug.WriteLine($"    시트: '{sheet.Key}', 활성화: {sheet.Value.Enabled}, 경로: '{sheet.Value.SavePath}'");
            }
        }

        /// <summary>
        /// 워크북 경로를 정규화합니다 (OneDrive/SharePoint URL 처리)
        /// </summary>
        private string NormalizeWorkbookPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;

            // OneDrive/SharePoint URL 경로 정규화
            if (IsOneDrivePath(path))
            {
                // URL 인코딩된 문자 디코딩
                string decoded = Uri.UnescapeDataString(path);
                
                // 슬래시 방향 통일
                decoded = decoded.Replace('\\', '/');
                
                // 중복 슬래시 제거 - https:// 부분은 유지
                if (decoded.StartsWith("http://"))
                {
                    string protocol = "http://";
                    string remaining = decoded.Substring(protocol.Length);
                    remaining = remaining.Replace("//", "/");
                    decoded = protocol + remaining;
                }
                else if (decoded.StartsWith("https://"))
                {
                    string protocol = "https://";
                    string remaining = decoded.Substring(protocol.Length);
                    remaining = remaining.Replace("//", "/");
                    decoded = protocol + remaining;
                }
                
                Debug.WriteLine($"[NormalizeWorkbookPath] 원본 경로: {path}");
                Debug.WriteLine($"[NormalizeWorkbookPath] 정규화된 경로: {decoded}");
                
                return decoded;
            }
            
            return path;
        }
        
        /// <summary>
        /// 경로가 OneDrive/SharePoint URL 경로인지 확인합니다
        /// </summary>
        private bool IsOneDrivePath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;
                
            return path.StartsWith("http://") || 
                   path.StartsWith("https://") || 
                   path.Contains("sharepoint.com") ||
                   path.Contains("onedrive.com");
        }

        /// <summary>
        /// 시트의 경로를 설정합니다.
        /// </summary>
        /// <param name="workbookName">워크북 이름</param>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="path">저장할 경로</param>
        public void SetSheetPath(string workbookName, string sheetName, string path)
        {
            if (string.IsNullOrEmpty(workbookName) || string.IsNullOrEmpty(sheetName))
                return;

            try
            {
                // 워크북 경로 정규화
                string normalizedPath = NormalizeWorkbookPath(workbookName);
                string fileName = Path.GetFileName(workbookName);
                
                // 전체 경로로 시트 경로 설정 시도
                bool setFullPathResult = SetSheetPathInternal(normalizedPath, sheetName, path);
                Debug.WriteLine($"[SetSheetPath] 전체 경로 '{normalizedPath}'에 시트 '{sheetName}' 경로 설정 {(setFullPathResult ? "성공" : "실패")}");
                
                // 파일명만으로 시트 경로 설정 시도 (성공한 경우만)
                if (setFullPathResult && !string.IsNullOrEmpty(fileName))
                {
                    bool setFileNameResult = SetSheetPathInternal(fileName, sheetName, path);
                    Debug.WriteLine($"[SetSheetPath] 파일명 '{fileName}'에 시트 '{sheetName}' 경로 설정 {(setFileNameResult ? "성공" : "실패")}");
                }
                
                // 설정 저장 - 항상 즉시 저장하도록 수정
                Debug.WriteLine($"[SetSheetPath] 설정 즉시 저장");
                SaveSheetPaths();
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SetSheetPath] 시트 경로 설정 중 예외 발생: {ex.Message}\n{ex.StackTrace}");
            }
        }
        
        /// <summary>
        /// 내부적으로 시트의 경로를 설정합니다.
        /// </summary>
        private bool SetSheetPathInternal(string workbookKey, string sheetName, string path)
        {
            try
            {
                // 잘못된 인수가 있으면 설정하지 않음
                if (string.IsNullOrEmpty(workbookKey) || string.IsNullOrEmpty(sheetName))
                {
                    Debug.WriteLine($"[SetSheetPathInternal] 잘못된 인수: workbookKey='{workbookKey}', sheetName='{sheetName}'");
                    return false;
                }

                // 해당 워크북 항목이 없으면 생성
                if (!LazyLoadSheetPaths().ContainsKey(workbookKey))
                {
                    Debug.WriteLine($"[SetSheetPathInternal] 워크북 '{workbookKey}'에 대한 새 딕셔너리 생성");
                    LazyLoadSheetPaths()[workbookKey] = new Dictionary<string, SheetPathInfo>();
                }

                // 해당 시트 항목이 없으면 생성 - 활성화 상태를 자동으로 변경하지 않음
                bool currentEnabled = false; // 기본값은 비활성화
                if (!LazyLoadSheetPaths()[workbookKey].ContainsKey(sheetName))
                {
                    LazyLoadSheetPaths()[workbookKey][sheetName] = new SheetPathInfo 
                    {
                        SavePath = "",
                        Enabled = false, // 시트를 새로 추가할 때는 기본적으로 비활성화 상태로 설정
                        YamlEmptyFields = false
                    };
                    Debug.WriteLine($"[SetSheetPathInternal] 시트 '{sheetName}'이 사전에 없어 새로 생성했으며 기본적으로 비활성화 상태로 설정");
                }
                else
                {
                    // 기존 활성화 상태 유지
                    currentEnabled = LazyLoadSheetPaths()[workbookKey][sheetName].Enabled;
                    Debug.WriteLine($"[SetSheetPathInternal] 시트 '{sheetName}'의 기존 활성화 상태({currentEnabled})를 유지합니다");
                }

                // 경로 설정 - 활성화 상태는 변경하지 않음
                if (path != null)
                {
                    LazyLoadSheetPaths()[workbookKey][sheetName].SavePath = path;
                    Debug.WriteLine($"[SetSheetPathInternal] 시트 경로 설정됨: workbookKey='{workbookKey}', sheetName='{sheetName}', path='{path}', 활성화 상태: {LazyLoadSheetPaths()[workbookKey][sheetName].Enabled}");
                }
                else
                {
                    LazyLoadSheetPaths()[workbookKey][sheetName].SavePath = "";
                    Debug.WriteLine($"[SetSheetPathInternal] 시트 경로 초기화됨: workbookKey='{workbookKey}', sheetName='{sheetName}', 활성화 상태: {LazyLoadSheetPaths()[workbookKey][sheetName].Enabled}");
                }

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SetSheetPathInternal] 경로 설정 중 예외 발생: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 시트의 저장 경로를 가져옵니다.
        /// </summary>
        /// <param name="sheetName">찾을 시트 이름</param>
        /// <returns>설정된 저장 경로 또는 빈 문자열</returns>
        public string GetSheetPath(string sheetName)
        {
            if (string.IsNullOrEmpty(sheetName))
                return "";

            Debug.WriteLine($"[GetSheetPath] 시작: 현재 워크북={_currentWorkbookPath}, 시트={sheetName}");

            try
            {
                // 현재 워크북 경로가 없으면 빈 문자열 반환
                if (string.IsNullOrEmpty(_currentWorkbookPath))
                    return "";
                
                string normalizedPath = NormalizeWorkbookPath(_currentWorkbookPath);
                string fileName = Path.GetFileName(_currentWorkbookPath);
                
                List<string> pathsToTry = new List<string>();
                
                // 검색할 경로 우선순위 설정
                pathsToTry.Add(normalizedPath);
                if (!string.IsNullOrEmpty(fileName) && fileName != normalizedPath)
                {
                    pathsToTry.Add(fileName);
                }
                
                // 다양한 시트 이름 형식 시도 (원본과 "!" 접두사 추가된 형식)
                List<string> sheetNamesToTry = new List<string>();
                sheetNamesToTry.Add(sheetName);
                if (!sheetName.StartsWith("!"))
                {
                    sheetNamesToTry.Add("!" + sheetName);
                }
                else if (sheetName.StartsWith("!"))
                {
                    sheetNamesToTry.Add(sheetName.Substring(1));
                }
                
                // 모든 경로와 시트 이름 조합 시도
                foreach (string pathToTry in pathsToTry)
                {
                    foreach (string sheetNameToTry in sheetNamesToTry)
                    {
                        if (TryGetSheetPathInternal(pathToTry, sheetNameToTry, out string pathResult) && 
                            !string.IsNullOrEmpty(pathResult))
                        {
                            Debug.WriteLine($"[GetSheetPath] 워크북 '{pathToTry}'에서 시트 '{sheetNameToTry}'의 경로 찾음: {pathResult}");
                            return pathResult;
                        }
                    }
                }
                
                Debug.WriteLine($"[GetSheetPath] 시트 경로 조회 실패: workbook={_currentWorkbookPath}, sheet={sheetName}");
                return "";
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetSheetPath] 예외 발생: {ex.Message}");
                return "";
            }
        }
        
        /// <summary>
        /// 지정한 워크북과 시트에 대한 경로를 내부적으로 찾는 메서드
        /// </summary>
        private bool TryGetSheetPathInternal(string workbookKey, string sheetName, out string path)
        {
            path = "";
            
            if (string.IsNullOrEmpty(workbookKey) || string.IsNullOrEmpty(sheetName))
                return false;
                
            try
            {
                if (LazyLoadSheetPaths().ContainsKey(workbookKey))
                {
                    Debug.WriteLine($"[GetSheetPath] 워크북 '{workbookKey}'가 딕셔너리에 있음");
                    Debug.WriteLine($"[GetSheetPath] 워크북 '{workbookKey}'에 등록된 시트 수: {LazyLoadSheetPaths()[workbookKey].Count}");
                    
                    var sheetDict = LazyLoadSheetPaths()[workbookKey];
                    if (sheetDict.ContainsKey(sheetName))
                    {
                        path = sheetDict[sheetName].SavePath;
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[TryGetSheetPathInternal] 예외: {ex.Message}");
                return false;
            }
        }

        // 특정 워크북의 시트 경로 활성화 상태 가져오기
        public bool GetSheetEnabled(string sheetName)
        {
            Debug.WriteLine($"[GetSheetEnabled] 호출: 시트='{sheetName}', 현재 워크북='{_currentWorkbookPath}'");
            
            if (string.IsNullOrEmpty(_currentWorkbookPath))
            {
                Debug.WriteLine($"[GetSheetEnabled] 현재 워크북이 설정되지 않았습니다.");
                return false;
            }
            
            // 0. 파일명과 정규화된 경로
            string normalizedPath = NormalizeWorkbookPath(_currentWorkbookPath);
            string fileName = Path.GetFileName(_currentWorkbookPath);
            
            Debug.WriteLine($"[GetSheetEnabled] 워크북 정보: 원본 경로='{_currentWorkbookPath}', 정규화된 경로='{normalizedPath}', 파일명='{fileName}'");
            
            // 원래 시트 이름과 ! 접두사를 추가/제거한 대체 시트 이름 준비
            string originalSheetName = sheetName;
            string alternateSheetName;
            
            if (sheetName.StartsWith("!"))
            {
                alternateSheetName = sheetName.Substring(1);
                Debug.WriteLine($"[GetSheetEnabled] 원본 시트 이름 '{sheetName}'에서 ! 접두사 제거한 대체 이름: '{alternateSheetName}'");
            }
            else
            {
                alternateSheetName = "!" + sheetName;
                Debug.WriteLine($"[GetSheetEnabled] 원본 시트 이름 '{sheetName}'에 ! 접두사 추가한 대체 이름: '{alternateSheetName}'");
            }
            
            // 원본 시트 이름으로 검색
            bool originalEnabled = CheckSheetEnabled(originalSheetName);
            if (originalEnabled)
            {
                Debug.WriteLine($"[GetSheetEnabled] 원본 시트 이름 '{originalSheetName}'로 활성화 상태 확인: true");
                return true;
            }
            
            // 대체 시트 이름으로 검색
            bool alternateEnabled = CheckSheetEnabled(alternateSheetName);
            if (alternateEnabled)
            {
                Debug.WriteLine($"[GetSheetEnabled] 대체 시트 이름 '{alternateSheetName}'로 활성화 상태 확인: true");
                return true;
            }
            
            Debug.WriteLine($"[GetSheetEnabled] 시트 '{sheetName}'의 활성화 상태를 찾을 수 없음. 기본값 'false' 반환");
            return false;
        }
        
        // 내부적으로 시트의 활성화 상태를 확인하는 헬퍼 메서드
        private bool CheckSheetEnabled(string sheetName)
        {
            // 1. 전체 워크북 경로로 확인
            if (LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath) && 
                LazyLoadSheetPaths()[_currentWorkbookPath].ContainsKey(sheetName))
            {
                bool enabled = LazyLoadSheetPaths()[_currentWorkbookPath][sheetName].Enabled;
                Debug.WriteLine($"[CheckSheetEnabled] 전체 경로 '{_currentWorkbookPath}'로 시트 '{sheetName}'의 활성화 상태 확인: {enabled}");
                return enabled;
            }
            
            // 1-1. 정규화된 경로로 확인 (원본과 다른 경우)
            string normalizedPath = NormalizeWorkbookPath(_currentWorkbookPath);
            if (_currentWorkbookPath != normalizedPath && 
                LazyLoadSheetPaths().ContainsKey(normalizedPath) && 
                LazyLoadSheetPaths()[normalizedPath].ContainsKey(sheetName))
            {
                bool enabled = LazyLoadSheetPaths()[normalizedPath][sheetName].Enabled;
                Debug.WriteLine($"[CheckSheetEnabled] 정규화된 경로 '{normalizedPath}'로 시트 '{sheetName}'의 활성화 상태 확인: {enabled}");
                return enabled;
            }
            
            // 2. 파일명으로 확인
            string fileName = Path.GetFileName(_currentWorkbookPath);
            if (!string.IsNullOrEmpty(fileName) && LazyLoadSheetPaths().ContainsKey(fileName) && 
                LazyLoadSheetPaths()[fileName].ContainsKey(sheetName))
            {
                bool enabled = LazyLoadSheetPaths()[fileName][sheetName].Enabled;
                Debug.WriteLine($"[CheckSheetEnabled] 파일명 '{fileName}'으로 시트 '{sheetName}'의 활성화 상태 확인: {enabled}");
                return enabled;
            }

            // 3. XML에 저장된 모든 워크북 경로 확인 (OneDrive URL 등 다양한 형식의 경로가 있을 수 있음)
            Debug.WriteLine($"[CheckSheetEnabled] 시트 '{sheetName}'의 활성화 상태를 전체/파일명으로 찾지 못함. 모든 경로 탐색 시작");
            
            foreach (var workbookEntry in LazyLoadSheetPaths())
            {
                string workbookKey = workbookEntry.Key;
                string entryFileName = Path.GetFileName(workbookKey);
                
                // 같은 파일명인지 확인
                if (workbookKey != _currentWorkbookPath && workbookKey != normalizedPath && workbookKey != fileName && 
                    !string.IsNullOrEmpty(fileName) && entryFileName == fileName && 
                    workbookEntry.Value.ContainsKey(sheetName))
                {
                    bool enabled = workbookEntry.Value[sheetName].Enabled;
                    Debug.WriteLine($"[CheckSheetEnabled] 다른 형식의 경로 '{workbookKey}'에서 시트 '{sheetName}'의 활성화 상태 발견: {enabled}");
                    return enabled;
                }
                
                // 시트 이름이 있는지 확인
                if (workbookEntry.Value.ContainsKey(sheetName))
                {
                    bool enabled = workbookEntry.Value[sheetName].Enabled;
                    Debug.WriteLine($"[CheckSheetEnabled] 다른 워크북 '{workbookKey}'에서 시트 '{sheetName}'의 활성화 상태 발견: {enabled}");
                    return enabled;
                }
            }
            
            return false;
        }

        /// <summary>
        /// 워크북의 특정 시트의 YAML 선택적 필드 처리 여부 가져오기
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>YAML 선택적 필드 처리 여부</returns>
        public bool GetYamlEmptyFieldsOption(string sheetName)
        {
            // XML에서 YAML 설정을 관리하지 않으므로 항상 false 반환
            // 이 설정은 ExcelConfigManager에서만 관리됨
            return false;
        }
        
        // 특정 시트의 YAML 선택적 필드 처리 설정
        public void SetYamlEmptyFieldsOption(string sheetName, bool yamlEmptyFields)
        {
            // XML에 저장하지 않고 ExcelConfigManager에서만 관리하므로 아무 작업도 수행하지 않음
            Debug.WriteLine($"[SheetPathManager] SetYamlEmptyFieldsOption: XML에 저장하지 않습니다. ExcelConfigManager를 사용하세요.");
        }

        // 특정 워크북의 시트에 대한 YAML 선택적 필드 처리 설정 (내부용)
        private void SetYamlEmptyFieldsOptionInternal(string workbookName, string sheetName, bool yamlEmptyFields)
        {
            // XML에 저장하지 않고 ExcelConfigManager에서만 관리하므로 아무 작업도 수행하지 않음
        }

        // 특정 워크북의 시트에 대한 활성화 상태 설정
        public void SetSheetEnabled(string workbookName, string sheetName, bool enabled)
        {
            Debug.WriteLine($"[SheetPathManager] SetSheetEnabled 오버로드 호출: 워크북={workbookName}, 시트={sheetName}, Enabled={enabled}");
            SetSheetEnabledInternal(workbookName, sheetName, enabled);
            
            // 변경 후 즉시 저장하여 설정이 유지되도록 함
            Debug.WriteLine($"[SheetPathManager] SetSheetEnabled 후 즉시 SaveSettings 호출");
            SaveSettings();
        }

        // 현재 워크북의 시트에 대한 활성화 상태 설정 (오버로드)
        public void SetSheetEnabled(string sheetName, bool enabled)
        {
            // 현재 워크북 이름 가져오기
            string currentWorkbook = _currentWorkbookPath;
            if (string.IsNullOrEmpty(currentWorkbook))
            {
                Debug.WriteLine($"[SheetPathManager] 경고: 현재 워크북 경로가 없습니다. 시트={sheetName}, Enabled={enabled}");
                return;
            }
            
            Debug.WriteLine($"[SheetPathManager] SetSheetEnabled(2-param) 호출: 현재 워크북={currentWorkbook}, 시트={sheetName}, Enabled={enabled}");
            
            // 3-parameter 오버로드 호출
            SetSheetEnabled(currentWorkbook, sheetName, enabled);
        }

        // 특정 워크북의 시트에 대한 활성화 상태 설정 (내부용)
        private void SetSheetEnabledInternal(string workbookName, string sheetName, bool enabled)
        {
            if (string.IsNullOrEmpty(workbookName) || string.IsNullOrEmpty(sheetName))
            {
                Debug.WriteLine($"[SetSheetEnabledInternal] 잘못된 인수: workbookName='{workbookName}', sheetName='{sheetName}'");
                return;
            }
            
            // 워크북이 사전에 없으면 추가
            if (!LazyLoadSheetPaths().ContainsKey(workbookName))
            {
                LazyLoadSheetPaths()[workbookName] = new Dictionary<string, SheetPathInfo>();
                Debug.WriteLine($"[SetSheetEnabledInternal] 워크북 '{workbookName}'에 대한 새 사전 생성");
            }
            
            // 시트가 사전에 없으면 추가
            if (!LazyLoadSheetPaths()[workbookName].ContainsKey(sheetName))
            {
                LazyLoadSheetPaths()[workbookName][sheetName] = new SheetPathInfo
                {
                    SavePath = "",
                    Enabled = enabled,
                    YamlEmptyFields = false
                };
                Debug.WriteLine($"[SetSheetEnabledInternal] 시트 '{sheetName}'이 사전에 없어 새로 생성: Enabled={enabled}");
            }
            else
            {
                // 기존 시트 정보가 있으면 활성화 상태만 업데이트
                LazyLoadSheetPaths()[workbookName][sheetName].Enabled = enabled;
                Debug.WriteLine($"[SetSheetEnabledInternal] 시트 '{sheetName}'의 활성화 상태 업데이트: {enabled}");
            }
            
            // 활성화 상태 저장 후 사전 내용 출력
            Debug.WriteLine($"[SetSheetEnabledInternal] 워크북 '{workbookName}'의 시트 '{sheetName}' 저장 후 상태:");
            if (LazyLoadSheetPaths().ContainsKey(workbookName) && LazyLoadSheetPaths()[workbookName].ContainsKey(sheetName))
            {
                var info = LazyLoadSheetPaths()[workbookName][sheetName];
                Debug.WriteLine($"[SetSheetEnabledInternal] 경로='{info.SavePath}', 활성화={info.Enabled}, YAML빈필드={info.YamlEmptyFields}");
            }
            else
            {
                Debug.WriteLine($"[SetSheetEnabledInternal] 사전에서 항목을 찾을 수 없음");
            }
        }

        /// <summary>
        /// 워크북의 특정 시트의 후처리용 키 경로 인수 값을 가져옵니다.
        /// XML에서는 이 설정을 관리하지 않으므로 항상 빈 문자열을 반환합니다.
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>빈 문자열 (이 설정은 ExcelConfigManager에서만 관리됨)</returns>
        public string GetMergeKeyPaths(string sheetName)
        {
            // XML에서 병합 키 경로 설정을 관리하지 않으므로 항상 빈 문자열 반환
            // 이 설정은 ExcelConfigManager에서만 관리됨
            Debug.WriteLine($"[SheetPathManager] GetMergeKeyPaths: XML에서 관리하지 않는 설정입니다. ExcelConfigManager를 사용하세요.");
            return "";
        }
        
        /// <summary>
        /// 특정 워크북과 시트의 후처리용 키 경로 인수 값을 가져옵니다.
        /// XML에서는 이 설정을 관리하지 않으므로 항상 빈 문자열을 반환합니다.
        /// </summary>
        /// <param name="workbookName">워크북 이름</param>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>빈 문자열 (이 설정은 ExcelConfigManager에서만 관리됨)</returns>
        public string GetMergeKeyPaths(string workbookName, string sheetName)
        {
            // XML에서 병합 키 경로 설정을 관리하지 않으므로 항상 빈 문자열 반환
            Debug.WriteLine($"[SheetPathManager] GetMergeKeyPaths: XML에서 관리하지 않는 설정입니다. ExcelConfigManager를 사용하세요.");
            return "";
        }

        /// <summary>
        /// 워크북의 특정 시트의 YAML Flow Style 설정을 가져옵니다.
        /// XML에서는 이 설정을 관리하지 않으므로 항상 빈 문자열을 반환합니다.
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>빈 문자열 (이 설정은 ExcelConfigManager에서만 관리됨)</returns>
        public string GetFlowStyleConfig(string sheetName)
        {
            // XML에서 Flow Style 설정을 관리하지 않으므로 항상 빈 문자열 반환
            // 이 설정은 ExcelConfigManager에서만 관리됨
            Debug.WriteLine($"[SheetPathManager] GetFlowStyleConfig: XML에서 관리하지 않는 설정입니다. ExcelConfigManager를 사용하세요.");
            return "";
        }
        
        /// <summary>
        /// 특정 워크북과 시트의 YAML Flow Style 설정을 가져옵니다.
        /// XML에서는 이 설정을 관리하지 않으므로 항상 빈 문자열을 반환합니다.
        /// </summary>
        /// <param name="workbookName">워크북 이름</param>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>빈 문자열 (이 설정은 ExcelConfigManager에서만 관리됨)</returns>
        public string GetFlowStyleConfig(string workbookName, string sheetName)
        {
            // XML에서 Flow Style 설정을 관리하지 않으므로 항상 빈 문자열 반환
            Debug.WriteLine($"[SheetPathManager] GetFlowStyleConfig: XML에서 관리하지 않는 설정입니다. ExcelConfigManager를 사용하세요.");
            return "";
        }

        /// <summary>
        /// 시트의 후처리용 키 경로 인수 값을 설정합니다.
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="mergeKeyPaths">설정할 후처리 키 경로 인수 값</param>
        public void SetMergeKeyPaths(string sheetName, string mergeKeyPaths)
        {
            // XML에 저장하지 않고 ExcelConfigManager에서만 관리하므로 아무 작업도 수행하지 않음
            Debug.WriteLine($"[SheetPathManager] SetMergeKeyPaths: XML에 저장하지 않습니다. ExcelConfigManager를 사용하세요.");
        }

        /// <summary>
        /// 특정 워크북의 시트에 대한 후처리용 키 경로 인수 값을 설정합니다.
        /// </summary>
        /// <param name="workbookName">워크북 이름 또는 경로</param>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="mergeKeyPaths">설정할 후처리 키 경로 인수 값</param>
        public void SetMergeKeyPaths(string workbookName, string sheetName, string mergeKeyPaths)
        {
            // XML에 저장하지 않고 ExcelConfigManager에서만 관리하므로 아무 작업도 수행하지 않음
            Debug.WriteLine($"[SheetPathManager] SetMergeKeyPaths: XML에 저장하지 않습니다. ExcelConfigManager를 사용하세요.");
        }

        /// <summary>
        /// 지정된 시트의 Flow Style 설정을 저장합니다.
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="flowStyleConfig">Flow Style 설정 문자열</param>
        public void SetFlowStyleConfig(string sheetName, string flowStyleConfig)
        {
            // XML에 저장하지 않고 ExcelConfigManager에서만 관리하므로 아무 작업도 수행하지 않음
            Debug.WriteLine($"[SheetPathManager] SetFlowStyleConfig: XML에 저장하지 않습니다. ExcelConfigManager를 사용하세요.");
        }
        
        /// <summary>
        /// 특정 워크북의 특정 시트의 Flow Style 설정을 저장합니다.
        /// </summary>
        /// <param name="workbookName">워크북 이름</param>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="flowStyleConfig">Flow Style 설정</param>
        public void SetFlowStyleConfig(string workbookName, string sheetName, string flowStyleConfig)
        {
            // XML에 저장하지 않고 ExcelConfigManager에서만 관리하므로 아무 작업도 수행하지 않음
            Debug.WriteLine($"[SheetPathManager] SetFlowStyleConfig: XML에 저장하지 않습니다. ExcelConfigManager를 사용하세요.");
        }

        /// <summary>
        /// 설정 파일 경로를 반환합니다.
        /// </summary>
        /// <returns>설정 파일의 전체 경로</returns>
        public static string GetConfigFilePath()
        {
            return ConfigFilePath;
        }

        // 특정 워크북의 시트 경로 사전 반환 (수정된 메서드)
        public Dictionary<string, string> GetSheetPaths(string workbookPath)
        {
            // 1. 먼저 전체 경로로 시도
            if (!string.IsNullOrEmpty(workbookPath) &&
                LazyLoadSheetPaths().ContainsKey(workbookPath))
            {
                Debug.WriteLine($"[GetSheetPaths] 전체 경로 '{workbookPath}'에서 시트 경로 발견: {LazyLoadSheetPaths()[workbookPath].Count}개");
                return LazyLoadSheetPaths()[workbookPath].ToDictionary(kvp => kvp.Key, kvp => kvp.Value.SavePath);
            }

            // 2. 파일 이름만으로도 시도
            string fileName = Path.GetFileName(workbookPath);
            if (!string.IsNullOrEmpty(fileName) &&
                LazyLoadSheetPaths().ContainsKey(fileName))
            {
                Debug.WriteLine($"[GetSheetPaths] 파일명 '{fileName}'에서 시트 경로 발견: {LazyLoadSheetPaths()[fileName].Count}개");
                return LazyLoadSheetPaths()[fileName].ToDictionary(kvp => kvp.Key, kvp => kvp.Value.SavePath);
            }

            Debug.WriteLine($"[GetSheetPaths] '{workbookPath}' 또는 '{fileName}'에 대한 시트 경로를 찾을 수 없습니다.");
            return new Dictionary<string, string>();
        }

        // 현재 워크북의 모든 시트 경로 가져오기 (활성화 여부 상관없이)
        public Dictionary<string, string> GetAllSheetPaths()
        {
            // 현재 워크북이 없거나 해당 워크북의 설정이 없으면 빈 딕셔너리 반환
            if (string.IsNullOrEmpty(_currentWorkbookPath) || !LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath))
            {
                Debug.WriteLine($"[GetAllSheetPaths] 파일명 '{_currentWorkbookPath}'에서 시트 경로 발견: 0개");
                
                // 전체 경로로도 시도
                if (_currentWorkbookPath != null)
                {
                    string normalizedPath = NormalizeWorkbookPath(_currentWorkbookPath);
                    if (normalizedPath != _currentWorkbookPath && LazyLoadSheetPaths().ContainsKey(normalizedPath))
                    {
                        Debug.WriteLine($"[GetAllSheetPaths] 정규화된 경로 '{normalizedPath}'에서 시트 경로 발견: {LazyLoadSheetPaths()[normalizedPath].Count}개");
                        var result = new Dictionary<string, string>();
                        foreach (var entry in LazyLoadSheetPaths()[normalizedPath])
                        {
                            result[entry.Key] = entry.Value.SavePath;
                        }
                        return result;
                    }
                }
                
                // 또 다른 키 검색
            foreach (var key in LazyLoadSheetPaths().Keys)
            {
                    if (Path.GetFileName(key) == Path.GetFileName(_currentWorkbookPath))
                {
                    Debug.WriteLine($"[GetAllSheetPaths] 다른 키 '{key}'에서 시트 경로 발견: {LazyLoadSheetPaths()[key].Count}개");
                        var result = new Dictionary<string, string>();
                    foreach (var entry in LazyLoadSheetPaths()[key])
                    {
                            result[entry.Key] = entry.Value.SavePath;
                        }
                        return result;
                    }
                }
                
                Debug.WriteLine($"[GetAllSheetPaths] 워크북 '{_currentWorkbookPath}'에 대한 시트 경로를 찾을 수 없습니다.");
                return new Dictionary<string, string>();
            }

            // 모든 경로 반환
            var paths = new Dictionary<string, string>();
            foreach (var sheetInfo in LazyLoadSheetPaths()[_currentWorkbookPath])
            {
                paths[sheetInfo.Key] = sheetInfo.Value.SavePath;
            }
            return paths;
        }

        // 시트 경로가 이미 설정되어 있는지 확인
        public bool HasSheetPath(string sheetName)
        {
            if (string.IsNullOrEmpty(_currentWorkbookPath))
            {
                Debug.WriteLine($"[HasSheetPath] 현재 워크북이 설정되지 않았습니다.");
                return false;
            }

            // 1. 파일명으로 먼저 확인
            if (LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath) &&
                LazyLoadSheetPaths()[_currentWorkbookPath].ContainsKey(sheetName))
            {
                Debug.WriteLine($"[HasSheetPath] 파일명 '{_currentWorkbookPath}'에서 시트 '{sheetName}' 경로 발견");
                return true;
            }

            // 2. 워크북 이름과 일치하는 다른 키 검색
            foreach (var key in LazyLoadSheetPaths().Keys)
            {
                if (Path.GetFileName(key) == _currentWorkbookPath &&
                    LazyLoadSheetPaths()[key].ContainsKey(sheetName))
                {
                    Debug.WriteLine($"[HasSheetPath] 경로 '{key}'에서 시트 '{sheetName}' 경로 발견");
                    return true;
                }
            }

            Debug.WriteLine($"[HasSheetPath] 시트 '{sheetName}'의 경로를 찾을 수 없습니다.");
            return false;
        }

        // 특정 시트의 경로 정보 삭제
        public void RemoveSheetPath(string workbookName, string sheetName)
        {
            Debug.WriteLine($"[RemoveSheetPath] 호출 시작: 워크북='{workbookName}', 시트='{sheetName}'");
            if (string.IsNullOrEmpty(workbookName) || string.IsNullOrEmpty(sheetName))
            {
                Debug.WriteLine($"[RemoveSheetPath] 오류: 워크북이나 시트 이름이 비어 있음");
                return;
            }

            if (!LazyLoadSheetPaths().ContainsKey(workbookName))
            {
                Debug.WriteLine($"[RemoveSheetPath] 오류: 워크북 '{workbookName}'을 찾을 수 없음");
                return;
            }

            if (!LazyLoadSheetPaths()[workbookName].ContainsKey(sheetName))
            {
                Debug.WriteLine($"[RemoveSheetPath] 오류: 워크북 '{workbookName}'에서 시트 '{sheetName}'을 찾을 수 없음");
                return;
            }

            Debug.WriteLine($"[RemoveSheetPath] 워크북 '{workbookName}'에서 시트 '{sheetName}' 제거");
            LazyLoadSheetPaths()[workbookName].Remove(sheetName);
            Debug.WriteLine($"[RemoveSheetPath] 제거 완료");
        }

        // 기존 메서드도 남겨두기 (호환성을 위해)
        public void RemoveSheetPath(string sheetName)
        {
            Debug.WriteLine($"[RemoveSheetPath] 현재 워크북에서 시트 '{sheetName}' 제거 시도");
            if (string.IsNullOrEmpty(_currentWorkbookPath))
            {
                Debug.WriteLine($"[RemoveSheetPath] 오류: 현재 워크북이 설정되지 않음");
                return;
            }
            
            if (!LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath))
            {
                Debug.WriteLine($"[RemoveSheetPath] 오류: 현재 워크북 '{_currentWorkbookPath}'을 찾을 수 없음");
                return;
            }
            
            if (!LazyLoadSheetPaths()[_currentWorkbookPath].ContainsKey(sheetName))
            {
                Debug.WriteLine($"[RemoveSheetPath] 오류: 현재 워크북 '{_currentWorkbookPath}'에서 시트 '{sheetName}'을 찾을 수 없음");
                return;
            }

            Debug.WriteLine($"[RemoveSheetPath] 현재 워크북 '{_currentWorkbookPath}'에서 시트 '{sheetName}' 제거");
            LazyLoadSheetPaths()[_currentWorkbookPath].Remove(sheetName);
            Debug.WriteLine($"[RemoveSheetPath] 제거 완료 및 설정 저장");
            SaveSheetPaths();
        }

        /// <summary>
        /// 모든 시트 경로 설정을 XML 파일에 저장합니다.
        /// </summary>
        public void SaveSheetPaths()
        {
            try
            {
                Debug.WriteLine($"[SaveSheetPaths] 시트 경로 설정 저장 시작");
                
                // 설정 디렉토리가 없으면 생성
                string settingsDir = GetSettingsDirectory();
                if (!Directory.Exists(settingsDir))
                {
                    Debug.WriteLine($"[SaveSheetPaths] 설정 디렉토리 생성: {settingsDir}");
                    Directory.CreateDirectory(settingsDir);
                }

                // 중복 데이터 방지를 위한 사전 - 시트이름을 키로 사용
                Dictionary<string, SheetPathData> uniqueEntries = new Dictionary<string, SheetPathData>();
                
                // 데이터 변환 및 디버깅 출력
                foreach (var workbook in LazyLoadSheetPaths())
                {
                    string workbookKey = workbook.Key;
                    bool isFullPath = workbookKey.Contains("/") || workbookKey.Contains("\\") || workbookKey.Contains(":");
                    string fileName = isFullPath ? Path.GetFileName(workbookKey) : workbookKey;
                    
                    Debug.WriteLine($"[SaveSheetPaths] 워크북 '{workbookKey}' 처리 중 ({workbook.Value.Count}개 시트)");
                    Debug.WriteLine($"[SaveSheetPaths] 워크북 타입: {(isFullPath ? "전체 경로" : "파일명")}, 파일명: {fileName}");
                    
                    foreach (var sheet in workbook.Value)
                    {
                        var sheetName = sheet.Key;
                        var sheetInfo = sheet.Value;
                        
                        // 이미 이 시트가 등록되어 있고, 현재 항목이 파일명만 사용한 경우 우선 등록
                        string uniqueKey = sheetName; // 시트 이름을 키로 사용
                        
                        if (uniqueEntries.ContainsKey(uniqueKey))
                        {
                            // 이미 등록된 항목이 파일명이고 현재 항목이 전체 경로인 경우 건너뜀
                            if (!isFullPath && uniqueEntries[uniqueKey].WorkbookPath.Contains("/"))
                            {
                                Debug.WriteLine($"[SaveSheetPaths] 시트 '{sheetName}'는 이미 전체 경로 항목이 등록되어 있어 파일명 항목은 무시합니다.");
                                continue;
                            }
                            
                            // 이미 등록된 항목이 전체 경로이고 현재 항목이 파일명인 경우 덮어씀
                            if (isFullPath)
                            {
                                Debug.WriteLine($"[SaveSheetPaths] 시트 '{sheetName}'의 기존 항목을 전체 경로 항목으로 교체합니다.");
                            }
                        }
                        
                        // 새 항목 생성
                        SheetPathData pathData = new SheetPathData
                        {
                            // 항상 파일명만 저장하여 중복 방지
                            WorkbookPath = fileName,
                            SheetName = sheetName,
                            SavePath = sheetInfo.SavePath,
                            Enabled = sheetInfo.Enabled,
                            YamlEmptyFields = sheetInfo.YamlEmptyFields,
                            MergeKeyPaths = "", // XML에는 저장하지 않음
                            FlowStyleConfig = "" // XML에는 저장하지 않음
                        };
                        
                        uniqueEntries[uniqueKey] = pathData;
                        Debug.WriteLine($"[SaveSheetPaths] 저장할 항목: 워크북='{fileName}', 시트='{sheetName}', 경로='{sheetInfo.SavePath}', 활성화={sheetInfo.Enabled}");
                    }
                }

                // 중복 제거된 항목을 리스트로 변환
                List<SheetPathData> pathsToSave = uniqueEntries.Values.ToList();
                Debug.WriteLine($"[SaveSheetPaths] 총 {pathsToSave.Count}개 항목 저장 준비 완료 (중복 제거 후)");
                
                // XML로 직렬화하여 저장
                XmlSerializer serializer = new XmlSerializer(typeof(List<SheetPathData>));
                using (StreamWriter writer = new StreamWriter(GetSettingsFilePath()))
                {
                    serializer.Serialize(writer, pathsToSave);
                }

                Debug.WriteLine($"[SaveSheetPaths] 설정 파일에 저장 완료: {GetSettingsFilePath()}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SaveSheetPaths] 시트 경로 설정 저장 중 오류 발생: {ex.Message}\n{ex.StackTrace}");
            }
        }

        // 외부에서 접근 가능한 설정 저장 메서드
        public void SaveSettings()
        {
            Debug.WriteLine($"[SaveSettings] 설정 저장 시작");
            SaveSheetPaths();
            Debug.WriteLine($"[SaveSettings] 설정 저장 완료");
        }

        // 설정 파일 로드
        private void LoadSheetPaths()
        {
            try
            {
                _sheetPaths = new Dictionary<string, Dictionary<string, SheetPathInfo>>();
                string settingsPath = GetSettingsFilePath();
                
                if (!File.Exists(settingsPath))
                {
                    Debug.WriteLine($"[LoadSheetPaths] 설정 파일이 존재하지 않습니다: {settingsPath}");
                    return;
                }
                
                Debug.WriteLine($"[LoadSheetPaths] 설정 파일 로드 시작: {settingsPath}");
                
                List<SheetPathData> loadedPaths = null;
                XmlSerializer serializer = new XmlSerializer(typeof(List<SheetPathData>));
                
                using (StreamReader reader = new StreamReader(settingsPath))
                {
                    loadedPaths = (List<SheetPathData>)serializer.Deserialize(reader);
                }
                
                if (loadedPaths == null)
                {
                    Debug.WriteLine($"[LoadSheetPaths] 설정 파일에서 로드된 데이터가 없습니다.");
                    return;
                }

                Debug.WriteLine($"[LoadSheetPaths] 로드된 항목 수: {loadedPaths.Count}");
                
                // 시트 이름별 중복 제거를 위한 사전
                Dictionary<string, SheetPathData> uniqueEntries = new Dictionary<string, SheetPathData>();
                
                // 첫 번째 패스: 중복 엔트리 처리
                foreach (SheetPathData pathData in loadedPaths)
                {
                    string sheetName = pathData.SheetName;
                    
                    // 워크북 경로가 파일명인지 전체 경로인지 확인
                    bool isFullPath = pathData.WorkbookPath.Contains("/") || 
                                      pathData.WorkbookPath.Contains("\\") || 
                                      pathData.WorkbookPath.Contains(":");
                                      
                    // 항상 파일명 부분만 추출
                    string fileName = isFullPath ? 
                        Path.GetFileName(pathData.WorkbookPath) : 
                        pathData.WorkbookPath;
                        
                    // 중복 항목이 있는 경우 처리
                    if (uniqueEntries.ContainsKey(sheetName))
                    {
                        // 이미 등록된 항목의 워크북 경로가 파일명인지 확인
                        bool existingIsFullPath = uniqueEntries[sheetName].WorkbookPath.Contains("/") || 
                                                 uniqueEntries[sheetName].WorkbookPath.Contains("\\") || 
                                                 uniqueEntries[sheetName].WorkbookPath.Contains(":");
                                                 
                        // 기존 파일명 항목이 있고 현재 전체 경로인 경우, 전체 경로 우선
                        if (isFullPath && !existingIsFullPath)
                        {
                            Debug.WriteLine($"[LoadSheetPaths] 시트 '{sheetName}'에 대해 전체 경로 항목으로 교체합니다.");
                            uniqueEntries[sheetName] = pathData;
                        }
                        else if (!isFullPath && existingIsFullPath)
                        {
                            // 현재 항목이 파일명이고 기존 항목이 전체 경로인 경우, 기존 항목 유지
                            Debug.WriteLine($"[LoadSheetPaths] 시트 '{sheetName}'에 대해 기존 전체 경로 항목을 유지합니다.");
                            continue;
                        }
                        else
                        {
                            // 둘 다 같은 타입이면 나중에 로드된 항목(현재 항목) 사용
                            Debug.WriteLine($"[LoadSheetPaths] 시트 '{sheetName}'에 대해 중복 항목을 교체합니다.");
                            uniqueEntries[sheetName] = pathData;
                        }
                    }
                    else
                    {
                        // 새 항목 추가
                        uniqueEntries[sheetName] = pathData;
                    }
                }
                
                // 두 번째 패스: 정리된 데이터를 _sheetPaths에 로드
                foreach (var entry in uniqueEntries.Values)
                {
                    string workbookPath = entry.WorkbookPath;
                    string sheetName = entry.SheetName;
                    
                    // 워크북 경로 정규화
                    bool isFullPath = workbookPath.Contains("/") || 
                                     workbookPath.Contains("\\") || 
                                     workbookPath.Contains(":");
                                     
                    // 파일명만 저장하여 중복 방지
                    string normalizedWorkbookPath = isFullPath ? 
                        Path.GetFileName(workbookPath) : 
                        workbookPath;
                    
                    if (!_sheetPaths.ContainsKey(normalizedWorkbookPath))
                    {
                        _sheetPaths[normalizedWorkbookPath] = new Dictionary<string, SheetPathInfo>();
                    }
                    
                    _sheetPaths[normalizedWorkbookPath][sheetName] = new SheetPathInfo
                    {
                        SavePath = entry.SavePath,
                        Enabled = entry.Enabled,
                        YamlEmptyFields = entry.YamlEmptyFields
                    };
                    
                    Debug.WriteLine($"[LoadSheetPaths] 로드된 항목: 워크북='{normalizedWorkbookPath}', 시트='{sheetName}', 경로='{entry.SavePath}', 활성화={entry.Enabled}");
                }
                
                // 로드 후 데이터를 저장하여 중복 제거
                Debug.WriteLine($"[LoadSheetPaths] 중복 제거 후 다시 저장합니다.");
                SaveSheetPaths();
                
                Debug.WriteLine($"[LoadSheetPaths] 설정 로드 완료. 총 {_sheetPaths.Count}개 워크북, {uniqueEntries.Count}개 시트 로드됨.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LoadSheetPaths] 설정 로드 중 오류 발생: {ex.Message}\n{ex.StackTrace}");
                _sheetPaths = new Dictionary<string, Dictionary<string, SheetPathInfo>>();
            }
        }

        // 지연 초기화 패턴을 적용한 LoadSheetPaths 호출
        private Dictionary<string, Dictionary<string, SheetPathInfo>> LazyLoadSheetPaths()
        {
            if (_sheetPaths == null)
            {
                Debug.WriteLine("[LazyLoadSheetPaths] _sheetPaths가 null이므로 새로 초기화합니다.");
                _sheetPaths = new Dictionary<string, Dictionary<string, SheetPathInfo>>();
                LoadSheetPaths();
                Debug.WriteLine($"[LazyLoadSheetPaths] 초기화 완료: 워크북 수={_sheetPaths.Count}");
            }
            return _sheetPaths;
        }

        // LazyLoadSheetPaths를 호출하는 메서드
        public void Initialize()
        {
            try
            {
                Debug.WriteLine($"[Initialize] SheetPathManager 초기화 시작");
                
                // 설정 디렉토리 생성
                string settingsDir = GetSettingsDirectory();
                if (!Directory.Exists(settingsDir))
                {
                    Debug.WriteLine($"[Initialize] 설정 디렉토리 생성: {settingsDir}");
                    Directory.CreateDirectory(settingsDir);
                }
                
                // 설정 파일이 없으면 생성
                string settingsPath = GetSettingsFilePath();
                if (!File.Exists(settingsPath))
                {
                    Debug.WriteLine($"[Initialize] 설정 파일이 없어 새로 생성합니다: {settingsPath}");
                    SaveSheetPaths();
                }
                else
                {
                    // 설정 파일이 있으면 로드하여 중복 항목 정리
                    Debug.WriteLine($"[Initialize] 기존 설정 파일이 있어 중복 항목을 정리합니다: {settingsPath}");
                    LoadSheetPaths(); // 로드 후 자동으로 중복 제거 및 저장됨
                }
                
                Debug.WriteLine($"[Initialize] SheetPathManager 초기화 완료");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Initialize] 초기화 중 오류 발생: {ex.Message}");
            }
        }

        // 모든 워크북 경로 목록 가져오기
        public List<string> GetAllWorkbookPaths()
        {
            Debug.WriteLine("[GetAllWorkbookPaths] 시작");
            LazyLoadSheetPaths();

            if (_sheetPaths == null || _sheetPaths.Count == 0)
            {
                Debug.WriteLine("[GetAllWorkbookPaths] 저장된 워크북이 없습니다");
                return new List<string>();
            }

            var result = new List<string>(_sheetPaths.Keys);
            Debug.WriteLine($"[GetAllWorkbookPaths] 총 {result.Count}개의 워크북 발견");
            foreach (var wb in result)
            {
                Debug.WriteLine($"[GetAllWorkbookPaths] 워크북: {wb}, 시트 수: {_sheetPaths[wb].Count}");
            }

            return result;
        }

        // 워크북 경로를 기준으로 모든 시트 경로 가져오기
        public Dictionary<string, string> GetAllSheetPathsByWorkbookPath(string workbookPath)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            
            try
            {
                // 먼저 전체 경로로 시도
                if (!string.IsNullOrEmpty(workbookPath))
                {
                    if (LazyLoadSheetPaths().ContainsKey(workbookPath))
                    {
                        var sheetDict = LazyLoadSheetPaths()[workbookPath];
                        Debug.WriteLine($"[GetAllSheetPaths] 파일명 '{workbookPath}'에서 시트 경로 발견: {sheetDict.Count}개");
                        
                        foreach (var entry in sheetDict)
                        {
                            result[entry.Key] = entry.Value.SavePath;
                        }
                    }
                    
                    // 정규화된 경로로 다시 시도
                    string normalizedPath = NormalizeWorkbookPath(workbookPath);
                    if (normalizedPath != workbookPath && LazyLoadSheetPaths().ContainsKey(normalizedPath))
                    {
                        var sheetDict = LazyLoadSheetPaths()[normalizedPath];
                        Debug.WriteLine($"[GetAllSheetPaths] 정규화된 경로 '{normalizedPath}'에서 시트 경로 발견: {sheetDict.Count}개");
                        
                        foreach (var entry in sheetDict)
                        {
                            if (!result.ContainsKey(entry.Key))
                            {
                                result[entry.Key] = entry.Value.SavePath;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[GetAllSheetPathsByWorkbookPath] 시트 경로 조회 중 예외 발생: {ex.Message}");
            }
            
            return result;
        }

        // 파일명만으로 모든 시트 경로 가져오기
        public Dictionary<string, string> GetAllSheetPathsByFileName(string fileName)
        {
            if (string.IsNullOrEmpty(fileName) || !LazyLoadSheetPaths().ContainsKey(fileName))
            {
                return new Dictionary<string, string>();
            }

            // 파일명을 기준으로 모든 시트의 경로 반환
            return LazyLoadSheetPaths()[fileName].ToDictionary(kvp => kvp.Key, kvp => kvp.Value.SavePath);
        }

        /// <summary>
        /// 설정 디렉토리 경로를 반환합니다.
        /// </summary>
        /// <returns>설정 디렉토리 경로</returns>
        private string GetSettingsDirectory()
        {
            return Path.GetDirectoryName(ConfigFilePath);
        }

        /// <summary>
        /// 설정 파일 경로를 반환합니다.
        /// </summary>
        /// <returns>설정 파일 경로</returns>
        private string GetSettingsFilePath()
        {
            return ConfigFilePath;
        }

        private string GetSheetPathInternal(string workbookName, string sheetName)
        {
            Debug.WriteLine($"[GetSheetPathInternal] 호출: 워크북 '{workbookName}', 시트 '{sheetName}'");

            if (string.IsNullOrEmpty(workbookName))
            {
                Debug.WriteLine("[GetSheetPathInternal] 실패: 워크북 이름이 비어있음");
                return string.Empty;
            }

            // 워크북 및 시트 정보 가져오기 시도
            var workbookInfo = GetSheetPathsInternal(workbookName);
            if (workbookInfo == null || !workbookInfo.ContainsKey(sheetName))
            {
                Debug.WriteLine($"[GetSheetPathInternal] 시트 정보 없음: 워크북 '{workbookName}', 시트 '{sheetName}'");
                return string.Empty;
            }

            var sheetInfo = workbookInfo[sheetName];
            if (sheetInfo == null)
            {
                Debug.WriteLine($"[GetSheetPathInternal] 시트 정보가 null: 워크북 '{workbookName}', 시트 '{sheetName}'");
                return string.Empty;
            }

            // 경로가 비어있는지 확인
            if (string.IsNullOrEmpty(sheetInfo.SavePath))
            {
                bool isEnabled = IsSheetEnabled(workbookName, sheetName);
                Debug.WriteLine($"[GetSheetPathInternal] 경로가 비어있음: 워크북 '{workbookName}', 시트 '{sheetName}', 활성화 상태: {isEnabled}");
                
                // 경로가 없더라도 시트가 활성화되어 있으면 로그 기록
                if (isEnabled)
                {
                    Debug.WriteLine($"[GetSheetPathInternal] 주의: 시트가 활성화되어 있지만 경로가 없음: 워크북 '{workbookName}', 시트 '{sheetName}'");
                }
                
                return string.Empty;
            }

            Debug.WriteLine($"[GetSheetPathInternal] 성공: 워크북 '{workbookName}', 시트 '{sheetName}', 경로 '{sheetInfo.SavePath}'");
            return sheetInfo.SavePath;
        }

        // 워크북의 시트 정보를 가져오는 내부 메서드
        private Dictionary<string, SheetPathInfo> GetSheetPathsInternal(string workbookName)
        {
            if (string.IsNullOrEmpty(workbookName))
            {
                Debug.WriteLine("[GetSheetPathsInternal] 실패: 워크북 이름이 비어있음");
                return null;
            }

            // 워크북 정보가 있는지 확인
            if (!LazyLoadSheetPaths().ContainsKey(workbookName))
            {
                Debug.WriteLine($"[GetSheetPathsInternal] 워크북 정보 없음: '{workbookName}'");
                return null;
            }

            return LazyLoadSheetPaths()[workbookName];
        }

        // 특정 워크북의 특정 시트가 활성화되었는지 확인
        private bool IsSheetEnabled(string workbookName, string sheetName)
        {
            if (string.IsNullOrEmpty(workbookName) || string.IsNullOrEmpty(sheetName))
            {
                return false;
            }

            var paths = GetSheetPathsInternal(workbookName);
            if (paths == null || !paths.ContainsKey(sheetName))
            {
                return false;
            }

            return paths[sheetName].Enabled;
        }

        // 현재 워크북 모든 경로 중 활성화된 것만 가져오기
        public Dictionary<string, string> GetAllEnabledSheetPaths()
        {
            Debug.WriteLine($"[GetAllEnabledSheetPaths] 호출: 현재 워크북 경로 '{_currentWorkbookPath}'");
            Dictionary<string, string> result = new Dictionary<string, string>();

            // 현재 워크북 경로가 비어있는 경우
            if (string.IsNullOrEmpty(_currentWorkbookPath))
            {
                Debug.WriteLine("[GetAllEnabledSheetPaths] 실패: 현재 워크북 경로가 비어있음");
                return result;
            }
            
            // 0. 현재 워크북 경로 정규화
            string normalizedPath = NormalizeWorkbookPath(_currentWorkbookPath);
            string fileName = Path.GetFileName(_currentWorkbookPath);
            
            Debug.WriteLine($"[GetAllEnabledSheetPaths] 워크북 정보: 원본='{_currentWorkbookPath}', 정규화='{normalizedPath}', 파일명='{fileName}'");

            // 1. 원본 워크북 경로로 시트 정보 확인
            if (LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath))
            {
                Debug.WriteLine($"[GetAllEnabledSheetPaths] 원본 경로로 시트 정보 확인: '{_currentWorkbookPath}'");
                AddEnabledSheetPathsToResult(LazyLoadSheetPaths()[_currentWorkbookPath], result);
            }
            
            // 2. 정규화된 경로로 시트 정보 확인 (원본과 다른 경우만)
            if (_currentWorkbookPath != normalizedPath && LazyLoadSheetPaths().ContainsKey(normalizedPath))
            {
                Debug.WriteLine($"[GetAllEnabledSheetPaths] 정규화된 경로로 시트 정보 확인: '{normalizedPath}'");
                AddEnabledSheetPathsToResult(LazyLoadSheetPaths()[normalizedPath], result);
            }
            
            // 3. 파일명으로도 시트 정보 확인
            if (!string.IsNullOrEmpty(fileName) && LazyLoadSheetPaths().ContainsKey(fileName))
            {
                Debug.WriteLine($"[GetAllEnabledSheetPaths] 파일명으로 시트 정보 확인: '{fileName}'");
                AddEnabledSheetPathsToResult(LazyLoadSheetPaths()[fileName], result);
            }
            
            // 4. 같은 파일명을 가진 다른 형식의 경로도 확인
            foreach (var workbookEntry in LazyLoadSheetPaths())
            {
                string workbookKey = workbookEntry.Key;
                if (workbookKey != _currentWorkbookPath && workbookKey != normalizedPath && workbookKey != fileName && 
                    !string.IsNullOrEmpty(fileName))
                {
                    // 파일명이 같은 경우 검사
                    string entryFileName = Path.GetFileName(workbookKey);
                    if (entryFileName == fileName)
                    {
                        Debug.WriteLine($"[GetAllEnabledSheetPaths] 같은 파일명을 가진 다른 경로 확인: '{workbookKey}'");
                        AddEnabledSheetPathsToResult(workbookEntry.Value, result);
                    }
                }
            }
            
            // 5. 결과가 비어있는 경우, 모든 워크북에서 시트 정보 찾기 (최후의 수단)
            if (result.Count == 0)
            {
                Debug.WriteLine($"[GetAllEnabledSheetPaths] 결과가 비어있어 모든 워크북에서 활성화된 시트 찾기 시도");
                foreach (var workbookEntry in LazyLoadSheetPaths())
                {
                    string workbookKey = workbookEntry.Key;
                    Debug.WriteLine($"[GetAllEnabledSheetPaths] 워크북 검사: '{workbookKey}'");
                    AddEnabledSheetPathsToResult(workbookEntry.Value, result);
                }
            }

            Debug.WriteLine($"[GetAllEnabledSheetPaths] 완료: 활성화된 시트 수 {result.Count}");
            foreach (var kvp in result)
            {
                Debug.WriteLine($"[GetAllEnabledSheetPaths] 활성화된 시트: '{kvp.Key}', 경로: '{kvp.Value}'");
            }
            return result;
        }
        
        // 활성화된 시트만 결과 사전에 추가하는 헬퍼 메서드
        private void AddEnabledSheetPathsToResult(Dictionary<string, SheetPathInfo> sheetInfos, Dictionary<string, string> result)
        {
            if (sheetInfos == null) return;
            
            Debug.WriteLine($"[AddEnabledSheetPathsToResult] 시트 정보 확인 시작 (총 {sheetInfos.Count}개)");
            
            foreach (var sheet in sheetInfos)
            {
                string sheetName = sheet.Key;
                var sheetInfo = sheet.Value;
                
                // 시트가 활성화되어 있는지 확인
                bool isEnabled = sheetInfo != null && sheetInfo.Enabled;
                string path = sheetInfo?.SavePath ?? string.Empty;
                
                Debug.WriteLine($"[AddEnabledSheetPathsToResult] 시트 확인: '{sheetName}', 활성화 상태: {isEnabled}, 경로: '{path}'");
                
                if (isEnabled)
                {
                    // 이미 결과에 추가된 시트가 아닌 경우에만 추가
                    if (!result.ContainsKey(sheetName))
                    {
                        Debug.WriteLine($"[AddEnabledSheetPathsToResult] 활성화된 시트 추가: '{sheetName}', 경로: '{path}'");
                        
                        // 경로가 비어 있더라도 활성화된 시트는 결과에 포함 (경로는 빈 문자열)
                        result[sheetName] = path;
                    }
                    else
                    {
                        Debug.WriteLine($"[AddEnabledSheetPathsToResult] 이미 결과에 추가된 시트: '{sheetName}'");
                    }
                }
                else
                {
                    Debug.WriteLine($"[AddEnabledSheetPathsToResult] 비활성화된 시트: '{sheetName}'");
                }
            }
            
            Debug.WriteLine($"[AddEnabledSheetPathsToResult] 시트 정보 확인 완료, 현재 결과에 추가된 시트 수: {result.Count}");
        }

        // 특정 시트가 활성화되었는지 확인하는 메서드
        public bool IsSheetEnabled(string sheetName)
        {
            Debug.WriteLine($"[IsSheetEnabled] 호출: 시트='{sheetName}'");
            
            // 1. GetSheetEnabled로 먼저 확인
            bool result = GetSheetEnabled(sheetName);
            Debug.WriteLine($"[IsSheetEnabled] GetSheetEnabled 결과: 시트 '{sheetName}'의 활성화 상태: {result}");
            
            // 2. 활성화된 모든 시트 경로에서도 확인 (결과가 false인 경우)
            if (!result)
            {
                // 현재 워크북의 모든 활성화된 시트 가져오기
                var allEnabledPaths = GetAllEnabledSheetPaths();
                if (allEnabledPaths.ContainsKey(sheetName))
                {
                    result = true;
                    Debug.WriteLine($"[IsSheetEnabled] GetAllEnabledSheetPaths에서 시트 '{sheetName}'이 활성화된 것으로 확인됨");
                }
                
                // 3. !로 시작하는 경우와 시작하지 않는 경우 모두 확인
                if (!result)
                {
                    string alternateSheetName;
                    if (sheetName.StartsWith("!"))
                    {
                        // !로 시작하면 접두사 제거한 버전 확인
                        alternateSheetName = sheetName.Substring(1);
                        Debug.WriteLine($"[IsSheetEnabled] ! 접두사 제거 후 다시 확인: '{alternateSheetName}'");
                    }
                    else
                    {
                        // !로 시작하지 않으면 접두사 추가한 버전 확인
                        alternateSheetName = "!" + sheetName;
                        Debug.WriteLine($"[IsSheetEnabled] ! 접두사 추가 후 다시 확인: '{alternateSheetName}'");
                    }
                    
                    // 접두사를 추가/제거한 이름으로 활성화 상태 다시 확인
                    bool alternateResult = GetSheetEnabled(alternateSheetName);
                    if (alternateResult)
                    {
                        Debug.WriteLine($"[IsSheetEnabled] 대체 이름 '{alternateSheetName}'로 활성화된 것으로 확인됨");
                        result = true;
                    }
                    
                    // GetAllEnabledSheetPaths에서도 확인
                    if (!result && allEnabledPaths.ContainsKey(alternateSheetName))
                    {
                        Debug.WriteLine($"[IsSheetEnabled] GetAllEnabledSheetPaths에서 대체 이름 '{alternateSheetName}'이 활성화된 것으로 확인됨");
                        result = true;
                    }
                }
            }
            
            // 최종 결과 반환
            Debug.WriteLine($"[IsSheetEnabled] 최종 결과: 시트 '{sheetName}'의 활성화 상태: {result}");
            return result;
        }
    }

    /// <summary>
    /// 시트 경로 정보를 저장하는 클래스
    /// </summary>
    public class SheetPathInfo
    {
        /// <summary>
        /// 저장 경로
        /// </summary>
        public string SavePath { get; set; } = "";
        
        /// <summary>
        /// 활성화 여부
        /// </summary>
        public bool Enabled { get; set; } = false; // 기본값을 false로 변경
        
        /// <summary>
        /// YAML 빈 필드 처리 옵션
        /// </summary>
        public bool YamlEmptyFields { get; set; } = false;
    }

    // XML 직렬화를 위한 클래스
    [Serializable]
    public class SheetPathData
    {
        public string WorkbookPath { get; set; }
        public string SheetName { get; set; }
        public string SavePath { get; set; }
        public bool Enabled { get; set; } = true;
        public bool YamlEmptyFields { get; set; } = false;
        public string MergeKeyPaths { get; set; } = ""; // 후처리용 키 경로 인수
        public string FlowStyleConfig { get; set; } = ""; // YAML Flow Style 설정
    }
}

