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

                // 해당 시트 항목이 없으면 생성
                if (!LazyLoadSheetPaths()[workbookKey].ContainsKey(sheetName))
                {
                    LazyLoadSheetPaths()[workbookKey][sheetName] = new SheetPathInfo();
                }

                // 경로 설정
                if (path != null)
                {
                    LazyLoadSheetPaths()[workbookKey][sheetName].SavePath = path;
                    Debug.WriteLine($"[SetSheetPathInternal] 시트 경로 설정됨: workbookKey='{workbookKey}', sheetName='{sheetName}', path='{path}'");
                }
                else
                {
                    LazyLoadSheetPaths()[workbookKey][sheetName].SavePath = "";
                    Debug.WriteLine($"[SetSheetPathInternal] 시트 경로 초기화됨: workbookKey='{workbookKey}', sheetName='{sheetName}'");
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
        public void SetSheetEnabled(string sheetName, bool enabled)
        {
            if (string.IsNullOrEmpty(_currentWorkbookPath))
            {
                Debug.WriteLine("[SheetPathManager] SetSheetEnabled: 현재 워크북이 설정되지 않음");
                return;
            }

            string workbookName = Path.GetFileName(_currentWorkbookPath);
            Debug.WriteLine($"[SheetPathManager] SetSheetEnabled 호출: 워크북={_currentWorkbookPath}, 시트={sheetName}, Enabled={enabled}");

            // 파일명과 전체 경로로 모두 업데이트
            SetSheetEnabledInternal(workbookName, sheetName, enabled);
            SetSheetEnabledInternal(_currentWorkbookPath, sheetName, enabled);
            
            // 활성화 상태 변경 후 즉시 저장
            SaveSheetPaths();
        }

        /// <summary>
        /// 특정 워크북의 시트에 대한 활성화 상태를 설정합니다. (오버로드)
        /// </summary>
        /// <param name="workbookName">워크북 경로 또는 이름</param>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="enabled">활성화 여부</param>
        public void SetSheetEnabled(string workbookName, string sheetName, bool enabled)
        {
            Debug.WriteLine($"[SheetPathManager] SetSheetEnabled 오버로드 호출: 워크북={workbookName}, 시트={sheetName}, Enabled={enabled}");
            SetSheetEnabledInternal(workbookName, sheetName, enabled);
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
            if (string.IsNullOrEmpty(workbookName) ||
                !LazyLoadSheetPaths().ContainsKey(workbookName) ||
                !LazyLoadSheetPaths()[workbookName].ContainsKey(sheetName))
            {
                return;
            }

            LazyLoadSheetPaths()[workbookName].Remove(sheetName);
        }

        // 기존 메서드도 남겨두기 (호환성을 위해)
        public void RemoveSheetPath(string sheetName)
        {
            if (string.IsNullOrEmpty(_currentWorkbookPath) ||
                !LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath) ||
                !LazyLoadSheetPaths()[_currentWorkbookPath].ContainsKey(sheetName))
            {
                return;
            }

            LazyLoadSheetPaths()[_currentWorkbookPath].Remove(sheetName);
            SaveSheetPaths();
        }

        /// <summary>
        /// 모든 시트 경로 설정을 XML 파일에 저장합니다.
        /// </summary>
        public void SaveSheetPaths()
        {
            try
            {
                // 설정 디렉토리가 없으면 생성
                if (!Directory.Exists(GetSettingsDirectory()))
                {
                    Directory.CreateDirectory(GetSettingsDirectory());
                    Debug.WriteLine($"[SaveSheetPaths] 설정 디렉토리 생성: {GetSettingsDirectory()}");
                }

                List<SheetPathData> allPaths = new List<SheetPathData>();
                Debug.WriteLine($"[SaveSheetPaths] 저장 시작: 워크북 수={LazyLoadSheetPaths().Count}");

                // 모든 워크북에 대해 반복
                foreach (var workbookEntry in LazyLoadSheetPaths())
                {
                    string workbookKey = workbookEntry.Key;
                    
                    // OneDrive URL이면 정규화 적용
                    if (IsOneDrivePath(workbookKey))
                    {
                        workbookKey = NormalizeWorkbookPath(workbookKey);
                    }
                    
                    Dictionary<string, SheetPathInfo> sheetInfos = workbookEntry.Value;
                    Debug.WriteLine($"[SaveSheetPaths] 워크북 '{workbookKey}': 시트 수={sheetInfos.Count}");

                    // 각 워크북의 모든 시트에 대해 반복
                    foreach (var sheetEntry in sheetInfos)
                    {
                        string sheetName = sheetEntry.Key;
                        SheetPathInfo info = sheetEntry.Value;

                        if (info != null) 
                        {
                            SheetPathData data = new SheetPathData()
                            {
                                WorkbookPath = workbookKey,
                                SheetName = sheetName, 
                                SavePath = info.SavePath,
                                Enabled = info.Enabled
                            };
                            
                            // XML에는 기본 설정만 저장 (YAML 관련 설정은 Excel에 저장)
                            // info.YamlEmptyFields는 Excel에 저장되므로 여기에 포함하지 않음
                            
                            allPaths.Add(data);
                            Debug.WriteLine($"[SaveSheetPaths] 추가: 워크북='{workbookKey}', 시트='{sheetName}', 경로='{info.SavePath}', 활성화={info.Enabled}");
                        }
                    }
                }

                // 직렬화 설정
                XmlSerializer serializer = new XmlSerializer(typeof(List<SheetPathData>));
                
                // XML 파일에 저장
                using (StreamWriter writer = new StreamWriter(GetSettingsFilePath()))
                {
                    serializer.Serialize(writer, allPaths);
                }

                Debug.WriteLine($"[SaveSheetPaths] 시트 경로 설정이 저장되었습니다: {GetSettingsFilePath()}, 총 항목 수: {allPaths.Count}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[SaveSheetPaths] 시트 경로 설정 저장 중 오류 발생: {ex.Message}\n{ex.StackTrace}");
            }
        }

        // 외부에서 접근 가능한 설정 저장 메서드
        public void SaveSettings()
        {
            SaveSheetPaths();
        }

        // 설정 파일 로드
        private void LoadSheetPaths()
        {
            try
            {
                if (File.Exists(GetSettingsFilePath()))
                {
                    Debug.WriteLine($"[LoadSheetPaths] 시트 경로 설정 파일 로드 시작: {GetSettingsFilePath()}");
                    XmlSerializer serializer = new XmlSerializer(typeof(List<SheetPathData>));
                    List<SheetPathData> loadedPaths;

                    using (StreamReader reader = new StreamReader(GetSettingsFilePath()))
                    {
                        loadedPaths = (List<SheetPathData>)serializer.Deserialize(reader);
                    }

                    Debug.WriteLine($"[LoadSheetPaths] 로드된 항목 수: {loadedPaths.Count}");

                    // 로드된 데이터를 딕셔너리에 추가
                    foreach (var pathData in loadedPaths)
                    {
                        // 워크북 경로가 없으면 건너뜀
                        if (string.IsNullOrEmpty(pathData.WorkbookPath))
                        {
                            Debug.WriteLine("[LoadSheetPaths] 워크북 경로가 없는 항목 발견, 건너뜀");
                            continue;
                        }

                        // 시트 이름이 없으면 건너뜀
                        if (string.IsNullOrEmpty(pathData.SheetName))
                        {
                            Debug.WriteLine($"[LoadSheetPaths] 시트 이름이 없는 항목 발견, 워크북: {pathData.WorkbookPath}, 건너뜀");
                            continue;
                        }

                        // 워크북 경로 정규화
                        string normalizedPath = NormalizeWorkbookPath(pathData.WorkbookPath);
                        
                        // 워크북 항목이 없으면 생성
                        if (!_sheetPaths.ContainsKey(normalizedPath))
                        {
                            _sheetPaths[normalizedPath] = new Dictionary<string, SheetPathInfo>();
                            Debug.WriteLine($"[LoadSheetPaths] 워크북 '{normalizedPath}'에 대한 새 사전 생성");
                        }

                        // 시트 정보 생성
                        SheetPathInfo info = new SheetPathInfo
                        {
                            SavePath = pathData.SavePath ?? "",
                            Enabled = pathData.Enabled,
                            YamlEmptyFields = pathData.YamlEmptyFields
                        };

                        // 시트 정보 추가
                        _sheetPaths[normalizedPath][pathData.SheetName] = info;
                        Debug.WriteLine($"[LoadSheetPaths] 시트 경로 추가: 워크북='{normalizedPath}', 시트='{pathData.SheetName}', 경로='{info.SavePath}', 활성화={info.Enabled}");

                        // 파일명만으로도 추가 (중복 방지를 위해 파일명이 이미 키로 존재하는지 확인)
                        string fileName = Path.GetFileName(pathData.WorkbookPath);
                        if (!string.IsNullOrEmpty(fileName))
                        {
                            // 파일명 항목이 없으면 생성
                            if (!_sheetPaths.ContainsKey(fileName))
                            {
                                _sheetPaths[fileName] = new Dictionary<string, SheetPathInfo>();
                                Debug.WriteLine($"[LoadSheetPaths] 파일명 '{fileName}'에 대한 새 사전 생성");
                            }

                            // 시트 정보 복제
                            _sheetPaths[fileName][pathData.SheetName] = new SheetPathInfo
                            {
                                SavePath = info.SavePath,
                                Enabled = info.Enabled,
                                YamlEmptyFields = info.YamlEmptyFields
                            };
                            Debug.WriteLine($"[LoadSheetPaths] 파일명으로 시트 경로 추가: 파일명='{fileName}', 시트='{pathData.SheetName}', 경로='{info.SavePath}', 활성화={info.Enabled}");
                        }
                    }

                    Debug.WriteLine($"[LoadSheetPaths] 시트 경로 설정 로드 완료: 워크북 수={_sheetPaths.Count}");
                    foreach (var wb in _sheetPaths.Keys)
                    {
                        Debug.WriteLine($"[LoadSheetPaths] 로드된 워크북: {wb}, 시트 수: {_sheetPaths[wb].Count}");
                    }
                }
                else
                {
                    Debug.WriteLine($"[LoadSheetPaths] 시트 경로 설정 파일이 존재하지 않습니다: {GetSettingsFilePath()}");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LoadSheetPaths] 시트 경로 설정 로드 중 오류 발생: {ex.Message}\n{ex.StackTrace}");
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
                // 설정 디렉토리가 없으면 생성
                string settingsDirectory = GetSettingsDirectory();
                if (!Directory.Exists(settingsDirectory))
                {
                    Directory.CreateDirectory(settingsDirectory);
                    Debug.WriteLine($"[Initialize] 설정 디렉토리 생성: {settingsDirectory}");
                }
                
                // 설정 파일 로드
                LoadSheetPaths();
                Debug.WriteLine("[Initialize] 시트 경로 설정 로드 완료");
                
                // 현재 워크북이 설정되어 있으면 출력
                if (!string.IsNullOrEmpty(_currentWorkbookPath))
                {
                    Debug.WriteLine($"[Initialize] 현재 워크북 경로: {_currentWorkbookPath}");
                }
                
                // 현재 로드된 워크북 정보 출력
                string fileName = !string.IsNullOrEmpty(_currentWorkbookPath) ? Path.GetFileName(_currentWorkbookPath) : "";
                
                Debug.WriteLine($"[Initialize] 현재 로드된 워크북 수: {LazyLoadSheetPaths().Count}");
                foreach (var wb in LazyLoadSheetPaths().Keys)
                {
                    Debug.WriteLine($"[Initialize] 워크북: {wb}, 시트 수: {LazyLoadSheetPaths()[wb].Count}");
                    
                    // 현재 워크북이거나 같은 파일명이면 시트 정보도 출력
                    if (wb == _currentWorkbookPath || 
                        (!string.IsNullOrEmpty(fileName) && Path.GetFileName(wb) == fileName))
                    {
                        DumpSheetInfo(LazyLoadSheetPaths()[wb]);
                    }
                }
                
                // 설정 저장 (데이터 동기화 및 정합성 유지)
                SaveSettings();
                Debug.WriteLine("[Initialize] 설정 저장 완료");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Initialize] 오류 발생: {ex.Message}");
                
                // 설정 초기화 실패시 기본 빈 사전으로 초기화
                if (_sheetPaths == null)
                {
                    _sheetPaths = new Dictionary<string, Dictionary<string, SheetPathInfo>>();
                    Debug.WriteLine("[Initialize] 오류로 인해 빈 사전으로 초기화");
                }
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
        /// 시트의 저장 경로
        /// </summary>
        public string SavePath { get; set; } = "";

        /// <summary>
        /// 활성화 여부
        /// </summary>
        public bool Enabled { get; set; } = true;

        /// <summary>
        /// YAML 선택적 필드 처리 여부 (더 이상 XML에 저장되지 않음, ExcelConfigManager에서 관리)
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

