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

        // 현재 워크북 설정
        public void SetCurrentWorkbook(string workbookPath)
        {
            // 전체 경로에서 파일 이름만 추출
            string fileName = Path.GetFileName(workbookPath);
            _currentWorkbookPath = fileName; // 파일 이름만 저장

            Debug.WriteLine($"[SetCurrentWorkbook] 전체 경로: {workbookPath}, 설정된 워크북 이름: {_currentWorkbookPath}");

            // 설정 파일 로드 확인
            if (_sheetPaths == null)
            {
                LoadSheetPaths();
                Debug.WriteLine($"[SetCurrentWorkbook] 설정 파일 로드됨, 총 워크북 수: {(_sheetPaths != null ? _sheetPaths.Count : 0)}");
            }

            // 현재 _sheetPaths 내용 로깅
            if (_sheetPaths != null && _sheetPaths.Count > 0)
            {
                Debug.WriteLine($"[SetCurrentWorkbook] 현재 로드된 워크북 수: {_sheetPaths.Count}");
                foreach (var wb in _sheetPaths.Keys)
                {
                    Debug.WriteLine($"[SetCurrentWorkbook] 로드된 워크북: {wb}, 시트 수: {_sheetPaths[wb].Count}");
                }
            }

            // 워크북 경로가 딕셔너리에 없으면 추가 (파일명)
            if (!LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath))
            {
                LazyLoadSheetPaths()[_currentWorkbookPath] = new Dictionary<string, SheetPathInfo>();
                Debug.WriteLine($"[SetCurrentWorkbook] 워크북 '{_currentWorkbookPath}'에 대한 새 사전 생성");
            }

            // 전체 경로도 확인 (파일명과 다른 경우)
            if (workbookPath != _currentWorkbookPath && !LazyLoadSheetPaths().ContainsKey(workbookPath))
            {
                LazyLoadSheetPaths()[workbookPath] = new Dictionary<string, SheetPathInfo>();
                Debug.WriteLine($"[SetCurrentWorkbook] 전체 경로 '{workbookPath}'에 대한 새 사전 생성");
            }

            // 전체 경로와 파일명이 다른 경우 내용 동기화
            if (workbookPath != _currentWorkbookPath &&
                LazyLoadSheetPaths().ContainsKey(workbookPath) &&
                LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath))
            {
                // 전체 경로 사전에 있는 모든 항목을 파일명 사전으로 복사
                foreach (var entry in LazyLoadSheetPaths()[workbookPath])
                {
                    if (!LazyLoadSheetPaths()[_currentWorkbookPath].ContainsKey(entry.Key))
                    {
                        LazyLoadSheetPaths()[_currentWorkbookPath][entry.Key] = entry.Value;
                        Debug.WriteLine($"[SetCurrentWorkbook] 경로 '{workbookPath}'에서 '{_currentWorkbookPath}'로 항목 '{entry.Key}'='{entry.Value}' 복사");
                    }
                }

                // 파일명 사전에 있는 모든 항목을 전체 경로 사전으로 복사
                foreach (var entry in LazyLoadSheetPaths()[_currentWorkbookPath])
                {
                    if (!LazyLoadSheetPaths()[workbookPath].ContainsKey(entry.Key))
                    {
                        LazyLoadSheetPaths()[workbookPath][entry.Key] = entry.Value;
                        Debug.WriteLine($"[SetCurrentWorkbook] 경로 '{_currentWorkbookPath}'에서 '{workbookPath}'로 항목 '{entry.Key}'='{entry.Value}' 복사");
                    }
                }
            }
        }

        // 특정 시트의 저장 경로 추가 또는 업데이트
        public void SetSheetPath(string sheetName, string path)
        {
            if (string.IsNullOrEmpty(_currentWorkbookPath))
                return;

            SetSheetPath(_currentWorkbookPath, sheetName, path);
        }

        // 특정 워크북의 시트 경로 설정
        public void SetSheetPath(string workbookName, string sheetName, string path)
        {
            // 기본값으로 활성화 상태를 true로 설정하고 YAML 선택적 필드를 false로 설정
            SetSheetPath(workbookName, sheetName, path, true, false);
        }

        // 특정 워크북의 시트 경로 및 활성화 상태 설정 (오버로드)
        public void SetSheetPath(string workbookName, string sheetName, string path, bool enabled)
        {
            // YAML 선택적 필드 처리 기본값으로 false 설정
            SetSheetPath(workbookName, sheetName, path, enabled, false);
        }

        // 특정 워크북의 시트 경로, 활성화 상태 및 YAML 선택적 필드 처리 상태 설정 (오버로드)
        public void SetSheetPath(string workbookName, string sheetName, string path, bool enabled, bool yamlEmptyFields)
        {
            Debug.WriteLine($"[SetSheetPath] 시트 경로 설정: 워크북 '{workbookName}', 시트 '{sheetName}', 경로 '{path}', 활성화: {enabled}, YAML 선택적 필드: {yamlEmptyFields}");

            if (!LazyLoadSheetPaths().ContainsKey(workbookName))
            {
                LazyLoadSheetPaths()[workbookName] = new Dictionary<string, SheetPathInfo>();
            }

            LazyLoadSheetPaths()[workbookName][sheetName] = new SheetPathInfo { 
                Path = path, 
                Enabled = enabled,
                YamlEmptyFields = yamlEmptyFields
            };
            SaveSheetPaths(); // 변경 즉시 저장
        }

        // 특정 시트의 저장 경로 가져오기
        public string GetSheetPath(string sheetName)
        {
            Debug.WriteLine($"[GetSheetPath] 시작: 현재 워크북={_currentWorkbookPath}, 시트={sheetName}");

            if (string.IsNullOrEmpty(_currentWorkbookPath) ||
                !LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath) ||
                !LazyLoadSheetPaths()[_currentWorkbookPath].ContainsKey(sheetName))
            {
                Debug.WriteLine($"[GetSheetPath] 시트 경로 조회 실패: workbook={_currentWorkbookPath}, sheet={sheetName}");

                // workbookPath가 _sheetPaths에 있는지 확인
                if (!string.IsNullOrEmpty(_currentWorkbookPath))
                {
                    Debug.WriteLine($"[GetSheetPath] 워크북 '{_currentWorkbookPath}'가 딕셔너리에 {(LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath) ? "있음" : "없음")}");

                    // 워크북은 있지만 시트가 없는 경우
                    if (LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath))
                    {
                        Debug.WriteLine($"[GetSheetPath] 워크북 '{_currentWorkbookPath}'에 등록된 시트 수: {LazyLoadSheetPaths()[_currentWorkbookPath].Count}");
                        foreach (var sheet in LazyLoadSheetPaths()[_currentWorkbookPath])
                        {
                            Debug.WriteLine($"[GetSheetPath] 등록된 시트: {sheet.Key}, 경로: {sheet.Value}");
                        }

                        // 시트 이름에 '!'가 포함된 경우와 그렇지 않은 경우 모두 시도
                        string altSheetName = sheetName.StartsWith("!") ? sheetName.Substring(1) : $"!{sheetName}";
                        Debug.WriteLine($"[GetSheetPath] 대체 시트 이름 '{altSheetName}' 확인");

                        if (LazyLoadSheetPaths()[_currentWorkbookPath].ContainsKey(altSheetName))
                        {
                            Debug.WriteLine($"[GetSheetPath] 대체 시트 이름 '{altSheetName}'로 경로 찾음: {LazyLoadSheetPaths()[_currentWorkbookPath][altSheetName]}");
                            return LazyLoadSheetPaths()[_currentWorkbookPath][altSheetName].Path;
                        }
                    }
                }

                return null;
            }

            var sheetInfo = LazyLoadSheetPaths()[_currentWorkbookPath][sheetName];
            Debug.WriteLine($"[GetSheetPath] 시트 경로 조회 성공: '{sheetName}' -> '{sheetInfo.Path}', 활성화: {sheetInfo.Enabled}");
            return sheetInfo.Path;
        }

        // 특정 워크북의 시트 경로 사전 반환 (수정된 메서드)
        public Dictionary<string, string> GetSheetPaths(string workbookPath)
        {
            // 1. 먼저 전체 경로로 시도
            if (!string.IsNullOrEmpty(workbookPath) &&
                LazyLoadSheetPaths().ContainsKey(workbookPath))
            {
                Debug.WriteLine($"[GetSheetPaths] 전체 경로 '{workbookPath}'에서 시트 경로 발견: {LazyLoadSheetPaths()[workbookPath].Count}개");
                return LazyLoadSheetPaths()[workbookPath].ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Path);
            }

            // 2. 파일 이름만으로도 시도
            string fileName = Path.GetFileName(workbookPath);
            if (!string.IsNullOrEmpty(fileName) &&
                LazyLoadSheetPaths().ContainsKey(fileName))
            {
                Debug.WriteLine($"[GetSheetPaths] 파일명 '{fileName}'에서 시트 경로 발견: {LazyLoadSheetPaths()[fileName].Count}개");
                return LazyLoadSheetPaths()[fileName].ToDictionary(kvp => kvp.Key, kvp => kvp.Value.Path);
            }

            Debug.WriteLine($"[GetSheetPaths] '{workbookPath}' 또는 '{fileName}'에 대한 시트 경로를 찾을 수 없습니다.");
            return new Dictionary<string, string>();
        }

        // 현재 선택된 워크북의 모든 시트 경로 반환
        // 파일 경로 디버깅 추가
        public Dictionary<string, string> GetAllSheetPaths()
        {
            if (string.IsNullOrEmpty(_currentWorkbookPath))
            {
                Debug.WriteLine("[GetAllSheetPaths] 현재 워크북이 설정되지 않았습니다.");
                return new Dictionary<string, string>();
            }

            var result = new Dictionary<string, string>();

            // 1. 먼저 파일명으로 시도 (_currentWorkbookPath에는 파일명만 저장)
            if (LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath))
            {
                Debug.WriteLine($"[GetAllSheetPaths] 파일명 '{_currentWorkbookPath}'에서 시트 경로 발견: {LazyLoadSheetPaths()[_currentWorkbookPath].Count}개");

                foreach (var entry in LazyLoadSheetPaths()[_currentWorkbookPath])
                {
                    Debug.WriteLine($"[GetAllSheetPaths] 추가됨: 시트 '{entry.Key}' -> 경로 '{entry.Value}'");
                    result[entry.Key] = entry.Value.Path;
                }
            }

            // 2. 전체 경로 검색 (워크북 이름과 일치하는 모든 항목 검색)
            var fileName = Path.GetFileName(_currentWorkbookPath);
            foreach (var key in LazyLoadSheetPaths().Keys)
            {
                // 현재 워크북 경로가 이미 체크됐으면 스킵
                if (key == _currentWorkbookPath) continue;

                // 파일명이 일치하는 전체 경로 처리
                if (Path.GetFileName(key) == fileName || key == fileName)
                {
                    Debug.WriteLine($"[GetAllSheetPaths] 다른 키 '{key}'에서 시트 경로 발견: {LazyLoadSheetPaths()[key].Count}개");

                    foreach (var entry in LazyLoadSheetPaths()[key])
                    {
                        if (!result.ContainsKey(entry.Key))
                        {
                            Debug.WriteLine($"[GetAllSheetPaths] 추가됨: 시트 '{entry.Key}' -> 경로 '{entry.Value}'");
                            result[entry.Key] = entry.Value.Path;
                        }
                    }
                }
            }

            if (result.Count > 0)
            {
                Debug.WriteLine($"[GetAllSheetPaths] 워크북 '{_currentWorkbookPath}'에 대한 시트 경로를 총 {result.Count}개 찾았습니다.");
                return result;
            }

            Debug.WriteLine($"[GetAllSheetPaths] 워크북 '{_currentWorkbookPath}'에 대한 시트 경로를 찾을 수 없습니다.");
            return new Dictionary<string, string>();
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

        // 설정 파일 저장
        private void SaveSheetPaths()
        {
            try
            {
                // 디렉토리 생성
                string directory = Path.GetDirectoryName(ConfigFilePath);
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // XML 직렬화를 위한 임시 클래스 생성
                var serializableData = new List<SheetPathData>();

                foreach (var workbook in LazyLoadSheetPaths())
                {
                    foreach (var sheet in workbook.Value)
                    {
                        serializableData.Add(new SheetPathData
                        {
                            WorkbookPath = workbook.Key,
                            SheetName = sheet.Key,
                            SavePath = sheet.Value.Path,
                            Enabled = sheet.Value.Enabled,
                            YamlEmptyFields = sheet.Value.YamlEmptyFields,
                            MergeKeyPaths = sheet.Value.MergeKeyPaths
                        });
                    }
                }

                // XML 직렬화
                var serializer = new XmlSerializer(typeof(List<SheetPathData>));
                using (var writer = new StreamWriter(ConfigFilePath))
                {
                    serializer.Serialize(writer, serializableData);
                }

                Debug.WriteLine($"시트 경로 설정이 저장되었습니다: {ConfigFilePath}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"시트 경로 저장 중 오류 발생: {ex.Message}");
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
                if (!File.Exists(ConfigFilePath))
                {
                    Debug.WriteLine("[LoadSheetPaths] 시트 경로 설정 파일이 없습니다. 새로 생성됩니다.");
                    _sheetPaths = new Dictionary<string, Dictionary<string, SheetPathInfo>>();
                    return;
                }

                Debug.WriteLine($"[LoadSheetPaths] 시트 경로 설정 파일 로드 시작: {ConfigFilePath}");

                // XML 역직렬화
                var serializer = new XmlSerializer(typeof(List<SheetPathData>));
                List<SheetPathData> serializableData;

                using (var reader = new StreamReader(ConfigFilePath))
                {
                    serializableData = (List<SheetPathData>)serializer.Deserialize(reader);
                }

                Debug.WriteLine($"[LoadSheetPaths] 불러온 SheetPathData 항목 수: {serializableData.Count}");

                // 기존 데이터 초기화
                _sheetPaths = new Dictionary<string, Dictionary<string, SheetPathInfo>>();

                foreach (var data in serializableData)
                {
                    Debug.WriteLine($"[LoadSheetPaths] 불러온 시트 경로 설정: 워크북='{data.WorkbookPath}', 시트='{data.SheetName}', 경로='{data.SavePath}'");

                    // 워크북 경로 딕셔너리가 없으면 생성
                    if (!_sheetPaths.ContainsKey(data.WorkbookPath))
                    {
                        _sheetPaths[data.WorkbookPath] = new Dictionary<string, SheetPathInfo>();
                        Debug.WriteLine($"[LoadSheetPaths] 워크북 '{data.WorkbookPath}'에 대한 새 딕셔너리 생성");
                    }

                    // 시트 경로 설정
                    _sheetPaths[data.WorkbookPath][data.SheetName] = new SheetPathInfo
                    {
                        Path = data.SavePath,
                        Enabled = data.Enabled,
                        YamlEmptyFields = data.YamlEmptyFields,
                        MergeKeyPaths = data.MergeKeyPaths
                    };
                    Debug.WriteLine($"[LoadSheetPaths] 시트 경로 저장: 워크북='{data.WorkbookPath}', 시트='{data.SheetName}', 경로='{data.SavePath}', YAML 선택적 필드: {data.YamlEmptyFields}");
                }

                // 저장된 시트 경로 항목 수 출력
                int totalSettings = 0;
                foreach (var workbook in _sheetPaths.Keys)
                {
                    totalSettings += _sheetPaths[workbook].Count;
                    Debug.WriteLine($"[LoadSheetPaths] 워크북 '{workbook}'에 등록된 시트 수: {_sheetPaths[workbook].Count}");
                    foreach (var sheet in _sheetPaths[workbook].Keys)
                    {
                        Debug.WriteLine($"[LoadSheetPaths]   시트: '{sheet}', 경로: '{_sheetPaths[workbook][sheet].Path}'");
                    }
                }
                Debug.WriteLine($"[LoadSheetPaths] 총 저장된 시트 경로 항목 수: {totalSettings}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[LoadSheetPaths] 시트 경로 설정 불러오기 중 오류 발생: {ex.Message}");
                Debug.WriteLine($"[LoadSheetPaths] 스택 트레이스: {ex.StackTrace}");
                // 오류 발생 시 새로운 딕셔너리 생성
                _sheetPaths = new Dictionary<string, Dictionary<string, SheetPathInfo>>();
            }
        }

        // 지연 초기화 패턴을 적용한 LoadSheetPaths 호출
        private Dictionary<string, Dictionary<string, SheetPathInfo>> LazyLoadSheetPaths()
        {
            if (_sheetPaths == null)
            {
                _sheetPaths = new Dictionary<string, Dictionary<string, SheetPathInfo>>();
                LoadSheetPaths();
            }
            return _sheetPaths;
        }

        // LazyLoadSheetPaths를 호출하는 메서드
        public void Initialize()
        {
            // 기존 데이터를 제거하고 새로 로드
            _sheetPaths = null;
            Debug.WriteLine("[Initialize] 시트 경로 설정을 다시 로드합니다.");
            LazyLoadSheetPaths();
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

        // 특정 워크북의 시트 경로 활성화 상태 가져오기
        public bool GetSheetEnabled(string sheetName)
        {
            if (string.IsNullOrEmpty(_currentWorkbookPath) ||
                !LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath) ||
                !LazyLoadSheetPaths()[_currentWorkbookPath].ContainsKey(sheetName))
            {
                return false;
            }

            return LazyLoadSheetPaths()[_currentWorkbookPath][sheetName].Enabled;
        }

        // 현재 워크북 모든 경로 중 활성화된 것만 가져오기
        public Dictionary<string, string> GetAllEnabledSheetPaths()
        {
            if (string.IsNullOrEmpty(_currentWorkbookPath) || !LazyLoadSheetPaths().ContainsKey(_currentWorkbookPath))
            {
                Debug.WriteLine($"[GetAllEnabledSheetPaths] 워크북 '{_currentWorkbookPath}'의 시트 경로가 없습니다.");
                return new Dictionary<string, string>();
            }

            // 활성화된 경로만 반환
            var result = new Dictionary<string, string>();
            foreach (var sheetInfo in LazyLoadSheetPaths()[_currentWorkbookPath])
            {
                if (sheetInfo.Value.Enabled)
                {
                    result[sheetInfo.Key] = sheetInfo.Value.Path;
                }
            }
            return result;
        }

        // 특정 시트가 활성화되었는지 확인하는 메서드 (IsSheetEnabled가 GetSheetEnabled를 호출)
        public bool IsSheetEnabled(string sheetName)
        {
            return GetSheetEnabled(sheetName);
        }

        // 워크북의 특정 시트의 YAML 선택적 필드 처리 여부 가져오기
        /// <param name="sheetName">시트 이름</param>
        /// <returns>YAML 선택적 필드 처리 여부</returns>
        public bool GetYamlEmptyFieldsOption(string sheetName)
        {
            // 현재 워크북 경로가 설정되어 있지 않으면 기본값 false 반환
            if (string.IsNullOrEmpty(_currentWorkbookPath))
            {
                Debug.WriteLine("[SheetPathManager] GetYamlEmptyFieldsOption: 현재 워크북이 설정되지 않음");
                return false;
            }

            // 워크북의 파일명과 전체 경로로 모두 확인
            string workbookName = Path.GetFileName(_currentWorkbookPath);
            
            Debug.WriteLine($"[SheetPathManager] GetYamlEmptyFieldsOption 호출: 시트={sheetName}, 워크북={workbookName}, 전체경로={_currentWorkbookPath}");
            
            // 모든 워크북 설정 정보 출력
            var sheetPaths = LazyLoadSheetPaths();
            Debug.WriteLine($"[SheetPathManager] 저장된 워크북 수: {sheetPaths.Count}");
            foreach (var workbook in sheetPaths.Keys)
            {
                Debug.WriteLine($"[SheetPathManager] 워크북: {workbook}, 시트 수: {sheetPaths[workbook].Count}");
                foreach (var sheet in sheetPaths[workbook].Keys)
                {
                    Debug.WriteLine($"[SheetPathManager] - 시트: {sheet}, YAML 설정: {sheetPaths[workbook][sheet].YamlEmptyFields}");
                }
            }
            
            // 먼저 파일명으로 확인
            if (sheetPaths.ContainsKey(workbookName) &&
                sheetPaths[workbookName].ContainsKey(sheetName))
            {
                bool result = sheetPaths[workbookName][sheetName].YamlEmptyFields;
                Debug.WriteLine($"[SheetPathManager] GetYamlEmptyFieldsOption: {workbookName} / {sheetName} -> {result}");
                return result;
            }
            else
            {
                Debug.WriteLine($"[SheetPathManager] 파일명으로 시트를 찾지 못함: {workbookName} / {sheetName}");
            }

            // 파일명으로 찾지 못하면 전체 경로로 확인
            if (sheetPaths.ContainsKey(_currentWorkbookPath) &&
                sheetPaths[_currentWorkbookPath].ContainsKey(sheetName))
            {
                bool result = sheetPaths[_currentWorkbookPath][sheetName].YamlEmptyFields;
                Debug.WriteLine($"[SheetPathManager] GetYamlEmptyFieldsOption: {_currentWorkbookPath} / {sheetName} -> {result}");
                return result;
            }
            else
            {
                Debug.WriteLine($"[SheetPathManager] 전체 경로로도 시트를 찾지 못함: {_currentWorkbookPath} / {sheetName}");
            }

            // 해당 시트에 대한 설정이 없으면 기본값 false 반환
            Debug.WriteLine($"[SheetPathManager] GetYamlEmptyFieldsOption: {sheetName} -> false (기본값)");
            return false;
        }

        /// <summary>
        /// 시트의 후처리용 키 경로 인수 값을 가져옵니다.
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>설정된 후처리 키 경로 인수 값 (없으면 빈 문자열)</returns>
        public string GetMergeKeyPaths(string sheetName)
        {
            if (string.IsNullOrEmpty(_currentWorkbookPath))
                return "";

            return GetMergeKeyPaths(_currentWorkbookPath, sheetName);
        }

        /// <summary>
        /// 특정 워크북의 시트에 대한 후처리용 키 경로 인수 값을 가져옵니다.
        /// </summary>
        /// <param name="workbookName">워크북 이름 또는 경로</param>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>설정된 후처리 키 경로 인수 값 (없으면 빈 문자열)</returns>
        public string GetMergeKeyPaths(string workbookName, string sheetName)
        {
            // 설정이 없으면 빈 문자열 반환
            if (!LazyLoadSheetPaths().ContainsKey(workbookName) ||
                !LazyLoadSheetPaths()[workbookName].ContainsKey(sheetName))
            {
                return "";
            }

            return LazyLoadSheetPaths()[workbookName][sheetName].MergeKeyPaths ?? "";
        }

        /// <summary>
        /// 시트의 후처리용 키 경로 인수 값을 설정합니다.
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="mergeKeyPaths">설정할 후처리 키 경로 인수 값</param>
        public void SetMergeKeyPaths(string sheetName, string mergeKeyPaths)
        {
            if (string.IsNullOrEmpty(_currentWorkbookPath))
                return;

            SetMergeKeyPaths(_currentWorkbookPath, sheetName, mergeKeyPaths);
        }

        /// <summary>
        /// 특정 워크북의 시트에 대한 후처리용 키 경로 인수 값을 설정합니다.
        /// </summary>
        /// <param name="workbookName">워크북 이름 또는 경로</param>
        /// <param name="sheetName">시트 이름</param>
        /// <param name="mergeKeyPaths">설정할 후처리 키 경로 인수 값</param>
        public void SetMergeKeyPaths(string workbookName, string sheetName, string mergeKeyPaths)
        {
            // 워크북과 시트가 없으면 무시
            if (!LazyLoadSheetPaths().ContainsKey(workbookName) ||
                !LazyLoadSheetPaths()[workbookName].ContainsKey(sheetName))
            {
                return;
            }

            // 값 설정 및 디버그 로그
            LazyLoadSheetPaths()[workbookName][sheetName].MergeKeyPaths = mergeKeyPaths;
            Debug.WriteLine($"[SetMergeKeyPaths] 후처리 키 경로 설정: 워크북='{workbookName}', 시트='{sheetName}', 값='{mergeKeyPaths}'");
        }
    }

    // 시트 경로 정보 클래스 (내부 용도)
    public class SheetPathInfo
    {
        public string Path { get; set; }
        public bool Enabled { get; set; } = true;
        public bool YamlEmptyFields { get; set; } = false;
        public string MergeKeyPaths { get; set; } = ""; // 후처리용 키 경로 인수 (예: "test:merge;test2:append")
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
    }
}
