using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using System.Text.RegularExpressions;

namespace ExcelToJsonAddin.Core.YamlPostProcessors
{
    /// <summary>
    /// YAML 파일에서 키 경로를 기반으로 객체를 병합하는 프로세서입니다.
    /// 파이썬의 merge_roles.py를 C#으로 구현한 버전입니다.
    /// </summary>
    public class YamlMergeKeyPathsProcessor
    {
        private readonly string idPath;
        private readonly string[] mergePaths;
        private readonly Dictionary<string, string> keyPathStrategies;

        /// <summary>
        /// 기본 생성자
        /// </summary>
        /// <param name="idPath">ID가 있는 경로 (기본값 없음)</param>
        /// <param name="mergePaths">병합할 경로들 (기본값 없음)</param>
        public YamlMergeKeyPathsProcessor(string idPath = "", string mergePaths = "")
        {
            // 기본값을 사용하지 않고 그대로 저장 (Python 로직과 일치)
            this.idPath = idPath;
            this.mergePaths = string.IsNullOrWhiteSpace(mergePaths) ? new string[0] : mergePaths.Split(',');
            this.keyPathStrategies = new Dictionary<string, string>();
        }

        /// <summary>
        /// 설정 문자열을 파싱하여 프로세서를 생성합니다.
        /// </summary>
        /// <param name="mergeKeyPathsConfig">설정 문자열 (형식: "idPath|mergePaths|keyPaths")</param>
        /// <returns>YamlMergeKeyPathsProcessor 인스턴스</returns>
        public static YamlMergeKeyPathsProcessor FromConfigString(string mergeKeyPathsConfig)
        {
            string idPath = "";
            string mergePaths = "";
            string keyPaths = "";
            
            if (!string.IsNullOrWhiteSpace(mergeKeyPathsConfig))
            {
                string[] parts = mergeKeyPathsConfig.Split('|');
                if (parts.Length >= 1)
                    idPath = parts[0]; // 빈 문자열도 그대로 사용
                if (parts.Length >= 2)
                    mergePaths = parts[1]; // 빈 문자열도 그대로 사용
                if (parts.Length >= 3)
                    keyPaths = parts[2]; // 빈 문자열도 그대로 사용
            }
            
            YamlMergeKeyPathsProcessor processor = new YamlMergeKeyPathsProcessor(idPath, mergePaths);
            processor.ParseKeyPaths(keyPaths); // 빈 문자열이어도 호출
            
            return processor;
        }

        /// <summary>
        /// YAML 파일을 처리하여 지정된 키 경로를 기반으로 항목을 병합합니다.
        /// </summary>
        /// <param name="yamlPath">처리할 YAML 파일 경로</param>
        /// <param name="keyPaths">키 경로:전략 문자열 (예: "level:merge;achievement:append")</param>
        /// <returns>처리 성공 여부</returns>
        public bool ProcessYamlFile(string yamlPath, string keyPaths)
        {
            try
            {
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] YAML 파일 처리: {yamlPath}");
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] ID 경로: {idPath}");
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 병합 경로: {string.Join(", ", mergePaths)}");
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 키 경로: {keyPaths}");

                // Python 구현과 일치: keyPaths가 비어있으면 처리하지 않음
                if (string.IsNullOrWhiteSpace(keyPaths))
                {
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 키 경로가 비어있어 처리를 중단합니다.");
                    return true; // 처리하지 않더라도 성공으로 처리
                }

                // ID 경로와 병합 경로가 모두 비어있으면 처리하지 않음
                if (string.IsNullOrWhiteSpace(idPath) && (mergePaths == null || mergePaths.Length == 0 || (mergePaths.Length == 1 && string.IsNullOrWhiteSpace(mergePaths[0]))))
                {
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] ID 경로와 병합 경로가 모두 비어있어 처리를 중단합니다.");
                    return true; // 처리하지 않더라도 성공으로 처리
                }

                if (!File.Exists(yamlPath))
                {
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 오류: YAML 파일을 찾을 수 없습니다.");
                    return false;
                }

                // 키 경로:전략 파싱
                ParseKeyPaths(keyPaths);
                
                // 키 경로 전략이 없으면 처리하지 않음 (Python 로직과 일치)
                if (keyPathStrategies.Count == 0)
                {
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 유효한 키 경로 전략이 없어 처리를 중단합니다.");
                    return true; // 처리하지 않더라도 성공으로 처리
                }

                // YAML 파일 읽기
                string yamlContent = File.ReadAllText(yamlPath);
                
                if (string.IsNullOrWhiteSpace(yamlContent))
                {
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 오류: YAML 파일이 비어 있습니다.");
                    return false;
                }

                // YAML 역직렬화
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(CamelCaseNamingConvention.Instance)
                    .Build();

                // YAML 내용을 리스트로 변환
                var yamlList = deserializer.Deserialize<List<Dictionary<string, object>>>(yamlContent);

                if (yamlList == null || yamlList.Count == 0)
                {
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 오류: YAML을 리스트로 변환할 수 없거나 비어 있습니다.");
                    return false;
                }

                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 총 {yamlList.Count}개 항목 로드됨");

                // 병합 처리
                var mergedList = GenericMerge(yamlList);

                // 직렬화 옵션 설정
                var serializer = new SerializerBuilder()
                    .WithNamingConvention(CamelCaseNamingConvention.Instance)
                    .ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitNull)
                    .Build();

                // 수정된 객체를 YAML로 직렬화
                string modifiedYaml = serializer.Serialize(mergedList);

                // 파일 저장
                File.WriteAllText(yamlPath, modifiedYaml);
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] YAML 파일 병합 처리 완료: {yamlPath}, 병합 후 항목 수: {mergedList.Count}");

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 처리 중 오류 발생: {ex.Message}");
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 스택 추적: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 설정 문자열로부터 YAML 파일을 처리하여 지정된 키 경로를 기반으로 항목을 병합합니다.
        /// </summary>
        /// <param name="yamlPath">처리할 YAML 파일 경로</param>
        /// <param name="mergeKeyPathsConfig">설정 문자열 (형식: "idPath|mergePaths|keyPaths")</param>
        /// <returns>처리 성공 여부</returns>
        public static bool ProcessYamlFileFromConfig(string yamlPath, string mergeKeyPathsConfig)
        {
            // 설정 문자열이 비어있으면 후처리 실행하지 않음
            if (string.IsNullOrWhiteSpace(mergeKeyPathsConfig))
            {
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 설정 문자열이 비어있어 후처리를 실행하지 않습니다: {yamlPath}");
                return true; // 후처리를 실행하지 않더라도 성공으로 처리
            }
            
            string idPath = "";
            string mergePaths = "";
            string keyPaths = "";
            
            string[] parts = mergeKeyPathsConfig.Split('|');
            if (parts.Length >= 1)
                idPath = parts[0]; // 빈 문자열도 그대로 사용
            if (parts.Length >= 2)
                mergePaths = parts[1]; // 빈 문자열도 그대로 사용
            if (parts.Length >= 3)
                keyPaths = parts[2]; // 빈 문자열도 그대로 사용
            
            // 모든 인수가 비어있으면 후처리 실행하지 않음
            if (string.IsNullOrWhiteSpace(idPath) && string.IsNullOrWhiteSpace(mergePaths) && string.IsNullOrWhiteSpace(keyPaths))
            {
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 모든 인수가 비어있어 후처리를 실행하지 않습니다: {yamlPath}");
                return true; // 후처리를 실행하지 않더라도 성공으로 처리
            }
            
            // Python 구현과 일치: keyPaths가 비어있으면 처리하지 않음
            if (string.IsNullOrWhiteSpace(keyPaths))
            {
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 키 경로가 비어있어 후처리를 실행하지 않습니다: {yamlPath}");
                return true; // 후처리를 실행하지 않더라도 성공으로 처리
            }
            
            Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 설정 문자열 파싱: ID 경로={idPath}, 병합 경로={mergePaths}, 키 경로={keyPaths}");
            
            YamlMergeKeyPathsProcessor processor = new YamlMergeKeyPathsProcessor(idPath, mergePaths);
            return processor.ProcessYamlFile(yamlPath, keyPaths);
        }

        /// <summary>
        /// 키 경로:전략 문자열을 파싱합니다.
        /// </summary>
        /// <param name="keyPaths">키 경로:전략 문자열 (예: "level:merge;achievement:append")</param>
        private void ParseKeyPaths(string keyPaths)
        {
            keyPathStrategies.Clear();
            if (string.IsNullOrWhiteSpace(keyPaths)) return;

            var pairs = keyPaths.Split(';');
            foreach (var pair in pairs)
            {
                var parts = pair.Split(':');
                if (parts.Length == 2)
                {
                    var keyPath = parts[0].Trim();
                    var strategy = parts[1].Trim();
                    keyPathStrategies[keyPath] = strategy;
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 키 경로 파싱: {keyPath} -> {strategy}");
                }
            }
        }

        /// <summary>
        /// 점 표기법으로 객체에서 값을 가져옵니다.
        /// </summary>
        /// <param name="data">데이터 객체</param>
        /// <param name="path">점 표기법 경로</param>
        /// <param name="defaultValue">기본값</param>
        /// <returns>경로에 해당하는 값 또는 기본값</returns>
        private object DeepGet(Dictionary<string, object> data, string path, object defaultValue = null)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(path))
                    return defaultValue;

                var keys = path.Split('.');
                object current = data;

                foreach (var key in keys)
                {
                    if (current is Dictionary<string, object> dict)
                    {
                        if (!dict.TryGetValue(key, out current))
                            return defaultValue;
                    }
                    // 배열 인덱스 참조 처리 (0, 1, 2 등 숫자 인덱스)
                    else if (current is List<object> list && int.TryParse(key, out int index) && index >= 0 && index < list.Count)
                    {
                        current = list[index];
                    }
                    else
                    {
                        return defaultValue;
                    }
                }

                return current;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] DeepGet 오류: {ex.Message} [경로: {path}]");
                return defaultValue;
            }
        }

        /// <summary>
        /// 두 객체를 재귀적으로 병합합니다. (Python deep_merge와 일치)
        /// </summary>
        /// <param name="target">대상 객체</param>
        /// <param name="source">소스 객체</param>
        private void DeepMerge(Dictionary<string, object> target, Dictionary<string, object> source)
        {
            foreach (var kvp in source)
            {
                if (!target.ContainsKey(kvp.Key))
                {
                    // 키가 없으면 새로 추가
                    target[kvp.Key] = CloneObject(kvp.Value);
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] DeepMerge - 새 키 추가: {kvp.Key}");
                }
                else if (kvp.Value is Dictionary<string, object> sourceDict && 
                         target[kvp.Key] is Dictionary<string, object> targetDict)
                {
                    // 양쪽 모두 딕셔너리인 경우 재귀 병합
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] DeepMerge - 딕셔너리 재귀 병합: {kvp.Key}");
                    DeepMerge(targetDict, sourceDict);
                }
                else if (kvp.Value is List<object> sourceList && 
                         target[kvp.Key] is List<object> targetList)
                {
                    // 양쪽 모두 리스트인 경우 확장 (Python과 일치)
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] DeepMerge - 리스트 확장: {kvp.Key}");
                    targetList.AddRange(sourceList.Select(CloneObject));
                }
                else if (target[kvp.Key] == null)
                {
                    // 대상이 null인 경우 소스 값으로 교체
                    target[kvp.Key] = CloneObject(kvp.Value);
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] DeepMerge - null 값 대체: {kvp.Key}");
                }
                else if (keyPathStrategies.TryGetValue(kvp.Key, out var strategy) && strategy.Equals("replace", StringComparison.OrdinalIgnoreCase))
                {
                    // 전략이 replace인 경우 값 대체
                    target[kvp.Key] = CloneObject(kvp.Value);
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] DeepMerge - 전략에 따른 값 대체: {kvp.Key}");
                }
                // 그 외 경우에는 기존 값 유지 (Python과 동일)
            }
        }

        /// <summary>
        /// 객체를 깊은 복사합니다.
        /// </summary>
        private object CloneObject(object obj)
        {
            if (obj is Dictionary<string, object> dict)
            {
                var newDict = new Dictionary<string, object>();
                foreach (var kvp in dict)
                {
                    newDict[kvp.Key] = CloneObject(kvp.Value);
                }
                return newDict;
            }
            else if (obj is List<object> list)
            {
                return list.Select(CloneObject).ToList();
            }
            else
            {
                return obj; // 기본 타입은 값 복사로 됨
            }
        }

        /// <summary>
        /// 점 표기법으로 객체에 값을 설정합니다.
        /// </summary>
        /// <param name="data">데이터 객체</param>
        /// <param name="path">점 표기법 경로</param>
        /// <param name="value">설정할 값</param>
        private void DeepSet(Dictionary<string, object> data, string path, object value)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;

            var keys = path.Split('.');
            Dictionary<string, object> current = data;

            // 마지막 키를 제외한 모든 키에 대해 경로 생성
            for (int i = 0; i < keys.Length - 1; i++)
            {
                var key = keys[i];
                if (!current.ContainsKey(key) || !(current[key] is Dictionary<string, object>))
                {
                    current[key] = new Dictionary<string, object>();
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] DeepSet - 새 딕셔너리 생성: {key}");
                }
                current = (Dictionary<string, object>)current[key];
            }

            // 마지막 키에 값 설정
            current[keys[keys.Length - 1]] = value;
            Debug.WriteLine($"[YamlMergeKeyPathsProcessor] DeepSet - 값 설정: {keys[keys.Length - 1]}");
        }

        /// <summary>
        /// 와일드카드 경로 확장 (Python 구현과 일치)
        /// </summary>
        /// <param name="data">데이터 객체</param>
        /// <param name="path">와일드카드를 포함한 경로</param>
        /// <returns>확장된 경로 리스트</returns>
        private List<string> ExpandWildcardPaths(Dictionary<string, object> data, string path)
        {
            if (!path.Contains(".*."))
            {
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] ExpandWildcardPaths - 와일드카드 없음: {path}");
                return new List<string> { path };
            }

            // Python 구현과 일치: 첫 번째 발견된 .*. 만 처리
            var parts = path.Split(new[] { ".*." }, 2, StringSplitOptions.None);
            var basePart = parts[0];
            var restPart = parts[1];

            var items = DeepGet(data, basePart) as List<object>;
            var result = new List<string>();

            if (items != null)
            {
                for (int i = 0; i < items.Count; i++)
                {
                    result.Add($"{basePart}.{i}.{restPart}");
                }
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] ExpandWildcardPaths - 원본 경로: {path}, 확장된 경로 수: {result.Count}");
            }

            // Python 구현과 일치: 결과가 없으면 빈 리스트 대신 원본 경로 반환
            return result.Count > 0 ? result : new List<string> { path };
        }

        /// <summary>
        /// 객체를 실제 사용할 수 있는 Dictionary로 변환하는 헬퍼 메서드
        /// </summary>
        private Dictionary<string, object> ConvertToDictionary(object obj)
        {
            // 이미 Dictionary<string, object>인 경우 그대로 반환
            if (obj is Dictionary<string, object> dict)
                return dict;

            // Dictionary<object, object>인 경우 변환
            if (obj is Dictionary<object, object> dictObj)
            {
                var newDict = new Dictionary<string, object>();
                foreach (var kvp in dictObj)
                {
                    if (kvp.Key != null)
                    {
                        string key = kvp.Key.ToString();
                        newDict[key] = kvp.Value;
                    }
                }
                return newDict;
            }

            // 그 외 다른 형태의 사전도 변환 시도
            try
            {
                var newDict = new Dictionary<string, object>();
                Type type = obj.GetType();
                
                // 리플렉션으로 속성 추출
                foreach (var prop in type.GetProperties())
                {
                    if (prop.CanRead)
                    {
                        newDict[prop.Name] = prop.GetValue(obj);
                    }
                }
                
                // 항목이 없으면 원본 개체를 문자열화하여 디버그 출력
                if (newDict.Count == 0)
                {
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 변환 실패한 객체: {obj}");
                }
                
                return newDict;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] ConvertToDictionary 오류: {ex.Message}");
                return new Dictionary<string, object>();
            }
        }

        /// <summary>
        /// 리스트 항목을 병합합니다. (수정된 버전)
        /// </summary>
        /// <param name="data">데이터 리스트</param>
        /// <returns>병합된 리스트</returns>
        private List<Dictionary<string, object>> GenericMerge(List<Dictionary<string, object>> data)
        {
            Debug.WriteLine("[YamlMergeKeyPathsProcessor] 병합 프로세스 시작");

            // Python 구현과 유사한 구조의 중첩 사전 사용
            var mergedItems = new Dictionary<object, Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>>>();
            var preservedFields = new Dictionary<object, Dictionary<string, object>>();

            // 각 항목을 처리
            foreach (var item in data)
            {
                // ID 경로에서 ID 값 가져오기
                var itemId = DeepGet(item, idPath);
                if (itemId == null)
                {
                    Debug.WriteLine("[YamlMergeKeyPathsProcessor] ID 누락 항목 건너뜀");
                    continue;
                }

                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 처리 중인 항목 ID: {itemId}");

                // ID가 mergedItems에 없으면 초기화
                if (!mergedItems.ContainsKey(itemId))
                {
                    mergedItems[itemId] = new Dictionary<string, Dictionary<string, List<Dictionary<string, object>>>>();
                    preservedFields[itemId] = new Dictionary<string, object>();
                }

                // 병합 대상이 아닌 최상위 필드들 보존
                foreach (var kvp in item)
                {
                    bool isIdField = kvp.Key == idPath || idPath.StartsWith(kvp.Key + ".");
                    bool isMergePath = mergePaths.Contains(kvp.Key);
                    
                    if (!isIdField && !isMergePath)
                    {
                        preservedFields[itemId][kvp.Key] = CloneObject(kvp.Value);
                        Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 보존 필드 저장 - ID: {itemId}, 필드: {kvp.Key}");
                    }
                }

                // 각 병합 경로 처리
                foreach (var mergePath in mergePaths)
                {
                    if (!item.ContainsKey(mergePath))
                    {
                        Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 병합 경로 데이터 없음 - ID: {itemId}, 경로: {mergePath}");
                        continue;
                    }
                    
                    // 리스트로 처리 시도
                    List<object> mergeData = null;
                    if (item[mergePath] is List<object> list)
                    {
                        mergeData = list;
                    }
                    else
                    {
                        // 다른 타입인 경우 로그 출력 후 계속 진행
                        Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 병합 경로 데이터가 리스트가 아님 - 타입: {item[mergePath]?.GetType().FullName ?? "null"}");
                        continue;
                    }
                    
                    Debug.WriteLine($"[YamlMergeKeyPathsProcessor] ID: {itemId}, 경로: {mergePath} - 항목 수: {mergeData.Count}");

                    if (!mergedItems[itemId].ContainsKey(mergePath))
                    {
                        mergedItems[itemId][mergePath] = new Dictionary<string, List<Dictionary<string, object>>>();
                    }

                    // 각 병합 데이터 항목 처리
                    foreach (var obj in mergeData)
                    {
                        // 타입 정보 출력
                        Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 항목 타입: {obj?.GetType().FullName ?? "null"}");

                        // 사전으로 변환 시도
                        Dictionary<string, object> dataItem = null;
                        try
                        {
                            dataItem = ConvertToDictionary(obj);
                            
                            if (dataItem.Count == 0)
                            {
                                Debug.WriteLine("[YamlMergeKeyPathsProcessor] 변환된a 사전이 비어있음, 건너뜀");
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 사전 변환 실패: {ex.Message}");
                            continue;
                        }

                        // 복합 키를 위한 값들 가져오기
                        var compoundKeyParts = new List<string>();
                        var strategies = new List<string>();
                        
                        foreach (var keyPathEntry in keyPathStrategies)
                        {
                            string keyPath = keyPathEntry.Key;
                            string strategy = keyPathEntry.Value;
                            
                            // 키 경로에서 실제 값 가져오기
                            var keyValue = string.Empty;
                            
                            // 여기서 실제로 category와 같은 값을 가져와야 함
                            if (dataItem.ContainsKey(keyPath))
                            {
                                keyValue = dataItem[keyPath]?.ToString() ?? string.Empty;
                                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 키값 찾음: {keyPath} -> {keyValue}");
                            }
                            else
                            {
                                // 키가 없는 경우 빈 문자열 사용
                                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 키를 찾을 수 없음: {keyPath}");
                            }
                            
                            compoundKeyParts.Add(keyValue);
                            strategies.Add(strategy);
                        }
                        
                        // 복합 키 생성
                        string compoundKey = string.Join("|", compoundKeyParts);
                        
                        Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 복합키: {compoundKey}");

                        // 복합 키가 없으면 초기화
                        if (!mergedItems[itemId][mergePath].ContainsKey(compoundKey))
                        {
                            mergedItems[itemId][mergePath][compoundKey] = new List<Dictionary<string, object>>();
                        }

                        var target = mergedItems[itemId][mergePath][compoundKey];
                        if (target.Count == 0)
                        {
                            // 딕셔너리 깊은 복사
                            var newItem = new Dictionary<string, object>();
                            foreach (var kvp in dataItem)
                            {
                                newItem[kvp.Key] = CloneObject(kvp.Value);
                            }
                            target.Add(newItem);
                            Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 새 항목 추가 - ID: {itemId}, 키: {compoundKey}");
                        }
                        else
                        {
                            // 기존 항목에 병합 (전략 적용)
                            DeepMerge(target[0], dataItem);
                            Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 항목 병합 - ID: {itemId}, 키: {compoundKey}");
                        }
                    }
                }
            }

            // 병합된 결과를 최종 리스트로 변환
            var result = new List<Dictionary<string, object>>();
            foreach (var kvp in mergedItems)
            {
                var itemId = kvp.Key;
                var mergeData = kvp.Value;

                var newItem = new Dictionary<string, object>();
                DeepSet(newItem, idPath, itemId);

                // 보존된 필드들 복원
                if (preservedFields.ContainsKey(itemId))
                {
                    foreach (var field in preservedFields[itemId])
                    {
                        newItem[field.Key] = field.Value;
                        Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 보존 필드 복원 - ID: {itemId}, 필드: {field.Key}");
                    }
                }

                // 각 병합 경로에 대한 병합 결과 설정
                foreach (var mergePath in mergePaths)
                {
                    if (mergeData.ContainsKey(mergePath))
                    {
                        var mergedList = new List<object>();
                        foreach (var items in mergeData[mergePath].Values)
                        {
                            mergedList.AddRange(items.Cast<object>());
                        }
                        newItem[mergePath] = mergedList;
                        Debug.WriteLine($"[YamlMergeKeyPathsProcessor] ID: {itemId} - 병합된 {mergePath} 항목 수: {mergedList.Count}");
                    }
                }

                result.Add(newItem);
            }

            Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 병합 완료 - 총 항목: {result.Count}");
            return result;
        }
    }
} 