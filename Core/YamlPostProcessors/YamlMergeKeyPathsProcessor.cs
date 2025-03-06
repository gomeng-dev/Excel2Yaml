using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using YamlDotNet.Serialization.EventEmitters;
using YamlDotNet.Core;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text;
using YamlDotNet.Serialization.ObjectFactories;
using ExcelToJsonAddin.Core;
using ExcelToJsonAddin.Config;

namespace ExcelToJsonAddin.Core.YamlPostProcessors
{
    /// <summary>
    /// YAML 파일의 항목들을 지정된 키 경로를 기반으로 병합하는 후처리기입니다.
    /// </summary>
    public class YamlMergeKeyPathsProcessor
    {
        // 기본 설정들
        private readonly string idPath;
        private readonly string[] mergePaths;
        private readonly Dictionary<string, string> keyPathStrategies;
        private readonly Dictionary<string, string> arrayFieldStrategies;
        
        // 추가 설정 옵션들
        private readonly char keyPathSeparator;
        private readonly char strategySeparator;
        private readonly char arrayPathSeparator;
        private readonly bool preserveArrayOrder;
        private readonly INamingConvention namingConvention;
        private readonly bool debugMode;
        private readonly bool includeEmptyFields;

        /// <summary>
        /// 따옴표 없이 YAML 값을 직렬화하기 위한 도우미 클래스
        /// </summary>
        private class YamlScalarValue
        {
            public string Value { get; }
            
            public YamlScalarValue(string value)
            {
                Value = value;
            }
            
            public override string ToString()
            {
                return Value;
            }
        }

        /// <summary>
        /// YamlMergeKeyPathsProcessor 설정을 위한 빌더 클래스
        /// </summary>
        public class Builder
        {
            private string idPath = "";
            private string mergePaths = "";
            private string keyPaths = "";
            private string arrayFieldPaths = "";
            private char keyPathSeparator = '.';
            private char strategySeparator = ':';
            private char arrayPathSeparator = ';';
            private bool preserveArrayOrder = false;
            private INamingConvention namingConvention = CamelCaseNamingConvention.Instance;
            private bool debugMode = false;
            private bool includeEmptyFields = false;

            public Builder WithIdPath(string idPath)
            {
                this.idPath = idPath;
                return this;
            }

            public Builder WithMergePaths(string mergePaths)
            {
                this.mergePaths = mergePaths;
                return this;
            }

            public Builder WithKeyPaths(string keyPaths)
            {
                this.keyPaths = keyPaths;
                return this;
            }

            public Builder WithArrayFieldPaths(string arrayFieldPaths)
            {
                this.arrayFieldPaths = arrayFieldPaths;
                return this;
            }

            public Builder WithKeyPathSeparator(char keyPathSeparator)
            {
                this.keyPathSeparator = keyPathSeparator;
                return this;
            }

            public Builder WithStrategySeparator(char strategySeparator)
            {
                this.strategySeparator = strategySeparator;
                return this;
            }

            public Builder WithArrayPathSeparator(char arrayPathSeparator)
            {
                this.arrayPathSeparator = arrayPathSeparator;
                return this;
            }

            public Builder WithPreserveArrayOrder(bool preserveArrayOrder)
            {
                this.preserveArrayOrder = preserveArrayOrder;
                return this;
            }

            public Builder WithNamingConvention(INamingConvention namingConvention)
            {
                this.namingConvention = namingConvention;
                return this;
            }

            public Builder WithDebugMode(bool debugMode)
            {
                this.debugMode = debugMode;
                return this;
            }

            public Builder WithIncludeEmptyFields(bool includeEmptyFields)
            {
                this.includeEmptyFields = includeEmptyFields;
                return this;
            }

            public YamlMergeKeyPathsProcessor Build()
            {
                return new YamlMergeKeyPathsProcessor(
                    idPath,
                    mergePaths,
                    keyPaths,
                    arrayFieldPaths,
                    keyPathSeparator,
                    strategySeparator,
                    arrayPathSeparator,
                    preserveArrayOrder,
                    namingConvention,
                    debugMode,
                    includeEmptyFields
                );
            }
        }

        /// <summary>
        /// 새로운 빌더 인스턴스를 생성합니다.
        /// </summary>
        /// <returns>빌더 인스턴스</returns>
        public static Builder CreateBuilder()
        {
            return new Builder();
        }

        /// <summary>
        /// 기본 생성자
        /// </summary>
        /// <param name="idPath">ID가 있는 경로 (기본값 없음)</param>
        /// <param name="mergePaths">병합할 경로들 (기본값 없음)</param>
        /// <param name="keyPathSeparator">키 경로 구분자 (기본값 '.')</param>
        /// <param name="strategySeparator">전략 구분자 (기본값 ':')</param>
        /// <param name="arrayPathSeparator">배열 경로 구분자 (기본값 ';')</param>
        /// <param name="preserveArrayOrder">배열 순서 유지 여부 (기본값 false)</param>
        /// <param name="namingConvention">YAML 네이밍 컨벤션 (기본값 CamelCase)</param>
        /// <param name="debugMode">디버그 모드 여부 (기본값 false)</param>
        /// <param name="includeEmptyFields">빈 필드 포함 여부 (기본값 false)</param>
        public YamlMergeKeyPathsProcessor(
            string idPath = "", 
            string mergePaths = "", 
            string keyPaths = "",
            string arrayFieldPaths = "",
            char keyPathSeparator = '.', 
            char strategySeparator = ':', 
            char arrayPathSeparator = ';', 
            bool preserveArrayOrder = false, 
            INamingConvention namingConvention = null, 
            bool debugMode = false,
            bool includeEmptyFields = false)
        {
            this.idPath = idPath;
            this.mergePaths = string.IsNullOrWhiteSpace(mergePaths) 
                ? new string[0] 
                : mergePaths.Split(new char[] { arrayPathSeparator, ',' }, StringSplitOptions.RemoveEmptyEntries);
            this.keyPathStrategies = new Dictionary<string, string>();
            this.arrayFieldStrategies = new Dictionary<string, string>();
            this.keyPathSeparator = keyPathSeparator;
            this.strategySeparator = strategySeparator;
            this.arrayPathSeparator = arrayPathSeparator;
            this.preserveArrayOrder = preserveArrayOrder;
            this.namingConvention = namingConvention ?? CamelCaseNamingConvention.Instance;
            this.debugMode = debugMode;
            this.includeEmptyFields = includeEmptyFields;

            // 키 경로와 배열 필드 경로 파싱
            ParseKeyPaths(keyPaths);
            ParseArrayFieldPaths(arrayFieldPaths);
        }

        /// <summary>
        /// 설정 문자열을 파싱하여 프로세서를 생성합니다.
        /// </summary>
        /// <param name="mergeKeyPathsConfig">설정 문자열 (형식: "idPath|mergePaths|keyPaths|arrayFieldPaths")</param>
        /// <param name="includeEmptyFields">빈 필드 포함 여부 (기본값 false)</param>
        /// <returns>YamlMergeKeyPathsProcessor 인스턴스</returns>
        public static YamlMergeKeyPathsProcessor FromConfigString(string mergeKeyPathsConfig, bool includeEmptyFields = false)
        {
            string idPath = "";
            string mergePaths = "";
            string keyPaths = "";
            string arrayFieldPaths = "";
            
            if (!string.IsNullOrWhiteSpace(mergeKeyPathsConfig))
            {
                string[] parts = mergeKeyPathsConfig.Split('|');
                if (parts.Length >= 1)
                    idPath = parts[0];
                if (parts.Length >= 2)
                    mergePaths = parts[1];
                if (parts.Length >= 3)
                    keyPaths = parts[2];
                if (parts.Length >= 4)
                    arrayFieldPaths = parts[3];
            }
            
            return new YamlMergeKeyPathsProcessor(
                idPath,
                mergePaths,
                keyPaths,
                arrayFieldPaths,
                '.',
                ':',
                ';',
                false,
                null,
                false,
                includeEmptyFields
            );
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
                LogMessage($"YAML 파일 처리: {yamlPath}");
                LogMessage($"ID 경로: {idPath}");
                LogMessage($"병합 경로: {string.Join(", ", mergePaths)}");
                LogMessage($"키 경로: {keyPaths}");

                // 기본 검증 (YAML 파일 존재 여부만 확인하도록 수정)
                if (!File.Exists(yamlPath))
                {
                    LogMessage($"오류: YAML 파일을 찾을 수 없습니다.");
                    return false;
                }

                // 키 경로가 비어있는 경우 아무런 처리를 하지 않고 성공으로 반환
                if (string.IsNullOrWhiteSpace(keyPaths))
                {
                    LogMessage("키 경로가 비어있어 YAML 파일을 처리하지 않고 종료합니다.");
                    return true;
                }

                // 키 경로:전략 파싱
                ParseKeyPaths(keyPaths);

                // YAML 파일 읽기
                string yamlContent = File.ReadAllText(yamlPath);
                
                if (string.IsNullOrWhiteSpace(yamlContent))
                {
                    LogMessage($"오류: YAML 파일이 비어 있습니다.");
                    return false;
                }

                // YAML을 JSON으로 변환하는 단계
                LogMessage("YAML을 JSON으로 변환 중...");
                JArray jsonArray = ConvertYamlToJson(yamlContent);
                
                if (jsonArray == null || !jsonArray.Any())
                {
                    LogMessage("오류: YAML을 JSON으로 변환하는 데 실패했습니다.");
                    return false;
                }

                LogMessage($"변환된 JSON 배열 항목 수: {jsonArray.Count}");
                
                // 디버깅용: 변환된 JSON 출력 (첫 번째 항목만)
                if (debugMode && jsonArray.Count > 0)
                {
                    LogMessage("첫 번째 JSON 항목 (요약):");
                    string firstItemJson = jsonArray[0].ToString(Formatting.None);
                    LogMessage(firstItemJson.Length <= 500 ? firstItemJson : firstItemJson.Substring(0, 500) + "...");
                    
                    // 디버깅을 위해 임시 파일에 JSON 저장
                    //string jsonDebugPath = Path.Combine(Path.GetDirectoryName(yamlPath), "debug_before_merge.json");
                    //File.WriteAllText(jsonDebugPath, jsonArray.ToString(Formatting.Indented));
                    //LogMessage($"디버깅용 JSON 저장됨: {jsonDebugPath}");
                }
                
                // JSON 병합 처리
                LogMessage("JSON 항목 병합 중...");
                JArray mergedJsonArray = MergeJsonItems(jsonArray);
                
                LogMessage($"병합 결과 항목 수: {mergedJsonArray.Count}");
                
                // 디버깅용: 병합된 JSON 출력 (첫 번째 항목만)
                if (debugMode && mergedJsonArray.Count > 0)
                {
                    LogMessage("첫 번째 병합 항목 (요약):");
                    string firstItemJson = mergedJsonArray[0].ToString(Formatting.None);
                    LogMessage(firstItemJson.Length <= 500 ? firstItemJson : firstItemJson.Substring(0, 500) + "...");
                    
                    // 디버깅을 위해 임시 파일에 JSON 저장
                    //string jsonDebugPath = Path.Combine(Path.GetDirectoryName(yamlPath), "debug_after_merge.json");
                    //File.WriteAllText(jsonDebugPath, mergedJsonArray.ToString(Formatting.Indented));
                    //LogMessage($"디버깅용 JSON 저장됨: {jsonDebugPath}");
                }
                
                // JSON을 YAML로 변환
                LogMessage("JSON을 YAML로 변환 중...");
                string mergedYamlContent = ConvertJsonToYaml(mergedJsonArray);
                
                // 파일에 쓰기
                File.WriteAllText(yamlPath, mergedYamlContent);
                LogMessage($"YAML 파일 처리 완료: {yamlPath}");

                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"처리 중 오류 발생: {ex.Message}");
                LogMessage($"스택 추적: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 설정 문자열을 사용하여 YAML 파일을 처리합니다.
        /// </summary>
        /// <param name="yamlPath">처리할 YAML 파일 경로</param>
        /// <param name="mergeKeyPathsConfig">설정 문자열 (형식: "idPath|mergePaths|keyPaths|arrayFieldPaths")</param>
        /// <param name="includeEmptyFields">빈 필드 포함 여부 (기본값 false)</param>
        /// <returns>처리 성공 여부</returns>
        public static bool ProcessYamlFileFromConfig(string yamlPath, string mergeKeyPathsConfig, bool includeEmptyFields = false)
        {
            try
            {
                // 설정 문자열이 비어있어도 기본값을 사용하여 처리 진행
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 설정 문자열 파싱: {mergeKeyPathsConfig}");

                // 비어있는 경우에도 FromConfigString에서 기본값 적용
                YamlMergeKeyPathsProcessor processor = FromConfigString(mergeKeyPathsConfig, includeEmptyFields);
            
                string idPath = "";
                string mergePaths = "";
                string keyPaths = "";
                string arrayFieldPaths = "";
            
                if (!string.IsNullOrWhiteSpace(mergeKeyPathsConfig))
                {
                    string[] parts = mergeKeyPathsConfig.Split('|');
                    if (parts.Length >= 1)
                        idPath = parts[0];
                    if (parts.Length >= 2)
                        mergePaths = parts[1];
                    if (parts.Length >= 3)
                        keyPaths = parts[2];
                    if (parts.Length >= 4)
                        arrayFieldPaths = parts[3];
                }
                
                return processor.ProcessYamlFile(yamlPath, keyPaths);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlMergeKeyPathsProcessor] 설정 문자열 처리 중 오류: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 키 경로:전략 문자열을 파싱합니다.
        /// </summary>
        /// <param name="keyPaths">키 경로:전략 문자열 (예: "level:merge;achievement:append")</param>
        private void ParseKeyPaths(string keyPaths)
        {
            keyPathStrategies.Clear();
            if (string.IsNullOrWhiteSpace(keyPaths)) return;

            // 콤마(,)도 구분자로 지원
            char[] separators = new char[] { arrayPathSeparator, ',' };
            var pairs = keyPaths.Split(separators);
            
            foreach (var pair in pairs)
            {
                if (string.IsNullOrWhiteSpace(pair)) continue;
                
                var parts = pair.Split(strategySeparator);
                if (parts.Length == 2)
                {
                    var keyPath = parts[0].Trim();
                    var strategy = parts[1].Trim();
                    keyPathStrategies[keyPath] = strategy;
                    LogMessage($"키 경로 파싱: {keyPath} -> {strategy}");
                }
                else if (parts.Length == 1)
                {
                    // 전략이 없는 경우 기본값 'match' 사용
                    var keyPath = parts[0].Trim();
                    keyPathStrategies[keyPath] = "match";
                    LogMessage($"키 경로 파싱 (기본 전략 사용): {keyPath} -> match");
                }
            }
        }

        /// <summary>
        /// 배열 필드 경로:전략 문자열을 파싱합니다.
        /// </summary>
        /// <param name="arrayFieldPaths">배열 필드 경로:전략 문자열 (예: "results;append,items;replace")</param>
        private void ParseArrayFieldPaths(string arrayFieldPaths)
        {
            arrayFieldStrategies.Clear();
            if (string.IsNullOrWhiteSpace(arrayFieldPaths)) return;

            // 콤마(,)도 구분자로 지원
            char[] separators = new char[] { arrayPathSeparator, ',' };
            var pairs = arrayFieldPaths.Split(separators);
            
            foreach (var pair in pairs)
            {
                if (string.IsNullOrWhiteSpace(pair)) continue;
                
                var parts = pair.Split(strategySeparator);
                if (parts.Length == 2)
                {
                    var fieldPath = parts[0].Trim();
                    var strategy = parts[1].Trim();
                    arrayFieldStrategies[fieldPath] = strategy;
                    LogMessage($"배열 필드 경로 파싱: {fieldPath} -> {strategy}");
                }
                else if (parts.Length == 1)
                {
                    // 전략이 없는 경우 기본값 'append' 사용
                    var fieldPath = parts[0].Trim();
                    arrayFieldStrategies[fieldPath] = "append";
                    LogMessage($"배열 필드 경로 파싱 (기본 전략 사용): {fieldPath} -> append");
                }
            }
        }

        /// <summary>
        /// YAML 문자열을 JSON 배열로 변환합니다.
        /// </summary>
        /// <param name="yamlContent">YAML 문자열</param>
        /// <returns>JSON 배열</returns>
        private JArray ConvertYamlToJson(string yamlContent)
        {
            try
            {
                LogMessage("YAML을 JSON으로 변환 시작...");
                
                // YAML 역직렬화기 설정
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(namingConvention)
                    .Build();
                
                // YAML을 동적 객체로 역직렬화
                dynamic yamlObject;
                
                // YAML 내용이 배열로 시작하는지 객체로 시작하는지 확인
                if (yamlContent.TrimStart().StartsWith("-"))
                {
                    // 배열인 경우
                    yamlObject = deserializer.Deserialize<List<object>>(new StringReader(yamlContent));
                }
                else
                {
                    // 객체인 경우
                    yamlObject = deserializer.Deserialize<Dictionary<object, object>>(new StringReader(yamlContent));
                }
                
                // JSON으로 변환
                string jsonString = JsonConvert.SerializeObject(yamlObject, Formatting.None);
                
                if (debugMode)
                {
                    LogMessage($"변환된 JSON: {jsonString.Substring(0, Math.Min(100, jsonString.Length))}...");
                }
                
                // 변환된 JSON 문자열을 JArray로 파싱
                JArray resultArray;
                
                if (jsonString.StartsWith("{"))
                {
                    // 단일 객체인 경우 배열로 변환
                    var jObject = JObject.Parse(jsonString);
                    resultArray = new JArray();
                    resultArray.Add(jObject);
                }
                else
                {
                    resultArray = JArray.Parse(jsonString);
                }
                
                // 숫자 문자열을 실제 숫자로 변환
                ProcessNumericValues(resultArray);
                
                return resultArray;
            }
            catch (Exception ex)
            {
                LogMessage($"YAML을 JSON으로 변환 중 오류: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// JToken 내의 문자열로 된 숫자 값을 실제 숫자 타입으로 변환합니다.
        /// </summary>
        private void ProcessNumericValues(JToken token)
        {
            if (token is JObject jObject)
            {
                // 객체의 각 속성을 처리
                foreach (var property in jObject.Properties().ToList())
                {
                    if (property.Value is JValue jValue && jValue.Type == JTokenType.String)
                    {
                        string valueStr = jValue.Value<string>();
                        
                        // 숫자 문자열인지 확인
                        if (int.TryParse(valueStr, out int intValue))
                        {
                            property.Value = new JValue(intValue);
                        }
                        else if (double.TryParse(valueStr, out double doubleValue))
                        {
                            property.Value = new JValue(doubleValue);
                        }
                    }
                    else
                    {
                        // 중첩된 객체/배열 처리
                        ProcessNumericValues(property.Value);
                    }
                }
            }
            else if (token is JArray jArray)
            {
                // 배열의 각 항목을 처리
                for (int i = 0; i < jArray.Count; i++)
                {
                    if (jArray[i] is JValue jValue && jValue.Type == JTokenType.String)
                    {
                        string valueStr = jValue.Value<string>();
                        
                        // 숫자 문자열인지 확인
                        if (int.TryParse(valueStr, out int intValue))
                        {
                            jArray[i] = new JValue(intValue);
                        }
                        else if (double.TryParse(valueStr, out double doubleValue))
                        {
                            jArray[i] = new JValue(doubleValue);
                        }
                    }
                    else
                    {
                        // 중첩된 객체/배열 처리
                        ProcessNumericValues(jArray[i]);
                    }
                }
            }
        }

        /// <summary>
        /// JSON 배열을 YAML 문자열로 변환합니다.
        /// </summary>
        /// <param name="jsonArray">JSON 배열</param>
        /// <returns>YAML 문자열</returns>
        private string ConvertJsonToYaml(JArray jsonArray)
        {
            try
            {
                // JArray를 OrderedYamlFactory가 사용하는 YamlArray로 변환
                YamlArray rootArray = (YamlArray)JsonToOrderedYaml(jsonArray);
                
                // OrderedYamlFactory를 사용하여 YAML 문자열로 직렬화
                // preserveQuotes를 false로 설정하여 필요한 경우에만 따옴표 사용
                // includeEmptyFields 옵션을 전달하여 빈 필드 포함 여부 제어
                return OrderedYamlFactory.SerializeToYaml(rootArray, 2, YamlStyle.Block, false, includeEmptyFields);
            }
            catch (Exception ex)
            {
                LogMessage($"JSON을 YAML로 변환하는 중 오류 발생: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// JSON 객체를 OrderedYaml 객체로 변환합니다.
        /// </summary>
        private object JsonToOrderedYaml(JToken token)
        {
            switch (token.Type)
            {
                case JTokenType.Object:
                    var yamlObj = OrderedYamlFactory.CreateObject();
                    foreach (JProperty prop in token.Children<JProperty>())
                    {
                        yamlObj.Add(prop.Name, JsonToOrderedYaml(prop.Value));
                    }
                    return yamlObj;

                case JTokenType.Array:
                    var yamlArr = OrderedYamlFactory.CreateArray();
                    foreach (JToken item in token)
                    {
                        yamlArr.Add(JsonToOrderedYaml(item));
                    }
                    return yamlArr;

                case JTokenType.String:
                    string stringValue = token.Value<string>();
                    
                    // 빈 문자열이거나 특수 문자가 포함된 경우에만 문자열로 처리
                    // 일반 텍스트는 따옴표 없이 그대로 반환하기 위해 YamlScalarValue 사용
                    if (ShouldQuoteString(stringValue))
                    {
                        return stringValue; // OrderedYamlFactory가 알아서 따옴표 처리
                    }
                    else
                    {
                        // 특수 문자가 없는 일반 텍스트는 YamlScalarValue로 감싸서 반환
                        // 이렇게 하면 OrderedYamlFactory가 따옴표를 붙이지 않음
                        return new YamlScalarValue(stringValue);
                    }

                case JTokenType.Integer:
                    return token.Value<long>();

                case JTokenType.Float:
                    return token.Value<double>();

                case JTokenType.Boolean:
                    return token.Value<bool>();

                case JTokenType.Null:
                    return null;

                default:
                    // 그 외 타입은 문자열로 변환
                    return token.ToString();
            }
        }

        /// <summary>
        /// 문자열에 따옴표가 필요한지 확인합니다.
        /// </summary>
        private bool ShouldQuoteString(string value)
        {
            if (string.IsNullOrEmpty(value))
                return true; // 빈 문자열은 따옴표 필요
            
            // 숫자처럼 보이는 문자열 체크 제거 (숫자는 JToken에서 이미 적절한 타입으로 처리됨)
            
            // true/false처럼 보이는 문자열이면 따옴표 필요
            if (bool.TryParse(value, out _))
                return true;
            
            // null 또는 ~처럼 보이는 문자열이면 따옴표 필요
            if (value.Equals("null", StringComparison.OrdinalIgnoreCase) || 
                value.Equals("~"))
                return true;
            
            // 특수 문자가 포함된 경우 따옴표 필요
            if (value.IndexOfAny(new[] { ':', '{', '}', '[', ']', ',', '&', '*', '#', '?', '|', '-', '<', '>', '=', '!', '%', '@', '\\', '\n', '\r', '\t' }) >= 0)
                return true;
            
            // 공백으로 시작하거나 끝나면 따옴표 필요
            if (value.StartsWith(" ") || value.EndsWith(" "))
                return true;
            
            return false;
        }

        /// <summary>
        /// JSON 배열의 항목들을 ID와 키 경로를 기반으로 병합합니다.
        /// </summary>
        /// <param name="jsonArray">병합할 JSON 배열</param>
        /// <returns>병합된 JSON 배열</returns>
        private JArray MergeJsonItems(JArray jsonArray)
        {
            try
            {
                // ID별로 항목 그룹화
                var groupedItems = new Dictionary<string, List<JObject>>();
                
                foreach (JObject item in jsonArray)
                {
                    // ID 값 추출
                    JToken idToken = null;
                    
                    if (!string.IsNullOrEmpty(idPath))
                    {
                        idToken = item.SelectToken(idPath);
                    }
                    
                    string id = idToken?.ToString() ?? "default";
                    
                    if (!groupedItems.ContainsKey(id))
                    {
                        groupedItems[id] = new List<JObject>();
                    }
                    
                    groupedItems[id].Add(item);
                }
                
                LogMessage($"고유 ID 개수: {groupedItems.Count}");
                
                // 결과 배열 생성
                var result = new JArray();
                
                // 각 ID 그룹에 대해 병합 처리
                foreach (var group in groupedItems)
                {
                    string id = group.Key;
                    List<JObject> items = group.Value;
                    
                    LogMessage($"ID '{id}'에 대한 항목 수: {items.Count}");
                    
                    if (items.Count == 1)
                    {
                        // 항목이 하나뿐이면 그대로 추가
                        result.Add(items[0]);
                        continue;
                    }
                    
                    // 병합을 위한 기본 항목 생성
                    JObject mergedItem = new JObject();
                    
                    // ID 속성 추가
                    if (!string.IsNullOrEmpty(idPath))
                    {
                        JToken idToken = items[0].SelectToken(idPath);
                        SetPropertyByPath(mergedItem, idPath, idToken);
                    }
                    
                    // 첫 번째 항목의 모든 속성을 기본값으로 복사
                    foreach (var property in items[0].Properties())
                    {
                        string propertyPath = property.Name;
                        if (propertyPath == idPath) continue; // ID는 이미 처리했으므로 건너뜀
                        
                        mergedItem[property.Name] = property.Value.DeepClone();
                    }
                    
                    // 나머지 항목의 속성들을 병합
                    for (int i = 1; i < items.Count; i++)
                    {
                        MergeProperties(mergedItem, items[i]);
                    }
                    
                    // 결과에 추가
                    result.Add(mergedItem);
                }
                
                return result;
            }
            catch (Exception ex)
            {
                LogMessage($"JSON 항목 병합 중 오류: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 두 JSON 객체의 속성들을 병합합니다.
        /// </summary>
        /// <param name="target">대상 JSON 객체</param>
        /// <param name="source">소스 JSON 객체</param>
        private void MergeProperties(JObject target, JObject source)
        {
            // 특별히 병합할 경로들만 처리
            if (mergePaths != null && mergePaths.Length > 0)
            {
                foreach (string mergePath in mergePaths)
                {
                    if (string.IsNullOrWhiteSpace(mergePath)) continue;
                    
                    JToken sourceValue = source.SelectToken(mergePath);
                    JToken targetValue = target.SelectToken(mergePath);
                    
                    if (sourceValue == null) continue;
                    
                    if (targetValue == null)
                    {
                        // 대상에 해당 경로가 없으면 새로 추가
                        SetPropertyByPath(target, mergePath, sourceValue.DeepClone());
                    }
                    else if (sourceValue.Type == JTokenType.Array && targetValue.Type == JTokenType.Array)
                    {
                        // 배열인 경우 키 경로를 기준으로 항목들 병합
                        JArray genericSourceArray = (JArray)sourceValue;
                        JArray genericTargetArray = (JArray)targetValue;
                        
                        foreach (JToken sourceItem in genericSourceArray)
                        {
                            bool shouldAdd = true;
                            
                            // 키 경로가 있으면 항목들을 비교하여 중복 확인
                            if (keyPathStrategies.Count > 0)
                            {
                                foreach (JToken targetItem in genericTargetArray)
                                {
                                    bool allKeysMatch = true;
                                    
                                    foreach (var keyPathStrategy in keyPathStrategies)
                                    {
                                        string keyPath = keyPathStrategy.Key;
                                        string strategy = keyPathStrategy.Value;
                                        
                                        JToken sourceKeyValue = sourceItem is JObject sourceObj ? sourceObj.SelectToken(keyPath) : null;
                                        JToken targetKeyValue = targetItem is JObject targetObj ? targetObj.SelectToken(keyPath) : null;
                                        
                                        if (sourceKeyValue == null || targetKeyValue == null)
                                        {
                                            allKeysMatch = false;
                                            break;
                                        }
                                        
                                        string sourceKeyStr = sourceKeyValue.ToString();
                                        string targetKeyStr = targetKeyValue.ToString();
                                        
                                        if (strategy.Equals("match", StringComparison.OrdinalIgnoreCase))
                                        {
                                            // 값이 일치하는지 확인
                                            if (!sourceKeyStr.Equals(targetKeyStr, StringComparison.Ordinal))
                                            {
                                                allKeysMatch = false;
                                                break;
                                            }
                                        }
                                        else if (strategy.Equals("contains", StringComparison.OrdinalIgnoreCase))
                                        {
                                            // 값이 포함되는지 확인
                                            if (!targetKeyStr.Contains(sourceKeyStr))
                                            {
                                                allKeysMatch = false;
                                                break;
                                            }
                                        }
                                    }
                                    
                                    if (allKeysMatch)
                                    {
                                        // 모든 키가 일치하면 해당 항목을 추가하지 않거나 업데이트
                                        shouldAdd = false;
                                        
                                        // 전략이 'update'이면 해당 항목을 업데이트
                                        if (keyPathStrategies.Any(x => x.Value.Equals("update", StringComparison.OrdinalIgnoreCase)))
                                        {
                                            if (targetItem is JObject targetObj && sourceItem is JObject sourceObj)
                                            {
                                                foreach (var property in sourceObj.Properties())
                                                {
                                                    targetObj[property.Name] = property.Value.DeepClone();
                                                }
                                            }
                                        }
                                        // 배열 필드 전략 처리
                                        else if (targetItem is JObject targetObj && sourceItem is JObject sourceObj)
                                        {
                                            // 각 배열 필드에 대해 전략 적용
                                            foreach (var arrayFieldStrategy in arrayFieldStrategies)
                                            {
                                                string fieldPath = arrayFieldStrategy.Key;
                                                string strategy = arrayFieldStrategy.Value;
                                                
                                                JToken sourceFieldToken = sourceObj.SelectToken(fieldPath);
                                                JToken targetFieldToken = targetObj.SelectToken(fieldPath);
                                                
                                                // 소스 필드가 null이거나 배열이 아니면 건너뜀
                                                if (sourceFieldToken == null || !(sourceFieldToken is JArray sourceArrayItems)) 
                                                    continue;
                                                    
                                                JArray targetArrayItems;
                                                if (targetFieldToken == null || !(targetFieldToken is JArray))
                                                {
                                                    // 타겟 필드가 없거나 배열이 아니면 생성
                                                    targetArrayItems = new JArray();
                                                    SetPropertyByPath(targetObj, fieldPath, targetArrayItems);
                            }
                            else
                            {
                                                    targetArrayItems = (JArray)targetFieldToken;
                                                }
                                                
                                                LogMessage($"{fieldPath} 배열 필드 병합 시작: 전략={strategy}, 기존={targetArrayItems.Count}개, 추가={sourceArrayItems.Count}개");
                                                
                                                if (strategy.Equals("append", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    // 배열 항목 추가
                                                    foreach (var item in sourceArrayItems)
                                                    {
                                                        targetArrayItems.Add(item.DeepClone());
                                                    }
                                                }
                                                else if (strategy.Equals("replace", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    // 배열 대체
                                                    targetArrayItems.Replace(sourceArrayItems.DeepClone());
                                                }
                                                else if (strategy.Equals("prepend", StringComparison.OrdinalIgnoreCase))
                                                {
                                                    // 배열 앞에 추가
                                                    JArray newArray = new JArray();
                                                    foreach (var item in sourceArrayItems)
                                                    {
                                                        newArray.Add(item.DeepClone());
                                                    }
                                                    foreach (var item in targetArrayItems)
                                                    {
                                                        newArray.Add(item.DeepClone());
                                                    }
                                                    targetArrayItems.Replace(newArray);
                                                }
                                                
                                                LogMessage($"{fieldPath} 배열 필드 병합 완료: 병합 후 {targetArrayItems.Count}개");
                                            }
                                        }
                                        
                                        break;
                                    }
                                }
                            }
                            
                            if (shouldAdd)
                            {
                                // 항목 추가
                                genericTargetArray.Add(sourceItem.DeepClone());
                            }
                        }
                        }
                        else
                        {
                        // 배열이 아닌 경우 (단일 값이나 객체) 소스 값으로 대체
                        SetPropertyByPath(target, mergePath, sourceValue.DeepClone());
                    }
                }
            }
            else
            {
                // 병합 경로가 지정되지 않은 경우 모든 속성을 대상에 추가/업데이트
                foreach (var property in source.Properties())
                {
                    string propertyName = property.Name;
                    JToken sourceValue = property.Value;
                    
                    if (target[propertyName] == null)
                    {
                        // 대상에 해당 속성이 없으면 새로 추가
                        target[propertyName] = sourceValue.DeepClone();
                    }
                    else if (target[propertyName].Type == JTokenType.Object && sourceValue.Type == JTokenType.Object)
                    {
                        // 객체인 경우 재귀적으로 병합
                        MergeProperties((JObject)target[propertyName], (JObject)sourceValue);
                    }
                    else if (target[propertyName].Type == JTokenType.Array && sourceValue.Type == JTokenType.Array)
                    {
                        // 배열인 경우 항목들을 추가
                        JArray genericTargetArray = (JArray)target[propertyName];
                        JArray genericSourceArray = (JArray)sourceValue;
                        
                        foreach (JToken item in genericSourceArray)
                        {
                            genericTargetArray.Add(item.DeepClone());
                        }
                    }
                    else
                    {
                        // 다른 타입은 소스 값으로 대체
                        target[propertyName] = sourceValue.DeepClone();
                    }
                }
            }
        }
        
        /// <summary>
        /// 경로를 기준으로 JSON 객체에 속성을 설정합니다.
        /// </summary>
        /// <param name="jObject">JSON 객체</param>
        /// <param name="path">경로</param>
        /// <param name="value">설정할 값</param>
        private void SetPropertyByPath(JObject jObject, string path, JToken value)
        {
            if (string.IsNullOrEmpty(path))
            {
                return;
            }
            
            var parts = path.Split(keyPathSeparator);
            
            if (parts.Length == 1)
            {
                // 단일 속성 처리
                jObject[parts[0]] = value;
            }
            else
            {
                // 중첩 속성 처리
                JObject current = jObject;
                
                for (int i = 0; i < parts.Length - 1; i++)
                {
                    var part = parts[i];
                    
                    if (current[part] == null || current[part].Type != JTokenType.Object)
                    {
                        current[part] = new JObject();
                    }
                    
                    current = (JObject)current[part];
                }
                
                current[parts[parts.Length - 1]] = value;
            }
        }
        
        /// <summary>
        /// 디버그 메시지를 출력합니다.
        /// </summary>
        /// <param name="message">출력할 메시지</param>
        private void LogMessage(string message)
        {
            Debug.WriteLine($"[YamlMergeKeyPathsProcessor] {message}");
        }
    }
} 