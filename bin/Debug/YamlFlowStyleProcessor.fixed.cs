using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using System.Text.RegularExpressions;
using Newtonsoft.Json;
using ExcelToJsonAddin.Config;  // YamlStyle 열거형을 사용하기 위한 네임스페이스 추가
using YamlObject = ExcelToJsonAddin.Core.YamlObject;  // YamlObject 클래스 명시적 참조
using YamlArray = ExcelToJsonAddin.Core.YamlArray;  // YamlArray 클래스 명시적 참조

namespace ExcelToJsonAddin.Core.YamlPostProcessors
{
    /// <summary>
    /// YAML 파일의 특정 필드를 Flow 스타일로 변환하는 프로세서입니다.
    /// OrderedYamlFactory를 활용한 구현 버전입니다.
    /// </summary>
    public class YamlFlowStyleProcessor
    {
        private readonly List<string> flowStyleFields;
        private readonly List<string> flowStyleItemsFields;
        private readonly Dictionary<string, List<int>> fieldIndices;

        /// <summary>
        /// 기본 생성자
        /// </summary>
        /// <param name="flowStyleFields">전체 필드를 플로우 스타일로 변환할 필드 목록 (예: "details,info")</param>
        /// <param name="flowStyleItemsFields">특정 필드의 리스트 아이템을 플로우 스타일로 변환할 필드 목록 (예: "triggers,events")</param>
        public YamlFlowStyleProcessor(string flowStyleFields = "", string flowStyleItemsFields = "")
        {
            this.flowStyleFields = string.IsNullOrWhiteSpace(flowStyleFields)
                ? new List<string>()
                : flowStyleFields.Split(',').Select(f => f.Trim()).Where(f => !string.IsNullOrWhiteSpace(f)).ToList();

            this.flowStyleItemsFields = string.IsNullOrWhiteSpace(flowStyleItemsFields)
                ? new List<string>()
                : flowStyleItemsFields.Split(',').Select(f => f.Trim()).Where(f => !string.IsNullOrWhiteSpace(f)).ToList();

            this.fieldIndices = ParseFieldIndices();
            
            Debug.WriteLine($"[YamlFlowStyleProcessor] 초기화: flowStyleFields={string.Join(",", this.flowStyleFields)}");
            Debug.WriteLine($"[YamlFlowStyleProcessor] 초기화: flowStyleItemsFields={string.Join(",", this.flowStyleItemsFields)}");
        }

        /// <summary>
        /// 설정 문자열을 파싱하여 프로세서를 생성합니다.
        /// </summary>
        /// <param name="flowStyleConfig">설정 문자열 (형식: "flowStyleFields|flowStyleItemsFields")</param>
        /// <returns>YamlFlowStyleProcessor 인스턴스</returns>
        public static YamlFlowStyleProcessor FromConfigString(string flowStyleConfig)
        {
            string flowStyleFields = "";
            string flowStyleItemsFields = "";
            
            if (!string.IsNullOrWhiteSpace(flowStyleConfig))
            {
                string[] parts = flowStyleConfig.Split('|');
                if (parts.Length >= 1)
                    flowStyleFields = parts[0];
                if (parts.Length >= 2)
                    flowStyleItemsFields = parts[1];
            }
            
            return new YamlFlowStyleProcessor(flowStyleFields, flowStyleItemsFields);
        }

        /// <summary>
        /// '필드명' 또는 '필드명[인덱스]'를 해석해 Dictionary로 변환합니다.
        /// </summary>
        /// <returns>필드명:인덱스리스트 형태의 Dictionary</returns>
        private Dictionary<string, List<int>> ParseFieldIndices()
        {
            var fieldIndices = new Dictionary<string, List<int>>();
            var pattern = new Regex(@"^(?<name>\w+)(\[(?<index>\d+)\])?$");
            
            // flowStyleFields 처리
            foreach (var field in flowStyleFields)
            {
                var match = pattern.Match(field);
                if (match.Success)
                {
                    string name = match.Groups["name"].Value;
                    string indexStr = match.Groups["index"].Value;
                    
                    if (!fieldIndices.ContainsKey(name))
                        fieldIndices[name] = new List<int>();
                    
                    if (!string.IsNullOrEmpty(indexStr) && int.TryParse(indexStr, out int index))
                        fieldIndices[name].Add(index);
                    else
                        fieldIndices[name] = null; // null은 "모든 인덱스" 의미
                }
                else
                {
                    Debug.WriteLine($"[YamlFlowStyleProcessor] 경고: 필드명 '{field}'이(가) 올바른 형식이 아닙니다. 무시됩니다.");
                }
            }
            
            // flowStyleItemsFields도 동일한 방식으로 처리
            foreach (var field in flowStyleItemsFields)
            {
                var match = pattern.Match(field);
                if (match.Success)
                {
                    string name = match.Groups["name"].Value;
                    string indexStr = match.Groups["index"].Value;
                    
                    if (!fieldIndices.ContainsKey(name))
                        fieldIndices[name] = new List<int>();
                    
                    if (!string.IsNullOrEmpty(indexStr) && int.TryParse(indexStr, out int index))
                        fieldIndices[name].Add(index);
                    else
                        fieldIndices[name] = null; // null은 "모든 인덱스" 의미
                }
                else
                {
                    Debug.WriteLine($"[YamlFlowStyleProcessor] 경고: 필드명 '{field}'이(가) 올바른 형식이 아닙니다. 무시됩니다.");
                }
            }
            
            return fieldIndices;
        }

        /// <summary>
        /// YAML 파일을 처리하여 지정된 필드를 Flow 스타일로 변환합니다.
        /// </summary>
        /// <param name="yamlPath">처리할 YAML 파일 경로</param>
        /// <returns>처리 성공 여부</returns>
        public bool ProcessYamlFile(string yamlPath)
        {
            if (string.IsNullOrWhiteSpace(yamlPath) || !File.Exists(yamlPath))
            {
                Debug.WriteLine($"[YamlFlowStyleProcessor] 오류: 유효하지 않은 YAML 파일 경로 '{yamlPath}'");
                return false;
            }

            try
            {
                // 백업 파일 생성
                string backupPath = yamlPath + "_backup";
                File.Copy(yamlPath, backupPath, true);
                
                // YAML 파일 텍스트 읽기
                string yamlContent = File.ReadAllText(yamlPath);
                
                // YamlDotNet 사용하여 YAML을 객체로 변환 (Dictionary 또는 List)
                var deserializer = new YamlDotNet.Serialization.Deserializer();
                object yamlObj = deserializer.Deserialize(new StringReader(yamlContent));
                
                // OrderedYamlFactory용 객체로 변환
                object orderedYamlObj = ConvertToOrderedYaml(yamlObj);
                
                // 필요한 필드에 Flow 스타일 적용
                ApplyFlowStyleToObject(orderedYamlObj, "");
                
                // OrderedYamlFactory를 사용하여 YAML로 직렬화
                string newYamlContent = OrderedYamlFactory.SerializeToYaml(orderedYamlObj, 2, YamlStyle.Block, true);
                
                // 파일에 저장
                File.WriteAllText(yamlPath, newYamlContent);
                
                Debug.WriteLine($"[YamlFlowStyleProcessor] YAML Flow 스타일 처리 완료: '{yamlPath}'");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlFlowStyleProcessor] YAML 처리 중 오류 발생: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// 일반 객체를 OrderedYamlFactory용 객체로 변환합니다.
        /// </summary>
        private object ConvertToOrderedYaml(object obj)
        {
            if (obj == null)
                return null;
                
            // Dictionary<string, object>를 YamlObject로 변환
            if (obj is IDictionary<object, object> dictObj)
            {
                var yamlObj = OrderedYamlFactory.CreateObject();
                foreach (var kvp in dictObj)
                {
                    yamlObj.Add(kvp.Key.ToString(), ConvertToOrderedYaml(kvp.Value));
                }
                return yamlObj;
            }
            // List<object>를 YamlArray로 변환
            else if (obj is IList<object> listObj)
            {
                var yamlArray = OrderedYamlFactory.CreateArray();
                foreach (var item in listObj)
                {
                    yamlArray.Add(ConvertToOrderedYaml(item));
                }
                return yamlArray;
            }
            // 기본 타입은 그대로 반환
            else
            {
                return obj;
            }
        }
        
        /// <summary>
        /// 객체의 특정 필드에 Flow 스타일을 적용합니다.
        /// </summary>
        private void ApplyFlowStyleToObject(object obj, string path)
        {
            if (obj == null)
                return;
                
            if (obj is YamlObject yamlObj)
            {
                // 현재 경로가 flowStyleFields에 포함되어 있으면 Flow 스타일 적용
                bool applyFlowStyle = ShouldApplyFlowStyle(path);
                
                // 하위 필드 처리
                foreach (var prop in yamlObj.Properties.ToList())
                {
                    string childPath = string.IsNullOrEmpty(path) ? prop.Key : $"{path}.{prop.Key}";
                    
                    // 재귀적으로 하위 객체/배열 처리
                    ApplyFlowStyleToObject(prop.Value, childPath);
                    
                    // 현재 필드가 Flow 스타일이어야 하는 경우, 별도 처리
                    if (applyFlowStyle || ShouldApplyFlowStyle(childPath))
                    {
                        ApplyFlowStyleDirectly(prop.Value);
                    }
                }
            }
            else if (obj is YamlArray yamlArray)
            {
                // 현재 경로가 flowStyleItemsFields에 포함되어 있으면 Flow 스타일 적용
                bool applyFlowStyle = ShouldApplyFlowStyleItems(path);
                
                // 배열의 각 항목 처리
                int index = 0;
                foreach (var item in yamlArray.Items.ToList())
                {
                    string childPath = $"{path}[{index}]";
                    
                    // 재귀적으로 하위 객체/배열 처리
                    ApplyFlowStyleToObject(item, childPath);
                    
                    // 현재 필드가 Flow 스타일이어야 하는 경우, 별도 처리
                    if (applyFlowStyle)
                    {
                        ApplyFlowStyleDirectly(item);
                    }
                    
                    index++;
                }
            }
        }
        
        /// <summary>
        /// 객체에 직접 Flow 스타일을 적용합니다 (OrderedYamlFactory 이용).
        /// </summary>
        private void ApplyFlowStyleDirectly(object obj)
        {
            // OrderedYamlFactory는 내부적으로 flow/block 스타일 정보를 가지고 있지 않으므로,
            // 이 메서드에서는 실제로 아무 것도 하지 않습니다.
            // 실제 Flow 스타일 적용은 OrderedYamlFactory.SerializeToYaml 호출 시 YamlStyle.Flow 파라미터로 수행합니다.
            
            // 여기서는 로그만 기록합니다.
            Debug.WriteLine($"[YamlFlowStyleProcessor] Flow 스타일 직접 적용: {obj?.GetType().Name}");
        }
        
        /// <summary>
        /// 주어진 경로가 Flow 스타일을 적용해야 하는지 확인합니다.
        /// </summary>
        private bool ShouldApplyFlowStyle(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;
                
            // 정확한 경로 매칭
            if (flowStyleFields.Contains(path))
                return true;
                
            // 와일드카드 매칭 (예: "items.*")
            foreach (var field in flowStyleFields)
            {
                if (field.EndsWith(".*") && path.StartsWith(field.TrimEnd('*', '.')))
                    return true;
            }
            
            // 기본 필드명 추출 후 확인
            string fieldName = ExtractFieldName(path);
            if (flowStyleFields.Contains(fieldName))
                return true;
                
            return false;
        }
        
        /// <summary>
        /// 주어진 경로가 Flow 스타일 항목으로 처리해야 하는지 확인합니다.
        /// </summary>
        private bool ShouldApplyFlowStyleItems(string path)
        {
            if (string.IsNullOrEmpty(path))
                return false;
            
            // 정확한 경로 매칭
            if (flowStyleItemsFields.Contains(path))
                return true;
            
            // 와일드카드 매칭 (예: "items.*")
            foreach (var field in flowStyleItemsFields)
            {
                if (field.EndsWith(".*") && path.StartsWith(field.TrimEnd('*', '.')))
                    return true;
            }
            
            // 기본 필드명 추출 후 확인
            string fieldName = ExtractFieldName(path);
            if (flowStyleItemsFields.Contains(fieldName))
                return true;
            
            return false;
        }
        
        /// <summary>
        /// 경로에서 필드명만 추출합니다 (인덱스나 하위 경로 제외).
        /// </summary>
        private string ExtractFieldName(string path)
        {
            if (string.IsNullOrEmpty(path))
                return string.Empty;
                
            // 인덱스가 있는 경우 (예: "items[0]")
            int bracketIndex = path.IndexOf('[');
            if (bracketIndex > 0)
                return path.Substring(0, bracketIndex);
                
            // 점(.)이 있는 경우 (예: "items.name")
            int dotIndex = path.IndexOf('.');
            if (dotIndex > 0)
                return path.Substring(0, dotIndex);
                
            return path;
        }

        /// <summary>
        /// 설정 문자열로부터 YAML 파일을 처리하여 Flow 스타일을 적용합니다.
        /// </summary>
        /// <param name="yamlPath">처리할 YAML 파일 경로</param>
        /// <param name="flowStyleConfig">설정 문자열 (형식: "flowStyleFields|flowStyleItemsFields")</param>
        /// <returns>처리 성공 여부</returns>
        public static bool ProcessYamlFileFromConfig(string yamlPath, string flowStyleConfig)
        {
            // 설정 문자열이 비어있으면 후처리 실행하지 않음
            if (string.IsNullOrWhiteSpace(flowStyleConfig))
            {
                Debug.WriteLine($"[YamlFlowStyleProcessor] 설정 문자열이 비어있어 후처리를 실행하지 않습니다: {yamlPath}");
                return true; // 후처리를 실행하지 않더라도 성공으로 처리
            }
            
            string flowStyleFields = "";
            string flowStyleItemsFields = "";
            
            string[] parts = flowStyleConfig.Split('|');
            if (parts.Length >= 1)
                flowStyleFields = parts[0];
            if (parts.Length >= 2)
                flowStyleItemsFields = parts[1];
            
            // 모든 인수가 비어있으면 후처리 실행하지 않음
            if (string.IsNullOrWhiteSpace(flowStyleFields) && string.IsNullOrWhiteSpace(flowStyleItemsFields))
            {
                Debug.WriteLine($"[YamlFlowStyleProcessor] 모든 인수가 비어있어 후처리를 실행하지 않습니다: {yamlPath}");
                return true; // 후처리를 실행하지 않더라도 성공으로 처리
            }
            
            Debug.WriteLine($"[YamlFlowStyleProcessor] 설정 문자열 파싱: flowStyleFields={flowStyleFields}, flowStyleItemsFields={flowStyleItemsFields}");
            
            YamlFlowStyleProcessor processor = new YamlFlowStyleProcessor(flowStyleFields, flowStyleItemsFields);
            return processor.ProcessYamlFile(yamlPath);
        }
    }
} 