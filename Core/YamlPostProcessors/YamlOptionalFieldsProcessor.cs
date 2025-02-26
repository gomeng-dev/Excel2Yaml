using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Diagnostics;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace ExcelToJsonAddin.Core.YamlPostProcessors
{
    /// <summary>
    /// YAML 파일에서 선택적 필드를 처리하는 클래스입니다.
    /// 모든 아이템에 모든 필드를 포함하도록 처리합니다.
    /// </summary>
    public class YamlOptionalFieldsProcessor
    {
        /// <summary>
        /// YAML 파일을 처리하여 모든 항목에 모든 필드를 포함시킵니다.
        /// </summary>
        /// <param name="yamlPath">처리할 YAML 파일 경로</param>
        /// <returns>처리 성공 여부</returns>
        public bool ProcessYamlFile(string yamlPath)
        {
            try
            {
                Debug.WriteLine($"[YamlOptionalFieldsProcessor] YAML 파일 처리: {yamlPath}");

                if (!File.Exists(yamlPath))
                {
                    Debug.WriteLine($"[YamlOptionalFieldsProcessor] 오류: YAML 파일을 찾을 수 없습니다.");
                    return false;
                }

                // YAML 파일 읽기
                string yamlContent = File.ReadAllText(yamlPath);
                
                if (string.IsNullOrWhiteSpace(yamlContent))
                {
                    Debug.WriteLine($"[YamlOptionalFieldsProcessor] 오류: YAML 파일이 비어 있습니다.");
                    return false;
                }

                // YAML 역직렬화
                var deserializer = new DeserializerBuilder()
                    .WithNamingConvention(CamelCaseNamingConvention.Instance)
                    .Build();

                // YAML 내용을 객체로 변환
                object yamlObject = deserializer.Deserialize<object>(yamlContent);

                if (yamlObject == null)
                {
                    Debug.WriteLine($"[YamlOptionalFieldsProcessor] 오류: YAML을 객체로 변환할 수 없습니다.");
                    return false;
                }

                // 루트 리스트 처리
                if (yamlObject is List<object> rootList)
                {
                    Debug.WriteLine($"[YamlOptionalFieldsProcessor] 루트 리스트 처리: {rootList.Count}개 항목");
                    
                    // 모든 필드 수집
                    Dictionary<string, object> allFields = CollectAllFields(rootList);
                    Debug.WriteLine($"[YamlOptionalFieldsProcessor] 수집된 필드: {allFields.Count}개");

                    // 모든 항목에 없는 필드 추가
                    ApplyFieldsToAllItems(rootList, allFields);
                }
                else
                {
                    Debug.WriteLine($"[YamlOptionalFieldsProcessor] 오류: 루트가 리스트가 아닙니다. 타입: {yamlObject.GetType().Name}");
                    return false;
                }

                // 직렬화 옵션 설정
                var serializer = new SerializerBuilder()
                    .WithNamingConvention(CamelCaseNamingConvention.Instance)
                    .ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitNull)
                    .Build();

                // 수정된 객체를 YAML로 직렬화
                string modifiedYaml = serializer.Serialize(yamlObject);

                // 파일 저장
                File.WriteAllText(yamlPath, modifiedYaml);
                Debug.WriteLine($"[YamlOptionalFieldsProcessor] YAML 파일 처리 완료: {yamlPath}");

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlOptionalFieldsProcessor] 처리 중 오류 발생: {ex.Message}");
                Debug.WriteLine($"[YamlOptionalFieldsProcessor] 스택 추적: {ex.StackTrace}");
                return false;
            }
        }

        /// <summary>
        /// 모든 항목에서 사용된 모든 필드를 수집합니다.
        /// </summary>
        /// <param name="items">항목 리스트</param>
        /// <returns>모든 필드와 기본값 딕셔너리</returns>
        private Dictionary<string, object> CollectAllFields(List<object> items)
        {
            Dictionary<string, object> allFields = new Dictionary<string, object>();

            foreach (var item in items)
            {
                if (item is Dictionary<object, object> itemDict)
                {
                    foreach (var kvp in itemDict)
                    {
                        string key = kvp.Key.ToString();
                        if (!allFields.ContainsKey(key))
                        {
                            // 기본 값으로 null 사용
                            allFields.Add(key, null);
                        }
                    }
                }
            }

            return allFields;
        }

        /// <summary>
        /// 모든 항목에 수집된 모든 필드를 적용합니다.
        /// </summary>
        /// <param name="items">항목 리스트</param>
        /// <param name="allFields">모든 필드와 기본값 딕셔너리</param>
        private void ApplyFieldsToAllItems(List<object> items, Dictionary<string, object> allFields)
        {
            for (int i = 0; i < items.Count; i++)
            {
                if (items[i] is Dictionary<object, object> itemDict)
                {
                    // 모든 필드를 현재 항목에 추가
                    foreach (var field in allFields)
                    {
                        if (!itemDict.ContainsKey(field.Key))
                        {
                            itemDict[field.Key] = field.Value;
                        }
                    }
                }
                else if (items[i] == null)
                {
                    // null 항목은 새 딕셔너리로 교체하고 모든 필드 추가
                    var newDict = new Dictionary<object, object>();
                    foreach (var field in allFields)
                    {
                        newDict[field.Key] = field.Value;
                    }
                    items[i] = newDict;
                }
            }
        }
    }
} 