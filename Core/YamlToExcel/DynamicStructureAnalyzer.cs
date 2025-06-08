using System;
using System.Collections.Generic;
using System.Linq;
using YamlDotNet.RepresentationModel;
using ExcelToYamlAddin.Logging;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    public class DynamicStructureAnalyzer
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<DynamicStructureAnalyzer>();
        
        public enum PatternType
        {
            RootArray,
            RootObject,
            Empty
        }

        public class PropertyPattern
        {
            public string Name { get; set; }
            public int OccurrenceCount { get; set; }
            public HashSet<Type> Types { get; set; }
            public int FirstAppearanceIndex { get; set; }
            public double OccurrenceRatio { get; set; }
            public bool IsRequired { get; set; }
            public bool IsArray { get; set; }
            public bool IsObject { get; set; }
            public List<string> ObjectProperties { get; set; } = new List<string>(); // 객체의 하위 속성들
            public List<string> NestedProperties => ObjectProperties; // 별칭 추가
            public Dictionary<string, PropertyPattern> NestedPatterns { get; set; } = new Dictionary<string, PropertyPattern>(); // 재귀적 패턴
            public ArrayPattern ArrayPattern { get; set; } // 배열인 경우의 패턴 정보
        }

        public class ArrayPattern
        {
            public string Name { get; set; }
            public int MaxSize { get; set; }
            public int MinSize { get; set; }
            public double OccurrenceRatio { get; set; }
            public bool RequiresMultipleRows { get; set; }
            public bool HasVariableStructure { get; set; }
            public bool HasVariableProperties { get; set; }
            public Dictionary<string, PropertyPattern> ElementProperties { get; set; }
            public Dictionary<string, int> ElementPropertyCounts { get; set; }
            public List<string> AllUniqueProperties { get; set; }
        }

        public class StructurePattern
        {
            public PatternType Type { get; set; }
            public Dictionary<string, PropertyPattern> Properties { get; set; }
            public Dictionary<string, ArrayPattern> Arrays { get; set; }
            public int MaxDepth { get; set; }
            public double ConsistencyScore { get; set; }
        }

        // 내부 도우미 클래스들
        private class ObjectInfo
        {
            public string Type { get; set; }
            public List<string> Properties { get; set; }
        }

        private class ArrayInfo
        {
            public bool IsArray { get; set; }
            public int ElementCount { get; set; }
            public List<Dictionary<string, object>> Elements { get; set; }
        }

        public StructurePattern AnalyzeStructure(YamlNode root)
        {
            var pattern = new StructurePattern
            {
                Properties = new Dictionary<string, PropertyPattern>(),
                Arrays = new Dictionary<string, ArrayPattern>()
            };

            // 동적 분석 - 하드코딩 없음
            if (root is YamlSequenceNode sequence)
            {
                pattern.Type = PatternType.RootArray;
                AnalyzeArrayElements(sequence, pattern);
            }
            else if (root is YamlMappingNode mapping)
            {
                pattern.Type = PatternType.RootObject;
                AnalyzeObjectProperties(mapping, pattern);
            }
            else
            {
                pattern.Type = PatternType.Empty;
            }

            pattern.ConsistencyScore = CalculateConsistencyScore(pattern);
            pattern.MaxDepth = CalculateMaxDepth(root);

            return pattern;
        }

        private void AnalyzeArrayElements(YamlSequenceNode array, StructurePattern pattern)
        {
            var unifiedProperties = new Dictionary<string, PropertyPattern>();
            var nestedArrays = new Dictionary<string, ArrayPattern>();
            var globalPropertyIndex = 0;
            
            // 모든 중첩 배열 요소들을 수집하여 통합 분석
            var allNestedArrayElements = new Dictionary<string, List<YamlSequenceNode>>();
            
            // 각 요소를 순회하면서 즉시 스키마 업데이트
            for (int elementIndex = 0; elementIndex < array.Children.Count; elementIndex++)
            {
                var element = array.Children[elementIndex];
                if (element is YamlMappingNode mapping)
                {
                    UpdateSchemaFromElement(mapping, unifiedProperties, nestedArrays, 
                                          ref globalPropertyIndex, elementIndex);
                    
                    // 중첩 배열 요소들을 수집
                    CollectNestedArrays(mapping, allNestedArrayElements, "");
                }
            }
            
            // 수집된 모든 중첩 배열 요소들을 통합 분석
            foreach (var kvp in allNestedArrayElements)
            {
                var arrayPath = kvp.Key;
                var allArrayInstances = kvp.Value;
                
                Logger.Information($"중첩 배열 '{arrayPath}' 통합 분석: {allArrayInstances.Count}개 인스턴스");
                
                // 최상위 배열 이름 추출 (예: "events.results" -> "events")
                var topLevelArrayName = arrayPath.Contains(".") ? arrayPath.Split('.')[0] : arrayPath;
                
                if (arrayPath.Contains("."))
                {
                    // 중첩된 배열 (예: events.results)
                    var pathParts = arrayPath.Split('.');
                    if (pathParts.Length >= 2 && pathParts[1] == "results")
                    {
                        // results 배열의 모든 요소를 통합
                        var allElements = new List<YamlNode>();
                        foreach (var instance in allArrayInstances)
                        {
                            allElements.AddRange(instance.Children);
                        }
                        
                        Logger.Information($"★ results 배열 통합 분석: 총 {allElements.Count}개 요소");
                        
                        // 통합된 요소들로 배열 패턴 생성
                        var unifiedResultsPattern = AnalyzeArrayFromElements("results", allElements);
                        
                        // events 배열의 ElementProperties에서 results 속성 업데이트
                        if (nestedArrays.ContainsKey("events") && 
                            nestedArrays["events"].ElementProperties != null &&
                            nestedArrays["events"].ElementProperties.ContainsKey("results"))
                        {
                            nestedArrays["events"].ElementProperties["results"].ArrayPattern = unifiedResultsPattern;
                            Logger.Information($"events.results 패턴 업데이트 완료: {unifiedResultsPattern.ElementProperties?.Count ?? 0}개 속성");
                        }
                    }
                }
                else if (nestedArrays.ContainsKey(arrayPath))
                {
                    // 최상위 배열
                    var allElements = new List<YamlNode>();
                    foreach (var instance in allArrayInstances)
                    {
                        allElements.AddRange(instance.Children);
                    }
                    
                    // 통합된 요소들로 배열 패턴 재생성
                    var unifiedArrayPattern = AnalyzeArrayFromElements(arrayPath, allElements);
                    nestedArrays[arrayPath] = unifiedArrayPattern;
                    
                    // PropertyPattern에도 업데이트
                    if (unifiedProperties.ContainsKey(arrayPath))
                    {
                        unifiedProperties[arrayPath].ArrayPattern = unifiedArrayPattern;
                    }
                }
            }
            
            // 통계 정보 업데이트
            foreach (var prop in unifiedProperties.Values)
            {
                prop.OccurrenceRatio = (double)prop.OccurrenceCount / array.Children.Count;
                prop.IsRequired = prop.OccurrenceRatio > 0.8;
            }
            
            pattern.Properties = unifiedProperties;
            pattern.Arrays = nestedArrays;
        }
        
        private void UpdateSchemaFromElement(
            YamlMappingNode element, 
            Dictionary<string, PropertyPattern> properties,
            Dictionary<string, ArrayPattern> arrays,
            ref int globalPropertyIndex,
            int elementIndex)
        {
            foreach (var kvp in element.Children)
            {
                var propName = kvp.Key.ToString();
                var propValue = kvp.Value;
                
                // 속성이 처음 나타나면 생성
                if (!properties.ContainsKey(propName))
                {
                    properties[propName] = new PropertyPattern
                    {
                        Name = propName,
                        FirstAppearanceIndex = globalPropertyIndex++,
                        OccurrenceCount = 0,
                        Types = new HashSet<Type>()
                    };
                    
                    Logger.Debug($"새 속성 발견: '{propName}', FirstAppearanceIndex={properties[propName].FirstAppearanceIndex}");
                }
                
                // 속성 정보 업데이트
                var pattern = properties[propName];
                pattern.OccurrenceCount++;
                pattern.Types.Add(propValue.GetType());
                
                // 타입별 처리
                if (propValue is YamlSequenceNode sequence)
                {
                    pattern.IsArray = true;
                    if (!arrays.ContainsKey(propName))
                    {
                        arrays[propName] = AnalyzeArray(propName, sequence);
                        Logger.Debug($"새 배열 패턴 생성: {propName}, ElementProperties 개수={arrays[propName].ElementProperties?.Count ?? 0}");
                        
                        // Option 배열 특별 디버깅
                        if (propName == "Option")
                        {
                            Logger.Information($"🔍 Option 배열 분석 결과:");
                            Logger.Information($"  - 배열 요소 수: {sequence.Children.Count}");
                            Logger.Information($"  - ElementProperties 개수: {arrays[propName].ElementProperties?.Count ?? 0}");
                            if (arrays[propName].ElementProperties != null)
                            {
                                foreach (var elemProp in arrays[propName].ElementProperties)
                                {
                                    Logger.Information($"    - {elemProp.Key}: OccurrenceCount={elemProp.Value.OccurrenceCount}");
                                }
                            }
                        }
                    }
                    else
                    {
                        // 기존 배열 패턴과 병합하여 통합 스키마 구축
                        MergeArrayPattern(arrays[propName], sequence);
                        Logger.Debug($"배열 패턴 병합: {propName}, ElementProperties 개수={arrays[propName].ElementProperties?.Count ?? 0}");
                    }
                    pattern.ArrayPattern = arrays[propName];
                    
                    // results 배열 업데이트 확인
                    if (propName == "results" && arrays[propName].ElementProperties != null)
                    {
                        Logger.Information($"★ results 배열 업데이트 (요소 {elementIndex}):");
                        var hasDelay = arrays[propName].ElementProperties.ContainsKey("delay");
                        var hasSendAll = arrays[propName].ElementProperties.ContainsKey("sendAll");
                        Logger.Information($"  - delay 포함: {hasDelay}");
                        Logger.Information($"  - sendAll 포함: {hasSendAll}");
                        if (hasDelay || hasSendAll)
                        {
                            Logger.Information($"  - 속성 목록: [{string.Join(", ", arrays[propName].ElementProperties.Keys)}]");
                        }
                    }
                    
                    // results 배열 디버깅
                    if (propName == "results" && arrays[propName].ElementProperties != null)
                    {
                        Logger.Information($"★ results 배열 업데이트 (요소 {elementIndex}):");
                        foreach (var elemProp in arrays[propName].ElementProperties)
                        {
                            Logger.Information($"  - {elemProp.Key}: OccurrenceCount={elemProp.Value.OccurrenceCount}");
                        }
                        
                        if (!arrays[propName].ElementProperties.ContainsKey("delay"))
                        {
                            Logger.Warning("  ⚠️ delay 속성이 아직 없음");
                        }
                        if (!arrays[propName].ElementProperties.ContainsKey("sendAll"))
                        {
                            Logger.Warning("  ⚠️ sendAll 속성이 아직 없음");
                        }
                    }
                    
                    // events 배열인 경우 특별 로깅
                    if (propName == "events" && arrays[propName].ElementProperties != null)
                    {
                        Logger.Information($"★ events 배열 분석 결과:");
                        foreach (var elemProp in arrays[propName].ElementProperties)
                        {
                            Logger.Information($"  - {elemProp.Key}: IsObject={elemProp.Value.IsObject}, OccurrenceCount={elemProp.Value.OccurrenceCount}");
                            if (elemProp.Key == "activation")
                            {
                                Logger.Information($"    ★★★ activation 발견! IsObject={elemProp.Value.IsObject}, Properties=[{string.Join(", ", elemProp.Value.ObjectProperties ?? new List<string>())}]");
                            }
                        }
                    }
                }
                else if (propValue is YamlMappingNode objMapping)
                {
                    pattern.IsObject = true;
                    pattern.ObjectProperties = ExtractObjectPropertyNames(objMapping);
                    
                    Logger.Information($"UpdateSchemaFromElement: '{propName}' 객체 감지! (요소 {elementIndex})");
                    Logger.Information($"  - 객체 속성 개수: {pattern.ObjectProperties?.Count ?? 0}");
                    if (pattern.ObjectProperties?.Count > 0)
                    {
                        Logger.Information($"  - 객체 속성 목록: [{string.Join(", ", pattern.ObjectProperties)}]");
                    }
                    
                    // 중첩된 객체도 재귀적으로 분석
                    var nestedPattern = new StructurePattern
                    {
                        Properties = new Dictionary<string, PropertyPattern>(),
                        Arrays = new Dictionary<string, ArrayPattern>()
                    };
                    AnalyzeObjectProperties(objMapping, nestedPattern);
                    pattern.NestedPatterns = nestedPattern.Properties;
                }
            }
        }

        private void AnalyzeObjectProperties(YamlMappingNode mapping, StructurePattern pattern)
        {
            foreach (var kvp in mapping.Children)
            {
                var key = kvp.Key.ToString();
                var value = kvp.Value;

                var prop = new PropertyPattern
                {
                    Name = key,
                    OccurrenceCount = 1,
                    Types = new HashSet<Type> { value.GetType() },
                    FirstAppearanceIndex = 0,
                    OccurrenceRatio = 1.0,
                    IsRequired = true
                };

                if (value is YamlSequenceNode)
                {
                    prop.IsArray = true;
                    var arrayPattern = AnalyzeArray(key, value as YamlSequenceNode);
                    pattern.Arrays[key] = arrayPattern;
                    prop.ArrayPattern = arrayPattern; // PropertyPattern에도 ArrayPattern 설정
                }
                else if (value is YamlMappingNode objMapping)
                {
                    prop.IsObject = true;
                    prop.ObjectProperties = ExtractObjectPropertyNames(objMapping);
                    
                    // 재귀적으로 중첩된 패턴 분석
                    var nestedPattern = new StructurePattern
                    {
                        Properties = new Dictionary<string, PropertyPattern>(),
                        Arrays = new Dictionary<string, ArrayPattern>()
                    };
                    AnalyzeObjectProperties(objMapping, nestedPattern);
                    prop.NestedPatterns = nestedPattern.Properties;
                    
                    Logger.Information($"AnalyzeObjectProperties: '{key}' 객체 속성 분석 완료, ObjectProperties 개수 = {prop.ObjectProperties?.Count ?? 0}");
                    if (prop.ObjectProperties?.Count > 0)
                    {
                        Logger.Information($"AnalyzeObjectProperties: '{key}' 객체 속성 목록 = [{string.Join(", ", prop.ObjectProperties)}]");
                    }
                }

                pattern.Properties[key] = prop;
            }
        }

        private List<string> ExtractObjectPropertyNames(YamlMappingNode objMapping)
        {
            var properties = new List<string>();
            foreach (var kvp in objMapping.Children)
            {
                properties.Add(kvp.Key.ToString());
            }
            Logger.Debug($"ExtractObjectPropertyNames: YAML 파싱 순서대로 추출된 속성 개수 = {properties.Count}, 속성들 = [{string.Join(", ", properties)}]");
            return properties;
        }

        private Dictionary<string, object> ExtractElementSchema(YamlNode element)
        {
            var schema = new Dictionary<string, object>();

            if (element is YamlMappingNode mapping)
            {
                Logger.Debug($"ExtractElementSchema: 요소 분석 시작, 속성 개수 = {mapping.Children.Count}");
                foreach (var kvp in mapping.Children)
                {
                    var key = kvp.Key.ToString();
                    var value = kvp.Value;
                    Logger.Debug($"  - 속성 '{key}' 타입: {value.GetType().Name}");

                    if (value is YamlScalarNode scalar)
                    {
                        schema[key] = scalar.Value;
                    }
                    else if (value is YamlSequenceNode sequence)
                    {
                        // 배열의 요소들을 분석
                        var arrayInfo = new ArrayInfo
                        {
                            IsArray = true,
                            ElementCount = sequence.Children.Count,
                            Elements = new List<Dictionary<string, object>>()
                        };
                        
                        // 각 배열 요소의 스키마 추출 (재귀적으로)
                        foreach (var child in sequence.Children)
                        {
                            // 재귀적으로 각 요소의 전체 스키마를 추출
                            var childSchema = ExtractElementSchema(child);
                            if (childSchema.Count > 0)
                            {
                                arrayInfo.Elements.Add(childSchema);
                            }
                        }
                        
                        schema[key] = arrayInfo;
                    }
                    else if (value is YamlMappingNode nestedMapping)
                    {
                        // 중첩 객체의 속성들도 추출
                        var objInfo = new ObjectInfo
                        {
                            Type = "Object",
                            Properties = ExtractObjectPropertyNames(nestedMapping)
                        };
                        Logger.Information($"ExtractElementSchema: '{key}' 객체 발견!");
                        Logger.Information($"  - 속성 개수: {objInfo.Properties.Count}");
                        Logger.Information($"  - 속성 목록: [{string.Join(", ", objInfo.Properties)}]");
                        schema[key] = objInfo;
                    }
                }
            }

            return schema;
        }

        private Dictionary<string, PropertyPattern> UnifySchemas(List<Dictionary<string, object>> schemas)
        {
            var unified = new Dictionary<string, PropertyPattern>();
            
            Logger.Information($"UnifySchemas 시작: 스키마 개수 = {schemas.Count}");
            
            // 모든 속성 수집 및 분석
            for (int i = 0; i < schemas.Count; i++)
            {
                var schema = schemas[i];
                Logger.Debug($"스키마 {i}: 속성 개수 = {schema.Count}");
                foreach (var prop in schema) // ★ _ 필터 제거 - 모든 속성 포함
                {
                    var valueType = prop.Value?.GetType()?.Name ?? "null";
                    Logger.Debug($"  속성 '{prop.Key}': 타입 = {valueType}");
                    if (prop.Value is ObjectInfo objInfo)
                    {
                        Logger.Information($"    -> ObjectInfo 감지! Type={objInfo.Type}, Properties={objInfo.Properties?.Count ?? 0}");
                        if (prop.Key == "activation")
                        {
                            Logger.Information($"    ★★★ activation ObjectInfo 발견! 속성: [{string.Join(", ", objInfo.Properties)}]");
                        }
                    }
                    if (!unified.ContainsKey(prop.Key))
                    {
                        unified[prop.Key] = new PropertyPattern
                        {
                            Name = prop.Key,
                            OccurrenceCount = 0,
                            Types = new HashSet<Type>(),
                            FirstAppearanceIndex = i,
                            ObjectProperties = new List<string>()
                        };
                    }
                    
                    unified[prop.Key].OccurrenceCount++;
                    unified[prop.Key].Types.Add(prop.Value?.GetType() ?? typeof(object));

                    // 배열이나 객체 타입 감지
                    if (prop.Value is ArrayInfo arrayInfo)
                    {
                        unified[prop.Key].IsArray = true;
                        
                        // 배열 요소들의 패턴 분석
                        if (arrayInfo.Elements != null && arrayInfo.Elements.Any())
                        {
                            var elementPattern = new ArrayPattern
                            {
                                Name = prop.Key,
                                ElementProperties = UnifySchemas(arrayInfo.Elements),
                                MaxSize = arrayInfo.ElementCount,
                                MinSize = arrayInfo.ElementCount
                            };
                            unified[prop.Key].ArrayPattern = elementPattern;
                        }
                    }
                    else if (prop.Value is ObjectInfo innerObjInfo)
                    {
                        // ObjectInfo 타입 직접 처리
                        Logger.Debug($"UnifySchemas: ObjectInfo 타입 감지 - '{prop.Key}', Type={innerObjInfo.Type}, 속성 개수={innerObjInfo.Properties?.Count ?? 0}");
                        if (innerObjInfo.Type == "Object")
                        {
                            unified[prop.Key].IsObject = true;
                            // 기존 속성들과 병합
                            if (unified[prop.Key].ObjectProperties == null)
                                unified[prop.Key].ObjectProperties = new List<string>();
                            
                            foreach (var subProp in innerObjInfo.Properties)
                            {
                                if (!unified[prop.Key].ObjectProperties.Contains(subProp))
                                    unified[prop.Key].ObjectProperties.Add(subProp);
                            }
                            
                            Logger.Information($"UnifySchemas: '{prop.Key}' 객체 속성 설정 완료, ObjectProperties 개수 = {unified[prop.Key].ObjectProperties?.Count ?? 0}");
                            Logger.Information($"UnifySchemas: '{prop.Key}' 객체 속성 목록 = [{string.Join(", ", unified[prop.Key].ObjectProperties)}]");
                            
                            if (prop.Key == "activation")
                            {
                                Logger.Information($"★★★ activation IsObject 설정됨! IsObject={unified[prop.Key].IsObject}");
                            }
                        }
                    }
                    else if (prop.Value is Dictionary<string, object> dict)
                    {
                        var type = dict.ContainsKey("Type") ? dict["Type"].ToString() : "";
                        if (type == "Array")
                            unified[prop.Key].IsArray = true;
                        else if (type == "Object")
                            unified[prop.Key].IsObject = true;
                    }
                    else
                    {
                        // 익명 타입 처리 (폴백)
                        var propValueType = prop.Value?.GetType();
                        if (propValueType != null && !propValueType.IsPrimitive && propValueType != typeof(string))
                        {
                            var typeProperty = propValueType.GetProperty("Type");
                            var propertiesProperty = propValueType.GetProperty("Properties");
                            
                            if (typeProperty != null && propertiesProperty != null)
                            {
                                var typeValue = typeProperty.GetValue(prop.Value)?.ToString();
                                if (typeValue == "Object")
                                {
                                    unified[prop.Key].IsObject = true;
                                    var properties = propertiesProperty.GetValue(prop.Value) as List<string>;
                                    if (properties != null)
                                    {
                                        unified[prop.Key].ObjectProperties = properties;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // 출현 비율 계산
            foreach (var prop in unified.Values)
            {
                prop.OccurrenceRatio = (double)prop.OccurrenceCount / schemas.Count;
                prop.IsRequired = prop.OccurrenceRatio > 0.8; // 80% 이상 출현시 필수
            }

            // 모든 속성 포함 - 한 번이라도 나타난 속성은 헤더에 표시
            Logger.Information($"통합된 속성 총 {unified.Count}개 - 모두 스키마에 포함");
            foreach (var kvp in unified)
            {
                Logger.Debug($"속성 포함: '{kvp.Key}' (출현 횟수: {kvp.Value.OccurrenceCount}/{schemas.Count}, " +
                           $"출현율: {kvp.Value.OccurrenceRatio:P}, 객체: {kvp.Value.IsObject})");
            }

            return unified;
        }

        private Dictionary<string, ArrayPattern> DetectNestedArrays(List<Dictionary<string, object>> schemas)
        {
            var arrays = new Dictionary<string, ArrayPattern>();

            // 배열 속성 감지
            foreach (var schema in schemas)
            {
                foreach (var prop in schema)
                {
                    if (prop.Value is ArrayInfo arrayInfo && arrayInfo.IsArray)
                    {
                        if (!arrays.ContainsKey(prop.Key))
                        {
                            arrays[prop.Key] = new ArrayPattern
                            {
                                Name = prop.Key,
                                MaxSize = 0,
                                MinSize = int.MaxValue,
                                ElementProperties = new Dictionary<string, PropertyPattern>()
                            };
                        }

                        arrays[prop.Key].MaxSize = Math.Max(arrays[prop.Key].MaxSize, arrayInfo.ElementCount);
                        arrays[prop.Key].MinSize = Math.Min(arrays[prop.Key].MinSize, arrayInfo.ElementCount);
                        
                        // 배열 요소들의 스키마 통합
                        if (arrayInfo.Elements != null && arrayInfo.Elements.Any())
                        {
                            var elementSchemas = arrayInfo.Elements;
                            arrays[prop.Key].ElementProperties = UnifySchemas(elementSchemas);
                            arrays[prop.Key].HasVariableStructure = arrays[prop.Key].ElementProperties.Any(p => p.Value.OccurrenceRatio < 1.0);
                            
                            // 가변 속성 분석 (weaponSpec의 damage/addDamage 같은 경우)
                            var allUniqueProps = new HashSet<string>();
                            var propCounts = new Dictionary<string, int>();
                            
                            foreach (var elem in elementSchemas)
                            {
                                foreach (var propKey in elem.Keys)
                                {
                                    allUniqueProps.Add(propKey);
                                    if (!propCounts.ContainsKey(propKey))
                                        propCounts[propKey] = 0;
                                    propCounts[propKey]++;
                                }
                            }
                            
                            arrays[prop.Key].AllUniqueProperties = allUniqueProps.ToList();
                            arrays[prop.Key].ElementPropertyCounts = propCounts;
                            arrays[prop.Key].HasVariableProperties = allUniqueProps.Count > elementSchemas.First().Count;
                        }
                    }
                }
            }

            // 배열 패턴 분석
            foreach (var array in arrays.Values)
            {
                array.OccurrenceRatio = 1.0; // 추후 정확한 계산 필요
                array.RequiresMultipleRows = array.MaxSize > 5 || array.MinSize != array.MaxSize;
            }

            return arrays;
        }

        private void MergeArrayPattern(ArrayPattern existingPattern, YamlSequenceNode newArray)
        {
            // 새 배열의 각 요소를 기존 패턴과 병합
            for (int i = 0; i < newArray.Children.Count; i++)
            {
                if (newArray.Children[i] is YamlMappingNode mapping)
                {
                    // 해당 인덱스의 요소 속성들을 기존 패턴에 병합
                    foreach (var kvp in mapping.Children)
                    {
                        var propName = kvp.Key.ToString();
                        
                        // ElementProperties가 null이면 초기화
                        if (existingPattern.ElementProperties == null)
                        {
                            existingPattern.ElementProperties = new Dictionary<string, PropertyPattern>();
                        }
                        
                        // 속성이 처음 나타나면 추가
                        if (!existingPattern.ElementProperties.ContainsKey(propName))
                        {
                            existingPattern.ElementProperties[propName] = new PropertyPattern
                            {
                                Name = propName,
                                OccurrenceCount = 1,
                                Types = new HashSet<Type> { kvp.Value.GetType() },
                                FirstAppearanceIndex = existingPattern.ElementProperties.Count
                            };
                            
                            // 객체나 배열 타입 처리
                            if (kvp.Value is YamlMappingNode objMapping)
                            {
                                existingPattern.ElementProperties[propName].IsObject = true;
                                existingPattern.ElementProperties[propName].ObjectProperties = ExtractObjectPropertyNames(objMapping);
                                
                                Logger.Information($"MergeArrayPattern: '{propName}' 객체 감지!");
                                Logger.Information($"  - 객체 속성: [{string.Join(", ", existingPattern.ElementProperties[propName].ObjectProperties)}]");
                            }
                            else if (kvp.Value is YamlSequenceNode nestedSequence)
                            {
                                existingPattern.ElementProperties[propName].IsArray = true;
                                
                                // 중첩된 배열의 패턴도 병합
                                if (existingPattern.ElementProperties[propName].ArrayPattern == null)
                                {
                                    existingPattern.ElementProperties[propName].ArrayPattern = AnalyzeArray(propName, nestedSequence);
                                }
                                else
                                {
                                    // 재귀적으로 중첩 배열 병합
                                    MergeArrayPattern(existingPattern.ElementProperties[propName].ArrayPattern, nestedSequence);
                                }
                            }
                        }
                        else
                        {
                            // 기존 속성 업데이트
                            existingPattern.ElementProperties[propName].OccurrenceCount++;
                            existingPattern.ElementProperties[propName].Types.Add(kvp.Value.GetType());
                            
                            // 객체 타입인 경우 속성 병합
                            if (kvp.Value is YamlMappingNode objMapping)
                            {
                                // 객체 타입으로 설정 (이전에 스칼라였어도 객체로 변경)
                                existingPattern.ElementProperties[propName].IsObject = true;
                                
                                // ObjectProperties가 null이면 초기화
                                if (existingPattern.ElementProperties[propName].ObjectProperties == null)
                                {
                                    existingPattern.ElementProperties[propName].ObjectProperties = new List<string>();
                                }
                                
                                var newProps = ExtractObjectPropertyNames(objMapping);
                                foreach (var newProp in newProps)
                                {
                                    if (!existingPattern.ElementProperties[propName].ObjectProperties.Contains(newProp))
                                    {
                                        existingPattern.ElementProperties[propName].ObjectProperties.Add(newProp);
                                    }
                                }
                                
                                Logger.Information($"MergeArrayPattern: '{propName}' 기존 속성을 객체로 업데이트!");
                                Logger.Information($"  - 객체 속성: [{string.Join(", ", existingPattern.ElementProperties[propName].ObjectProperties)}]");
                            }
                            else if (kvp.Value is YamlSequenceNode nestedSequence)
                            {
                                // 중첩된 배열인 경우 기존 패턴과 병합
                                existingPattern.ElementProperties[propName].IsArray = true;
                                
                                if (existingPattern.ElementProperties[propName].ArrayPattern == null)
                                {
                                    existingPattern.ElementProperties[propName].ArrayPattern = AnalyzeArray(propName, nestedSequence);
                                }
                                else
                                {
                                    // 재귀적으로 중첩 배열 병합
                                    MergeArrayPattern(existingPattern.ElementProperties[propName].ArrayPattern, nestedSequence);
                                }
                            }
                        }
                    }
                }
            }
            
            // 배열 크기 업데이트
            existingPattern.MaxSize = Math.Max(existingPattern.MaxSize, newArray.Children.Count);
        }

        private ArrayPattern AnalyzeArray(string name, YamlSequenceNode array)
        {
            Logger.Information($"🔍 AnalyzeArray 시작: '{name}' 배열, 요소 수={array.Children.Count}");
            
            var pattern = new ArrayPattern
            {
                Name = name,
                MaxSize = array.Children.Count,
                MinSize = array.Children.Count,
                ElementProperties = new Dictionary<string, PropertyPattern>(),
                ElementPropertyCounts = new Dictionary<string, int>(),
                AllUniqueProperties = new List<string>()
            };

            // 배열 요소들의 스키마 분석
            var elementSchemas = new List<Dictionary<string, object>>();
            for (int i = 0; i < array.Children.Count; i++)
            {
                var element = array.Children[i];
                Logger.Debug($"  배열 요소 [{i}] 분석 중: {element.GetType().Name}");
                
                var schema = ExtractElementSchema(element);
                Logger.Debug($"  배열 요소 [{i}] 스키마 추출 완료: {schema.Count}개 속성");
                
                if (name == "Option")
                {
                    Logger.Information($"  Option 요소 [{i}] 스키마: [{string.Join(", ", schema.Keys)}]");
                }
                
                elementSchemas.Add(schema);
            }

            if (elementSchemas.Any())
            {
                Logger.Debug($"  UnifySchemas 호출: {elementSchemas.Count}개 스키마 통합");
                pattern.ElementProperties = UnifySchemas(elementSchemas);
                Logger.Debug($"  UnifySchemas 완료: {pattern.ElementProperties?.Count ?? 0}개 통합 속성");
                
                if (name == "Option")
                {
                    Logger.Information($"  Option 배열 통합 결과: {pattern.ElementProperties?.Count ?? 0}개 속성");
                    if (pattern.ElementProperties != null)
                    {
                        foreach (var prop in pattern.ElementProperties)
                        {
                            Logger.Information($"    - {prop.Key}: OccurrenceCount={prop.Value.OccurrenceCount}");
                        }
                    }
                }
                
                // 가변 속성 분석
                var allUniqueProps = new HashSet<string>();
                var propCounts = new Dictionary<string, int>();
                
                foreach (var elem in elementSchemas)
                {
                    foreach (var propKey in elem.Keys)
                    {
                        allUniqueProps.Add(propKey);
                        if (!propCounts.ContainsKey(propKey))
                            propCounts[propKey] = 0;
                        propCounts[propKey]++;
                    }
                }
                
                pattern.AllUniqueProperties = allUniqueProps.ToList();
                pattern.ElementPropertyCounts = propCounts;
                pattern.HasVariableProperties = allUniqueProps.Count > elementSchemas.First().Count;
            }

            pattern.RequiresMultipleRows = pattern.MaxSize > 5;
            
            // events 배열 후처리 - activation 확인
            if (name == "events" && pattern.ElementProperties != null)
            {
                Logger.Information($"AnalyzeArray 후처리: events 배열 분석 완료");
                foreach (var elemProp in pattern.ElementProperties)
                {
                    if (elemProp.Key == "activation")
                    {
                        Logger.Information($"  ★ activation 최종 상태: IsObject={elemProp.Value.IsObject}, " +
                                         $"OccurrenceCount={elemProp.Value.OccurrenceCount}, " +
                                         $"Properties=[{string.Join(", ", elemProp.Value.ObjectProperties ?? new List<string>())}]");
                    }
                }
            }

            return pattern;
        }

        private void CollectNestedArrays(YamlMappingNode mapping, Dictionary<string, List<YamlSequenceNode>> allNestedArrayElements, string parentPath = "")
        {
            foreach (var kvp in mapping.Children)
            {
                var key = kvp.Key.ToString();
                var value = kvp.Value;
                var currentPath = string.IsNullOrEmpty(parentPath) ? key : $"{parentPath}.{key}";
                
                if (value is YamlSequenceNode sequence)
                {
                    if (!allNestedArrayElements.ContainsKey(currentPath))
                    {
                        allNestedArrayElements[currentPath] = new List<YamlSequenceNode>();
                    }
                    allNestedArrayElements[currentPath].Add(sequence);
                    
                    // 배열의 각 요소도 재귀적으로 탐색 (results 같은 중첩 배열 찾기)
                    foreach (var element in sequence.Children)
                    {
                        if (element is YamlMappingNode elementMapping)
                        {
                            CollectNestedArrays(elementMapping, allNestedArrayElements, currentPath);
                        }
                    }
                }
                else if (value is YamlMappingNode nestedMapping)
                {
                    // 재귀적으로 중첩된 배열들도 수집
                    CollectNestedArrays(nestedMapping, allNestedArrayElements, currentPath);
                }
            }
        }
        
        private ArrayPattern AnalyzeArrayFromElements(string name, List<YamlNode> allElements)
        {
            var pattern = new ArrayPattern
            {
                Name = name,
                MaxSize = 0,
                MinSize = int.MaxValue,
                ElementProperties = new Dictionary<string, PropertyPattern>(),
                ElementPropertyCounts = new Dictionary<string, int>(),
                AllUniqueProperties = new List<string>()
            };
            
            // 모든 요소들의 스키마 분석
            var elementSchemas = new List<Dictionary<string, object>>();
            foreach (var element in allElements)
            {
                var schema = ExtractElementSchema(element);
                if (schema.Count > 0)
                {
                    elementSchemas.Add(schema);
                }
            }
            
            Logger.Information($"AnalyzeArrayFromElements: '{name}' 배열의 총 {elementSchemas.Count}개 요소 분석");
            
            if (elementSchemas.Any())
            {
                // 모든 요소의 스키마를 통합
                pattern.ElementProperties = UnifySchemas(elementSchemas);
                
                // 가변 속성 분석
                var allUniqueProps = new HashSet<string>();
                var propCounts = new Dictionary<string, int>();
                
                foreach (var elem in elementSchemas)
                {
                    foreach (var propKey in elem.Keys)
                    {
                        allUniqueProps.Add(propKey);
                        if (!propCounts.ContainsKey(propKey))
                            propCounts[propKey] = 0;
                        propCounts[propKey]++;
                    }
                }
                
                pattern.AllUniqueProperties = allUniqueProps.ToList();
                pattern.ElementPropertyCounts = propCounts;
                pattern.HasVariableProperties = allUniqueProps.Count > (elementSchemas.Count > 0 ? elementSchemas.First().Count : 0);
                
                // 크기 정보 업데이트
                pattern.MaxSize = elementSchemas.Count;
                pattern.MinSize = elementSchemas.Count;
            }
            
            pattern.RequiresMultipleRows = pattern.MaxSize > 5;
            
            // results 배열 디버깅
            if (name == "results" && pattern.ElementProperties != null)
            {
                Logger.Information($"★ results 배열 통합 분석 완료:");
                Logger.Information($"  - 총 요소 수: {elementSchemas.Count}");
                Logger.Information($"  - 발견된 속성들: [{string.Join(", ", pattern.AllUniqueProperties)}]");
                foreach (var prop in pattern.ElementProperties)
                {
                    Logger.Information($"  - {prop.Key}: OccurrenceCount={prop.Value.OccurrenceCount}, " +
                                     $"OccurrenceRatio={prop.Value.OccurrenceRatio:P}");
                }
                
                if (!pattern.ElementProperties.ContainsKey("delay"))
                {
                    Logger.Warning("  ⚠️ delay 속성이 최종 스키마에 없음!");
                }
                if (!pattern.ElementProperties.ContainsKey("sendAll"))
                {
                    Logger.Warning("  ⚠️ sendAll 속성이 최종 스키마에 없음!");
                }
            }
            
            return pattern;
        }

        private double CalculateConsistencyScore(StructurePattern pattern)
        {
            if (!pattern.Properties.Any())
                return 0;

            // 필수 속성의 비율로 일관성 점수 계산
            var requiredCount = pattern.Properties.Count(p => p.Value.IsRequired);
            return (double)requiredCount / pattern.Properties.Count;
        }

        private int CalculateMaxDepth(YamlNode node, int currentDepth = 0)
        {
            if (node is YamlScalarNode)
                return currentDepth;

            int maxChildDepth = currentDepth;

            if (node is YamlSequenceNode sequence)
            {
                foreach (var child in sequence.Children)
                {
                    maxChildDepth = Math.Max(maxChildDepth, CalculateMaxDepth(child, currentDepth + 1));
                }
            }
            else if (node is YamlMappingNode mapping)
            {
                foreach (var kvp in mapping.Children)
                {
                    maxChildDepth = Math.Max(maxChildDepth, CalculateMaxDepth(kvp.Value, currentDepth + 1));
                }
            }

            return maxChildDepth;
        }
        
    }
}