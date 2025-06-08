using System;
using System.Collections.Generic;
using System.Linq;
using YamlDotNet.RepresentationModel;
using ExcelToYamlAddin.Logging;
using static ExcelToYamlAddin.Core.YamlToExcel.DynamicStructureAnalyzer;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    public class DynamicHorizontalExpander
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<DynamicHorizontalExpander>();
        
        public class ElementLayout
        {
            public int Index { get; set; }
            public List<string> Properties { get; set; }
            public int RequiredColumns { get; set; }
            public List<string> UnifiedProperties { get; set; }
            public Dictionary<string, int> PropertyColumnMap { get; set; }
        }

        public class DynamicArrayLayout
        {
            public string ArrayPath { get; set; }
            public int ElementCount { get; set; }
            public List<ElementLayout> Elements { get; set; }
            public int TotalColumns { get; set; }
            public int ActualUsedColumns { get; set; }
            public bool OptimizeColumns { get; set; }
            public Dictionary<string, PropertyPattern> UnifiedSchema { get; set; }
            public List<string> OrderedProperties { get; set; }
        }

        private readonly DynamicPropertyOrderer _propertyOrderer;

        public DynamicHorizontalExpander()
        {
            _propertyOrderer = new DynamicPropertyOrderer();
        }

        public DynamicArrayLayout CalculateArrayLayout(
            string arrayPath, 
            YamlSequenceNode array,
            Dictionary<string, PropertyPattern> unifiedSchema)
        {
            var layout = new DynamicArrayLayout
            {
                ArrayPath = arrayPath,
                ElementCount = array.Children.Count,
                Elements = new List<ElementLayout>(),
                UnifiedSchema = unifiedSchema ?? new Dictionary<string, PropertyPattern>(),
                OptimizeColumns = false  // 개별 스키마 사용을 위해 false로 변경
            };

            // 배열 요소들의 모든 속성 수집
            var allElementProperties = CollectAllElementProperties(array);
            
            // 통합 스키마가 없으면 생성
            if (!layout.UnifiedSchema.Any())
            {
                layout.UnifiedSchema = CreateUnifiedSchemaFromElements(allElementProperties);
            }
            
            // UnifiedSchema가 이미 전체 루트의 통합 스키마인 경우에도
            // 현재 배열의 실제 속성들을 기반으로 보완
            var currentArraySchema = CreateUnifiedSchemaFromElements(allElementProperties);
            foreach (var prop in currentArraySchema)
            {
                if (!layout.UnifiedSchema.ContainsKey(prop.Key))
                {
                    layout.UnifiedSchema[prop.Key] = prop.Value;
                }
            }

            Logger.Debug($"CalculateArrayLayout: {arrayPath} 통합 스키마 속성 {layout.UnifiedSchema.Count}개");
            foreach (var prop in layout.UnifiedSchema)
            {
                Logger.Debug($"  - {prop.Key}: FirstAppearanceIndex={prop.Value.FirstAppearanceIndex}");
            }

            // 속성 순서 결정
            layout.OrderedProperties = _propertyOrderer.OrderPropertiesForArrayElement(
                layout.UnifiedSchema, 
                new ArrayPattern { ElementProperties = layout.UnifiedSchema });

            // 각 요소별 레이아웃 계산
            for (int i = 0; i < array.Children.Count; i++)
            {
                if (array.Children[i] is YamlMappingNode element)
                {
                    var elementLayout = CreateElementLayout(i, element, layout.OrderedProperties);
                    layout.Elements.Add(elementLayout);
                }
            }

            // 전체 컬럼 수 계산
            layout.TotalColumns = CalculateTotalColumns(layout);
            layout.ActualUsedColumns = CalculateActualUsedColumns(layout);

            return layout;
        }

        private List<Dictionary<string, YamlNode>> CollectAllElementProperties(YamlSequenceNode array)
        {
            var allProperties = new List<Dictionary<string, YamlNode>>();

            foreach (var element in array.Children)
            {
                if (element is YamlMappingNode mapping)
                {
                    var properties = new Dictionary<string, YamlNode>();
                    foreach (var kvp in mapping.Children)
                    {
                        properties[kvp.Key.ToString()] = kvp.Value;
                    }
                    allProperties.Add(properties);
                }
            }

            return allProperties;
        }

        private Dictionary<string, PropertyPattern> CreateUnifiedSchemaFromElements(
            List<Dictionary<string, YamlNode>> elements)
        {
            var schema = new Dictionary<string, PropertyPattern>();

            Logger.Debug($"CreateUnifiedSchemaFromElements: {elements.Count}개 요소에서 통합 스키마 생성");
            
            for (int i = 0; i < elements.Count; i++)
            {
                Logger.Debug($"  요소 {i}: {elements[i].Count}개 속성");
                foreach (var prop in elements[i])
                {
                    if (!schema.ContainsKey(prop.Key))
                    {
                        schema[prop.Key] = new PropertyPattern
                        {
                            Name = prop.Key,
                            OccurrenceCount = 0,
                            Types = new HashSet<System.Type>(),
                            FirstAppearanceIndex = i
                        };
                        Logger.Debug($"    새 속성 '{prop.Key}' 추가 (FirstAppearanceIndex={i})");
                    }

                    schema[prop.Key].OccurrenceCount++;
                    schema[prop.Key].Types.Add(prop.Value.GetType());

                    if (prop.Value is YamlSequenceNode sequenceNode)
                    {
                        schema[prop.Key].IsArray = true;
                        // 배열 요소의 구조 분석
                        schema[prop.Key].ArrayPattern = AnalyzeArrayPattern(prop.Key, sequenceNode);
                    }
                    else if (prop.Value is YamlMappingNode mappingNode)
                    {
                        schema[prop.Key].IsObject = true;
                        // 객체의 하위 속성 수집
                        schema[prop.Key].ObjectProperties = ExtractObjectProperties(mappingNode);
                    }
                }
            }

            // 출현 비율 계산
            Logger.Debug("통합 스키마 완성:");
            foreach (var prop in schema.Values)
            {
                prop.OccurrenceRatio = (double)prop.OccurrenceCount / elements.Count;
                prop.IsRequired = prop.OccurrenceRatio > 0.8;
                Logger.Debug($"  '{prop.Name}': 출현 {prop.OccurrenceCount}/{elements.Count} ({prop.OccurrenceRatio:P}), FirstAppearanceIndex={prop.FirstAppearanceIndex}");
            }

            return schema;
        }

        private ElementLayout CreateElementLayout(
            int index, 
            YamlMappingNode element, 
            List<string> orderedProperties)
        {
            var elementLayout = new ElementLayout
            {
                Index = index,
                Properties = new List<string>(),
                PropertyColumnMap = new Dictionary<string, int>(),
                UnifiedProperties = orderedProperties
            };

            // 요소가 가진 속성들 수집
            foreach (var kvp in element.Children)
            {
                elementLayout.Properties.Add(kvp.Key.ToString());
            }

            // 통합 스키마 순서에 따라 컬럼 할당
            int columnOffset = 0;
            foreach (var prop in orderedProperties)
            {
                if (elementLayout.Properties.Contains(prop))
                {
                    elementLayout.PropertyColumnMap[prop] = columnOffset;
                }
                columnOffset++;
            }

            // 스키마에 없는 추가 속성 처리
            var extraProperties = elementLayout.Properties.Except(orderedProperties).ToList();
            foreach (var extraProp in extraProperties)
            {
                elementLayout.PropertyColumnMap[extraProp] = columnOffset++;
            }

            // 각 요소별로 실제 필요한 컬럼 수 설정 (가변 속성 지원)
            elementLayout.RequiredColumns = columnOffset;

            return elementLayout;
        }

        private int CalculateTotalColumns(DynamicArrayLayout layout)
        {
            if (!layout.Elements.Any())
                return 0;

            // 배열의 구조를 분석하여 중첩된 구조가 있는지 확인
            if (layout.UnifiedSchema != null && layout.UnifiedSchema.Any())
            {
                int totalColumnCount = 0;
                
                // 각 속성이 중첩된 구조인지 확인
                foreach (var prop in layout.UnifiedSchema)
                {
                    if (prop.Value.IsObject && prop.Value.ObjectProperties?.Count > 0)
                    {
                        // 객체의 하위 속성 수
                        totalColumnCount += prop.Value.ObjectProperties.Count;
                    }
                    else if (prop.Value.IsArray && prop.Value.ArrayPattern?.ElementProperties != null)
                    {
                        // 배열의 요소 속성 수 계산
                        var elementProps = prop.Value.ArrayPattern.ElementProperties;
                        if (elementProps.Any())
                        {
                            // results 배열처럼 요소가 단순 속성들을 가진 경우
                            totalColumnCount += elementProps.Count;
                        }
                        else
                        {
                            totalColumnCount += 1;
                        }
                    }
                    else
                    {
                        // 단순 속성
                        totalColumnCount += 1;
                    }
                }
                
                // events 배열처럼 단일 요소인 경우, 요소 수를 곱하지 않음
                if (layout.ElementCount == 1 && HasComplexNestedStructure(layout))
                {
                    return totalColumnCount;
                }
                
                // 여러 요소가 있는 일반적인 배열의 경우
                return layout.ElementCount * totalColumnCount;
            }

            // 폴백: 각 요소의 실제 속성 수를 합산
            int totalColumns = 0;
            foreach (var element in layout.Elements)
            {
                totalColumns += element.Properties.Count;
            }
            
            // 요소 수가 Elements.Count보다 많은 경우 마지막 요소 반복
            if (layout.ElementCount > layout.Elements.Count)
            {
                var lastElementColumns = layout.Elements.Last().Properties.Count;
                totalColumns += (layout.ElementCount - layout.Elements.Count) * lastElementColumns;
            }
            
            return totalColumns;
        }

        public int GetElementStartColumn(DynamicArrayLayout layout, int elementIndex)
        {
            if (elementIndex < 0 || elementIndex >= layout.Elements.Count)
                return -1;

            int startColumn = 0;
            for (int i = 0; i < elementIndex; i++)
            {
                startColumn += layout.Elements[i].RequiredColumns;
            }

            return startColumn;
        }

        public Dictionary<string, int> GetGlobalColumnMapping(DynamicArrayLayout layout)
        {
            var globalMapping = new Dictionary<string, int>();
            
            for (int i = 0; i < layout.Elements.Count; i++)
            {
                var element = layout.Elements[i];
                var startColumn = GetElementStartColumn(layout, i);
                
                foreach (var prop in element.PropertyColumnMap)
                {
                    var globalKey = $"{layout.ArrayPath}[{i}].{prop.Key}";
                    globalMapping[globalKey] = startColumn + prop.Value;
                }
            }

            return globalMapping;
        }

        // 수평 레이아웃 정보를 담는 클래스
        public class HorizontalLayout
        {
            public Dictionary<string, DynamicArrayLayout> ArrayLayouts { get; set; }
            public int TotalColumns { get; set; }

            public HorizontalLayout()
            {
                ArrayLayouts = new Dictionary<string, DynamicArrayLayout>();
            }
        }

        // YAML 시퀀스를 분석하여 수평 레이아웃 생성
        public HorizontalLayout AnalyzeHorizontalLayout(YamlSequenceNode rootSequence, StructurePattern pattern)
        {
            var layout = new HorizontalLayout();

            // 루트 배열의 각 요소에서 중첩 배열 찾기
            foreach (var array in pattern.Arrays)
            {
                if (array.Value.ElementProperties != null)
                {
                    // 모든 루트 요소에서 해당 배열을 수집하여 통합 분석
                    var allArrayElements = new List<YamlSequenceNode>();
                    
                    foreach (var rootElement in rootSequence.Children)
                    {
                        if (rootElement is YamlMappingNode mapping)
                        {
                            foreach (var prop in mapping.Children)
                            {
                                if (prop.Key.ToString() == array.Key && prop.Value is YamlSequenceNode seq)
                                {
                                    allArrayElements.Add(seq);
                                    break;
                                }
                            }
                        }
                    }

                    if (allArrayElements.Any())
                    {
                        // 인덱스별 개별 스키마를 위해 첫 번째 배열 사용
                        var firstArray = allArrayElements.First();
                        var arrayLayout = CalculateArrayLayoutWithIndexSchemas(array.Key, firstArray, allArrayElements);
                        layout.ArrayLayouts[array.Key] = arrayLayout;
                        // events 같은 복잡한 중첩 구조의 경우 실제 필요한 컬럼 수로 설정
                        if (HasComplexNestedStructure(arrayLayout) && arrayLayout.ElementCount == 1)
                        {
                            // 복잡한 단일 요소 배열의 경우, 중첩된 구조의 실제 컬럼 수를 사용
                            layout.TotalColumns += CalculateComplexArrayColumns(arrayLayout);
                        }
                        else
                        {
                            layout.TotalColumns += arrayLayout.TotalColumns;
                        }
                    }
                }
            }

            return layout;
        }

        private void OptimizeArrayColumns(DynamicArrayLayout layout)
        {
            // UnifiedSchema가 이미 모든 속성을 포함하고 있으므로 이를 사용
            if (layout.UnifiedSchema == null || !layout.UnifiedSchema.Any())
            {
                Logger.Warning("OptimizeArrayColumns: UnifiedSchema가 비어있음");
                return;
            }

            // 통합 스키마의 속성을 FirstAppearanceIndex 순서로 정렬
            var orderedProperties = layout.UnifiedSchema
                .OrderBy(p => p.Value.FirstAppearanceIndex)
                .Select(p => p.Key)
                .ToList();

            Logger.Debug($"OptimizeArrayColumns: 통합 스키마 속성 {orderedProperties.Count}개");
            Logger.Debug($"  속성 목록: {string.Join(", ", orderedProperties)}");

            layout.OrderedProperties = orderedProperties;

            // 각 요소의 컬럼 수를 통합 스키마 크기로 맞춤
            foreach (var element in layout.Elements)
            {
                element.RequiredColumns = orderedProperties.Count;
                element.UnifiedProperties = orderedProperties;
                
                // 속성 컬럼 맵 재계산
                element.PropertyColumnMap.Clear();
                for (int i = 0; i < orderedProperties.Count; i++)
                {
                    // 요소가 해당 속성을 가지고 있든 없든 모든 속성에 대해 컬럼 할당
                    element.PropertyColumnMap[orderedProperties[i]] = i;
                }
            }
        }

        private int CalculateActualUsedColumns(DynamicArrayLayout layout)
        {
            // 실제 사용된 컬럼 수를 계산합니다.
            int maxColumns = 0;
            foreach (var element in layout.Elements)
            {
                maxColumns = System.Math.Max(maxColumns, element.RequiredColumns);
            }
            return maxColumns * layout.ElementCount;
        }
        
        private bool HasComplexNestedStructure(DynamicArrayLayout layout)
        {
            // 복잡한 중첩 구조를 가진 배열인지 확인
            if (layout.UnifiedSchema == null)
                return false;
                
            foreach (var prop in layout.UnifiedSchema.Values)
            {
                // 객체나 배열을 포함하는 경우 복잡한 구조로 판단
                if (prop.IsObject && prop.ObjectProperties?.Count > 0)
                    return true;
                if (prop.IsArray && prop.ArrayPattern?.ElementProperties != null && prop.ArrayPattern.ElementProperties.Any())
                    return true;
            }
            
            return false;
        }
        
        private int CalculateComplexArrayColumns(DynamicArrayLayout layout)
        {
            // events 같은 복잡한 중첩 구조의 실제 컬럼 수 계산
            if (layout.UnifiedSchema == null || !layout.UnifiedSchema.Any())
                return 0;
                
            // 배열 자체가 차지하는 컬럼 수를 정확히 계산
            // 단일 요소 배열이고 복잡한 중첩 구조를 가진 경우
            if (layout.ElementCount == 1)
            {
                // 배열 속성들 중 가장 단순한 속성의 수를 기준으로 계산
                var simplePropsCount = layout.UnifiedSchema.Count(p => !p.Value.IsObject && !p.Value.IsArray);
                if (simplePropsCount > 0)
                {
                    return simplePropsCount;
                }
                
                // 모든 속성이 복잡한 경우 최소 컬럼 수 반환
                return Math.Min(2, layout.UnifiedSchema.Count);
            }
            
            return layout.TotalColumns;
        }
        
        private ArrayPattern AnalyzeArrayPattern(string name, YamlSequenceNode array)
        {
            var pattern = new ArrayPattern
            {
                Name = name,
                MaxSize = array.Children.Count,
                MinSize = array.Children.Count,
                ElementProperties = new Dictionary<string, PropertyPattern>()
            };
            
            // 배열 요소들의 스키마 분석
            var elementSchemas = new List<Dictionary<string, YamlNode>>();
            foreach (var element in array.Children)
            {
                if (element is YamlMappingNode mapping)
                {
                    var elemProps = new Dictionary<string, YamlNode>();
                    foreach (var kvp in mapping.Children)
                    {
                        elemProps[kvp.Key.ToString()] = kvp.Value;
                    }
                    elementSchemas.Add(elemProps);
                }
            }
            
            // 요소 속성들의 통합 스키마 생성
            if (elementSchemas.Any())
            {
                pattern.ElementProperties = CreateUnifiedSchemaFromElements(elementSchemas);
            }
            
            return pattern;
        }
        
        private List<string> ExtractObjectProperties(YamlMappingNode mapping)
        {
            var properties = new List<string>();
            foreach (var kvp in mapping.Children)
            {
                properties.Add(kvp.Key.ToString());
            }
            return properties;
        }

        /// <summary>
        /// 여러 배열을 통합하여 전체 스키마를 분석하기 위한 가상 배열 생성
        /// </summary>
        private YamlSequenceNode MergeArraysForAnalysis(List<YamlSequenceNode> arrays)
        {
            var mergedSequence = new YamlSequenceNode();
            
            Logger.Information($"MergeArraysForAnalysis: {arrays.Count}개 배열 통합 시작");
            
            // 모든 배열의 각 인덱스별로 통합 스키마 생성
            int maxLength = arrays.Max(a => a.Children.Count);
            Logger.Information($"  최대 배열 길이: {maxLength}");
            
            for (int i = 0; i < maxLength; i++)
            {
                // 각 인덱스에서 모든 배열의 요소를 수집
                var elementsAtIndex = new List<YamlMappingNode>();
                
                foreach (var array in arrays)
                {
                    if (i < array.Children.Count && array.Children[i] is YamlMappingNode mapping)
                    {
                        elementsAtIndex.Add(mapping);
                    }
                }
                
                if (elementsAtIndex.Any())
                {
                    Logger.Debug($"  인덱스 {i}: {elementsAtIndex.Count}개 요소 병합");
                    
                    // 해당 인덱스의 모든 속성을 통합한 요소 생성
                    var mergedElement = MergeElementsAtIndex(elementsAtIndex);
                    mergedSequence.Add(mergedElement);
                    
                    // 병합된 요소의 속성 로깅
                    var propNames = new List<string>();
                    foreach (var kvp in mergedElement.Children)
                    {
                        propNames.Add(kvp.Key.ToString());
                    }
                    Logger.Debug($"    병합된 속성: {string.Join(", ", propNames)}");
                }
            }
            
            return mergedSequence;
        }

        /// <summary>
        /// 특정 인덱스의 모든 요소를 통합하여 전체 속성을 포함하는 요소 생성
        /// </summary>
        private YamlMappingNode MergeElementsAtIndex(List<YamlMappingNode> elements)
        {
            var mergedMapping = new YamlMappingNode();
            var allProperties = new Dictionary<string, YamlNode>();
            
            // 모든 요소의 속성을 수집
            foreach (var element in elements)
            {
                foreach (var kvp in element.Children)
                {
                    var key = kvp.Key.ToString();
                    // 속성이 처음 나타나면 추가 (첫 번째 값 유지)
                    if (!allProperties.ContainsKey(key))
                    {
                        allProperties[key] = kvp.Value;
                    }
                }
            }
            
            // 통합된 속성들을 새 매핑에 추가
            foreach (var kvp in allProperties)
            {
                mergedMapping.Add(kvp.Key, kvp.Value);
            }
            
            return mergedMapping;
        }

        /// <summary>
        /// 인덱스별 개별 스키마를 사용하여 배열 레이아웃 계산
        /// </summary>
        private DynamicArrayLayout CalculateArrayLayoutWithIndexSchemas(
            string arrayPath, 
            YamlSequenceNode firstArray,
            List<YamlSequenceNode> allArrays)
        {
            var layout = new DynamicArrayLayout
            {
                ArrayPath = arrayPath,
                ElementCount = firstArray.Children.Count,
                Elements = new List<ElementLayout>(),
                OptimizeColumns = false  // 인덱스별 개별 스키마
            };

            Logger.Information($"CalculateArrayLayoutWithIndexSchemas: {arrayPath}, 첫 배열 요소 수={firstArray.Children.Count}");

            // 각 인덱스별로 모든 루트의 해당 인덱스 요소를 수집하여 통합 스키마 생성
            for (int i = 0; i < firstArray.Children.Count; i++)
            {
                var elementsAtIndex = new List<Dictionary<string, YamlNode>>();
                
                // 모든 배열에서 i번째 요소 수집
                foreach (var array in allArrays)
                {
                    if (i < array.Children.Count && array.Children[i] is YamlMappingNode mapping)
                    {
                        var elementProps = new Dictionary<string, YamlNode>();
                        foreach (var kvp in mapping.Children)
                        {
                            elementProps[kvp.Key.ToString()] = kvp.Value;
                        }
                        elementsAtIndex.Add(elementProps);
                    }
                }

                // 해당 인덱스의 통합 스키마 생성
                var indexSchema = CreateUnifiedSchemaFromElements(elementsAtIndex);
                var orderedProps = indexSchema.Keys.OrderBy(k => indexSchema[k].FirstAppearanceIndex).ToList();
                
                Logger.Debug($"  인덱스 {i}: {orderedProps.Count}개 속성 - {string.Join(", ", orderedProps)}");

                // 요소 레이아웃 생성
                var elementLayout = new ElementLayout
                {
                    Index = i,
                    Properties = orderedProps,
                    RequiredColumns = orderedProps.Count,
                    UnifiedProperties = orderedProps,
                    PropertyColumnMap = new Dictionary<string, int>()
                };

                // 속성별 컬럼 인덱스 매핑
                for (int j = 0; j < orderedProps.Count; j++)
                {
                    elementLayout.PropertyColumnMap[orderedProps[j]] = j;
                }

                layout.Elements.Add(elementLayout);
            }

            // 전체 컬럼 수 계산
            layout.TotalColumns = layout.Elements.Sum(e => e.RequiredColumns);
            layout.ActualUsedColumns = layout.TotalColumns;

            Logger.Information($"인덱스별 개별 스키마 완료: 전체 {layout.TotalColumns}개 컬럼");

            return layout;
        }
    }
}