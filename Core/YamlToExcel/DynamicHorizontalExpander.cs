using System.Collections.Generic;
using System.Linq;
using YamlDotNet.RepresentationModel;
using static ExcelToYamlAddin.Core.YamlToExcel.DynamicStructureAnalyzer;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    public class DynamicHorizontalExpander
    {
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
                OptimizeColumns = true
            };

            // 배열 요소들의 모든 속성 수집
            var allElementProperties = CollectAllElementProperties(array);
            
            // 통합 스키마가 없으면 생성
            if (!layout.UnifiedSchema.Any())
            {
                layout.UnifiedSchema = CreateUnifiedSchemaFromElements(allElementProperties);
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

            // 컬럼 최적화: 가변 속성을 가진 배열의 경우
            if (layout.OptimizeColumns)
            {
                OptimizeArrayColumns(layout);
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

            for (int i = 0; i < elements.Count; i++)
            {
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
                    }

                    schema[prop.Key].OccurrenceCount++;
                    schema[prop.Key].Types.Add(prop.Value.GetType());

                    if (prop.Value is YamlSequenceNode)
                        schema[prop.Key].IsArray = true;
                    else if (prop.Value is YamlMappingNode)
                        schema[prop.Key].IsObject = true;
                }
            }

            // 출현 비율 계산
            foreach (var prop in schema.Values)
            {
                prop.OccurrenceRatio = (double)prop.OccurrenceCount / elements.Count;
                prop.IsRequired = prop.OccurrenceRatio > 0.8;
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

            elementLayout.RequiredColumns = columnOffset;

            return elementLayout;
        }

        private int CalculateTotalColumns(DynamicArrayLayout layout)
        {
            if (!layout.Elements.Any())
                return 0;

            // 모든 요소가 통합 스키마를 따르므로 최대 컬럼 수 사용
            var baseColumns = layout.OrderedProperties.Count;
            
            // 각 요소의 추가 속성 고려
            var maxExtraColumns = 0;
            foreach (var element in layout.Elements)
            {
                var extraCount = element.Properties.Count(p => !layout.OrderedProperties.Contains(p));
                maxExtraColumns = System.Math.Max(maxExtraColumns, extraCount);
            }

            // 각 요소당 필요한 컬럼 수
            var columnsPerElement = baseColumns + maxExtraColumns;

            // 전체 컬럼 수 = 요소 수 × 요소당 컬럼 수
            return layout.ElementCount * columnsPerElement;
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
                    // 더미 시퀀스 생성 (실제 데이터에서 배열 찾기)
                    YamlSequenceNode arrayNode = null;
                    if (rootSequence.Children.Count > 0 && rootSequence.Children[0] is YamlMappingNode firstElement)
                    {
                        foreach (var prop in firstElement.Children)
                        {
                            if (prop.Key.ToString() == array.Key && prop.Value is YamlSequenceNode seq)
                            {
                                arrayNode = seq;
                                break;
                            }
                        }
                    }

                    if (arrayNode != null)
                    {
                        var arrayLayout = CalculateArrayLayout(array.Key, arrayNode, array.Value.ElementProperties);
                        layout.ArrayLayouts[array.Key] = arrayLayout;
                        layout.TotalColumns += arrayLayout.TotalColumns;
                    }
                }
            }

            return layout;
        }

        private void OptimizeArrayColumns(DynamicArrayLayout layout)
        {
            // 모든 요소에서 사용되는 고유 속성 수집
            var allProperties = new HashSet<string>();
            foreach (var element in layout.Elements)
            {
                allProperties.UnionWith(element.Properties);
            }

            // 통합 스키마 생성
            var unifiedProperties = allProperties.OrderBy(p => 
            {
                // 우선순위: 1) 스키마에 있는 속성, 2) 출현 빈도, 3) 알파벳 순
                if (layout.UnifiedSchema.ContainsKey(p))
                    return layout.UnifiedSchema[p].FirstAppearanceIndex;
                return int.MaxValue;
            }).ToList();

            // level, material, mineral은 앞쪽에, damage/addDamage는 뒤쪽에 배치
            var orderedProperties = new List<string>();
            
            // 공통 속성 먼저
            var commonProps = unifiedProperties.Where(p => 
                layout.Elements.All(e => e.Properties.Contains(p))).ToList();
            orderedProperties.AddRange(commonProps);
            
            // 가변 속성
            var variableProps = unifiedProperties.Except(commonProps).ToList();
            orderedProperties.AddRange(variableProps);

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
                    if (element.Properties.Contains(orderedProperties[i]))
                    {
                        element.PropertyColumnMap[orderedProperties[i]] = i;
                    }
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
    }
}