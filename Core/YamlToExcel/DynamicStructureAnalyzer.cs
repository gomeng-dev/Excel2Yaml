using System;
using System.Collections.Generic;
using System.Linq;
using YamlDotNet.RepresentationModel;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    public class DynamicStructureAnalyzer
    {
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
        }

        public class ArrayPattern
        {
            public string Name { get; set; }
            public int MaxSize { get; set; }
            public int MinSize { get; set; }
            public double OccurrenceRatio { get; set; }
            public bool RequiresMultipleRows { get; set; }
            public bool HasVariableStructure { get; set; }
            public Dictionary<string, PropertyPattern> ElementProperties { get; set; }
        }

        public class StructurePattern
        {
            public PatternType Type { get; set; }
            public Dictionary<string, PropertyPattern> Properties { get; set; }
            public Dictionary<string, ArrayPattern> Arrays { get; set; }
            public int MaxDepth { get; set; }
            public double ConsistencyScore { get; set; }
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
            var elementSchemas = new List<Dictionary<string, object>>();
            
            // 모든 배열 요소 분석
            foreach (var element in array.Children)
            {
                var schema = ExtractElementSchema(element);
                elementSchemas.Add(schema);
            }

            // 동적 패턴 인식
            pattern.Properties = UnifySchemas(elementSchemas);
            pattern.Arrays = DetectNestedArrays(elementSchemas);
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
                }
                else if (value is YamlMappingNode)
                {
                    prop.IsObject = true;
                }

                pattern.Properties[key] = prop;
            }
        }

        private Dictionary<string, object> ExtractElementSchema(YamlNode element)
        {
            var schema = new Dictionary<string, object>();

            if (element is YamlMappingNode mapping)
            {
                foreach (var kvp in mapping.Children)
                {
                    var key = kvp.Key.ToString();
                    var value = kvp.Value;

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
                        
                        // 각 배열 요소의 스키마 추출
                        foreach (var child in sequence.Children)
                        {
                            if (child is YamlMappingNode childMapping)
                            {
                                var elementSchema = new Dictionary<string, object>();
                                foreach (var childKvp in childMapping.Children)
                                {
                                    elementSchema[childKvp.Key.ToString()] = childKvp.Value.ToString();
                                }
                                arrayInfo.Elements.Add(elementSchema);
                            }
                        }
                        
                        schema[key] = arrayInfo;
                    }
                    else if (value is YamlMappingNode)
                    {
                        schema[key] = new { Type = "Object" };
                    }
                }
            }

            return schema;
        }

        private Dictionary<string, PropertyPattern> UnifySchemas(List<Dictionary<string, object>> schemas)
        {
            var unified = new Dictionary<string, PropertyPattern>();
            
            // 모든 속성 수집 및 분석
            for (int i = 0; i < schemas.Count; i++)
            {
                var schema = schemas[i];
                foreach (var prop in schema.Where(p => !p.Key.StartsWith("_")))
                {
                    if (!unified.ContainsKey(prop.Key))
                    {
                        unified[prop.Key] = new PropertyPattern
                        {
                            Name = prop.Key,
                            OccurrenceCount = 0,
                            Types = new HashSet<Type>(),
                            FirstAppearanceIndex = i
                        };
                    }
                    
                    unified[prop.Key].OccurrenceCount++;
                    unified[prop.Key].Types.Add(prop.Value?.GetType() ?? typeof(object));

                    // 배열이나 객체 타입 감지
                    if (prop.Value is Dictionary<string, object> dict)
                    {
                        var type = dict.ContainsKey("Type") ? dict["Type"].ToString() : "";
                        if (type == "Array")
                            unified[prop.Key].IsArray = true;
                        else if (type == "Object")
                            unified[prop.Key].IsObject = true;
                    }
                }
            }

            // 출현 비율 계산
            foreach (var prop in unified.Values)
            {
                prop.OccurrenceRatio = (double)prop.OccurrenceCount / schemas.Count;
                prop.IsRequired = prop.OccurrenceRatio > 0.8; // 80% 이상 출현시 필수
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

        private ArrayPattern AnalyzeArray(string name, YamlSequenceNode array)
        {
            var pattern = new ArrayPattern
            {
                Name = name,
                MaxSize = array.Children.Count,
                MinSize = array.Children.Count,
                ElementProperties = new Dictionary<string, PropertyPattern>()
            };

            // 배열 요소들의 스키마 분석
            var elementSchemas = new List<Dictionary<string, object>>();
            foreach (var element in array.Children)
            {
                var schema = ExtractElementSchema(element);
                elementSchemas.Add(schema);
            }

            if (elementSchemas.Any())
            {
                pattern.ElementProperties = UnifySchemas(elementSchemas);
            }

            pattern.RequiresMultipleRows = pattern.MaxSize > 5;

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
        
        private class ArrayInfo
        {
            public bool IsArray { get; set; }
            public int ElementCount { get; set; }
            public List<Dictionary<string, object>> Elements { get; set; }
        }
    }
}