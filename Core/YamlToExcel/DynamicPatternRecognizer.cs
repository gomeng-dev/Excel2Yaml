using System.Linq;
using static ExcelToYamlAddin.Core.YamlToExcel.DynamicStructureAnalyzer;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    public class DynamicPatternRecognizer
    {
        public enum LayoutStrategy
        {
            Simple,              // 단순 구조
            VerticalNesting,     // 수직 중첩 (각 요소가 여러 행)
            HorizontalExpansion, // 수평 확장 (배열을 컬럼으로)
            Mixed               // 혼합 전략
        }

        public class StructureMetrics
        {
            public bool IsSimpleStructure { get; set; }
            public bool HasLargeNestedArrays { get; set; }
            public int ArrayElementCount { get; set; }
            public bool HasVariableDepth { get; set; }
            public bool HasOptionalNesting { get; set; }
            public bool HasVariableArrayProperties { get; set; }
            public int TotalProperties { get; set; }
            public int TotalArrays { get; set; }
            public double AverageArraySize { get; set; }
        }

        public LayoutStrategy DetermineStrategy(StructurePattern pattern)
        {
            // 동적 전략 결정 - 하드코딩 없음
            var metrics = CalculateMetrics(pattern);
            
            if (metrics.IsSimpleStructure)
            {
                return LayoutStrategy.Simple;
            }
            
            // Weapons.yaml처럼 루트가 배열이고 중첩 배열이 있는 경우
            if (pattern.Type == PatternType.RootArray && pattern.Arrays.Any())
            {
                return LayoutStrategy.HorizontalExpansion;
            }
            
            if (metrics.HasLargeNestedArrays && metrics.ArrayElementCount > 5)
            {
                return LayoutStrategy.HorizontalExpansion;
            }
            
            if (metrics.HasVariableDepth || metrics.HasOptionalNesting)
            {
                return LayoutStrategy.VerticalNesting;
            }
            
            return LayoutStrategy.Mixed;
        }

        private StructureMetrics CalculateMetrics(StructurePattern pattern)
        {
            var metrics = new StructureMetrics
            {
                IsSimpleStructure = pattern.MaxDepth <= 2 && !pattern.Arrays.Any(),
                HasLargeNestedArrays = pattern.Arrays.Any(a => a.Value.MaxSize > 3),
                ArrayElementCount = pattern.Arrays.Sum(a => a.Value.MaxSize),
                HasVariableDepth = pattern.Properties.Any(p => p.Value.OccurrenceRatio < 0.5),
                HasOptionalNesting = pattern.Arrays.Any(a => a.Value.OccurrenceRatio < 1.0),
                HasVariableArrayProperties = pattern.Arrays.Any(a => a.Value.HasVariableProperties),
                TotalProperties = pattern.Properties.Count,
                TotalArrays = pattern.Arrays.Count
            };

            // 평균 배열 크기 계산
            if (pattern.Arrays.Any())
            {
                metrics.AverageArraySize = pattern.Arrays.Average(a => a.Value.MaxSize);
            }

            return metrics;
        }

        public StructureMetrics GetMetrics(StructurePattern pattern)
        {
            return CalculateMetrics(pattern);
        }

        public string GetStrategyDescription(LayoutStrategy strategy)
        {
            switch (strategy)
            {
                case LayoutStrategy.Simple:
                    return "Simple structure with direct property mapping";
                case LayoutStrategy.VerticalNesting:
                    return "Complex structure requiring vertical expansion";
                case LayoutStrategy.HorizontalExpansion:
                    return "Array-heavy structure requiring horizontal expansion";
                case LayoutStrategy.Mixed:
                    return "Mixed structure requiring combined strategies";
                default:
                    return "Unknown strategy";
            }
        }

        public bool RequiresSchemaOptimization(StructurePattern pattern)
        {
            var metrics = CalculateMetrics(pattern);
            
            // 스키마 최적화가 필요한 경우
            return metrics.ArrayElementCount > 10 || 
                   metrics.TotalProperties > 20 ||
                   pattern.MaxDepth > 4;
        }

        public bool ShouldMergeArrayElements(ArrayPattern arrayPattern)
        {
            // 배열 요소를 병합해야 하는 경우
            return arrayPattern.MaxSize > 8 || 
                   (arrayPattern.ElementProperties != null && 
                    arrayPattern.ElementProperties.Count > 5);
        }

        public int EstimateRequiredColumns(StructurePattern pattern, LayoutStrategy strategy)
        {
            var baseColumns = pattern.Properties.Count(p => !p.Value.IsArray);
            
            switch (strategy)
            {
                case LayoutStrategy.Simple:
                    return baseColumns + 1; // +1 for ^ marker
                    
                case LayoutStrategy.HorizontalExpansion:
                    var arrayColumns = pattern.Arrays.Sum(a => 
                        a.Value.MaxSize * (a.Value.ElementProperties?.Count ?? 1));
                    return baseColumns + arrayColumns + 1;
                    
                case LayoutStrategy.VerticalNesting:
                    var maxNestedColumns = pattern.Arrays.Max(a => 
                        a.Value.ElementProperties?.Count ?? 0);
                    return System.Math.Max(baseColumns, maxNestedColumns) + 1;
                    
                case LayoutStrategy.Mixed:
                    // 혼합 전략은 더 복잡한 계산 필요
                    return EstimateComplexLayoutColumns(pattern);
                    
                default:
                    return baseColumns + 1;
            }
        }

        private int EstimateComplexLayoutColumns(StructurePattern pattern)
        {
            var columns = pattern.Properties.Count + 1; // 기본 속성 + ^ 마커
            
            // 각 배열에 대해 최적 전략 결정
            foreach (var array in pattern.Arrays.Values)
            {
                if (ShouldMergeArrayElements(array))
                {
                    // 수평 확장
                    columns += array.MaxSize * (array.ElementProperties?.Count ?? 1);
                }
                else
                {
                    // 수직 확장 - 최대 너비만 고려
                    columns = System.Math.Max(columns, array.ElementProperties?.Count ?? 0);
                }
            }
            
            return columns;
        }
    }
}