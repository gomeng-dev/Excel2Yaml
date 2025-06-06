using System;
using System.Collections.Generic;
using System.Linq;
using static ExcelToYamlAddin.Core.YamlToExcel.DynamicStructureAnalyzer;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    public class DynamicPropertyOrderer
    {
        public class PropertyGrouping
        {
            public List<string> Properties { get; set; }
            public double Strength { get; set; }
        }

        public List<string> DeterminePropertyOrder(Dictionary<string, PropertyPattern> properties)
        {
            // 동적 우선순위 계산 - 하드코딩 없음
            return properties
                .OrderByDescending(p => p.Value.OccurrenceRatio)  // 출현 빈도
                .ThenBy(p => p.Value.FirstAppearanceIndex)        // 첫 등장 순서
                .ThenBy(p => p.Key.Length)                        // 이름 길이 (짧은 것 우선)
                .ThenBy(p => p.Key)                               // 알파벳 순
                .Select(p => p.Key)
                .ToList();
        }

        public List<string> OptimizeForHorizontalLayout(
            Dictionary<string, PropertyPattern> properties,
            List<Dictionary<string, object>> samples)
        {
            // 샘플 데이터 분석을 통한 최적화
            var groupings = AnalyzePropertyGroupings(samples);
            var ordered = new List<string>();

            // 함께 나타나는 속성들을 그룹화
            foreach (var group in groupings.OrderByDescending(g => g.Strength))
            {
                var groupProperties = group.Properties
                    .Where(p => properties.ContainsKey(p))
                    .OrderByDescending(p => properties[p].OccurrenceRatio)
                    .ToList();
                ordered.AddRange(groupProperties);
            }

            // 그룹에 속하지 않은 속성들 추가
            var ungrouped = properties.Keys.Except(ordered);
            ordered.AddRange(DeterminePropertyOrder(
                properties.Where(p => ungrouped.Contains(p.Key))
                         .ToDictionary(p => p.Key, p => p.Value)));

            return ordered;
        }

        private List<PropertyGrouping> AnalyzePropertyGroupings(
            List<Dictionary<string, object>> samples)
        {
            var cooccurrence = new Dictionary<(string, string), int>();
            
            // 속성 동시 출현 분석
            foreach (var sample in samples)
            {
                var props = sample.Keys.ToList();
                for (int i = 0; i < props.Count; i++)
                {
                    for (int j = i + 1; j < props.Count; j++)
                    {
                        var pair = OrderPair(props[i], props[j]);
                        if (!cooccurrence.ContainsKey(pair))
                            cooccurrence[pair] = 0;
                        cooccurrence[pair]++;
                    }
                }
            }

            // 강한 연관성을 가진 속성 그룹 생성
            return CreatePropertyGroups(cooccurrence, samples.Count);
        }

        private (string, string) OrderPair(string a, string b)
        {
            return string.Compare(a, b, StringComparison.Ordinal) < 0 ? (a, b) : (b, a);
        }

        private List<PropertyGrouping> CreatePropertyGroups(
            Dictionary<(string, string), int> cooccurrence, 
            int sampleCount)
        {
            var groups = new List<PropertyGrouping>();
            var processedProperties = new HashSet<string>();
            
            // 강한 연결을 가진 속성들을 그룹화
            var strongPairs = cooccurrence
                .Where(kvp => kvp.Value > sampleCount * 0.7) // 70% 이상 함께 출현
                .OrderByDescending(kvp => kvp.Value)
                .ToList();

            foreach (var pair in strongPairs)
            {
                var (prop1, prop2) = pair.Key;
                
                // 이미 처리된 속성인지 확인
                if (processedProperties.Contains(prop1) || processedProperties.Contains(prop2))
                    continue;

                // 새 그룹 생성 또는 기존 그룹에 추가
                var existingGroup = groups.FirstOrDefault(g => 
                    g.Properties.Contains(prop1) || g.Properties.Contains(prop2));
                
                if (existingGroup != null)
                {
                    if (!existingGroup.Properties.Contains(prop1))
                        existingGroup.Properties.Add(prop1);
                    if (!existingGroup.Properties.Contains(prop2))
                        existingGroup.Properties.Add(prop2);
                }
                else
                {
                    groups.Add(new PropertyGrouping
                    {
                        Properties = new List<string> { prop1, prop2 },
                        Strength = (double)pair.Value / sampleCount
                    });
                }

                processedProperties.Add(prop1);
                processedProperties.Add(prop2);
            }

            return groups;
        }

        public List<string> OrderPropertiesForArrayElement(
            Dictionary<string, PropertyPattern> properties,
            ArrayPattern arrayPattern)
        {
            // 배열 요소의 속성 순서 결정
            var elementProps = arrayPattern.ElementProperties ?? new Dictionary<string, PropertyPattern>();
            
            // 배열 요소에 특화된 순서 결정
            return elementProps
                .OrderByDescending(p => p.Value.IsRequired)        // 필수 속성 우선
                .ThenByDescending(p => p.Value.OccurrenceRatio)   // 출현 빈도
                .ThenBy(p => IsIdentifierProperty(p.Key))         // ID류 속성 우선
                .ThenBy(p => p.Key.Length)                         // 짧은 이름 우선
                .ThenBy(p => p.Key)                                // 알파벳 순
                .Select(p => p.Key)
                .ToList();
        }

        private bool IsIdentifierProperty(string propertyName)
        {
            var lowerName = propertyName.ToLower();
            // ID나 식별자로 보이는 속성들을 우선 배치
            return lowerName.Contains("id") || 
                   lowerName.Contains("key") || 
                   lowerName.Contains("name") ||
                   lowerName.Contains("code");
        }

        public Dictionary<string, int> CreatePropertyPriorityMap(
            List<string> orderedProperties)
        {
            var priorityMap = new Dictionary<string, int>();
            for (int i = 0; i < orderedProperties.Count; i++)
            {
                priorityMap[orderedProperties[i]] = i;
            }
            return priorityMap;
        }
    }
}