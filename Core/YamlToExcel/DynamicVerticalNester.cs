using System.Collections.Generic;
using System.Linq;
using YamlDotNet.RepresentationModel;
using static ExcelToYamlAddin.Core.YamlToExcel.DynamicStructureAnalyzer;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    public class DynamicVerticalNester
    {
        public class RowGroup
        {
            public string GroupKey { get; set; }
            public List<Dictionary<string, object>> Rows { get; set; }
            public int StartRow { get; set; }
            public int EndRow { get; set; }
        }

        public class VerticalLayout
        {
            public List<RowGroup> RowGroups { get; set; }
            public Dictionary<string, int> ColumnMapping { get; set; }
            public bool RequiresMerging { get; set; }
            public string MergeKey { get; set; }
            public int MaxDepth { get; set; }
            public List<string> OrderedColumns { get; set; }
        }

        private readonly DynamicPropertyOrderer _propertyOrderer;

        public DynamicVerticalNester()
        {
            _propertyOrderer = new DynamicPropertyOrderer();
        }

        public VerticalLayout GenerateVerticalLayout(
            StructurePattern pattern,
            List<YamlNode> items)
        {
            var layout = new VerticalLayout
            {
                RowGroups = new List<RowGroup>(),
                ColumnMapping = new Dictionary<string, int>()
            };

            // 병합 키 자동 감지
            layout.MergeKey = DetectMergeKey(items);
            layout.RequiresMerging = !string.IsNullOrEmpty(layout.MergeKey);

            // 속성 순서 결정
            layout.OrderedColumns = _propertyOrderer.DeterminePropertyOrder(pattern.Properties);

            // 컬럼 매핑 생성
            int currentColumn = 1; // 첫 컬럼은 ^ 마커용
            foreach (var prop in layout.OrderedColumns)
            {
                layout.ColumnMapping[prop] = currentColumn++;
            }

            // 중첩 구조를 위한 추가 컬럼 할당
            AssignNestedColumns(pattern, layout, ref currentColumn);

            // 행 그룹 생성
            if (layout.RequiresMerging)
            {
                CreateMergedRowGroups(items, layout);
            }
            else
            {
                CreateSimpleRowGroups(items, layout);
            }

            return layout;
        }

        private string DetectMergeKey(List<YamlNode> items)
        {
            // 중복 값을 가진 속성 찾기
            var propertyValues = new Dictionary<string, List<object>>();
            
            foreach (var item in items.OfType<YamlMappingNode>())
            {
                foreach (var kvp in item.Children)
                {
                    var key = kvp.Key.ToString();
                    var value = kvp.Value.ToString();
                    
                    if (!propertyValues.ContainsKey(key))
                        propertyValues[key] = new List<object>();
                    
                    propertyValues[key].Add(value);
                }
            }

            // 중복 값이 있는 속성 찾기
            foreach (var prop in propertyValues)
            {
                var uniqueValues = prop.Value.Distinct().Count();
                if (uniqueValues < prop.Value.Count && uniqueValues < items.Count * 0.5)
                {
                    return prop.Key; // 병합 키 후보
                }
            }

            return null;
        }

        private void AssignNestedColumns(StructurePattern pattern, VerticalLayout layout, ref int currentColumn)
        {
            // 중첩 배열과 객체를 위한 추가 컬럼 할당
            foreach (var array in pattern.Arrays.Values)
            {
                if (array.RequiresMultipleRows)
                {
                    // 수직 확장이 필요한 배열
                    layout.ColumnMapping[$"{array.Name}$[]"] = currentColumn++;
                    
                    // 배열 요소의 속성들
                    if (array.ElementProperties != null)
                    {
                        foreach (var prop in array.ElementProperties.Keys)
                        {
                            if (!layout.ColumnMapping.ContainsKey($"{array.Name}.{prop}"))
                            {
                                layout.ColumnMapping[$"{array.Name}.{prop}"] = currentColumn++;
                            }
                        }
                    }
                }
            }
        }

        private void CreateMergedRowGroups(List<YamlNode> items, VerticalLayout layout)
        {
            var groups = new Dictionary<string, RowGroup>();
            
            foreach (var item in items.OfType<YamlMappingNode>())
            {
                var mergeValue = GetPropertyValue(item, layout.MergeKey);
                var groupKey = mergeValue?.ToString() ?? "null";
                
                if (!groups.ContainsKey(groupKey))
                {
                    groups[groupKey] = new RowGroup
                    {
                        GroupKey = groupKey,
                        Rows = new List<Dictionary<string, object>>()
                    };
                }
                
                var rowData = ExtractRowData(item, layout);
                groups[groupKey].Rows.Add(rowData);
            }
            
            layout.RowGroups.AddRange(groups.Values);
        }

        private void CreateSimpleRowGroups(List<YamlNode> items, VerticalLayout layout)
        {
            foreach (var item in items)
            {
                var group = new RowGroup
                {
                    GroupKey = null,
                    Rows = new List<Dictionary<string, object>>()
                };
                
                if (item is YamlMappingNode mapping)
                {
                    var rowData = ExtractRowData(mapping, layout);
                    group.Rows.Add(rowData);
                    
                    // 중첩 구조 처리
                    ExtractNestedRows(mapping, layout, group);
                }
                
                layout.RowGroups.Add(group);
            }
        }

        private Dictionary<string, object> ExtractRowData(YamlMappingNode node, VerticalLayout layout)
        {
            var rowData = new Dictionary<string, object>();
            
            foreach (var kvp in node.Children)
            {
                var key = kvp.Key.ToString();
                
                if (kvp.Value is YamlScalarNode scalar)
                {
                    rowData[key] = scalar.Value;
                }
                else if (kvp.Value is YamlSequenceNode)
                {
                    rowData[key] = "[Array]";
                }
                else if (kvp.Value is YamlMappingNode)
                {
                    rowData[key] = "[Object]";
                }
            }
            
            return rowData;
        }

        private void ExtractNestedRows(YamlMappingNode parent, VerticalLayout layout, RowGroup group)
        {
            foreach (var kvp in parent.Children)
            {
                if (kvp.Value is YamlSequenceNode array)
                {
                    var arrayName = kvp.Key.ToString();
                    
                    foreach (var element in array.Children)
                    {
                        if (element is YamlMappingNode elementMapping)
                        {
                            var nestedRow = new Dictionary<string, object>();
                            nestedRow["_parent"] = arrayName;
                            
                            foreach (var prop in elementMapping.Children)
                            {
                                var propKey = $"{arrayName}.{prop.Key}";
                                nestedRow[propKey] = GetNodeValue(prop.Value);
                            }
                            
                            group.Rows.Add(nestedRow);
                        }
                    }
                }
            }
        }

        private object GetPropertyValue(YamlMappingNode node, string propertyName)
        {
            var key = new YamlScalarNode(propertyName);
            if (node.Children.ContainsKey(key))
            {
                return GetNodeValue(node.Children[key]);
            }
            return null;
        }

        private object GetNodeValue(YamlNode node)
        {
            if (node is YamlScalarNode scalar)
                return scalar.Value;
            else if (node is YamlSequenceNode)
                return "[Array]";
            else if (node is YamlMappingNode)
                return "[Object]";
            else
                return null;
        }

        public int CalculateRequiredRows(VerticalLayout layout)
        {
            int totalRows = 0;
            
            foreach (var group in layout.RowGroups)
            {
                totalRows += group.Rows.Count;
                
                // 그룹 간 구분을 위한 추가 행
                if (layout.RequiresMerging && group != layout.RowGroups.Last())
                {
                    totalRows++; // 빈 행 추가
                }
            }
            
            return totalRows;
        }

        public void AssignRowNumbers(VerticalLayout layout, int startRow)
        {
            int currentRow = startRow;
            
            foreach (var group in layout.RowGroups)
            {
                group.StartRow = currentRow;
                currentRow += group.Rows.Count;
                group.EndRow = currentRow - 1;
                
                // 그룹 간 구분
                if (layout.RequiresMerging && group != layout.RowGroups.Last())
                {
                    currentRow++; // 빈 행
                }
            }
        }
    }
}