using System;
using System.Collections.Generic;
using System.Linq;
using YamlDotNet.RepresentationModel;
using ExcelToYamlAddin.Logging;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    /// <summary>
    /// YAML 구조를 분석하여 Excel 스키마 트리를 생성하는 역 스키마 빌더 (개선된 버전)
    /// </summary>
    public class ReverseSchemeBuilder
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<ReverseSchemeBuilder>();

        public class ExcelSchemeNode
        {
            public string Key { get; set; }
            public string SchemeMarker { get; set; }
            public SchemeNode.SchemeNodeType NodeType { get; set; }
            public int RowIndex { get; set; }
            public int ColumnIndex { get; set; }
            public int ColumnSpan { get; set; } = 1;
            public ExcelSchemeNode Parent { get; set; }
            public List<ExcelSchemeNode> Children { get; set; } = new List<ExcelSchemeNode>();
            public bool IsMergedCell { get; set; } = false;
            public string OriginalYamlPath { get; set; }
        }

        public class SchemeBuildResult
        {
            public ExcelSchemeNode RootNode { get; set; }
            public int TotalRows { get; set; }
            public int TotalColumns { get; set; }
            public Dictionary<int, List<ExcelSchemeNode>> RowMap { get; set; } = new Dictionary<int, List<ExcelSchemeNode>>();
            public List<(int row, int col, int colspan)> MergedCells { get; set; } = new List<(int, int, int)>();
            public Dictionary<string, int> ColumnMappings { get; set; } = new Dictionary<string, int>();
        }

        private int currentRow = 2;
        private int currentColumn = 1;
        private int maxColumn = 0;

        public SchemeBuildResult BuildSchemaTree(YamlNode yamlRoot)
        {
            Logger.Information("========== 스키마 빌드 시작 (v2) ==========");
            
            currentRow = 2;
            currentColumn = 1;
            maxColumn = 0;
            
            var result = new SchemeBuildResult();
            
            // 루트 노드 처리
            result.RootNode = ProcessRootNode(yamlRoot);
            
            // 행별로 노드 매핑
            BuildRowMap(result.RootNode, result.RowMap);
            
            // 병합 셀 정보 계산
            CalculateMergedCells(result);
            
            // 컬럼 매핑 생성
            BuildColumnMappings(result.RootNode, result.ColumnMappings);
            
            // $scheme_end 행 추가
            result.TotalRows = currentRow;
            result.TotalColumns = maxColumn;
            
            Logger.Information($"스키마 빌드 완료: 총 {result.TotalRows}행, {result.TotalColumns}열");
            
            return result;
        }

        private ExcelSchemeNode ProcessRootNode(YamlNode node)
        {
            if (node is YamlSequenceNode rootSequence)
            {
                // 루트가 배열인 경우
                var rootArrayNode = new ExcelSchemeNode
                {
                    Key = "",
                    SchemeMarker = "$[]",
                    NodeType = SchemeNode.SchemeNodeType.ARRAY,
                    RowIndex = currentRow,
                    ColumnIndex = 1,
                    OriginalYamlPath = ""
                };

                // 배열의 첫 번째 요소로 전체 구조 분석
                if (rootSequence.Children.Count > 0 && rootSequence.Children[0] is YamlMappingNode firstMapping)
                {
                    var columns = CalculateObjectColumns(firstMapping);
                    rootArrayNode.ColumnSpan = columns;
                    rootArrayNode.IsMergedCell = columns > 1;
                    maxColumn = columns;
                    
                    currentRow++;
                    
                    // ^ 마커와 ${} 추가
                    var caretNode = new ExcelSchemeNode
                    {
                        Key = "^",
                        SchemeMarker = "",
                        NodeType = SchemeNode.SchemeNodeType.IGNORE,
                        Parent = rootArrayNode,
                        RowIndex = currentRow,
                        ColumnIndex = 1
                    };
                    rootArrayNode.Children.Add(caretNode);
                    
                    var elementNode = new ExcelSchemeNode
                    {
                        Key = "",
                        SchemeMarker = "${}",
                        NodeType = SchemeNode.SchemeNodeType.MAP,
                        Parent = rootArrayNode,
                        RowIndex = currentRow,
                        ColumnIndex = 2,
                        ColumnSpan = columns - 1,
                        IsMergedCell = true,
                        OriginalYamlPath = "[0]"
                    };
                    rootArrayNode.Children.Add(elementNode);
                    
                    currentRow++;
                    
                    // Stage 처리
                    ProcessObjectProperties(elementNode, firstMapping, 2, "[0]");
                }
                
                return rootArrayNode;
            }
            else if (node is YamlMappingNode rootMapping)
            {
                // 루트가 객체인 경우
                var rootObjectNode = new ExcelSchemeNode
                {
                    Key = "",
                    SchemeMarker = "${}",
                    NodeType = SchemeNode.SchemeNodeType.MAP,
                    RowIndex = currentRow,
                    ColumnIndex = 1,
                    OriginalYamlPath = ""
                };
                
                var columns = CalculateObjectColumns(rootMapping);
                rootObjectNode.ColumnSpan = columns;
                rootObjectNode.IsMergedCell = columns > 1;
                maxColumn = columns;
                
                currentRow++;
                ProcessObjectProperties(rootObjectNode, rootMapping, 1, "");
                
                return rootObjectNode;
            }
            
            throw new InvalidOperationException("지원하지 않는 루트 노드 타입");
        }

        private void ProcessObjectProperties(ExcelSchemeNode parentNode, YamlMappingNode mapping, int startColumn, string yamlPath)
        {
            int col = startColumn;
            
            // 루트 배열의 요소인 경우 ^ 마커 추가
            if (parentNode.Parent != null && 
                parentNode.Parent.NodeType == SchemeNode.SchemeNodeType.ARRAY &&
                parentNode.Parent.Parent == null &&
                startColumn > 1)
            {
                var caretNode = new ExcelSchemeNode
                {
                    Key = "^",
                    SchemeMarker = "",
                    NodeType = SchemeNode.SchemeNodeType.IGNORE,
                    Parent = parentNode,
                    RowIndex = currentRow,
                    ColumnIndex = 1
                };
                parentNode.Children.Add(caretNode);
            }
            
            foreach (var kvp in mapping.Children)
            {
                var key = kvp.Key.ToString();
                var value = kvp.Value;
                var propPath = string.IsNullOrEmpty(yamlPath) ? key : $"{yamlPath}.{key}";
                
                if (value is YamlScalarNode)
                {
                    // 단순 속성
                    var propNode = new ExcelSchemeNode
                    {
                        Key = key,
                        SchemeMarker = "",
                        NodeType = SchemeNode.SchemeNodeType.PROPERTY,
                        Parent = parentNode,
                        RowIndex = currentRow,
                        ColumnIndex = col++,
                        OriginalYamlPath = propPath
                    };
                    parentNode.Children.Add(propNode);
                }
                else if (value is YamlSequenceNode sequence)
                {
                    // 배열 속성 처리
                    var arrayColumns = CalculateArrayColumns(sequence);
                    
                    var arrayNode = new ExcelSchemeNode
                    {
                        Key = key,
                        SchemeMarker = "$[]",
                        NodeType = SchemeNode.SchemeNodeType.ARRAY,
                        Parent = parentNode,
                        RowIndex = currentRow,
                        ColumnIndex = col,
                        ColumnSpan = arrayColumns,
                        IsMergedCell = arrayColumns > 1,
                        OriginalYamlPath = propPath
                    };
                    parentNode.Children.Add(arrayNode);
                    
                    if (sequence.Children.Count > 0 && sequence.Children[0] is YamlMappingNode)
                    {
                        // 배열 요소가 객체인 경우
                        currentRow++;
                        ProcessArrayElements(arrayNode, sequence, col, propPath);
                        currentRow--;
                    }
                    
                    col += arrayColumns;
                }
                else if (value is YamlMappingNode childMapping)
                {
                    // 중첩 객체 속성
                    var objectColumns = CalculateObjectColumns(childMapping);
                    
                    var objectNode = new ExcelSchemeNode
                    {
                        Key = key,
                        SchemeMarker = "${}",
                        NodeType = SchemeNode.SchemeNodeType.MAP,
                        Parent = parentNode,
                        RowIndex = currentRow,
                        ColumnIndex = col,
                        ColumnSpan = objectColumns,
                        IsMergedCell = objectColumns > 1,
                        OriginalYamlPath = propPath
                    };
                    parentNode.Children.Add(objectNode);
                    
                    currentRow++;
                    ProcessObjectProperties(objectNode, childMapping, col, propPath);
                    currentRow--;
                    
                    col += objectColumns;
                }
            }
            
            currentRow++;
            maxColumn = Math.Max(maxColumn, col - 1);
        }

        private void ProcessArrayElements(ExcelSchemeNode arrayNode, YamlSequenceNode sequence, int startColumn, string yamlPath)
        {
            if (sequence.Children.Count == 0)
                return;
                
            // 배열 요소가 객체인 경우만 처리
            if (sequence.Children[0] is YamlMappingNode)
            {
                // 모든 배열 요소의 속성을 수집하여 통합된 스키마 생성
                var mergedStructure = MergeArrayElementStructures(sequence);
                var singleElementColumns = CalculateObjectColumns(mergedStructure);
                
                // 배열의 실제 요소 수 계산
                int displayCount = sequence.Children.Count;
                
                // 각 배열 요소를 위한 ${} 마커를 한 행에 나란히 생성
                for (int i = 0; i < displayCount; i++)
                {
                    var elementNode = new ExcelSchemeNode
                    {
                        Key = "",
                        SchemeMarker = "${}",
                        NodeType = SchemeNode.SchemeNodeType.MAP,
                        Parent = arrayNode,
                        RowIndex = currentRow,
                        ColumnIndex = startColumn + (i * singleElementColumns),
                        ColumnSpan = singleElementColumns,
                        IsMergedCell = singleElementColumns > 1,
                        OriginalYamlPath = $"{yamlPath}[*]"  // 모든 요소가 동일한 구조
                    };
                    arrayNode.Children.Add(elementNode);
                }
                
                // 통합된 구조를 한 번만 처리 (모든 요소가 동일한 구조를 공유)
                currentRow++;
                var savedRow = currentRow;
                
                // 첫 번째 요소 위치에서 통합 구조 처리
                ProcessObjectProperties(arrayNode.Children[0], mergedStructure, startColumn, $"{yamlPath}[*]");
                
                // 나머지 요소들에 대해 동일한 구조를 동일한 행에 복사
                for (int i = 1; i < displayCount; i++)
                {
                    currentRow = savedRow; // 같은 행에서 시작
                    ProcessArrayElementCopy(arrayNode.Children[i], mergedStructure, startColumn + (i * singleElementColumns), $"{yamlPath}[*]");
                }
            }
        }
        
        // 배열 요소의 구조를 복사하는 헬퍼 메서드
        private void ProcessArrayElementCopy(ExcelSchemeNode parentNode, YamlMappingNode mapping, int startColumn, string yamlPath)
        {
            int col = startColumn;
            
            foreach (var kvp in mapping.Children)
            {
                var key = kvp.Key.ToString();
                var value = kvp.Value;
                var propPath = $"{yamlPath}.{key}";
                
                if (value is YamlScalarNode)
                {
                    // 단순 속성
                    var propNode = new ExcelSchemeNode
                    {
                        Key = key,
                        SchemeMarker = "",
                        NodeType = SchemeNode.SchemeNodeType.PROPERTY,
                        Parent = parentNode,
                        RowIndex = currentRow,
                        ColumnIndex = col++,
                        OriginalYamlPath = propPath
                    };
                    parentNode.Children.Add(propNode);
                }
                else if (value is YamlSequenceNode sequence)
                {
                    // 배열 속성 처리
                    var arrayColumns = CalculateObjectColumns(MergeArrayElementStructures(sequence));
                    
                    var arrayNode = new ExcelSchemeNode
                    {
                        Key = key,
                        SchemeMarker = "$[]",
                        NodeType = SchemeNode.SchemeNodeType.ARRAY,
                        Parent = parentNode,
                        RowIndex = currentRow,
                        ColumnIndex = col,
                        ColumnSpan = arrayColumns,
                        IsMergedCell = arrayColumns > 1,
                        OriginalYamlPath = propPath
                    };
                    parentNode.Children.Add(arrayNode);
                    
                    if (sequence.Children.Count > 0 && sequence.Children[0] is YamlMappingNode)
                    {
                        // 중첩 배열은 ProcessArrayElements와 동일한 방식으로 처리
                        currentRow++;
                        ProcessArrayElements(arrayNode, sequence, col, propPath);
                        currentRow--;
                    }
                    
                    col += arrayColumns;
                }
                else if (value is YamlMappingNode childMapping)
                {
                    // 중첩 객체 속성
                    var objectColumns = CalculateObjectColumns(childMapping);
                    
                    var objectNode = new ExcelSchemeNode
                    {
                        Key = key,
                        SchemeMarker = "${}",
                        NodeType = SchemeNode.SchemeNodeType.MAP,
                        Parent = parentNode,
                        RowIndex = currentRow,
                        ColumnIndex = col,
                        ColumnSpan = objectColumns,
                        IsMergedCell = objectColumns > 1,
                        OriginalYamlPath = propPath
                    };
                    parentNode.Children.Add(objectNode);
                    
                    currentRow++;
                    ProcessArrayElementCopy(objectNode, childMapping, col, propPath);
                    currentRow--;
                    
                    col += objectColumns;
                }
            }
        }
        
        // 배열의 모든 요소를 분석하여 통합된 구조 생성
        private YamlMappingNode MergeArrayElementStructures(YamlSequenceNode sequence)
        {
            var mergedMapping = new YamlMappingNode();
            var allKeys = new HashSet<string>();
            
            // 모든 요소에서 키 수집
            foreach (var element in sequence.Children)
            {
                if (element is YamlMappingNode mapping)
                {
                    foreach (var kvp in mapping.Children)
                    {
                        var key = kvp.Key.ToString();
                        if (!allKeys.Contains(key))
                        {
                            allKeys.Add(key);
                            mergedMapping.Add(kvp.Key, kvp.Value);
                        }
                    }
                }
            }
            
            return mergedMapping;
        }

        private int CalculateObjectColumns(YamlMappingNode mapping)
        {
            int columns = 0;
            
            foreach (var kvp in mapping.Children)
            {
                if (kvp.Value is YamlScalarNode)
                {
                    columns += 1;
                }
                else if (kvp.Value is YamlSequenceNode sequence)
                {
                    columns += CalculateArrayColumns(sequence);
                }
                else if (kvp.Value is YamlMappingNode childMapping)
                {
                    columns += CalculateObjectColumns(childMapping);
                }
            }
            
            return Math.Max(1, columns);
        }

        private int CalculateArrayColumns(YamlSequenceNode sequence)
        {
            if (sequence.Children.Count == 0)
                return 1;
                
            if (sequence.Children[0] is YamlMappingNode firstMapping)
            {
                // 배열 요소가 객체인 경우: 각 요소의 컬럼 수 * 표시할 요소 수
                var mergedStructure = MergeArrayElementStructures(sequence);
                var singleElementColumns = CalculateObjectColumns(mergedStructure);
                int displayCount = sequence.Children.Count;
                return singleElementColumns * displayCount;
            }
            
            // 단순 배열
            return 1;
        }

        private void BuildRowMap(ExcelSchemeNode node, Dictionary<int, List<ExcelSchemeNode>> rowMap)
        {
            if (!rowMap.ContainsKey(node.RowIndex))
            {
                rowMap[node.RowIndex] = new List<ExcelSchemeNode>();
            }
            rowMap[node.RowIndex].Add(node);
            
            foreach (var child in node.Children)
            {
                BuildRowMap(child, rowMap);
            }
        }

        private void CalculateMergedCells(SchemeBuildResult result)
        {
            foreach (var kvp in result.RowMap)
            {
                foreach (var node in kvp.Value)
                {
                    if (node.IsMergedCell && node.ColumnSpan > 1)
                    {
                        result.MergedCells.Add((node.RowIndex, node.ColumnIndex, node.ColumnSpan));
                    }
                }
            }
            
            // $scheme_end 행 병합
            if (result.TotalColumns > 0)
            {
                result.MergedCells.Add((result.TotalRows, 1, result.TotalColumns));
            }
        }

        private void BuildColumnMappings(ExcelSchemeNode node, Dictionary<string, int> mappings)
        {
            if (node.NodeType == SchemeNode.SchemeNodeType.PROPERTY && !string.IsNullOrEmpty(node.OriginalYamlPath))
            {
                mappings[node.OriginalYamlPath] = node.ColumnIndex;
                
                // 디버깅: 매핑 추가 로깅
                Logger.Information($"Column mapping: {node.OriginalYamlPath} -> Column {node.ColumnIndex}");
            }
            
            foreach (var child in node.Children)
            {
                BuildColumnMappings(child, mappings);
            }
        }

        /// <summary>
        /// 디버깅용 스키마 트리 출력
        /// </summary>
        public void PrintSchemaTree(ExcelSchemeNode rootNode)
        {
            Logger.Information("========== 스키마 트리 구조 ==========");
            PrintNode(rootNode, 0);
            Logger.Information("=====================================");
        }

        private void PrintNode(ExcelSchemeNode node, int depth)
        {
            if (node == null) return;

            var indent = new string(' ', depth * 2);
            var mergeInfo = node.IsMergedCell ? $" [병합:{node.ColumnSpan}]" : "";
            var pathInfo = !string.IsNullOrEmpty(node.OriginalYamlPath) ? $" (경로:{node.OriginalYamlPath})" : "";
            
            Logger.Information($"{indent}[{node.RowIndex},{node.ColumnIndex}] '{node.Key}'{node.SchemeMarker} ({node.NodeType}){mergeInfo}{pathInfo}");
            
            foreach (var child in node.Children)
            {
                PrintNode(child, depth + 1);
            }
        }
    }
}