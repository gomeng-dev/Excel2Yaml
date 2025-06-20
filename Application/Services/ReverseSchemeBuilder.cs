using System;
using System.Collections.Generic;
using System.Linq;
using YamlDotNet.RepresentationModel;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Domain.Constants;

namespace ExcelToYamlAddin.Application.Services
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
            public SchemeNodeType NodeType { get; set; }
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

        private int currentRow = SchemeConstants.Sheet.SchemaStartRow;
        private int maxColumn = 0;

        public SchemeBuildResult BuildSchemaTree(YamlNode yamlRoot)
        {
            Logger.Information("========== 스키마 빌드 시작 (v2) ==========");
            
            currentRow = SchemeConstants.Sheet.SchemaStartRow;
            maxColumn = 0;
            
            var result = new SchemeBuildResult();
            
            // 루트 노드 처리 (병합된 YAML 사용)
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
                    SchemeMarker = SchemeConstants.Markers.ArrayStart,
                    NodeType = SchemeNodeType.Array,
                    RowIndex = currentRow,
                    ColumnIndex = SchemeConstants.Position.RootNodeColumn,
                    OriginalYamlPath = ""
                };

                // 배열의 첫 번째 요소로 전체 구조 분석 (병합된 YAML 사용)
                if (rootSequence.Children.Count > 0 && rootSequence.Children[0] is YamlMappingNode firstMapping)
                {
                    // 병합된 YAML로 컬럼 계산 (모든 필드가 포함된 완전한 구조)
                    var totalColumns = CalculateObjectColumns(firstMapping);
                    
                    // HTML 참조에서 확인: 2행 $[]는 A-K (11컬럼), L에 ^
                    // 따라서 totalColumns + 1이 전체 컬럼 수
                    var fullColumnCount = totalColumns + 1;
                    
                    // 루트 배열 노드는 A-K 컬럼을 차지 (totalColumns 만큼)
                    rootArrayNode.ColumnSpan = fullColumnCount;
                    rootArrayNode.IsMergedCell = totalColumns > 1;
                    maxColumn = fullColumnCount;
                    
                    Logger.Information($"루트 배열 노드 설정: ColumnSpan={totalColumns}, maxColumn={maxColumn}");
                    
                    currentRow++;
                    
                    // ^ 마커와 ${} 추가
                    var caretNode = new ExcelSchemeNode
                    {
                        Key = SchemeConstants.Markers.Ignore,
                        SchemeMarker = "",
                        NodeType = SchemeNodeType.Ignore,
                        Parent = rootArrayNode,
                        RowIndex = currentRow,
                        ColumnIndex = fullColumnCount  // 마지막 컬럼에 ^
                    };
                    rootArrayNode.Children.Add(caretNode);
                    
                    // HTML 참조에서 확인: 3행 ${}는 B-K (10컬럼)
                    var elementNode = new ExcelSchemeNode
                    {
                        Key = "",
                        SchemeMarker = SchemeConstants.Markers.MapStart,
                        NodeType = SchemeNodeType.Map,
                        Parent = rootArrayNode,
                        RowIndex = currentRow,
                        ColumnIndex = 2,
                        ColumnSpan = totalColumns,  // B부터 K까지 (totalColumns - 1)
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
                    SchemeMarker = SchemeConstants.Markers.MapStart,
                    NodeType = SchemeNodeType.Map,
                    RowIndex = currentRow,
                    ColumnIndex = SchemeConstants.Position.RootNodeColumn,
                    OriginalYamlPath = ""
                };
                
                var columns = CalculateObjectColumns(rootMapping);
                rootObjectNode.ColumnSpan = columns;
                rootObjectNode.IsMergedCell = columns > 1;
                maxColumn = columns;
                
                currentRow++;
                ProcessObjectProperties(rootObjectNode, rootMapping, SchemeConstants.Position.RootNodeColumn, "");
                
                return rootObjectNode;
            }
            
            throw new InvalidOperationException("지원하지 않는 루트 노드 타입");
        }

        private void ProcessObjectProperties(ExcelSchemeNode parentNode, YamlMappingNode mapping, int startColumn, string yamlPath)
        {
            Logger.Information($"ProcessObjectProperties 시작: parentKey={parentNode.Key}, currentRow={currentRow}, startColumn={startColumn}, yamlPath={yamlPath}, 자식수={mapping.Children.Count}");
            
            int col = startColumn;
            
            // 루트 배열의 요소인 경우 ^ 마커 추가
            if (parentNode.Parent != null && 
                parentNode.Parent.NodeType == SchemeNodeType.Array &&
                parentNode.Parent.Parent == null &&
                startColumn > SchemeConstants.Position.RootNodeColumn)
            {
                var caretNode = new ExcelSchemeNode
                {
                    Key = SchemeConstants.Markers.Ignore,
                    SchemeMarker = "",
                    NodeType = SchemeNodeType.Ignore,
                    Parent = parentNode,
                    RowIndex = currentRow,
                    ColumnIndex = SchemeConstants.Position.RootNodeColumn
                };
                parentNode.Children.Add(caretNode);
            }
            
            var baseRow = currentRow; // 현재 행 저장
            var childNodesToProcess = new List<(ExcelSchemeNode node, YamlNode value, string path)>();
            
            Logger.Information($"1단계: 형제 마커들을 행 {baseRow}에 배치");
            
            // 1단계: 모든 마커를 같은 행에 배치
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
                        NodeType = SchemeNodeType.Property,
                        Parent = parentNode,
                        RowIndex = baseRow,
                        ColumnIndex = col++,
                        OriginalYamlPath = propPath
                    };
                    parentNode.Children.Add(propNode);
                    Logger.Information($"  속성 {key} -> 행{baseRow}, 열{propNode.ColumnIndex}");
                }
                else if (value is YamlSequenceNode sequence)
                {
                    // 배열 속성 - 마커만 현재 행에 배치
                    var arrayColumns = CalculateArrayColumns(sequence);
                    
                    var arrayNode = new ExcelSchemeNode
                    {
                        Key = key,
                        SchemeMarker = SchemeConstants.Markers.ArrayStart,
                        NodeType = SchemeNodeType.Array,
                        Parent = parentNode,
                        RowIndex = baseRow,
                        ColumnIndex = col,
                        ColumnSpan = arrayColumns,
                        IsMergedCell = arrayColumns > 1,
                        OriginalYamlPath = propPath
                    };
                    parentNode.Children.Add(arrayNode);
                    Logger.Information($"  배열 {key}$[] -> 행{baseRow}, 열{col}-{col + arrayColumns - 1}");
                    
                    // 하위 구조 처리를 위해 저장
                    if (sequence.Children.Count > 0 && sequence.Children[0] is YamlMappingNode)
                    {
                        childNodesToProcess.Add((arrayNode, sequence, propPath));
                    }
                    
                    col += arrayColumns;
                }
                else if (value is YamlMappingNode childMapping)
                {
                    // 중첩 객체 - 마커만 현재 행에 배치
                    var objectColumns = CalculateObjectColumns(childMapping);
                    
                    var objectNode = new ExcelSchemeNode
                    {
                        Key = key,
                        SchemeMarker = SchemeConstants.Markers.MapStart,
                        NodeType = SchemeNodeType.Map,
                        Parent = parentNode,
                        RowIndex = baseRow,
                        ColumnIndex = col,
                        ColumnSpan = objectColumns,
                        IsMergedCell = objectColumns > 1,
                        OriginalYamlPath = propPath
                    };
                    parentNode.Children.Add(objectNode);
                    Logger.Information($"  객체 {key}${{}} -> 행{baseRow}, 열{col}-{col + objectColumns - 1}");
                    
                    // 하위 구조 처리를 위해 저장
                    childNodesToProcess.Add((objectNode, childMapping, propPath));
                    
                    col += objectColumns;
                }
            }
            
            // 2단계: 하위 구조들을 처리 (모든 형제 노드가 같은 행에서 시작하도록 개선)
            if (childNodesToProcess.Count > 0)
            {
                var nextRow = baseRow + 1; // 모든 형제 노드의 자식들이 시작할 행
                var maxRowUsed = nextRow; // 실제로 사용된 최대 행 추적
                
                Logger.Information($"2단계: 하위 구조 처리 시작: baseRow={baseRow}, nextRow={nextRow}, 자식 노드 수={childNodesToProcess.Count}");
                
                foreach (var (childNode, childValue, childPath) in childNodesToProcess)
                {
                    currentRow = nextRow; // 모든 형제 노드가 같은 행에서 시작
                    Logger.Information($"  형제 노드 [{childNode.Key}{childNode.SchemeMarker}] 처리 시작: currentRow={currentRow} (nextRow로 리셋)");
                    
                    if (childValue is YamlSequenceNode sequence)
                    {
                        ProcessArrayElements(childNode, sequence, childNode.ColumnIndex, childPath);
                    }
                    else if (childValue is YamlMappingNode mapping2)
                    {
                        ProcessObjectProperties(childNode, mapping2, childNode.ColumnIndex, childPath);
                    }
                    
                    Logger.Information($"  형제 노드 [{childNode.Key}{childNode.SchemeMarker}] 처리 완료: currentRow={currentRow}");
                    
                    // 현재 노드 처리 후 사용된 최대 행 업데이트
                    maxRowUsed = Math.Max(maxRowUsed, currentRow);
                    Logger.Information($"  maxRowUsed 업데이트: {maxRowUsed}");
                }
                
                // 모든 형제 노드 처리 완료 후 최대 행으로 설정
                Logger.Information($"모든 형제 노드 처리 완료: maxRowUsed={maxRowUsed}");
                currentRow = maxRowUsed;
            }
            else
            {
                // 하위 구조가 없으면 다음 행으로 이동
                currentRow++;
                Logger.Information($"하위 구조 없음, currentRow++: {currentRow}");
            }
            
            Logger.Information($"ProcessObjectProperties 완료: parentKey={parentNode.Key}, currentRow={currentRow}");
            
            maxColumn = Math.Max(maxColumn, col - 1);
        }

        private void ProcessArrayElements(ExcelSchemeNode arrayNode, YamlSequenceNode sequence, int startColumn, string yamlPath)
        {
            Logger.Information($"ProcessArrayElements 시작: arrayKey={arrayNode.Key}, currentRow={currentRow}, startColumn={startColumn}, 요소수={sequence.Children.Count}");
            
            if (sequence.Children.Count == 0)
                return;
                
            // 배열 요소가 객체인 경우만 처리
            if (sequence.Children[0] is YamlMappingNode)
            {
                Logger.Information($"배열 요소 ${{}}마커들을 행 {currentRow}에 배치 (개별 계산)");
                
                // 각 배열 요소별로 개별 컬럼 수 계산
                var elementColumns = new List<int>();
                var elementStructures = new List<YamlMappingNode>();
                
                for (int i = 0; i < sequence.Children.Count; i++)
                {
                    if (sequence.Children[i] is YamlMappingNode elementMapping)
                    {
                        // 각 요소를 개별적으로 계산
                        var singleColumns = CalculateObjectColumns(elementMapping);
                        elementColumns.Add(singleColumns);
                        elementStructures.Add(elementMapping);
                        
                        Logger.Information($"  요소[{i}] 개별 컬럼 계산: {singleColumns}개");
                    }
                }
                
                // 각 배열 요소를 위한 ${} 마커를 한 행에 나란히 생성
                int currentCol = startColumn;
                for (int i = 0; i < sequence.Children.Count; i++)
                {
                    var singleElementColumns = elementColumns[i];
                    
                    var elementNode = new ExcelSchemeNode
                    {
                        Key = "",
                        SchemeMarker = SchemeConstants.Markers.MapStart,
                        NodeType = SchemeNodeType.Map,
                        Parent = arrayNode,
                        RowIndex = currentRow,
                        ColumnIndex = currentCol,
                        ColumnSpan = singleElementColumns,
                        IsMergedCell = singleElementColumns > 1,
                        OriginalYamlPath = $"{yamlPath}[{i}]"  // 정확한 인덱스 사용
                    };
                    arrayNode.Children.Add(elementNode);
                    Logger.Information($"  요소[{i}] ${{}} -> 행{currentRow}, 열{currentCol}-{currentCol + singleElementColumns - 1} (경로: {elementNode.OriginalYamlPath})");
                    
                    currentCol += singleElementColumns;
                }
                
                // 다음 행에서 자식 구조 처리
                currentRow++;
                var childrenStartRow = currentRow;
                
                Logger.Information($"배열 요소들의 자식 구조를 행 {childrenStartRow}에서 처리 (개별 구조 사용)");
                
                // 각 요소를 개별 구조로 처리
                currentCol = startColumn;
                var maxRowUsed = currentRow;
                
                for (int i = 0; i < sequence.Children.Count; i++)
                {
                    currentRow = childrenStartRow; // 모든 요소가 같은 행에서 시작
                    var elementStructure = elementStructures[i];
                    var elementColumnCount = elementColumns[i];
                    
                    Logger.Information($"  요소[{i}] 개별 구조 처리 시작: currentRow={currentRow}, 컬럼={currentCol}");
                    ProcessObjectProperties(arrayNode.Children[i], elementStructure, currentCol, $"{yamlPath}[{i}]");
                    
                    maxRowUsed = Math.Max(maxRowUsed, currentRow);
                    currentCol += elementColumnCount;
                    Logger.Information($"  요소[{i}] 개별 구조 처리 완료: currentRow={currentRow}, maxRowUsed={maxRowUsed}");
                }
                
                // 배열 처리가 끝난 후 실제 사용된 최대 행으로 currentRow 설정
                Logger.Information($"ProcessArrayElements 완료: arrayKey={arrayNode.Key}, maxRowUsed={maxRowUsed}");
                currentRow = maxRowUsed;
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
                        NodeType = SchemeNodeType.Property,
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
                        SchemeMarker = SchemeConstants.Markers.ArrayStart,
                        NodeType = SchemeNodeType.Array,
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
                        SchemeMarker = SchemeConstants.Markers.MapStart,
                        NodeType = SchemeNodeType.Map,
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
        
        // 변경된 merge_yaml_complete.py의 merge_items_force_with_array_index 방식으로 구조 생성
        private YamlMappingNode MergeArrayElementStructures(YamlSequenceNode sequence)
        {
            Logger.Information($"MergeArrayElementStructures 시작 (인덱스별 배열 병합): 요소 수={sequence.Children.Count}");
            
            if (sequence.Children.Count == 0)
                return new YamlMappingNode();
            
            // merge_items_force_with_array_index 로직 구현
            var items = new List<YamlMappingNode>();
            foreach (var element in sequence.Children)
            {
                if (element is YamlMappingNode mapping)
                {
                    items.Add(mapping);
                }
            }
            
            if (items.Count == 0)
            {
                Logger.Information("  유효한 매핑 요소가 없음, 빈 구조 반환");
                return new YamlMappingNode();
            }
            
            if (items.Count == 1)
            {
                Logger.Information("  단일 항목, 복사하여 반환");
                return DeepCloneNode(items[0]) as YamlMappingNode;
            }
            
            Logger.Information($"  🔄 {items.Count}개 항목 병합 시작 (모든 배열은 인덱스별 병합)");
            
            // 첫 번째 항목을 기준으로 모든 항목 병합
            var merged = DeepCloneNode(items[0]) as YamlMappingNode;
            int mergeCount = 0;
            
            for (int i = 1; i < items.Count; i++)
            {
                merged = DeepMergeObjectsComplete(merged, items[i]);
                mergeCount++;
            }
            
            var finalKeys = merged.Children.Keys.Select(k => k.ToString()).ToList();
            Logger.Information($"  → {items.Count}개 항목을 1개로 병합 완료 (배열은 인덱스별 병합, 병합된 항목: {mergeCount}개)");
            Logger.Information($"  최종 병합 완료: {string.Join(", ", finalKeys.Take(5))}... (총 {finalKeys.Count}개 키)");
            Logger.Information($"MergeArrayElementStructures 완료 (인덱스별 배열 병합): 병합된 키 수={merged.Children.Count}");
            return merged;
        }
        
        // 변경된 merge_yaml_complete.py의 deep_merge_objects와 동일한 구현 (인덱스별 배열 병합)
        private YamlMappingNode DeepMergeObjectsComplete(YamlMappingNode obj1, YamlMappingNode obj2)
        {
            var result = new YamlMappingNode();
            
            // obj1의 모든 키 복사
            foreach (var kvp in obj1.Children)
            {
                result.Add(kvp.Key, DeepCloneNode(kvp.Value));
            }
            
            // obj2의 키들 병합
            foreach (var kvp in obj2.Children)
            {
                var key = kvp.Key;
                var value = kvp.Value;
                
                if (!result.Children.ContainsKey(key))
                {
                    // 새로운 키 추가
                    result.Add(key, DeepCloneNode(value));
                }
                else
                {
                    // 기존 키 병합
                    var existing = result.Children[key];
                    
                    if (existing is YamlMappingNode existingObj && value is YamlMappingNode valueObj)
                    {
                        // 둘 다 객체 - 재귀 병합
                        result.Children[key] = DeepMergeObjectsComplete(existingObj, valueObj);
                    }
                    else if (existing is YamlSequenceNode existingArray && value is YamlSequenceNode valueArray)
                    {
                        // 둘 다 배열 - 구조만 보고 병합 (값은 무시)
                        Logger.Information($"    🔀 배열 구조 기반 병합: [{existingArray.Children.Count}개] + [{valueArray.Children.Count}개]");
                        result.Children[key] = MergeArraysByStructure(new List<YamlSequenceNode> { existingArray, valueArray });
                    }
                    else if ((existing is YamlMappingNode || existing is YamlSequenceNode) && 
                             (value is YamlMappingNode || value is YamlSequenceNode))
                    {
                        // 혼합 타입: 객체와 배열이 섞인 경우 - 배열로 통일
                        Logger.Information($"    🔄 혼합 타입 감지: {existing.GetType().Name} + {value.GetType().Name} -> 배열로 통일");
                        
                        var mixedArray = new YamlSequenceNode();
                        
                        // 기존 값을 배열에 추가
                        if (existing is YamlSequenceNode existingSeq)
                        {
                            foreach (var item in existingSeq.Children)
                            {
                                mixedArray.Add(item);
                            }
                        }
                        else
                        {
                            mixedArray.Add(existing);
                        }
                        
                        // 새 값을 배열에 추가
                        if (value is YamlSequenceNode valueSeq)
                        {
                            foreach (var item in valueSeq.Children)
                            {
                                mixedArray.Add(item);
                            }
                        }
                        else
                        {
                            mixedArray.Add(value);
                        }
                        
                        result.Children[key] = mixedArray;
                        Logger.Information($"    → {key} 혼합 타입 병합 완료: 총 {mixedArray.Children.Count}개 요소");
                    }
                    // 스칼라 값은 첫 번째 값 유지 (기존 값 우선 - merge_yaml_complete.py의 "first" 전략)
                }
            }
            
            return result;
        }
        
        // 변경된 merge_yaml_complete.py의 merge_arrays_by_index와 동일한 구현
        private YamlSequenceNode MergeArraysByIndex(List<YamlSequenceNode> arrays)
        {
            if (arrays == null || arrays.Count == 0)
                return new YamlSequenceNode();
            
            // 빈 배열 제거
            var validArrays = arrays.Where(arr => arr != null && arr.Children.Count > 0).ToList();
            if (validArrays.Count == 0)
                return new YamlSequenceNode();
            
            // 가장 긴 배열의 길이를 찾습니다
            int maxLength = validArrays.Max(arr => arr.Children.Count);
            var mergedArray = new YamlSequenceNode();
            
            Logger.Information($"      📝 인덱스별 배열 병합 상세:");
            Logger.Information($"        - 입력 배열 개수: {validArrays.Count}");
            Logger.Information($"        - 각 배열 길이: [{string.Join(", ", validArrays.Select(arr => arr.Children.Count))}]");
            Logger.Information($"        - 최대 길이: {maxLength}");
            
            for (int i = 0; i < maxLength; i++)
            {
                // 인덱스 i에 있는 모든 항목들을 수집
                var itemsAtIndex = new List<YamlNode>();
                for (int j = 0; j < validArrays.Count; j++)
                {
                    var arr = validArrays[j];
                    if (i < arr.Children.Count)
                    {
                        itemsAtIndex.Add(arr.Children[i]);
                        var nodeType = arr.Children[i].GetType().Name;
                        var keys = arr.Children[i] is YamlMappingNode mapping ? 
                            string.Join(", ", mapping.Children.Keys.Take(3).Select(k => k.ToString())) : "N/A";
                        Logger.Information($"        - 배열 {j}[{i}]: {nodeType} (키: {keys})");
                    }
                }
                
                if (itemsAtIndex.Count > 0)
                {
                    Logger.Information($"        - 인덱스 {i}: {itemsAtIndex.Count}개 항목 병합");
                    
                    // 인덱스 i의 모든 항목들을 병합
                    var mergedItem = DeepCloneNode(itemsAtIndex[0]);
                    for (int k = 1; k < itemsAtIndex.Count; k++)
                    {
                        var item = itemsAtIndex[k];
                        Logger.Information($"          🔄 병합 중: {mergedItem.GetType().Name} + {item.GetType().Name}");
                        mergedItem = DeepMergeObjectsAny(mergedItem, item);
                    }
                    mergedArray.Add(mergedItem);
                    Logger.Information($"        - 인덱스 {i} 병합 완료: {mergedItem.GetType().Name}");
                }
            }
            
            Logger.Information($"      ✅ 최종 배열 길이: {mergedArray.Children.Count}");
            return mergedArray;
        }
        
        // 모든 타입의 YAML 노드를 병합하는 헬퍼 메서드 (Python의 deep_merge_objects와 동일)
        private YamlNode DeepMergeObjectsAny(YamlNode obj1, YamlNode obj2)
        {
            if (obj1 == null) return obj2;
            if (obj2 == null) return obj1;
            
            // 둘 다 딕셔너리인 경우
            if (obj1 is YamlMappingNode mapping1 && obj2 is YamlMappingNode mapping2)
            {
                return DeepMergeObjectsComplete(mapping1, mapping2);
            }
            
            // 둘 다 배열인 경우 - 인덱스별 병합
            if (obj1 is YamlSequenceNode seq1 && obj2 is YamlSequenceNode seq2)
            {
                return MergeArraysByIndex(new List<YamlSequenceNode> { seq1, seq2 });
            }
            
            // 값이 다른 경우 - 첫 번째 값 유지 (first 전략)
            return obj1;
        }
        
        // 두 YAML 노드가 같은지 비교하는 헬퍼 메서드
        private bool NodesEqual(YamlNode node1, YamlNode node2)
        {
            if (node1.GetType() != node2.GetType())
                return false;
                
            if (node1 is YamlScalarNode scalar1 && node2 is YamlScalarNode scalar2)
            {
                return scalar1.Value == scalar2.Value;
            }
            else if (node1 is YamlMappingNode mapping1 && node2 is YamlMappingNode mapping2)
            {
                if (mapping1.Children.Count != mapping2.Children.Count)
                    return false;
                    
                foreach (var kvp in mapping1.Children)
                {
                    if (!mapping2.Children.ContainsKey(kvp.Key) || 
                        !NodesEqual(kvp.Value, mapping2.Children[kvp.Key]))
                    {
                        return false;
                    }
                }
                return true;
            }
            else if (node1 is YamlSequenceNode seq1 && node2 is YamlSequenceNode seq2)
            {
                if (seq1.Children.Count != seq2.Children.Count)
                    return false;
                    
                for (int i = 0; i < seq1.Children.Count; i++)
                {
                    if (!NodesEqual(seq1.Children[i], seq2.Children[i]))
                        return false;
                }
                return true;
            }
            
            return false;
        }
        
        /// <summary>
        /// 구조만 보고 배열 병합 (값은 무시하고 키 구조만 통합)
        /// </summary>
        private YamlSequenceNode MergeArraysByStructure(List<YamlSequenceNode> arrays)
        {
            if (arrays == null || arrays.Count == 0)
                return new YamlSequenceNode();
            
            var validArrays = arrays.Where(arr => arr != null && arr.Children.Count > 0).ToList();
            if (validArrays.Count == 0)
                return new YamlSequenceNode();
            
            // 가장 긴 배열의 길이를 기준으로 함
            int maxLength = validArrays.Max(arr => arr.Children.Count);
            var result = new YamlSequenceNode();
            
            Logger.Information($"      📝 구조 기반 배열 병합:");
            Logger.Information($"        - 입력 배열 개수: {validArrays.Count}");
            Logger.Information($"        - 각 배열 길이: [{string.Join(", ", validArrays.Select(arr => arr.Children.Count))}]");
            Logger.Information($"        - 최대 길이: {maxLength}");
            
            // 모든 고유한 구조를 수집
            var allStructures = new List<YamlNode>();
            var seenStructures = new HashSet<string>();
            
            foreach (var array in validArrays)
            {
                foreach (var element in array.Children)
                {
                    var structureKey = GetNodeStructureKey(element);
                    if (!seenStructures.Contains(structureKey))
                    {
                        allStructures.Add(element);
                        seenStructures.Add(structureKey);
                        Logger.Information($"        - 고유 구조 발견: {structureKey}");
                    }
                }
            }
            
            // 최대 길이만큼 구조 반복
            for (int i = 0; i < maxLength; i++)
            {
                // 사용 가능한 구조 중에서 선택
                var templateNode = allStructures[i % allStructures.Count];
                var mergedNode = DeepCloneNode(templateNode);
                result.Add(mergedNode);
                
                var structureInfo = templateNode is YamlMappingNode mapping ? 
                    $"키: {string.Join(", ", mapping.Children.Keys.Take(3).Select(k => k.ToString()))}" : "단순값";
                
                Logger.Information($"        - 인덱스 {i}: 구조 '{GetNodeStructureKey(templateNode)}' 사용 ({structureInfo})");
            }
            
            Logger.Information($"      ✅ 최종 배열 길이: {result.Children.Count} (고유 구조: {allStructures.Count}개)");
            return result;
        }
        
        /// <summary>
        /// 노드의 구조적 특성을 나타내는 키 생성 (스키마 중복 방지용)
        /// </summary>
        private string GetNodeStructureKey(YamlNode node)
        {
            if (node is YamlMappingNode mapping)
            {
                // 객체의 키 구조로 고유성 판단
                var keys = mapping.Children.Keys.Select(k => k.ToString()).OrderBy(k => k).ToList();
                return $"Map[{string.Join(",", keys)}]";
            }
            else if (node is YamlSequenceNode sequence)
            {
                // 배열의 첫 번째 요소 구조로 고유성 판단
                if (sequence.Children.Count > 0)
                {
                    return $"Array[{GetNodeStructureKey(sequence.Children[0])}]";
                }
                return "Array[empty]";
            }
            else if (node is YamlScalarNode scalar)
            {
                // 스칼라는 값으로 고유성 판단 (스키마에서는 타입만 중요)
                return "Scalar";
            }
            
            return node.GetType().Name;
        }
        
        // 기존 DeepMergeObjects 메서드 (하위 호환성을 위해 유지)
        private YamlMappingNode DeepMergeObjects(YamlMappingNode obj1, YamlMappingNode obj2)
        {
            return DeepMergeObjectsComplete(obj1, obj2);
        }
        
        // 노드 깊은 복사
        private YamlNode DeepCloneNode(YamlNode node)
        {
            if (node is YamlMappingNode mapping)
            {
                var cloned = new YamlMappingNode();
                foreach (var kvp in mapping.Children)
                {
                    cloned.Add(kvp.Key, DeepCloneNode(kvp.Value));
                }
                return cloned;
            }
            else if (node is YamlSequenceNode sequence)
            {
                var cloned = new YamlSequenceNode();
                foreach (var child in sequence.Children)
                {
                    cloned.Add(DeepCloneNode(child));
                }
                return cloned;
            }
            else
            {
                // 스칼라 노드는 그대로 반환
                return node;
            }
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
                
            if (sequence.Children[0] is YamlMappingNode)
            {
                // 배열 요소가 객체인 경우: 각 요소를 개별적으로 계산해서 합계
                int totalColumns = 0;
                
                Logger.Information($"    배열 컬럼 계산 (개별): 요소수={sequence.Children.Count}");
                
                for (int i = 0; i < sequence.Children.Count; i++)
                {
                    if (sequence.Children[i] is YamlMappingNode elementMapping)
                    {
                        var elementColumns = CalculateObjectColumns(elementMapping);
                        totalColumns += elementColumns;
                        
                        var keys = elementMapping.Children.Keys.Take(3).Select(k => k.ToString());
                        Logger.Information($"      요소[{i}]: {elementColumns}개 컬럼 (키: {string.Join(", ", keys)})");
                    }
                }
                
                Logger.Information($"    배열 총 컬럼: {totalColumns}개");
                return totalColumns;
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
                result.MergedCells.Add((result.TotalRows, SchemeConstants.Position.RootNodeColumn, result.TotalColumns));
            }
        }

        private void BuildColumnMappings(ExcelSchemeNode node, Dictionary<string, int> mappings)
        {
            if (node.NodeType == SchemeNodeType.Property && !string.IsNullOrEmpty(node.OriginalYamlPath))
            {
                mappings[node.OriginalYamlPath] = node.ColumnIndex;
                
                // 디버깅: 매핑 추가 로깅
                Logger.Information($"Column mapping: {node.OriginalYamlPath} -> Column {node.ColumnIndex}");
            }
            else if (node.NodeType == SchemeNodeType.Array && !string.IsNullOrEmpty(node.OriginalYamlPath))
            {
                // 스칼라 배열인지 확인 (자식이 없거나 자식이 모두 스칼라인 경우)
                bool isScalarArray = node.Children.Count == 0 || 
                                   node.Children.All(child => child.NodeType == SchemeNodeType.Property);
                
                if (isScalarArray && node.ColumnSpan == 1)
                {
                    // 스칼라 배열의 [0] 인덱스 경로에 대한 컬럼 매핑 생성
                    string scalarArrayPath = $"{node.OriginalYamlPath}[0]";
                    mappings[scalarArrayPath] = node.ColumnIndex;
                    
                    Logger.Information($"Column mapping (스칼라 배열): {scalarArrayPath} -> Column {node.ColumnIndex}");
                }
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
