using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Logging;
using YamlDotNet.RepresentationModel;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    /// <summary>
    /// YAML 파일을 Excel로 변환하는 메인 컨버터
    /// ReverseSchemeBuilder를 사용하여 스키마를 생성하고 데이터를 매핑
    /// </summary>
    public class YamlToExcelConverter
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<YamlToExcelConverter>();

        private readonly ReverseSchemeBuilder _schemeBuilder;
        private readonly ExcelDataMapper _dataMapper;

        public YamlToExcelConverter()
        {
            _schemeBuilder = new ReverseSchemeBuilder();
            _dataMapper = new ExcelDataMapper();
        }

        /// <summary>
        /// YAML 파일을 Excel 파일로 변환
        /// </summary>
        public void Convert(string yamlPath, string excelPath)
        {
            try
            {
                Logger.Information($"YAML to Excel 변환 시작: {yamlPath} -> {excelPath}");

                // 1. YAML 로드
                var yamlContent = File.ReadAllText(yamlPath);
                var yaml = new YamlStream();
                yaml.Load(new StringReader(yamlContent));

                if (yaml.Documents.Count == 0)
                {
                    throw new InvalidOperationException("YAML 파일에 문서가 없습니다.");
                }

                var originalRootNode = yaml.Documents[0].RootNode;

                // 2. 스키마 생성용 병합된 YAML 생성
                Logger.Information("스키마 생성을 위한 병합된 YAML 생성 중...");
                var mergedRootNode = CreateMergedYamlForSchema(originalRootNode);
                
                Logger.Information("Excel 스키마 생성 중...");
                var schemaResult = _schemeBuilder.BuildSchemaTree(mergedRootNode);
                
                // 디버깅용 스키마 트리 출력
                _schemeBuilder.PrintSchemaTree(schemaResult.RootNode);

                // 3. Excel 워크북 생성
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");

                    // 4. 스키마 작성
                    WriteSchema(worksheet, schemaResult);

                    // 5. 데이터 매핑 및 작성 (원본 YAML 사용)
                    var dataStartRow = schemaResult.TotalRows + 1;
                    WriteData(worksheet, originalRootNode, schemaResult, dataStartRow);

                    // 6. 스타일 적용
                    ApplyStyles(worksheet, schemaResult.TotalRows);

                    // 7. 저장
                    var directory = Path.GetDirectoryName(excelPath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    workbook.SaveAs(excelPath);
                }

                Logger.Information($"변환 완료: {excelPath}");
            }
            catch (Exception ex)
            {
                Logger.Error($"변환 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// YAML 내용을 직접 Excel로 변환
        /// </summary>
        public void ConvertFromContent(string yamlContent, string excelPath)
        {
            try
            {
                Logger.Information($"YAML 내용을 Excel로 변환: -> {excelPath}");

                // YAML 로드
                var yaml = new YamlStream();
                yaml.Load(new StringReader(yamlContent));

                if (yaml.Documents.Count == 0)
                {
                    throw new InvalidOperationException("YAML 내용에 문서가 없습니다.");
                }

                var originalRootNode = yaml.Documents[0].RootNode;

                // 스키마 생성용 병합된 YAML 생성
                Logger.Information("스키마 생성을 위한 병합된 YAML 생성 중...");
                var mergedRootNode = CreateMergedYamlForSchema(originalRootNode);
                var schemaResult = _schemeBuilder.BuildSchemaTree(mergedRootNode);

                // Excel 생성
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Sheet1");
                    WriteSchema(worksheet, schemaResult);
                    
                    var dataStartRow = schemaResult.TotalRows + 1;
                    // 데이터는 원본 YAML 사용
                    WriteData(worksheet, originalRootNode, schemaResult, dataStartRow);
                    
                    ApplyStyles(worksheet, schemaResult.TotalRows);
                    
                    workbook.SaveAs(excelPath);
                }

                Logger.Information($"변환 완료: {excelPath}");
            }
            catch (Exception ex)
            {
                Logger.Error($"변환 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// YAML을 Excel 워크북으로 변환
        /// </summary>
        public IXLWorkbook ConvertToWorkbook(string yamlContent, string sheetName = "Sheet1")
        {
            try
            {
                Logger.Information("YAML을 워크북으로 변환");

                var yaml = new YamlStream();
                yaml.Load(new StringReader(yamlContent));

                if (yaml.Documents.Count == 0)
                {
                    throw new InvalidOperationException("YAML 내용에 문서가 없습니다.");
                }

                var originalRootNode = yaml.Documents[0].RootNode;
                
                // 스키마 생성용 병합된 YAML 생성
                Logger.Information("스키마 생성을 위한 병합된 YAML 생성 중...");
                var mergedRootNode = CreateMergedYamlForSchema(originalRootNode);
                var schemaResult = _schemeBuilder.BuildSchemaTree(mergedRootNode);

                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add(sheetName);
                
                WriteSchema(worksheet, schemaResult);
                
                var dataStartRow = schemaResult.TotalRows + 1;
                // 데이터는 원본 YAML 사용
                WriteData(worksheet, originalRootNode, schemaResult, dataStartRow);
                
                ApplyStyles(worksheet, schemaResult.TotalRows);

                return workbook;
            }
            catch (Exception ex)
            {
                Logger.Error($"워크북 변환 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// 스키마를 워크시트에 작성
        /// </summary>
        private void WriteSchema(IXLWorksheet worksheet, ReverseSchemeBuilder.SchemeBuildResult schemaResult)
        {
            Logger.Information("스키마 작성 시작");

            // 행별로 노드 작성
            foreach (var rowKvp in schemaResult.RowMap.OrderBy(r => r.Key))
            {
                int rowNum = rowKvp.Key;
                var nodes = rowKvp.Value.OrderBy(n => n.ColumnIndex).ToList();

                foreach (var node in nodes)
                {
                    string cellValue = node.Key;
                    if (!string.IsNullOrEmpty(node.SchemeMarker))
                    {
                        cellValue = string.IsNullOrEmpty(node.Key) ? node.SchemeMarker : $"{node.Key}{node.SchemeMarker}";
                    }

                    worksheet.Cell(node.RowIndex, node.ColumnIndex).Value = cellValue;
                }
            }

            // 병합 셀 처리
            foreach (var merge in schemaResult.MergedCells)
            {
                // 병합 범위가 유효한지 확인
                if (merge.row > 0 && merge.col > 0 && merge.colspan > 1)
                {
                    int endCol = merge.col + merge.colspan - 1;
                    // ClosedXML은 최대 16384 컬럼까지 지원
                    if (endCol <= 16384)
                    {
                        try
                        {
                            var range = worksheet.Range(merge.row, merge.col, merge.row, endCol);
                            range.Merge();
                        }
                        catch (Exception ex)
                        {
                            Logger.Warning($"병합 셀 생성 실패: 행={merge.row}, 열={merge.col}-{endCol}, 오류={ex.Message}");
                        }
                    }
                    else
                    {
                        Logger.Warning($"병합 셀 범위 초과: 행={merge.row}, 열={merge.col}-{endCol}");
                    }
                }
            }

            // $scheme_end 추가
            var schemeEndRow = schemaResult.TotalRows;
            if (schemeEndRow > 0 && schemaResult.TotalColumns > 0 && schemaResult.TotalColumns <= 16384)
            {
                worksheet.Cell(schemeEndRow, 1).Value = "$scheme_end";
                if (schemaResult.TotalColumns > 1)
                {
                    try
                    {
                        var schemeEndRange = worksheet.Range(schemeEndRow, 1, schemeEndRow, schemaResult.TotalColumns);
                        schemeEndRange.Merge();
                    }
                    catch (Exception ex)
                    {
                        Logger.Warning($"$scheme_end 병합 실패: {ex.Message}");
                    }
                }
            }

            Logger.Information($"스키마 작성 완료: {schemaResult.TotalRows}행");
        }

        /// <summary>
        /// 데이터를 워크시트에 작성
        /// </summary>
        private void WriteData(IXLWorksheet worksheet, YamlNode rootNode, ReverseSchemeBuilder.SchemeBuildResult schemaResult, int startRow)
        {
            Logger.Information("데이터 작성 시작");

            if (rootNode is YamlSequenceNode rootSequence)
            {
                // 루트가 배열인 경우
                int currentRow = startRow;
                foreach (var element in rootSequence.Children)
                {
                    WriteNodeData(worksheet, element, schemaResult.ColumnMappings, currentRow, "");
                    currentRow++;
                }
            }
            else if (rootNode is YamlMappingNode rootMapping)
            {
                // 루트가 객체인 경우
                WriteNodeData(worksheet, rootNode, schemaResult.ColumnMappings, startRow, "");
            }

            Logger.Information("데이터 작성 완료");
        }

        /// <summary>
        /// 노드 데이터를 재귀적으로 작성
        /// </summary>
        private void WriteNodeData(IXLWorksheet worksheet, YamlNode node, Dictionary<string, int> columnMappings, int row, string path)
        {
            if (node is YamlMappingNode mapping)
            {
                foreach (var kvp in mapping.Children)
                {
                    var key = kvp.Key.ToString();
                    var value = kvp.Value;
                    var fullPath = string.IsNullOrEmpty(path) ? key : $"{path}.{key}";

                    if (value is YamlScalarNode scalar)
                    {
                        // 단순 값 작성
                        if (columnMappings.ContainsKey(fullPath))
                        {
                            var col = columnMappings[fullPath];
                            worksheet.Cell(row, col).Value = scalar.Value;
                        }
                    }
                    else if (value is YamlMappingNode childMapping)
                    {
                        // 중첩 객체
                        WriteNodeData(worksheet, childMapping, columnMappings, row, fullPath);
                    }
                    else if (value is YamlSequenceNode childSequence)
                    {
                        // 배열 처리
                        WriteArrayData(worksheet, childSequence, columnMappings, row, fullPath);
                    }
                }
            }
            else if (node is YamlSequenceNode sequence)
            {
                WriteArrayData(worksheet, sequence, columnMappings, row, path);
            }
            else if (node is YamlScalarNode scalar)
            {
                if (columnMappings.ContainsKey(path))
                {
                    var col = columnMappings[path];
                    worksheet.Cell(row, col).Value = scalar.Value;
                }
            }
        }

        /// <summary>
        /// 배열 데이터 작성
        /// </summary>
        private void WriteArrayData(IXLWorksheet worksheet, YamlSequenceNode sequence, Dictionary<string, int> columnMappings, int row, string path)
        {
            // 배열의 각 요소를 처리
            for (int i = 0; i < sequence.Children.Count; i++)
            {
                var element = sequence.Children[i];
                
                if (element is YamlMappingNode elementMapping)
                {
                    // 배열 요소가 객체인 경우
                    foreach (var kvp in elementMapping.Children)
                    {
                        var key = kvp.Key.ToString();
                        var value = kvp.Value;
                        
                        // 통합 스키마 경로 사용 ([*] 대신 실제 인덱스 사용)
                        var propPath = $"{path}[*].{key}";
                        
                        // 실제 데이터를 쓸 때는 정확한 인덱스 경로도 시도
                        var indexedPath = $"{path}[{i}].{key}";

                        if (value is YamlScalarNode scalar)
                        {
                            // 통합 경로로 먼저 시도
                            if (columnMappings.ContainsKey(propPath))
                            {
                                var col = columnMappings[propPath];
                                worksheet.Cell(row, col).Value = scalar.Value;
                            }
                            // 인덱스 경로로도 시도 (하위 호환성)
                            else if (columnMappings.ContainsKey(indexedPath))
                            {
                                var col = columnMappings[indexedPath];
                                worksheet.Cell(row, col).Value = scalar.Value;
                            }
                        }
                        else if (value is YamlMappingNode childMapping)
                        {
                            // 중첩 객체는 통합 경로 사용
                            WriteNodeData(worksheet, childMapping, columnMappings, row, $"{path}[*].{key}");
                        }
                        else if (value is YamlSequenceNode childSequence)
                        {
                            // 중첩 배열은 통합 경로 사용
                            WriteArrayData(worksheet, childSequence, columnMappings, row, $"{path}[*].{key}");
                        }
                    }
                }
                else if (element is YamlScalarNode scalar)
                {
                    // 배열 요소가 단순 값인 경우
                    var propPath = $"{path}[*]";
                    var indexedPath = $"{path}[{i}]";
                    
                    if (columnMappings.ContainsKey(propPath))
                    {
                        var col = columnMappings[propPath];
                        worksheet.Cell(row, col).Value = scalar.Value;
                    }
                    else if (columnMappings.ContainsKey(indexedPath))
                    {
                        var col = columnMappings[indexedPath];
                        worksheet.Cell(row, col).Value = scalar.Value;
                    }
                }
            }
        }

        /// <summary>
        /// 스타일 적용
        /// </summary>
        private void ApplyStyles(IXLWorksheet worksheet, int schemaEndRow)
        {
            // 스키마 영역 스타일
            var schemaRange = worksheet.Range(1, 1, schemaEndRow, worksheet.LastColumnUsed().ColumnNumber());
            schemaRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            schemaRange.Style.Font.Bold = true;
            schemaRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            schemaRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // 데이터 영역 테두리
            if (worksheet.LastRowUsed() != null && worksheet.LastRowUsed().RowNumber() > schemaEndRow)
            {
                var dataRange = worksheet.Range(
                    schemaEndRow + 1, 1,
                    worksheet.LastRowUsed().RowNumber(),
                    worksheet.LastColumnUsed().ColumnNumber());
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }

            // 자동 너비 조정
            worksheet.Columns().AdjustToContents();
        }

        /// <summary>
        /// 스키마 생성을 위한 병합된 YAML 생성 (merge_yaml_complete.py 로직 활용)
        /// </summary>
        private YamlNode CreateMergedYamlForSchema(YamlNode originalNode)
        {
            Logger.Information("merge_yaml_complete.py 방식으로 스키마용 병합된 YAML 생성");
            
            if (originalNode is YamlSequenceNode rootSequence)
            {
                // 루트가 배열인 경우: 모든 요소를 병합하여 완전한 스키마 생성
                Logger.Information($"루트 배열 병합: {rootSequence.Children.Count}개 요소");
                
                if (rootSequence.Children.Count == 0)
                {
                    return originalNode;
                }
                
                if (rootSequence.Children.Count == 1)
                {
                    Logger.Information("단일 요소, 원본 반환");
                    return originalNode;
                }
                
                // merge_items_force_with_array_index 로직 적용
                var mergedArray = new YamlSequenceNode();
                var mergedItem = MergeAllSequenceElements(rootSequence);
                mergedArray.Add(mergedItem);
                
                Logger.Information($"배열 병합 완료: {rootSequence.Children.Count}개 → 1개 (완전한 스키마 포함)");
                return mergedArray;
            }
            else if (originalNode is YamlMappingNode rootMapping)
            {
                // 루트가 객체인 경우: 그대로 사용
                Logger.Information("루트 객체, 원본 사용");
                return originalNode;
            }
            
            Logger.Information("기타 노드 타입, 원본 사용");
            return originalNode;
        }
        
        /// <summary>
        /// 배열의 모든 요소를 병합하여 완전한 스키마를 가진 단일 요소 생성
        /// </summary>
        private YamlNode MergeAllSequenceElements(YamlSequenceNode sequence)
        {
            if (sequence.Children.Count == 0)
                return new YamlMappingNode();
            
            if (sequence.Children.Count == 1)
                return DeepCloneNode(sequence.Children[0]);
            
            Logger.Information($"  🔄 {sequence.Children.Count}개 배열 요소 병합 시작 (스키마용)");
            
            // 첫 번째 요소를 기준으로 시작
            var merged = DeepCloneNode(sequence.Children[0]);
            int mergeCount = 0;
            
            for (int i = 1; i < sequence.Children.Count; i++)
            {
                merged = DeepMergeNodesForSchema(merged, sequence.Children[i]);
                mergeCount++;
            }
            
            Logger.Information($"  → {sequence.Children.Count}개 요소를 1개로 병합 완료 (스키마용, 병합된 항목: {mergeCount}개)");
            return merged;
        }
        
        /// <summary>
        /// 스키마 생성용 노드 병합 (merge_yaml_complete.py의 deep_merge_objects 로직)
        /// </summary>
        private YamlNode DeepMergeNodesForSchema(YamlNode node1, YamlNode node2)
        {
            if (node1 == null) return node2 != null ? DeepCloneNode(node2) : null;
            if (node2 == null) return DeepCloneNode(node1);
            
            // 둘 다 매핑인 경우
            if (node1 is YamlMappingNode mapping1 && node2 is YamlMappingNode mapping2)
            {
                var result = new YamlMappingNode();
                
                // node1의 모든 키 복사
                foreach (var kvp in mapping1.Children)
                {
                    result.Add(kvp.Key, DeepCloneNode(kvp.Value));
                }
                
                // node2의 키들 병합
                foreach (var kvp in mapping2.Children)
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
                        result.Children[key] = DeepMergeNodesForSchema(existing, value);
                    }
                }
                
                return result;
            }
            
            // 둘 다 시퀀스인 경우 - 인덱스별 병합
            if (node1 is YamlSequenceNode seq1 && node2 is YamlSequenceNode seq2)
            {
                return MergeSequencesByIndexForSchema(new List<YamlSequenceNode> { seq1, seq2 });
            }
            
            // 기타 경우: 첫 번째 값 유지 (스키마에서는 구조가 중요)
            return DeepCloneNode(node1);
        }
        
        /// <summary>
        /// 스키마 생성용 시퀀스 인덱스별 병합
        /// </summary>
        private YamlSequenceNode MergeSequencesByIndexForSchema(List<YamlSequenceNode> sequences)
        {
            if (sequences == null || sequences.Count == 0)
                return new YamlSequenceNode();
            
            var validSequences = sequences.Where(seq => seq != null && seq.Children.Count > 0).ToList();
            if (validSequences.Count == 0)
                return new YamlSequenceNode();
            
            int maxLength = validSequences.Max(seq => seq.Children.Count);
            var result = new YamlSequenceNode();
            
            Logger.Information($"    [스키마용] 인덱스별 시퀀스 병합: 최대 길이 {maxLength}");
            
            for (int i = 0; i < maxLength; i++)
            {
                var itemsAtIndex = new List<YamlNode>();
                foreach (var seq in validSequences)
                {
                    if (i < seq.Children.Count)
                    {
                        itemsAtIndex.Add(seq.Children[i]);
                    }
                }
                
                if (itemsAtIndex.Count > 0)
                {
                    var mergedItem = itemsAtIndex[0];
                    for (int j = 1; j < itemsAtIndex.Count; j++)
                    {
                        mergedItem = DeepMergeNodesForSchema(mergedItem, itemsAtIndex[j]);
                    }
                    result.Add(mergedItem);
                }
            }
            
            return result;
        }
        
        /// <summary>
        /// YAML 노드 깊은 복사
        /// </summary>
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
    }
}