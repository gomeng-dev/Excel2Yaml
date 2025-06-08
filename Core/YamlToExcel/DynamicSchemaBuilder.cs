using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Logging;
using static ExcelToYamlAddin.Core.YamlToExcel.DynamicDataMapper;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    /// <summary>
    /// YAML 구조 분석 결과를 기반으로 Excel 스키마를 동적으로 생성하는 빌더
    /// </summary>
    public class DynamicSchemaBuilder
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<DynamicSchemaBuilder>();

        /// <summary>
        /// Excel 스키마 정보를 담는 클래스
        /// </summary>
        public class ExcelScheme
        {
            private readonly List<SchemaRow> rows = new List<SchemaRow>();
            private readonly Dictionary<string, int> columnMapping = new Dictionary<string, int>();
            private readonly Dictionary<string, int> arrayStartColumns = new Dictionary<string, int>();
            private readonly List<MergedCellInfo> mergedCells = new List<MergedCellInfo>();

            public int LastSchemaRow { get; private set; }

            public void AddCell(int row, int col, string value)
            {
                var schemaRow = GetOrCreateRow(row);
                schemaRow.Cells[col] = value;
                LastSchemaRow = Math.Max(LastSchemaRow, row);
            }

            public void AddMergedCell(int row, int startCol, int endCol, string value)
            {
                var schemaRow = GetOrCreateRow(row);
                schemaRow.MergedCells.Add(new MergedCell
                {
                    StartColumn = startCol,
                    EndColumn = endCol,
                    Value = value
                });
                mergedCells.Add(new MergedCellInfo
                {
                    Row = row,
                    StartColumn = startCol,
                    EndColumn = endCol
                });
                LastSchemaRow = Math.Max(LastSchemaRow, row);
            }

            public void AddSchemeEndRow(int row)
            {
                var schemaRow = GetOrCreateRow(row);
                // $scheme_end는 모든 열을 병합
                schemaRow.IsSchemeEnd = true;
                LastSchemaRow = row;
            }

            public int GetColumnIndex(string propertyName)
            {
                var result = columnMapping.ContainsKey(propertyName) ? columnMapping[propertyName] : -1;
                SimpleLoggerFactory.CreateLogger<ExcelScheme>()
                    .Debug($"GetColumnIndex('{propertyName}') = {result}");
                return result;
            }

            public int GetArrayStartColumn(string arrayName)
            {
                return arrayStartColumns.ContainsKey(arrayName) ? arrayStartColumns[arrayName] : -1;
            }

            public void SetColumnMapping(string propertyName, int column)
            {
                columnMapping[propertyName] = column;
                SimpleLoggerFactory.CreateLogger<ExcelScheme>()
                    .Information($"★ SetColumnMapping: '{propertyName}' -> 컬럼 {column}");
            }
            
            public void DebugAllMappings()
            {
                var logger = SimpleLoggerFactory.CreateLogger<ExcelScheme>();
                logger.Information("========== 모든 컬럼 매핑 상황 ==========");
                foreach (var mapping in columnMapping.OrderBy(m => m.Value))
                {
                    logger.Information($"  '{mapping.Key}' -> 컬럼 {mapping.Value}");
                }
                logger.Information($"총 {columnMapping.Count}개 매핑");
                logger.Information("======================================");
            }

            public void SetArrayStartColumn(string arrayName, int column)
            {
                arrayStartColumns[arrayName] = column;
            }

            private SchemaRow GetOrCreateRow(int rowNumber)
            {
                var row = rows.FirstOrDefault(r => r.RowNumber == rowNumber);
                if (row == null)
                {
                    row = new SchemaRow { RowNumber = rowNumber };
                    rows.Add(row);
                }
                return row;
            }

            public void WriteToWorksheet(IXLWorksheet worksheet)
            {
                foreach (var row in rows.OrderBy(r => r.RowNumber))
                {
                    if (row.IsSchemeEnd)
                    {
                        // $scheme_end는 모든 열을 병합
                        var lastCol = row.Cells.Keys.DefaultIfEmpty(1).Max();
                        foreach (var kvp in columnMapping)
                        {
                            lastCol = Math.Max(lastCol, kvp.Value);
                        }

                        // 병합 정보에서 실제 범위 찾기
                        var mergeInfo = mergedCells.FirstOrDefault(m => m.Row == row.RowNumber);
                        if (mergeInfo != null)
                        {
                            lastCol = mergeInfo.EndColumn;
                        }

                        worksheet.Range(row.RowNumber, 1, row.RowNumber, lastCol).Merge();
                        worksheet.Cell(row.RowNumber, 1).Value = "$scheme_end";
                    }
                    else
                    {
                        // 일반 셀 쓰기
                        foreach (var cell in row.Cells)
                        {
                            worksheet.Cell(row.RowNumber, cell.Key).Value = cell.Value;
                        }

                        // 병합 셀 처리
                        foreach (var merged in row.MergedCells)
                        {
                            var range = worksheet.Range(row.RowNumber, merged.StartColumn,
                                                      row.RowNumber, merged.EndColumn);
                            range.Merge();
                            range.FirstCell().Value = merged.Value;
                        }
                    }
                }
            }

            private class SchemaRow
            {
                public int RowNumber { get; set; }
                public Dictionary<int, string> Cells { get; } = new Dictionary<int, string>();
                public List<MergedCell> MergedCells { get; } = new List<MergedCell>();
                public bool IsSchemeEnd { get; set; }
            }

            private class MergedCell
            {
                public int StartColumn { get; set; }
                public int EndColumn { get; set; }
                public string Value { get; set; }
            }

            public class MergedCellInfo
            {
                public int Row { get; set; }
                public int StartColumn { get; set; }
                public int EndColumn { get; set; }
            }

            public List<MergedCellInfo> GetMergedCellsInRow(int row)
            {
                return mergedCells.Where(m => m.Row == row).ToList();
            }

            public void UpdateMergedCell(int row, int startColumn, int newEndColumn)
            {
                var merged = mergedCells.FirstOrDefault(m => m.Row == row && m.StartColumn == startColumn);
                if (merged != null)
                {
                    merged.EndColumn = newEndColumn;
                }
            }

            public void UpdateSchemeEndMarker(int actualUsedColumns)
            {
                // $scheme_end 행 찾기
                var schemeEndRow = rows.FirstOrDefault(r => r.IsSchemeEnd);
                if (schemeEndRow != null)
                {
                    // 기존 병합 정보 업데이트
                    var existingMerge = mergedCells.FirstOrDefault(m => m.Row == schemeEndRow.RowNumber);
                    if (existingMerge != null)
                    {
                        existingMerge.EndColumn = actualUsedColumns;
                    }
                    else
                    {
                        mergedCells.Add(new MergedCellInfo
                        {
                            Row = schemeEndRow.RowNumber,
                            StartColumn = 1,
                            EndColumn = actualUsedColumns
                        });
                    }
                }
            }

            public void OptimizeForDuplicates(Dictionary<string, int> duplicateCounts)
            {
                // 중복 요소 분석에 따른 스키마 최적화
                // 예: weaponSpec배열의 각 요소가 damage/addDamage를 가지는 경우
                foreach (var dup in duplicateCounts)
                {
                    Logger.Information($"중복 요소 감지: {dup.Key} = {dup.Value}개");
                }
            }

            public int CalculateActualUsedColumns(List<DynamicDataMapper.ExcelRow> rows)
            {
                int maxCol = 1;
                
                // rows가 null이거나 비어있으면 columnMapping에서 최대값 찾기
                if (rows == null || rows.Count == 0)
                {
                    if (columnMapping.Count > 0)
                    {
                        maxCol = columnMapping.Values.Max();
                    }
                    return maxCol;
                }
                
                // 데이터 행에서 사용된 최대 컬럼 찾기
                foreach (var row in rows)
                {
                    maxCol = Math.Max(maxCol, row.GetMaxColumn());
                }
                return maxCol;
            }
        }

        public ExcelScheme BuildScheme(
            DynamicStructureAnalyzer.StructurePattern pattern,
            DynamicPatternRecognizer.LayoutStrategy strategy,
            dynamic layoutInfo)
        {
            Logger.Information($"Excel 스키마 생성 시작: 전략={strategy}");

            var scheme = new ExcelScheme();
            int currentRow = 2; // 2행부터 시작 (1행은 헤더용으로 비워둠)

            switch (strategy)
            {
                case DynamicPatternRecognizer.LayoutStrategy.Simple:
                    BuildSimpleScheme(scheme, pattern, currentRow);
                    break;

                case DynamicPatternRecognizer.LayoutStrategy.HorizontalExpansion:
                    BuildHorizontalScheme(scheme, pattern, layoutInfo as DynamicHorizontalExpander.HorizontalLayout, currentRow);
                    break;

                case DynamicPatternRecognizer.LayoutStrategy.VerticalNesting:
                    BuildVerticalScheme(scheme, pattern, layoutInfo as DynamicVerticalNester.VerticalLayout, currentRow);
                    break;

                case DynamicPatternRecognizer.LayoutStrategy.Mixed:
                    BuildMixedScheme(scheme, pattern, layoutInfo, currentRow);
                    break;
            }

            // $scheme_end 추가
            scheme.AddSchemeEndRow(scheme.LastSchemaRow + 1);

            Logger.Information($"Excel 스키마 생성 완료: 마지막 행={scheme.LastSchemaRow}");
            return scheme;
        }

        private void BuildSimpleScheme(ExcelScheme scheme, DynamicStructureAnalyzer.StructurePattern pattern, int startRow)
        {
            Logger.Debug("단순 스키마 생성");

            if (pattern.Type == DynamicStructureAnalyzer.PatternType.RootArray)
            {
                // 루트가 배열인 경우
                scheme.AddCell(startRow, 1, "$[]");

                int row = startRow + 1;
                scheme.AddCell(row, 1, "^");

                int col = 2;
                var orderer = new DynamicPropertyOrderer();
                foreach (var prop in orderer.DeterminePropertyOrder(pattern.Properties))
                {
                    scheme.AddCell(row, col, prop);
                    scheme.SetColumnMapping(prop, col);
                    col++;
                }
            }
            else
            {
                // 루트가 객체인 경우
                scheme.AddCell(startRow, 1, "${}");

                int row = startRow + 1;
                int col = 1;
                foreach (var prop in pattern.Properties.Keys)
                {
                    scheme.AddCell(row, col, prop);
                    scheme.SetColumnMapping(prop, col);
                    col++;
                }
            }
        }

        private void BuildHorizontalScheme(
            ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern,
            DynamicHorizontalExpander.HorizontalLayout arrayLayout,
            int startRow)
        {
            Logger.Debug("수평 확장 스키마 생성");

            // 2행: 루트 배열 마커 $[]
            int totalColumns = EstimateTotalColumns(pattern, arrayLayout);
            scheme.AddMergedCell(startRow, 1, totalColumns, "$[]");

            // 3행: ^ 마커와 객체 마커 ${}
            int row = startRow + 1;
            scheme.AddCell(row, 1, "^");
            scheme.AddMergedCell(row, 2, totalColumns, "${}");

            // 4행: ^ 마커와 기본 속성들
            row++;
            scheme.AddCell(row, 1, "^");

            // 사용된 행과 열을 추적
            var usedCells = new HashSet<(int row, int col)>();
            usedCells.Add((row, 1)); // 4행 1열의 ^ 마커

            int col = 2;
            var orderer = new DynamicPropertyOrderer();

            // 속성들을 동적으로 처리
            // 먼저 전체 속성들 디버깅
            Logger.Information($"전체 속성 개수: {pattern.Properties.Count}");
            foreach (var prop in pattern.Properties)
            {
                Logger.Information($"  - '{prop.Key}': IsObject={prop.Value.IsObject}, IsArray={prop.Value.IsArray}, " +
                                 $"ObjectProperties={prop.Value.ObjectProperties?.Count ?? 0}개");
            }

            // 1. 먼저 모든 객체의 하위 속성명을 수집
            var objectSubProperties = new HashSet<string>();
            var objectProps = pattern.Properties
                .Where(p => p.Value.IsObject && !p.Value.IsArray)
                .ToDictionary(p => p.Key, p => p.Value);

            Logger.Information($"객체 속성 목록:");
            foreach (var objProp in objectProps)
            {
                var objProperties = objProp.Value.ObjectProperties ?? objProp.Value.NestedProperties ?? new List<string>();
                Logger.Information($"  - '{objProp.Key}': 하위 속성 {objProperties.Count}개 [{string.Join(", ", objProperties)}]");
                foreach (var subProp in objProperties)
                {
                    objectSubProperties.Add(subProp);
                }
            }

            // 2. 단순 속성들 처리 (객체 내부 속성과 중복되지 않는 것만)
            var simpleProps = pattern.Properties
                .Where(p => !p.Value.IsObject && !p.Value.IsArray)
                .OrderBy(p => p.Value.FirstAppearanceIndex)
                .ToDictionary(p => p.Key, p => p.Value);

            Logger.Information($"단순 속성 개수: {simpleProps.Count}");
            foreach (var prop in simpleProps)
            {
                Logger.Information($"단순 속성 처리: '{prop.Key}'");

                // 주석 처리 - 루트 레벨과 객체 내부 속성은 별개임
                // if (objectSubProperties.Contains(prop.Key))
                // {
                //     Logger.Information($"  -> 스킵됨 - 객체 내부 속성과 중복");
                //     continue;
                // }

                Logger.Information($"  -> 컬럼 {col}에 추가");
                scheme.AddCell(row, col, prop.Key);
                scheme.SetColumnMapping(prop.Key, col);
                usedCells.Add((row, col));
                col++;
            }

            // 3. 객체 속성들 처리 (동적으로)
            foreach (var objProp in objectProps.OrderBy(p => p.Value.FirstAppearanceIndex))
            {
                var objProperties = objProp.Value.ObjectProperties ?? objProp.Value.NestedProperties ?? new List<string>();

                Logger.Information($"{objProp.Key} 객체 처리 시작");
                Logger.Information($"  - IsObject: {objProp.Value.IsObject}");
                Logger.Information($"  - IsArray: {objProp.Value.IsArray}");
                Logger.Information($"  - ObjectProperties: {objProp.Value.ObjectProperties?.Count ?? 0}개");
                Logger.Information($"  - NestedProperties: {objProp.Value.NestedProperties?.Count ?? 0}개");
                Logger.Information($"  - 최종 속성 수: {objProperties.Count}개");
                Logger.Information($"  - 현재 col: {col}");

                if (objProperties.Count > 0)
                {
                    Logger.Information($"  - 속성 목록: [{string.Join(", ", objProperties)}]");
                }

                if (objProperties.Count > 0)
                {
                    int objStartCol = col;
                    int objEndCol = col + objProperties.Count - 1;

                    Logger.Information($"{objProp.Key} 객체 스키마 생성:");
                    Logger.Information($"  - 시작 컬럼: {objStartCol}");
                    Logger.Information($"  - 종료 컬럼: {objEndCol}");
                    Logger.Information($"  - 속성 개수: {objProperties.Count}");
                    Logger.Information($"  - 현재 row: {row}");

                    // 4행: 객체${} 마커 (병합)
                    scheme.AddMergedCell(row, objStartCol, objEndCol, $"{objProp.Key}${{}}");
                    for (int c = objStartCol; c <= objEndCol; c++)
                    {
                        usedCells.Add((row, c));
                    }

                    // 5행: 객체 속성들
                    Logger.Information($"5행 생성 중 (row+1={row + 1}):");
                    int propCol = objStartCol;
                    foreach (var prop in objProperties)
                    {
                        Logger.Information($"  - 컬럼 {propCol}: '{prop}' 속성 추가");
                        scheme.AddCell(row + 1, propCol, prop);
                        scheme.SetColumnMapping($"{objProp.Key}.{prop}", propCol);
                        usedCells.Add((row + 1, propCol));
                        propCol++;
                    }

                    col = objEndCol + 1;
                }
                else
                {
                    // 빈 객체인 경우에도 객체 마커 표시
                    Logger.Information($"{objProp.Key} 빈 객체 스키마 생성: 컬럼 {col}");
                    scheme.AddCell(row, col, $"{objProp.Key}${{}}");
                    scheme.SetColumnMapping(objProp.Key, col);
                    usedCells.Add((row, col));
                    col++;
                }
            }

            // 3. 배열 속성들은 아래에서 처리하므로 여기서는 스킵

            // 4. 배열 속성들 처리
            if (arrayLayout != null && arrayLayout.ArrayLayouts != null)
            {
                Logger.Information($"배열 속성 처리 시작, 현재 col={col}");
                foreach (var array in arrayLayout.ArrayLayouts)
                {
                    Logger.Information($"배열 '{array.Key}' 처리:");
                    var arrayStartCol = col;
                    var arrayTotalColumns = array.Value.TotalColumns;
                    var arrayPattern = pattern.Properties.ContainsKey(array.Key) ? pattern.Properties[array.Key] : null;
                    Logger.Information($"  - 시작 컬럼: {arrayStartCol}");
                    Logger.Information($"  - TotalColumns: {arrayTotalColumns}");

                    if (arrayTotalColumns > 0)
                    {
                        var arrayEndCol = col + arrayTotalColumns - 1;

                        // 복잡한 중첩 구조를 가진 배열의 경우 특별 처리
                        if (arrayPattern != null && arrayPattern.ArrayPattern != null &&
                            arrayPattern.ArrayPattern.ElementProperties != null &&
                            arrayPattern.ArrayPattern.ElementProperties.Any() &&
                            arrayPattern.ArrayPattern.ElementProperties.Any(p => p.Value.IsObject || p.Value.IsArray))
                        {
                            Logger.Information($"복잡한 배열 '{array.Key}' 처리 시작");
                            Logger.Information($"  - arrayStartCol: {arrayStartCol}");
                            Logger.Information($"  - arrayEndCol: {arrayEndCol}");
                            Logger.Information($"  - arrayTotalColumns: {arrayTotalColumns}");

                            // 복잡한 배열의 실제 필요 컬럼 수를 동적으로 계산
                            int actualNeededColumns = CalculateArrayHeaderColumns(arrayPattern, array.Value);
                            var actualEndCol = arrayStartCol + actualNeededColumns - 1;

                            Logger.Information($"  - actualNeededColumns: {actualNeededColumns}");
                            Logger.Information($"  - actualEndCol: {actualEndCol}");

                            // 배열 자체의 마커는 전체 배열 영역을 포함해야 함
                            scheme.AddMergedCell(row, arrayStartCol, actualEndCol, $"{array.Key}$[]");
                            scheme.SetArrayStartColumn(array.Key, arrayStartCol);
                            for (int c = arrayStartCol; c <= actualEndCol; c++)
                            {
                                usedCells.Add((row, c));
                            }

                            // events 배열 요소의 ${} 마커
                            scheme.AddMergedCell(row + 1, arrayStartCol, actualEndCol, "${}");
                            for (int c = arrayStartCol; c <= actualEndCol; c++)
                            {
                                usedCells.Add((row + 1, c));
                            }

                            // events 배열 요소의 속성들 처리
                            int elementCol = arrayStartCol;
                            int nestedRow = row + 2;

                            var specificElementProps = arrayPattern.ArrayPattern.ElementProperties;

                            foreach (var elemProp in specificElementProps.OrderBy(p => p.Value.FirstAppearanceIndex))
                            {
                                Logger.Debug($"  배열 요소 속성 처리: {elemProp.Key}, IsObject={elemProp.Value.IsObject}, ObjectProperties={elemProp.Value.ObjectProperties?.Count ?? 0}");
                                
                                if (elemProp.Key == "activation")
                                {
                                    Logger.Information($"  ★★★ activation 처리 중: IsObject={elemProp.Value.IsObject}, ObjectProperties=[{string.Join(", ", elemProp.Value.ObjectProperties ?? new List<string>())}]");
                                }
                                
                                if (elemProp.Value.IsObject && elemProp.Value.ObjectProperties?.Count > 0)
                                {
                                    // trigger, activation 같은 중첩 객체
                                    int objCols = elemProp.Value.ObjectProperties.Count;
                                    scheme.AddMergedCell(nestedRow, elementCol, elementCol + objCols - 1, $"{elemProp.Key}${{}}");
                                    for (int c = elementCol; c < elementCol + objCols; c++)
                                    {
                                        usedCells.Add((nestedRow, c));
                                    }

                                    // 객체의 속성들
                                    int subCol = elementCol;
                                    foreach (var subProp in elemProp.Value.ObjectProperties)
                                    {
                                        scheme.AddCell(nestedRow + 1, subCol, subProp);
                                        
                                        // ★ 모든 배열 요소에 대해 객체 속성 매핑 생성
                                        for (int arrayElementIndex = 0; arrayElementIndex < array.Value.ElementCount; arrayElementIndex++)
                                        {
                                            string objectPropPath = $"{array.Key}[{arrayElementIndex}].{elemProp.Key}.{subProp}";
                                            scheme.SetColumnMapping(objectPropPath, subCol);
                                            Logger.Debug($"    복잡 배열 객체 매핑: {objectPropPath} -> 컬럼 {subCol}");
                                        }
                                        
                                        usedCells.Add((nestedRow + 1, subCol));
                                        subCol++;
                                    }
                                    elementCol += objCols;
                                }
                                else if (elemProp.Value.IsArray)
                                {
                                    // 중첩 배열 처리 (Option 배열 등)
                                    Logger.Information($"🔧 중첩 배열 '{elemProp.Key}' 처리 시작");
                                    
                                    var nestedArrayPattern = elemProp.Value.ArrayPattern;
                                    Logger.Information($"  ArrayPattern null 여부: {nestedArrayPattern == null}");
                                    if (nestedArrayPattern != null)
                                    {
                                        Logger.Information($"  ElementProperties null 여부: {nestedArrayPattern.ElementProperties == null}");
                                        Logger.Information($"  ElementProperties 개수: {nestedArrayPattern.ElementProperties?.Count ?? 0}");
                                    }
                                    
                                    if (nestedArrayPattern?.ElementProperties != null && nestedArrayPattern.ElementProperties.Any())
                                    {
                                        int arrayCols = nestedArrayPattern.ElementProperties.Count;
                                        scheme.AddMergedCell(nestedRow, elementCol, elementCol + arrayCols - 1, $"{elemProp.Key}$[]");
                                        for (int c = elementCol; c < elementCol + arrayCols; c++)
                                        {
                                            usedCells.Add((nestedRow, c));
                                        }

                                        // 배열 요소의 ${} 마커
                                        scheme.AddMergedCell(nestedRow + 1, elementCol, elementCol + arrayCols - 1, "${}");
                                        for (int c = elementCol; c < elementCol + arrayCols; c++)
                                        {
                                            usedCells.Add((nestedRow + 1, c));
                                        }

                                        // 배열 요소의 속성들
                                        int subCol = elementCol;
                                        foreach (var subProp in nestedArrayPattern.ElementProperties.OrderBy(p => p.Value.FirstAppearanceIndex))
                                        {
                                            scheme.AddCell(nestedRow + 2, subCol, subProp.Key);
                                            
                                            // ★ 모든 상위 배열 요소에 대해 중첩 배열 속성 매핑 생성
                                            for (int arrayElementIndex = 0; arrayElementIndex < array.Value.ElementCount; arrayElementIndex++)
                                            {
                                                // 중첩 배열의 모든 요소에 대해서도 매핑 생성
                                                if (elemProp.Value.ArrayPattern?.MaxSize > 0)
                                                {
                                                    for (int nestedIndex = 0; nestedIndex < elemProp.Value.ArrayPattern.MaxSize; nestedIndex++)
                                                    {
                                                        string nestedArrayPropPath = $"{array.Key}[{arrayElementIndex}].{elemProp.Key}[{nestedIndex}].{subProp.Key}";
                                                        scheme.SetColumnMapping(nestedArrayPropPath, subCol);
                                                        Logger.Debug($"    복잡 배열 중첩 배열 매핑: {nestedArrayPropPath} -> 컬럼 {subCol}");
                                                    }
                                                }
                                                else
                                                {
                                                    // 중첩 배열 요소 개수를 알 수 없는 경우 기본값으로 처리
                                                    string nestedArrayPropPath = $"{array.Key}[{arrayElementIndex}].{elemProp.Key}[0].{subProp.Key}";
                                                    scheme.SetColumnMapping(nestedArrayPropPath, subCol);
                                                    Logger.Debug($"    복잡 배열 중첩 배열 기본 매핑: {nestedArrayPropPath} -> 컬럼 {subCol}");
                                                }
                                            }
                                            
                                            usedCells.Add((nestedRow + 2, subCol));
                                            subCol++;
                                        }
                                        elementCol += arrayCols;
                                    }
                                    else
                                    {
                                        // 빈 배열이거나 분석되지 않은 배열인 경우
                                        Logger.Warning($"⚠️ '{elemProp.Key}' 배열이 빈 배열로 처리됨 - 실제 구조 분석 확인 필요");
                                        scheme.AddCell(nestedRow, elementCol, elemProp.Key);
                                        
                                        // ★ 모든 배열 요소에 대해 빈 배열 매핑 생성
                                        for (int arrayElementIndex = 0; arrayElementIndex < array.Value.ElementCount; arrayElementIndex++)
                                        {
                                            string emptyArrayPath = $"{array.Key}[{arrayElementIndex}].{elemProp.Key}";
                                            scheme.SetColumnMapping(emptyArrayPath, elementCol);
                                            Logger.Debug($"    복잡 배열 빈 배열 매핑: {emptyArrayPath} -> 컬럼 {elementCol}");
                                        }
                                        
                                        usedCells.Add((nestedRow, elementCol));
                                        elementCol++;
                                    }
                                }
                                else
                                {
                                    // 단순 속성 또는 일부에만 있는 객체
                                    if (elemProp.Key == "activation")
                                    {
                                        Logger.Warning($"  ⚠️ activation이 단순 속성으로 처리됨! IsObject={elemProp.Value.IsObject}, ObjectProperties={elemProp.Value.ObjectProperties?.Count ?? 0}");
                                    }
                                    scheme.AddCell(nestedRow, elementCol, elemProp.Key);
                                    
                                    // ★ 모든 배열 요소에 대해 단순 속성 매핑 생성
                                    for (int arrayElementIndex = 0; arrayElementIndex < array.Value.ElementCount; arrayElementIndex++)
                                    {
                                        string simplePropPath = $"{array.Key}[{arrayElementIndex}].{elemProp.Key}";
                                        scheme.SetColumnMapping(simplePropPath, elementCol);
                                        Logger.Debug($"    복잡 배열 단순 속성 매핑: {simplePropPath} -> 컬럼 {elementCol}");
                                    }
                                    
                                    usedCells.Add((nestedRow, elementCol));
                                    elementCol++;
                                }
                            }

                            // 복잡한 배열의 경우 전체 배열 영역 끝까지 col 증가
                            Logger.Information($"복잡한 배열 처리 완료: col 변경 전={col}, actualEndCol={actualEndCol}, arrayEndCol={arrayEndCol}");
                            col = actualEndCol + 1;
                            Logger.Information($"복잡한 배열 처리 완료: col 변경 후={col}");

                            // 나머지 빈 공간은 나중에 FillEmptyCellsWithCaretMarker에서 처리됨
                        }
                        else
                        {
                            // 기존 로직 유지 - 단순 배열
                            scheme.AddMergedCell(row, arrayStartCol, arrayEndCol, $"{array.Key}$[]");
                            scheme.SetArrayStartColumn(array.Key, arrayStartCol);
                            for (int c = arrayStartCol; c <= arrayEndCol; c++)
                            {
                                usedCells.Add((row, c));
                            }

                            // 각 요소별 처리
                            BuildArrayElementScheme(scheme, array.Value, row + 1, arrayStartCol, usedCells, pattern);

                            col = arrayEndCol + 1;
                        }
                    }
                }
            }

            // 5. 모든 데이터 헤더 작성 후, 빈 셀에 ^ 마커 추가
            Logger.Information($"FillEmptyCellsWithCaretMarker 호출: startRow={startRow}, totalColumns={totalColumns}, col={col}");
            FillEmptyCellsWithCaretMarker(scheme, startRow, totalColumns, usedCells);
        }

        private void BuildArrayElementScheme(
            ExcelScheme scheme,
            DynamicHorizontalExpander.DynamicArrayLayout layout,
            int startRow,
            int startCol,
            HashSet<(int row, int col)> usedCells,
            DynamicStructureAnalyzer.StructurePattern pattern = null)
        {
            Logger.Debug($"배열 요소 스키마 생성: {layout.ArrayPath}, 요소 수={layout.ElementCount}");
            Logger.Information($"  OptimizeColumns: {layout.OptimizeColumns}");
            Logger.Information($"  TotalColumns: {layout.TotalColumns}");
            Logger.Information($"  Elements.Count: {layout.Elements?.Count ?? 0}");

            int currentCol = startCol;

            // OptimizeColumns가 false면 인덱스별 개별 스키마 사용
            if (!layout.OptimizeColumns && layout.Elements.Any())
            {
                // 인덱스별 개별 스키마 모드
                Logger.Debug($"인덱스별 개별 스키마 모드");

                for (int i = 0; i < layout.Elements.Count; i++)
                {
                    var element = layout.Elements[i];
                    var elementProps = element.UnifiedProperties ?? element.Properties;
                    int elementColumns = element.RequiredColumns;

                    Logger.Debug($"  요소 [{i}]: {elementColumns}개 컬럼, 속성: {string.Join(", ", elementProps)}");

                    // 5행: ${} 마커
                    scheme.AddMergedCell(startRow, currentCol, currentCol + elementColumns - 1, "${}");
                    for (int c = currentCol; c < currentCol + elementColumns; c++)
                    {
                        usedCells.Add((startRow, c));
                    }

                    // 6행: 해당 인덱스의 속성들을 헤더로 추가
                    int propCol = currentCol;
                    foreach (var prop in elementProps)
                    {
                        scheme.AddCell(startRow + 1, propCol, prop);
                        scheme.SetColumnMapping($"{layout.ArrayPath}[{i}].{prop}", propCol);
                        
                        // 중첩 배열인 경우 하위 요소들도 매핑
                        if (element.NestedArrays != null && element.NestedArrays.ContainsKey(prop))
                        {
                            var nestedArray = element.NestedArrays[prop];
                            AddNestedArrayMappings(scheme, $"{layout.ArrayPath}[{i}].{prop}", nestedArray, propCol);
                        }
                        
                        usedCells.Add((startRow + 1, propCol));
                        propCol++;
                    }

                    currentCol += elementColumns;
                }

                return;
            }
            else if (layout.OptimizeColumns && layout.Elements.Any())
            {
                // 통합 스키마 모드: 모든 요소를 하나의 통합된 스키마로 표현
                Logger.Debug($"통합 스키마 모드 - 모든 배열 요소에 대해 매핑 생성");
                
                // 첫 번째 요소를 기준으로 스키마 생성
                var firstElement = layout.Elements[0];
                var orderedProps = firstElement.UnifiedProperties ?? firstElement.Properties;
                int totalRequiredColumns = firstElement.RequiredColumns;
                
                Logger.Information($"통합 스키마 - 기준 요소: {orderedProps.Count}개 속성, {totalRequiredColumns}개 컬럼");
                Logger.Information($"  속성 목록: [{string.Join(", ", orderedProps)}]");

                // ${} 마커
                scheme.AddMergedCell(startRow, currentCol, currentCol + totalRequiredColumns - 1, "${}");
                for (int c = currentCol; c < currentCol + totalRequiredColumns; c++)
                {
                    usedCells.Add((startRow, c));
                }

                // 속성들
                int propCol = currentCol;
                foreach (var prop in orderedProps)
                {
                    scheme.AddCell(startRow + 1, propCol, prop);
                    
                    // ★ 중요: 모든 배열 요소(인덱스)에 대해 매핑 생성
                    for (int elementIndex = 0; elementIndex < layout.Elements.Count; elementIndex++)
                    {
                        string elementPath = $"{layout.ArrayPath}[{elementIndex}].{prop}";
                        scheme.SetColumnMapping(elementPath, propCol);
                        Logger.Debug($"  통합 매핑: {elementPath} -> 컬럼 {propCol}");
                        
                        // 중첩 배열인 경우 하위 요소들도 매핑
                        var currentElement = layout.Elements[elementIndex];
                        if (currentElement.NestedArrays != null && currentElement.NestedArrays.ContainsKey(prop))
                        {
                            var nestedArray = currentElement.NestedArrays[prop];
                            AddNestedArrayMappings(scheme, elementPath, nestedArray, propCol);
                        }
                    }
                    
                    usedCells.Add((startRow + 1, propCol));
                    propCol++;
                }

                currentCol += totalRequiredColumns;
            }
            else 
            {
                // 배열 요소가 없는 경우나 기타 예외 상황 처리
                Logger.Warning($"배열 요소가 없거나 처리할 수 없는 상황: Elements={layout.Elements?.Count ?? 0}");
                
                // 최소한의 빈 배열 스키마라도 생성
                scheme.AddCell(startRow, currentCol, "${}");
                usedCells.Add((startRow, currentCol));
                
                // 기본 컬럼 매핑 추가
                scheme.SetColumnMapping($"{layout.ArrayPath}[0]", currentCol);
            }
        }

        private void AddNestedArrayMappings(ExcelScheme scheme, string arrayPath, DynamicHorizontalExpander.DynamicArrayLayout nestedArray, int columnIndex)
        {
            Logger.Debug($"중첩 배열 매핑 추가: {arrayPath}, 컬럼 {columnIndex}");
            Logger.Information($"  중첩 배열 요소 수: {nestedArray.Elements?.Count ?? 0}");
            
            // 중첩 배열 전체를 해당 컬럼에 매핑
            scheme.SetColumnMapping(arrayPath, columnIndex);
            
            // 중첩 배열의 각 요소별 세부 매핑도 추가
            if (nestedArray.Elements != null && nestedArray.Elements.Count > 0)
            {
                // 모든 중첩 배열 요소에 대해 매핑 생성
                Logger.Information($"  중첩 배열의 모든 요소({nestedArray.Elements.Count}개)에 대해 매핑 생성");
                
                for (int i = 0; i < nestedArray.Elements.Count; i++)
                {
                    var element = nestedArray.Elements[i];
                    var properties = element.UnifiedProperties ?? element.Properties;
                    
                    Logger.Debug($"    요소 [{i}]: {properties.Count}개 속성");
                    
                    foreach (var prop in properties)
                    {
                        string elementPath = $"{arrayPath}[{i}].{prop}";
                        scheme.SetColumnMapping(elementPath, columnIndex);
                        Logger.Debug($"      중첩 요소 매핑: {elementPath} -> 컬럼 {columnIndex}");
                    }
                    
                    // 중첩 배열 자체도 매핑 (예: SpawnData[0], SpawnData[1], SpawnData[2])
                    string nestedElementPath = $"{arrayPath}[{i}]";
                    scheme.SetColumnMapping(nestedElementPath, columnIndex);
                    Logger.Debug($"      중첩 배열 요소 매핑: {nestedElementPath} -> 컬럼 {columnIndex}");
                }
            }
            else
            {
                // 빈 중첩 배열인 경우에도 기본 매핑 생성
                Logger.Debug($"  빈 중첩 배열 - 기본 매핑만 생성");
                string defaultPath = $"{arrayPath}[0]";
                scheme.SetColumnMapping(defaultPath, columnIndex);
                Logger.Debug($"    기본 매핑: {defaultPath} -> 컬럼 {columnIndex}");
            }
        }

        private void BuildVerticalScheme(
            ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern,
            DynamicVerticalNester.VerticalLayout verticalLayout,
            int startRow)
        {
            Logger.Debug("수직 중첩 스키마 생성");

            // 루트 배열 마커
            scheme.AddCell(startRow, 1, "$[]");

            int row = startRow + 1;

            // 기본 속성들 배치
            foreach (var mapping in verticalLayout.ColumnMapping.OrderBy(m => m.Value))
            {
                scheme.AddCell(row, mapping.Value, mapping.Key);
                scheme.SetColumnMapping(mapping.Key, mapping.Value);
            }

            // 중첩 구조가 있는 경우 추가 마커
            if (verticalLayout.RequiresMerging && !string.IsNullOrEmpty(verticalLayout.MergeKey))
            {
                // 병합 키 표시를 위한 특별 마커 추가 가능
                Logger.Debug($"병합 키 감지: {verticalLayout.MergeKey}");
            }
        }

        private void BuildMixedScheme(
            ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern,
            dynamic layoutInfo,
            int startRow)
        {
            Logger.Debug("혼합 스키마 생성");

            // 혼합 전략은 수평과 수직의 조합
            // 기본적으로 수평 확장을 사용하되, 특정 조건에서 수직 확장 추가

            if (layoutInfo is DynamicHorizontalExpander.HorizontalLayout horizontalLayout)
            {
                BuildHorizontalScheme(scheme, pattern, horizontalLayout, startRow);
            }
            else
            {
                // 폴백: 단순 스키마 생성
                BuildSimpleScheme(scheme, pattern, startRow);
            }
        }


        private int BuildNestedPropertyScheme(
            ExcelScheme scheme,
            string propertyName,
            string parentPath,
            int row,
            int col,
            HashSet<(int row, int col)> usedCells,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            var fullPath = string.IsNullOrEmpty(parentPath) ? propertyName : $"{parentPath}.{propertyName}";

            // 패턴에서 속성 정보 찾기
            var propPattern = pattern?.Properties?.ContainsKey(propertyName) == true ? pattern.Properties[propertyName] : null;

            if (propPattern != null && propPattern.IsObject && propPattern.ObjectProperties?.Count > 0)
            {
                // 중첩된 객체 처리
                int objStartCol = col;
                int totalCols = 0;

                // 객체 마커
                scheme.AddMergedCell(row, objStartCol, objStartCol + propPattern.ObjectProperties.Count - 1, $"{propertyName}${{}}");
                for (int c = objStartCol; c < objStartCol + propPattern.ObjectProperties.Count; c++)
                {
                    usedCells.Add((row, c));
                }

                // 객체의 각 속성을 재귀적으로 처리
                int subCol = objStartCol;
                foreach (var objProp in propPattern.ObjectProperties)
                {
                    // 중첩된 패턴 정보 가져오기
                    var subPropPattern = propPattern.NestedPatterns?.ContainsKey(objProp) == true
                        ? propPattern.NestedPatterns[objProp]
                        : null;

                    // 하위 속성에 대한 패턴 정보 생성
                    var subPattern = new DynamicStructureAnalyzer.StructurePattern
                    {
                        Properties = subPropPattern != null
                            ? new Dictionary<string, DynamicStructureAnalyzer.PropertyPattern> { { objProp, subPropPattern } }
                            : new Dictionary<string, DynamicStructureAnalyzer.PropertyPattern>()
                    };

                    // 재귀 호출로 더 깊은 중첩 처리
                    var subCols = BuildNestedPropertyScheme(
                        scheme,
                        objProp,
                        fullPath,
                        row + 1,
                        subCol,
                        usedCells,
                        subPattern
                    );

                    subCol += subCols;
                    totalCols += subCols;
                }

                return Math.Max(totalCols, propPattern.ObjectProperties.Count);
            }
            else if (propPattern != null && propPattern.IsArray)
            {
                // 중첩된 배열 처리
                if (propPattern.ArrayPattern != null && propPattern.ArrayPattern.ElementProperties != null && propPattern.ArrayPattern.ElementProperties.Count > 0)
                {
                    // 배열 요소가 객체인 경우, 배열의 첫 번째 요소의 속성들을 헤더로 추가
                    int arrayStartCol = col;
                    var elementProps = propPattern.ArrayPattern.ElementProperties;
                    int totalCols = 0;

                    // 배열 마커
                    scheme.AddMergedCell(row, arrayStartCol, arrayStartCol + elementProps.Count - 1, $"{propertyName}$[]");
                    for (int c = arrayStartCol; c < arrayStartCol + elementProps.Count; c++)
                    {
                        usedCells.Add((row, c));
                    }

                    // 배열 요소의 ${} 마커
                    scheme.AddMergedCell(row + 1, arrayStartCol, arrayStartCol + elementProps.Count - 1, "${}");
                    for (int c = arrayStartCol; c < arrayStartCol + elementProps.Count; c++)
                    {
                        usedCells.Add((row + 1, c));
                    }

                    // 배열 요소의 속성들
                    int subCol = arrayStartCol;
                    foreach (var elemProp in elementProps)
                    {
                        // 배열의 첫 번째 요소 기준으로 스키마 생성
                        scheme.AddCell(row + 2, subCol, elemProp.Key);
                        scheme.SetColumnMapping($"{fullPath}[0].{elemProp.Key}", subCol);
                        usedCells.Add((row + 2, subCol));
                        Logger.Debug($"배열 요소 속성 매핑: {fullPath}[0].{elemProp.Key} -> 컬럼 {subCol}");
                        subCol++;
                        totalCols++;
                    }

                    return Math.Max(totalCols, elementProps.Count);
                }
                else
                {
                    // 단순 배열
                    scheme.AddCell(row, col, propertyName);
                    scheme.SetColumnMapping(fullPath, col);
                    usedCells.Add((row, col));
                    return 1;
                }
            }
            else
            {
                // 일반 속성
                scheme.AddCell(row, col, propertyName);
                scheme.SetColumnMapping(fullPath, col);
                usedCells.Add((row, col));
                Logger.Debug($"속성 매핑: {fullPath} -> 컬럼 {col}");
                return 1;
            }
        }

        private void FillEmptyCellsWithCaretMarker(ExcelScheme scheme, int startRow, int totalColumns, HashSet<(int row, int col)> usedCells)
        {
            Logger.Debug("빈 셀에 ^ 마커 추가 시작");

            // 마지막 데이터 행 찾기
            int lastDataRow = scheme.LastSchemaRow;

            // startRow부터 lastDataRow까지 모든 빈 셀에 ^ 마커 추가
            for (int row = startRow; row <= lastDataRow; row++)
            {
                // 현재 행의 병합된 셀 정보 가져오기
                var mergedCells = scheme.GetMergedCellsInRow(row);

                for (int col = 1; col <= totalColumns; col++)
                {
                    // 병합된 셀 범위에 포함되는지 확인
                    bool isInMergedRange = false;
                    foreach (var merged in mergedCells)
                    {
                        if (col >= merged.StartColumn && col <= merged.EndColumn)
                        {
                            isInMergedRange = true;
                            break;
                        }
                    }

                    // 병합된 셀 범위에 포함되지 않고, 사용되지 않은 셀에만 ^ 마커 추가
                    if (!isInMergedRange && !usedCells.Contains((row, col)))
                    {
                        Logger.Debug($"빈 셀 발견: 행={row}, 열={col} - ^ 마커 추가");
                        scheme.AddCell(row, col, "^");
                    }
                }
            }

            Logger.Debug("빈 셀에 ^ 마커 추가 완료");
        }

        private int CalculateArrayHeaderColumns(DynamicStructureAnalyzer.PropertyPattern arrayPattern, DynamicHorizontalExpander.DynamicArrayLayout arrayLayout)
        {
            // 배열 헤더에 필요한 실제 컬럼 수를 계산
            Logger.Information($"CalculateArrayHeaderColumns 시작: 배열 요소 수={arrayLayout.ElementCount}, TotalColumns={arrayLayout.TotalColumns}");

            // 복잡한 중첩 구조를 가진 배열의 경우, 모든 하위 요소의 컬럼을 포함해야 함
            if (arrayPattern?.ArrayPattern?.ElementProperties != null &&
                arrayPattern.ArrayPattern.ElementProperties.Any(p => p.Value.IsObject || p.Value.IsArray))
            {
                int totalColumns = 0;
                var elementProps = arrayPattern.ArrayPattern.ElementProperties;

                Logger.Information($"복잡한 중첩 배열, 요소 속성 수: {elementProps.Count}");

                // 각 속성의 실제 컬럼 수 계산
                foreach (var prop in elementProps)
                {
                    if (prop.Value.IsObject && prop.Value.ObjectProperties?.Count > 0)
                    {
                        // 객체의 속성 수
                        totalColumns += prop.Value.ObjectProperties.Count;
                        Logger.Information($"  - 객체 '{prop.Key}': {prop.Value.ObjectProperties.Count}개 컬럼");
                    }
                    else if (prop.Value.IsArray && prop.Value.ArrayPattern?.ElementProperties != null)
                    {
                        // 중첩 배열의 요소 속성 수
                        int nestedColumns = 0;
                        foreach (var nestedProp in prop.Value.ArrayPattern.ElementProperties)
                        {
                            if (nestedProp.Value.IsObject)
                            {
                                nestedColumns += nestedProp.Value.ObjectProperties?.Count ?? 1;
                            }
                            else
                            {
                                nestedColumns += 1;
                            }
                        }
                        totalColumns += nestedColumns;
                        Logger.Information($"  - 배열 '{prop.Key}': {nestedColumns}개 컬럼");
                    }
                    else
                    {
                        // 단순 속성
                        totalColumns += 1;
                        Logger.Information($"  - 단순 '{prop.Key}': 1개 컬럼");
                    }
                }

                Logger.Information($"복잡한 중첩 배열 총 컬럼 수: {totalColumns}");
                return totalColumns;
            }

            // 일반적인 배열의 경우 기존 로직 사용
            Logger.Information($"일반 배열 - TotalColumns 반환: {arrayLayout.TotalColumns}");
            return arrayLayout.TotalColumns;
        }

        private int EstimateTotalColumns(DynamicStructureAnalyzer.StructurePattern pattern, DynamicHorizontalExpander.HorizontalLayout layout)
        {
            Logger.Information("EstimateTotalColumns 계산 시작");

            // ^ 마커
            int columns = 1;
            Logger.Information($"  ^ 마커: 1");

            // 단순 속성들
            int simplePropsCount = pattern.Properties.Count(p => !p.Value.IsObject && !p.Value.IsArray);
            columns += simplePropsCount;
            Logger.Information($"  단순 속성: {simplePropsCount}");

            // 객체 속성들의 하위 속성들
            var objectProps = pattern.Properties.Where(p => p.Value.IsObject && !p.Value.IsArray);
            foreach (var objProp in objectProps)
            {
                var objProperties = objProp.Value.ObjectProperties ?? objProp.Value.NestedProperties ?? new List<string>();
                // 빈 객체도 1개의 컬럼을 차지함
                int objColumns = objProperties.Count > 0 ? objProperties.Count : 1;
                columns += objColumns;
                Logger.Information($"  객체 '{objProp.Key}'의 속성: {objColumns} (실제 하위 속성: {objProperties.Count}개)");
            }

            // 배열 속성들 - 더 정확한 계산
            if (layout != null && layout.ArrayLayouts.Any())
            {
                foreach (var array in layout.ArrayLayouts)
                {
                    Logger.Information($"  배열 '{array.Key}' 처리");

                    // 복잡한 중첩 구조를 가진 배열의 경우 직접 계산
                    var arrayPattern = pattern.Properties.ContainsKey(array.Key) ? pattern.Properties[array.Key] : null;
                    if (arrayPattern?.ArrayPattern?.ElementProperties != null &&
                        arrayPattern.ArrayPattern.ElementProperties.Any(p => p.Value.IsObject || p.Value.IsArray))
                    {
                        int complexColumns = 0;
                        foreach (var elemProp in arrayPattern.ArrayPattern.ElementProperties)
                        {
                            if (elemProp.Value.IsObject && elemProp.Value.ObjectProperties?.Count > 0)
                            {
                                complexColumns += elemProp.Value.ObjectProperties.Count;
                                Logger.Information($"    - 객체 '{elemProp.Key}': {elemProp.Value.ObjectProperties.Count}개");
                            }
                            else if (elemProp.Value.IsArray && elemProp.Value.ArrayPattern?.ElementProperties != null)
                            {
                                complexColumns += elemProp.Value.ArrayPattern.ElementProperties.Count;
                                Logger.Information($"    - 배열 '{elemProp.Key}': {elemProp.Value.ArrayPattern.ElementProperties.Count}개");
                            }
                            else
                            {
                                complexColumns += 1;
                                Logger.Information($"    - 단순 '{elemProp.Key}': 1개");
                            }
                        }
                        columns += complexColumns;
                        Logger.Information($"    복잡한 배열 총 컬럼: {complexColumns}");
                    }
                    else
                    {
                        columns += array.Value.TotalColumns;
                        Logger.Information($"    일반 배열 TotalColumns: {array.Value.TotalColumns}");
                    }
                }
            }

            Logger.Information($"EstimateTotalColumns 최종 결과: {columns}");
            return columns;
        }
    }
}