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
                    .Information($"SetColumnMapping: '{propertyName}' -> 컬럼 {column}");
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

            int col = 2;
            var orderer = new DynamicPropertyOrderer();

            // 단순 속성들 순서: id, role, damage 등
            var simpleProps = pattern.Properties
                .Where(p => !p.Value.IsArray && !p.Value.IsObject)
                .ToDictionary(p => p.Key, p => p.Value);
            
            // 객체 속성들 (desc 같은 것)
            var objectProps = pattern.Properties
                .Where(p => p.Value.IsObject && !p.Value.IsArray)
                .ToDictionary(p => p.Key, p => p.Value);

            foreach (var prop in orderer.DeterminePropertyOrder(simpleProps))
            {
                scheme.AddCell(row, col, prop);
                scheme.SetColumnMapping(prop, col);
                col++;
            }
            
            // 객체 속성 처리 (동적으로)
            foreach (var objProp in objectProps)
            {
                // 객체의 하위 속성들을 동적으로 가져옴
                var objProperties = objProp.Value.NestedProperties ?? new List<string>();
                int objPropCount = objProperties.Count;
                
                if (objPropCount > 0)
                {
                    int objEndCol = col + objPropCount - 1;
                    
                    // 객체 마커 (병합)
                    scheme.AddMergedCell(row, col, objEndCol, $"{objProp.Key}${{}}");
                    
                    // 5행: 객체의 하위 속성들
                    int propCol = col;
                    foreach (var prop in objProperties)
                    {
                        scheme.AddCell(row + 1, propCol, prop);
                        scheme.SetColumnMapping($"{objProp.Key}.{prop}", propCol);
                        Logger.Debug($"객체 속성 매핑: {objProp.Key}.{prop} -> 컬럼 {propCol}");
                        propCol++;
                    }
                    
                    col = objEndCol + 1;
                }
                else
                {
                    // 속성이 없는 객체
                    scheme.AddCell(row, col, objProp.Key);
                    scheme.SetColumnMapping(objProp.Key, col);
                    col++;
                }
            }

            // 배열 속성 처리
            if (arrayLayout != null && arrayLayout.ArrayLayouts != null)
            {
                foreach (var array in arrayLayout.ArrayLayouts)
                {
                    var arrayStartCol = col;
                    var arrayTotalColumns = array.Value.TotalColumns;
                    
                    if (arrayTotalColumns > 0)
                    {
                        var arrayEndCol = col + arrayTotalColumns - 1;

                        // 배열 마커 (병합)
                        scheme.AddMergedCell(row, arrayStartCol, arrayEndCol, $"{array.Key}$[]");
                        scheme.SetArrayStartColumn(array.Key, arrayStartCol);

                        // 각 요소별 처리
                        BuildArrayElementScheme(scheme, array.Value, row + 1, arrayStartCol);

                        col = arrayEndCol + 1;
                    }
                }
            }
        }

        private void BuildArrayElementScheme(
            ExcelScheme scheme,
            DynamicHorizontalExpander.DynamicArrayLayout layout,
            int startRow,
            int startCol)
        {
            Logger.Debug($"배열 요소 스키마 생성: {layout.ArrayPath}, 요소 수={layout.ElementCount}");
            
            int currentCol = startCol;

            // 가변 속성을 가진 배열의 경우 통합 스키마 사용
            if (layout.OptimizeColumns && layout.Elements.Any())
            {
                // 각 요소별로 동적으로 처리
                for (int i = 0; i < layout.ElementCount; i++)
                {
                    var element = i < layout.Elements.Count ? layout.Elements[i] : layout.Elements.Last();
                    var elementProps = element.UnifiedProperties ?? element.Properties;
                    int elementColumns = elementProps.Count;
                    
                    // 5행: ${} 마커
                    scheme.AddMergedCell(startRow, currentCol, currentCol + elementColumns - 1, "${}");

                    // 6행: 요소의 속성들 (동적으로)
                    int propCol = currentCol;
                    foreach (var prop in elementProps)
                    {
                        scheme.AddCell(startRow + 1, propCol, prop);
                        scheme.SetColumnMapping($"{layout.ArrayPath}[{i}].{prop}", propCol);
                        Logger.Debug($"배열 요소 매핑: {layout.ArrayPath}[{i}].{prop} -> 컬럼 {propCol}");
                        propCol++;
                    }
                    
                    currentCol += elementColumns;
                }
            }
            else
            {
                // 기존 로직: 각 요소별 개별 스키마
                foreach (var element in layout.Elements)
                {
                    if (element.RequiredColumns > 0)
                    {
                        // ${} 마커
                        scheme.AddMergedCell(startRow, currentCol,
                            currentCol + element.RequiredColumns - 1, "${}");

                        // 속성들
                        int propCol = currentCol;
                        var orderedProps = element.UnifiedProperties ?? element.Properties;
                        foreach (var prop in orderedProps)
                        {
                            scheme.AddCell(startRow + 1, propCol++, prop);
                        }

                        currentCol += element.RequiredColumns;
                    }
                }
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
        
        private void AddObjectProperties(ExcelScheme scheme, int row, int startCol, int count, string objectName = "desc")
        {
            // desc 객체의 속성들
            string[] descProps = { "keyword", "name", "skillName", "skillKeyword", "skillExplanation", "skillRange", "skillCooltime" };
            
            // ^ 마커들
            scheme.AddCell(row, 1, "^");
            
            // 속성들
            for (int i = 0; i < descProps.Length && i < count; i++)
            {
                scheme.AddCell(row, startCol + i, descProps[i]);
                // 객체명.속성명 형태로 컬럼 매핑 추가
                scheme.SetColumnMapping($"{objectName}.{descProps[i]}", startCol + i);
            }
        }
        
        private void AddDynamicObjectProperties(ExcelScheme scheme, int row, int startCol, string objectName, List<string> properties)
        {
            // ^ 마커
            scheme.AddCell(row, 1, "^");
            
            // 속성들
            for (int i = 0; i < properties.Count; i++)
            {
                scheme.AddCell(row, startCol + i, properties[i]);
                // 객체명.속성명 형태로 컬럼 매핑 추가
                scheme.SetColumnMapping($"{objectName}.{properties[i]}", startCol + i);
            }
        }
        
        private List<string> GetObjectProperties(DynamicStructureAnalyzer.StructurePattern pattern, string objectName)
        {
            // 패턴에서 객체 속성 찾기
            if (pattern.Properties.ContainsKey(objectName))
            {
                var prop = pattern.Properties[objectName];
                if (prop.IsObject && prop.ObjectProperties != null)
                {
                    return prop.ObjectProperties;
                }
            }
            
            // 기본값: desc 객체의 경우 하드코듩된 속성 반환
            if (objectName == "desc")
            {
                return new List<string> { "keyword", "name", "skillName", "skillKeyword", "skillExplanation", "skillRange", "skillCooltime" };
            }
            
            return new List<string>();
        }
        
        private int EstimateTotalColumns(DynamicStructureAnalyzer.StructurePattern pattern, DynamicHorizontalExpander.HorizontalLayout layout)
        {
            // 기본 컬럼: ^ + 단순 속성들
            int baseColumns = 1 + pattern.Properties.Count(p => !p.Value.IsArray && !p.Value.IsObject);
            
            // 객체 속성 (desc)
            int objectColumns = 7; // desc의 속성 수
            
            // 배열 속성 (weaponSpec)
            int arrayColumns = 0;
            if (layout != null && layout.ArrayLayouts.Any())
            {
                // weaponSpec: 5개는 4컬럼, 3개는 5컬럼 = 35컬럼
                arrayColumns = 35;
            }
            
            return baseColumns + objectColumns + arrayColumns;
        }
    }
}