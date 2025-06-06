using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Logging;

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
                return columnMapping.ContainsKey(propertyName) ? columnMapping[propertyName] : -1;
            }

            public int GetArrayStartColumn(string arrayName)
            {
                return arrayStartColumns.ContainsKey(arrayName) ? arrayStartColumns[arrayName] : -1;
            }

            public void SetColumnMapping(string propertyName, int column)
            {
                columnMapping[propertyName] = column;
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

            Logger.Information($"Excel 스키마 생성 완료: 마지막 행={scheme.LastSchemaRow + 1}");
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

            // 루트 배열 마커
            scheme.AddCell(startRow, 1, "$[]");

            // ^ 마커와 기본 속성들
            int row = startRow + 1;
            scheme.AddCell(row, 1, "^");

            int col = 2;
            var orderer = new DynamicPropertyOrderer();

            // 단순 속성들
            var simpleProps = pattern.Properties
                .Where(p => !p.Value.IsArray)
                .ToDictionary(p => p.Key, p => p.Value);

            foreach (var prop in orderer.DeterminePropertyOrder(simpleProps))
            {
                scheme.AddCell(row, col, prop);
                scheme.SetColumnMapping(prop, col);
                col++;
            }

            // 배열 속성 처리
            if (arrayLayout != null && arrayLayout.ArrayLayouts != null)
            {
                foreach (var array in arrayLayout.ArrayLayouts)
                {
                    var arrayStartCol = col;
                    var totalColumns = array.Value.TotalColumns;
                    
                    if (totalColumns > 0)
                    {
                        var arrayEndCol = col + totalColumns - 1;

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
            int currentCol = startCol;

            // 각 배열 요소에 대한 스키마
            foreach (var element in layout.Elements)
            {
                if (element.RequiredColumns > 0)
                {
                    // ${} 마커
                    scheme.AddMergedCell(startRow, currentCol,
                        currentCol + element.RequiredColumns - 1, "${}");

                    // 속성들
                    int propCol = currentCol;
                    foreach (var prop in element.Properties)
                    {
                        scheme.AddCell(startRow + 1, propCol++, prop);
                    }

                    currentCol += element.RequiredColumns;
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
    }
}