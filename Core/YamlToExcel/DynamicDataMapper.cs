using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Logging;
using YamlDotNet.RepresentationModel;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    /// <summary>
    /// YAML 데이터를 Excel 행으로 매핑하는 클래스
    /// </summary>
    public class DynamicDataMapper
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<DynamicDataMapper>();

        /// <summary>
        /// Excel 행 데이터를 표현하는 클래스
        /// </summary>
        public class ExcelRow
        {
            private readonly Dictionary<int, object> cells = new Dictionary<int, object>();

            public void SetCell(int column, object value)
            {
                cells[column] = value;
            }

            public object GetCell(int column)
            {
                return cells.ContainsKey(column) ? cells[column] : null;
            }

            public void WriteToWorksheet(IXLWorksheet worksheet, int rowNumber)
            {
                foreach (var cell in cells)
                {
                    var xlCell = worksheet.Cell(rowNumber, cell.Key);
                    
                    if (cell.Value == null)
                    {
                        xlCell.Value = "";
                    }
                    else if (cell.Value is bool boolValue)
                    {
                        xlCell.Value = boolValue;
                    }
                    else if (cell.Value is int intValue)
                    {
                        xlCell.Value = intValue;
                    }
                    else if (cell.Value is double doubleValue)
                    {
                        xlCell.Value = doubleValue;
                    }
                    else if (cell.Value is DateTime dateValue)
                    {
                        xlCell.Value = dateValue;
                    }
                    else
                    {
                        xlCell.Value = cell.Value.ToString();
                    }
                }
            }

            public Dictionary<int, object> GetAllCells()
            {
                return new Dictionary<int, object>(cells);
            }
        }

        public List<ExcelRow> MapToExcelRows(
            YamlNode data,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            Logger.Information("YAML 데이터를 Excel 행으로 매핑 시작");
            var rows = new List<ExcelRow>();

            if (data is YamlSequenceNode sequence)
            {
                Logger.Debug($"시퀀스 노드 처리: {sequence.Children.Count}개 항목");
                foreach (var item in sequence.Children)
                {
                    var mappedRows = MapItem(item, scheme, pattern);
                    rows.AddRange(mappedRows);
                }
            }
            else if (data is YamlMappingNode mapping)
            {
                Logger.Debug("단일 매핑 노드 처리");
                var row = MapSingleItem(mapping, scheme, pattern);
                rows.Add(row);
            }

            Logger.Information($"매핑 완료: {rows.Count}개 행 생성");
            return rows;
        }

        private List<ExcelRow> MapItem(
            YamlNode item,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            var rows = new List<ExcelRow>();

            // 수직 확장이 필요한지 확인
            bool needsVerticalExpansion = pattern.Arrays.Any(a => 
                a.Value.RequiresMultipleRows || 
                (a.Value.HasVariableStructure && a.Value.MaxSize > 3));

            if (needsVerticalExpansion)
            {
                // 수직 확장 필요
                rows.AddRange(ExpandVertically(item, scheme, pattern));
            }
            else
            {
                // 단일 행 매핑
                rows.Add(MapHorizontally(item, scheme, pattern));
            }

            return rows;
        }

        private ExcelRow MapHorizontally(
            YamlNode item,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            var row = new ExcelRow();

            if (item is YamlMappingNode mapping)
            {
                // ^ 마커 (무시 마커)
                row.SetCell(1, "^");

                // 속성 매핑
                foreach (var prop in mapping.Children)
                {
                    var key = prop.Key.ToString();
                    var columnIndex = scheme.GetColumnIndex(key);

                    if (columnIndex > 0)
                    {
                        if (prop.Value is YamlSequenceNode nestedArray)
                        {
                            // 중첩 배열은 별도 처리
                            MapNestedArray(row, nestedArray, key, scheme, pattern);
                        }
                        else
                        {
                            var value = ConvertValue(prop.Value);
                            row.SetCell(columnIndex, value);
                        }
                    }
                }
            }
            else if (item is YamlScalarNode scalar)
            {
                // 스칼라 값인 경우
                row.SetCell(1, "^");
                row.SetCell(2, ConvertValue(scalar));
            }

            return row;
        }

        private void MapNestedArray(
            ExcelRow row,
            YamlSequenceNode array,
            string arrayName,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            var startColumn = scheme.GetArrayStartColumn(arrayName);
            if (startColumn < 0)
            {
                Logger.Warning($"배열 '{arrayName}'의 시작 컬럼을 찾을 수 없음");
                return;
            }

            int currentCol = startColumn;
            
            // 배열의 각 요소 처리
            foreach (var element in array.Children)
            {
                if (element is YamlMappingNode elementMapping)
                {
                    // 객체의 각 속성을 순서대로 매핑
                    var orderer = new DynamicPropertyOrderer();
                    var elementProps = elementMapping.Children.Keys
                        .Select(k => k.ToString())
                        .ToList();

                    // 통합 스키마가 있으면 그 순서를 따름
                    if (pattern.Arrays.ContainsKey(arrayName))
                    {
                        var arrayPattern = pattern.Arrays[arrayName];
                        var orderedProps = orderer.DeterminePropertyOrder(arrayPattern.ElementProperties);

                        foreach (var prop in orderedProps)
                        {
                            if (elementMapping.Children.ContainsKey(new YamlScalarNode(prop)))
                            {
                                var value = ConvertValue(elementMapping.Children[new YamlScalarNode(prop)]);
                                row.SetCell(currentCol++, value);
                            }
                            else
                            {
                                // 속성이 없으면 빈 셀
                                row.SetCell(currentCol++, "");
                            }
                        }
                    }
                    else
                    {
                        // 스키마 정보가 없으면 있는 속성만 순서대로
                        foreach (var prop in elementMapping.Children)
                        {
                            var value = ConvertValue(prop.Value);
                            row.SetCell(currentCol++, value);
                        }
                    }
                }
                else
                {
                    // 단순 값인 경우
                    var value = ConvertValue(element);
                    row.SetCell(currentCol++, value);
                }
            }
        }

        private List<ExcelRow> ExpandVertically(
            YamlNode item,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            var rows = new List<ExcelRow>();

            if (item is YamlMappingNode mapping)
            {
                // 기본 행 생성
                var baseRow = new ExcelRow();
                baseRow.SetCell(1, "^");

                // 단순 속성들 먼저 처리
                var simpleProps = new Dictionary<string, object>();
                var arrayProps = new Dictionary<string, YamlSequenceNode>();

                foreach (var prop in mapping.Children)
                {
                    var key = prop.Key.ToString();
                    if (prop.Value is YamlSequenceNode array)
                    {
                        arrayProps[key] = array;
                    }
                    else
                    {
                        var columnIndex = scheme.GetColumnIndex(key);
                        if (columnIndex > 0)
                        {
                            baseRow.SetCell(columnIndex, ConvertValue(prop.Value));
                        }
                    }
                }

                // 수직 확장이 필요한 배열 찾기
                var verticalArrays = arrayProps.Where(a => 
                    pattern.Arrays.ContainsKey(a.Key) && 
                    pattern.Arrays[a.Key].RequiresMultipleRows).ToList();

                if (verticalArrays.Any())
                {
                    // 가장 긴 배열 기준으로 행 생성
                    int maxRows = verticalArrays.Max(a => a.Value.Children.Count);
                    
                    for (int i = 0; i < maxRows; i++)
                    {
                        var newRow = new ExcelRow();
                        
                        // 기본 행의 데이터 복사
                        foreach (var cell in baseRow.GetAllCells())
                        {
                            newRow.SetCell(cell.Key, cell.Value);
                        }

                        // 각 배열의 i번째 요소 추가
                        foreach (var array in verticalArrays)
                        {
                            if (i < array.Value.Children.Count)
                            {
                                var element = array.Value.Children[i];
                                var columnIndex = scheme.GetColumnIndex(array.Key);
                                if (columnIndex > 0)
                                {
                                    newRow.SetCell(columnIndex, ConvertValue(element));
                                }
                            }
                        }

                        rows.Add(newRow);
                    }
                }
                else
                {
                    // 수직 확장이 필요 없으면 기본 행만 반환
                    rows.Add(baseRow);
                }
            }

            return rows;
        }

        private ExcelRow MapSingleItem(
            YamlMappingNode mapping,
            DynamicSchemaBuilder.ExcelScheme scheme,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            return MapHorizontally(mapping, scheme, pattern);
        }

        private object ConvertValue(YamlNode node)
        {
            if (node is YamlScalarNode scalar)
                return ConvertScalar(scalar);
            else if (node is YamlSequenceNode)
                return "[Array]";
            else if (node is YamlMappingNode)
                return "[Object]";
            else
                return null;
        }

        private object ConvertScalar(YamlScalarNode scalar)
        {
            var value = scalar.Value;

            if (string.IsNullOrEmpty(value))
                return "";

            // 동적 타입 추론
            // bool
            if (bool.TryParse(value, out bool boolResult))
                return boolResult;

            // int
            if (int.TryParse(value, out int intResult))
                return intResult;

            // double (소수점이 있는 경우)
            if (value.Contains('.') && double.TryParse(value, NumberStyles.Any,
                CultureInfo.InvariantCulture, out double doubleResult))
                return doubleResult;

            // DateTime
            if (DateTime.TryParse(value, out DateTime dateResult))
                return dateResult;

            // 기본값은 문자열
            return value;
        }
    }
}