using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Logging;
using YamlDotNet.RepresentationModel;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    /// <summary>
    /// YAML 데이터를 Excel 셀에 매핑하는 클래스
    /// </summary>
    public class ExcelDataMapper
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<ExcelDataMapper>();

        /// <summary>
        /// YAML 데이터를 Excel 행으로 매핑
        /// </summary>
        public List<ExcelRow> MapToExcelRows(YamlNode rootNode, Dictionary<string, int> columnMappings)
        {
            var rows = new List<ExcelRow>();

            if (rootNode is YamlSequenceNode rootSequence)
            {
                // 루트가 배열인 경우 각 요소를 행으로 변환
                foreach (var element in rootSequence.Children)
                {
                    var row = new ExcelRow();
                    MapNodeToRow(element, row, columnMappings, "");
                    rows.Add(row);
                }
            }
            else if (rootNode is YamlMappingNode rootMapping)
            {
                // 루트가 객체인 경우 단일 행으로 변환
                var row = new ExcelRow();
                MapNodeToRow(rootNode, row, columnMappings, "");
                rows.Add(row);
            }

            Logger.Information($"매핑 완료: {rows.Count}개 행");
            return rows;
        }

        /// <summary>
        /// 노드를 행에 매핑
        /// </summary>
        private void MapNodeToRow(YamlNode node, ExcelRow row, Dictionary<string, int> columnMappings, string path)
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
                        if (columnMappings.ContainsKey(fullPath))
                        {
                            var col = columnMappings[fullPath];
                            row.SetCell(col, scalar.Value);
                        }
                    }
                    else if (value is YamlMappingNode childMapping)
                    {
                        MapNodeToRow(childMapping, row, columnMappings, fullPath);
                    }
                    else if (value is YamlSequenceNode childSequence)
                    {
                        MapArrayToRow(childSequence, row, columnMappings, fullPath);
                    }
                }
            }
        }

        /// <summary>
        /// 배열을 행에 매핑
        /// </summary>
        private void MapArrayToRow(YamlSequenceNode sequence, ExcelRow row, Dictionary<string, int> columnMappings, string path)
        {
            for (int i = 0; i < sequence.Children.Count; i++)
            {
                var element = sequence.Children[i];
                var elementPath = $"{path}[{i}]";

                if (element is YamlMappingNode elementMapping)
                {
                    MapNodeToRow(elementMapping, row, columnMappings, elementPath);
                }
                else if (element is YamlScalarNode scalar)
                {
                    if (columnMappings.ContainsKey(elementPath))
                    {
                        var col = columnMappings[elementPath];
                        row.SetCell(col, scalar.Value);
                    }
                }
            }
        }

        /// <summary>
        /// Excel 행을 나타내는 클래스
        /// </summary>
        public class ExcelRow
        {
            private readonly Dictionary<int, object> _cells = new Dictionary<int, object>();

            public void SetCell(int column, object value)
            {
                _cells[column] = value;
            }

            public object GetCell(int column)
            {
                return _cells.ContainsKey(column) ? _cells[column] : null;
            }

            public void WriteToWorksheet(IXLWorksheet worksheet, int rowNumber)
            {
                foreach (var kvp in _cells)
                {
                    // ClosedXML에서 안전하게 값을 설정
                    var cell = worksheet.Cell(rowNumber, kvp.Key);
                    if (kvp.Value != null)
                    {
                        cell.Value = XLCellValue.FromObject(kvp.Value);
                    }
                }
            }

            public int GetMaxColumn()
            {
                return _cells.Count > 0 ? _cells.Keys.Max() : 0;
            }
        }
    }
}