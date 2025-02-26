using ExcelToJsonAddin.Logging;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace ExcelToJsonAddin.Core
{
    public class SchemeParser
    {
        // 로깅 방식 변경
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<SchemeParser>();

        private const int ILLEGAL_ROW_NUM = -1;
        private const int COMMENT_ROW_NUM = 0;
        private const string SCHEME_END = "$scheme_end";

        private readonly IXLWorksheet _sheet;
        private readonly IXLRow _schemeStartRow;
        private readonly int _firstCellNum;
        private readonly int _lastCellNum;
        private int _schemeEndRowNum;

        public SchemeParser(IXLWorksheet sheet)
        {
            _sheet = sheet;
            Logger.Information($"SchemeParser initialized: sheet name={sheet.Name}");

            // ClosedXML에서는 행 인덱스가 1부터 시작함
            _schemeEndRowNum = ILLEGAL_ROW_NUM;

            // 스키마 끝 마커 찾기
            foreach (var row in _sheet.Rows())
            {
                Logger.Debug($"Row inspection: {row.RowNumber()}");

                if (row.RowNumber() == (COMMENT_ROW_NUM + 1) || !ContainsEndMarker(row))
                {
                    continue;
                }
                _schemeEndRowNum = row.RowNumber();
                Logger.Information($"Scheme end marker found: row={_schemeEndRowNum}");
                break;
            }

            if (_schemeEndRowNum == ILLEGAL_ROW_NUM)
            {
                Logger.Error("Scheme end marker not found.");
                throw new InvalidOperationException("Scheme end marker not found.");
            }

            // ClosedXML에서는 행 번호가 1부터 시작하므로 2번 행이 스키마 시작 행임
            _schemeStartRow = _sheet.Row(2);
            if (_schemeStartRow == null)
            {
                Logger.Error("Scheme start row (2) not found.");
                throw new InvalidOperationException("Scheme start row (2) not found.");
            }

            // ClosedXML에서는 첫 번째 셀과 마지막 셀을 다르게 찾음
            _firstCellNum = _schemeStartRow.FirstCellUsed()?.Address.ColumnNumber ?? 1;
            _lastCellNum = _schemeStartRow.LastCellUsed()?.Address.ColumnNumber ?? 1;

            Logger.Information($"Scheme range: start={_firstCellNum}, end={_lastCellNum}");
        }

        private List<IXLRange> GetMergedRegionsInRow(int rowNum)
        {
            List<IXLRange> regions = new List<IXLRange>();
            foreach (var range in _sheet.MergedRanges)
            {
                if (range.FirstRow().RowNumber() <= rowNum && range.LastRow().RowNumber() >= rowNum)
                {
                    regions.Add(range);
                    Logger.Debug($"Merged region found: {range.RangeAddress.ToString()}, row={rowNum}");
                }
            }
            return regions;
        }

        // scheme parsing result
        public class SchemeParsingResult
        {
            public SchemeNode Root { get; set; }
            public int ContentStartRowNum { get; set; }
            public int EndRowNum { get; set; }
            public List<SchemeNode> LinearNodes { get; private set; } = new List<SchemeNode>();

            public List<SchemeNode> GetLinearNodes()
            {
                if (LinearNodes.Count == 0 && Root != null)
                {
                    CollectLinearNodes(Root);
                }
                return LinearNodes;
            }

            private void CollectLinearNodes(SchemeNode node)
            {
                LinearNodes.Add(node);
                foreach (var child in node.Children)
                {
                    CollectLinearNodes(child);
                }
            }
        }

        public SchemeParsingResult Parse()
        {
            Logger.Information("Scheme parsing started");

            // 처음부터 끝까지 모든 셀을 검사하여 스키마 구조 파악
            SchemeNode rootNode = null;

            // 스키마 시작 행부터 스키마 끝 행까지 검사
            for (int rowNum = _schemeStartRow.RowNumber(); rowNum < _schemeEndRowNum; rowNum++)
            {
                IXLRow row = _sheet.Row(rowNum);
                if (row == null || !row.CellsUsed().Any())
                {
                    continue;
                }

                // 첫 번째 유효한 행에서 루트 노드 생성 시도
                if (rootNode == null)
                {
                    // 첫 번째 셀 검사
                    for (int cellNum = _firstCellNum; cellNum <= _lastCellNum; cellNum++)
                    {
                        IXLCell cell = row.Cell(cellNum);
                        if (cell == null || cell.IsEmpty())
                        {
                            continue;
                        }

                        string value = cell.GetString();
                        if (string.IsNullOrEmpty(value))
                        {
                            continue;
                        }

                        // 루트 노드 생성
                        if (value.StartsWith("$"))
                        {
                            // 기본적으로 MAP 타입으로 설정
                            Logger.Information($"Default MAP type root node created: value={value}");
                            try
                            {
                                rootNode = new SchemeNode(_sheet, rowNum, cellNum, value.Contains("{}") ? value : value + "${}");
                            }
                            catch (Exception ex)
                            {
                                Logger.Error($"루트 노드 생성 중 오류: {ex.Message}");
                                // 안전한 기본값으로 설정
                                rootNode = new SchemeNode(_sheet, rowNum, cellNum, "${}");
                            }
                            break;
                        }
                        else
                        {
                            // 기본적으로 MAP 타입으로 설정
                            Logger.Information($"Default MAP type root node created: value={value}");
                            rootNode = new SchemeNode(_sheet, rowNum, cellNum, value + "${}");
                            break;
                        }
                    }

                    if (rootNode != null)
                    {
                        // 루트 노드를 찾았으므로 하위 구조 파싱
                        Parse(rootNode, rowNum, _firstCellNum, _lastCellNum);
                        break;
                    }
                }
            }

            // 루트 노드가 여전히 null이면 기본 ARRAY 노드 생성
            if (rootNode == null)
            {
                Logger.Error("Root node is null. Creating default ARRAY node.");
                try {
                    Logger.Debug("ARRAY 형식의 루트 노드 생성 시도");
                    rootNode = new SchemeNode(_sheet, _schemeStartRow.RowNumber(), _firstCellNum, "$[]");
                    Logger.Debug("ARRAY 형식의 루트 노드 생성 성공");
                    // 추가로 루트 노드의 자식 노드를 파싱
                    Parse(rootNode, _schemeStartRow.RowNumber(), _firstCellNum, _lastCellNum);
                }
                catch (Exception ex) {
                    Logger.Error("루트 노드 생성 중 오류: " + ex.Message);
                    Logger.Error("예외 스택 트레이스: " + ex.StackTrace);
                    // 안전한 기본값으로 설정
                    try {
                        Logger.Debug("기본 MAP 노드 생성 시도 (예외 복구)");
                        rootNode = new SchemeNode(_sheet, _schemeStartRow.RowNumber(), _firstCellNum, "${}");
                        Logger.Debug("기본 MAP 노드 생성 성공 (예외 복구)");
                    }
                    catch (Exception fallbackEx) {
                        Logger.Error("기본 노드 생성 중 추가 오류: " + fallbackEx.Message);
                        throw new Exception("스키마 파싱 중 복구할 수 없는 오류: " + ex.Message + ", 추가 오류: " + fallbackEx.Message, ex);
                    }
                }
            }

            var result = new SchemeParsingResult
            {
                Root = rootNode,
                ContentStartRowNum = _schemeEndRowNum + 1,
                EndRowNum = _sheet.LastRowUsed()?.RowNumber() ?? _schemeEndRowNum + 1
            };

            // 결과 로깅
            Logger.Information($"Scheme parsing completed: root={rootNode.Key}, data start row={result.ContentStartRowNum}, end row={result.EndRowNum}");

            return result;
        }

        private SchemeNode Parse(SchemeNode parent, int rowNum, int startCellNum, int endCellNum)
        {
            Logger.Debug($"Parse called: row={rowNum}, start column={startCellNum}, end column={endCellNum}, parent={parent?.Key ?? "null"}");

            for (int cellNum = startCellNum; cellNum <= endCellNum; cellNum++)
            {
                IXLCell cell = _sheet.Cell(rowNum, cellNum);

                if (cell == null || cell.IsEmpty())
                {
                    continue;
                }

                string value = cell.GetString();

                if (string.IsNullOrEmpty(value) || value.Equals("^", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Logger.Debug($"Cell value processed: row={rowNum}, column={cellNum}, value={value}");
                SchemeNode child = new SchemeNode(_sheet, rowNum, cellNum, value);

                if (parent == null)
                {
                    parent = child;
                    Logger.Debug($"Parent node set: {parent.Key}");
                }
                else
                {
                    parent.AddChild(child);
                    Logger.Debug($"Child node added: parent={parent.Key}, child={child.Key}");

                    if (child.NodeType == SchemeNode.SchemeNodeType.KEY)
                    {
                        cellNum++;
                        Parse(child, rowNum, cellNum, cellNum);
                        continue;
                    }
                }

                List<IXLRange> mergedRegionsInRow = GetMergedRegionsInRow(rowNum);

                if (child.IsContainer)
                {
                    int firstCellInRange = cellNum;
                    int lastCellInRange = cellNum;

                    foreach (var region in mergedRegionsInRow)
                    {
                        if (region.Contains(cell))
                        {
                            firstCellInRange = region.FirstColumn().ColumnNumber();
                            lastCellInRange = region.LastColumn().ColumnNumber();
                            Logger.Debug($"Merged region: {region.RangeAddress.ToString()}, first column={firstCellInRange}, last column={lastCellInRange}");
                            break;
                        }
                    }

                    // 다음 행에서 자식 노드 파싱
                    if (rowNum + 1 < _schemeEndRowNum)
                    {
                        Parse(child, rowNum + 1, firstCellInRange, lastCellInRange);
                    }

                    cellNum = lastCellInRange;
                }
            }

            return parent;
        }

        private bool ContainsEndMarker(IXLRow row)
        {
            if (row == null) return false;

            IXLCell cell = row.Cell(1);
            return cell != null && !cell.IsEmpty() &&
                   cell.DataType == XLDataType.Text &&
                   cell.GetString().Equals(SCHEME_END, StringComparison.OrdinalIgnoreCase);
        }
    }
}
