using ClosedXML.Excel;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Infrastructure.Logging;
using System;
using System.Collections.Generic;
using ExcelToYamlAddin.Domain.ValueObjects;
namespace ExcelToYamlAddin.Infrastructure.Excel
{
    public class SchemeParser
    {
        // 로깅 방식 변경
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<SchemeParser>();

        private const int ILLEGAL_ROW_NUM = SchemeConstants.RowNumbers.IllegalRow;
        private const int COMMENT_ROW_NUM = SchemeConstants.RowNumbers.CommentRow;
        private const string SCHEME_END = SchemeConstants.Markers.SchemeEnd;

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
                throw new InvalidOperationException(ErrorMessages.Schema.SchemeEndNotFound);
            }

            // ClosedXML에서는 행 번호가 1부터 시작하므로 2번 행이 스키마 시작 행임
            _schemeStartRow = _sheet.Row(SchemeConstants.Sheet.SchemaStartRow);
            if (_schemeStartRow == null)
            {
                Logger.Error($"Scheme start row ({SchemeConstants.Sheet.SchemaStartRow}) not found.");
                throw new InvalidOperationException(ErrorMessages.Schema.SchemeStartRowNotFound);
            }

            // ClosedXML에서는 첫 번째 셀과 마지막 셀을 다르게 찾음
            _firstCellNum = _schemeStartRow.FirstCellUsed()?.Address.ColumnNumber ?? SchemeConstants.Position.FirstColumnIndex;
            _lastCellNum = _schemeStartRow.LastCellUsed()?.Address.ColumnNumber ?? SchemeConstants.Position.FirstColumnIndex;

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

            // 자바 코드와 유사하게 단일 호출로 전체 파싱 처리
            SchemeNode rootNode = Parse(null, _schemeStartRow.RowNumber(), _firstCellNum, _lastCellNum);

            if (rootNode == null)
            {
                Logger.Error("Root node is null. Creating default ARRAY node.");
                try
                {
                    rootNode = SchemeNode.Create(SchemeConstants.NodeTypes.Array, _schemeStartRow.RowNumber(), _firstCellNum);
                    Logger.Debug("기본 ARRAY 형식의 루트 노드 생성");
                }
                catch (Exception ex)
                {
                    Logger.Error("루트 노드 생성 중 오류: " + ex.Message);
                    throw new InvalidOperationException("스키마 파싱 실패", ex);
                }
            }

            var result = new SchemeParsingResult
            {
                Root = rootNode,
                ContentStartRowNum = _schemeEndRowNum + SchemeConstants.Position.DataRowOffset,
                EndRowNum = _sheet.LastRowUsed()?.RowNumber() ?? _schemeEndRowNum + SchemeConstants.Position.DataRowOffset
            };

            // 결과 로깅
            Logger.Information($"Scheme parsing completed: root={rootNode.Key}, type={rootNode.NodeType}, data start row={result.ContentStartRowNum}, end row={result.EndRowNum}");
            Logger.Debug($"Root node child count: {rootNode.ChildCount}");

            return result;
        }

        private SchemeNode Parse(SchemeNode parent, int rowNum, int startCellNum, int endCellNum)
        {
            Logger.Debug($"Parse called: row={rowNum}, start column={startCellNum}, end column={endCellNum}, parent={parent?.Key ?? "null"}, parent type={parent?.NodeType}");

            for (int cellNum = startCellNum; cellNum <= endCellNum; cellNum++)
            {
                IXLRow currentRow = _sheet.Row(rowNum);
                if (currentRow == null) continue;

                IXLCell cell = currentRow.Cell(cellNum);

                if (cell == null || cell.IsEmpty())
                {
                    continue;
                }

                string value = cell.GetString();

                if (string.IsNullOrEmpty(value) || value.Equals(SchemeConstants.Markers.Ignore, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                Logger.Debug($"Cell value processed: row={rowNum}, column={cellNum}, value={value}");
                SchemeNode child = SchemeNode.Create(value, rowNum, cellNum);

                if (parent == null)
                {
                    parent = child;
                    Logger.Debug($"Parent node set: {parent.Key}, type={parent.NodeType}");
                }
                else
                {
                    // 자바 코드와 유사하게 자식 노드 추가 처리
                    parent.AddChild(child);
                    Logger.Debug($"Child node added: parent={parent.Key} ({parent.NodeType}), child={child.Key} ({child.NodeType})");

                    // KEY 타입 노드 처리 - 즉시 다음 셀 처리로 이동
                    if (child.NodeType == SchemeNodeType.Key)
                    {
                        cellNum++;
                        Parse(child, rowNum, cellNum, cellNum);
                        continue;
                    }
                }

                // 병합된 셀 영역 확인
                List<IXLRange> mergedRegionsInRow = GetMergedRegionsInRow(rowNum);

                // 컨테이너 타입 노드(MAP, ARRAY) 처리
                if (child.IsContainer)
                {
                    int firstCellInRange = cellNum;
                    int lastCellInRange = cellNum;

                    // 병합된 셀 영역 확인
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

                    // 컨테이너 노드는 항상 다음 행에서 자식 파싱 수행
                    if (rowNum + SchemeConstants.Position.NextRowOffset < _schemeEndRowNum)
                    {
                        // 자바 코드와 유사하게 단순화된 호출로 변경
                        Parse(child, rowNum + SchemeConstants.Position.NextRowOffset, firstCellInRange, lastCellInRange);
                    }

                    cellNum = lastCellInRange; // 병합된 셀 영역 끝까지 이동
                }
            }

            return parent;
        }

        private bool ContainsEndMarker(IXLRow row)
        {
            if (row == null) return false;

            IXLCell cell = row.Cell(SchemeConstants.Position.FirstColumnIndex);
            return cell != null && !cell.IsEmpty() &&
                   cell.DataType == XLDataType.Text &&
                   cell.GetString().Equals(SCHEME_END, StringComparison.OrdinalIgnoreCase);
        }
    }
}
