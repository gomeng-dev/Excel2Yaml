using ClosedXML.Excel;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Infrastructure.Excel.Parsing;
using System;
using System.Collections.Generic;

namespace ExcelToYamlAddin.Infrastructure.Excel
{
    /// <summary>
    /// 스키마 파서
    /// </summary>
    public class SchemeParser
    {
        private readonly ISimpleLogger _logger;
        private readonly ISchemeEndMarkerFinder _endMarkerFinder;
        private readonly IMergedCellHandler _mergedCellHandler;
        private readonly ISchemeNodeBuilder _nodeBuilder;
        private readonly IXLWorksheet _worksheet;

        public SchemeParser(IXLWorksheet worksheet)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            
            // 의존성 생성 - 팩토리를 통해 생성할 수도 있음
            _logger = SimpleLoggerFactory.CreateLogger<SchemeParser>();
            _endMarkerFinder = new SchemeEndMarkerFinder(_logger);
            _mergedCellHandler = new MergedCellHandler(_logger);
            _nodeBuilder = new SchemeNodeBuilder(_logger);
            
            _logger.Information($"SchemeParser 초기화됨: 시트명={worksheet.Name}");
        }

        /// <summary>
        /// 의존성 주입을 위한 생성자
        /// </summary>
        internal SchemeParser(
            IXLWorksheet worksheet,
            ISimpleLogger logger,
            ISchemeEndMarkerFinder endMarkerFinder,
            IMergedCellHandler mergedCellHandler,
            ISchemeNodeBuilder nodeBuilder)
        {
            _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
            _endMarkerFinder = endMarkerFinder ?? throw new ArgumentNullException(nameof(endMarkerFinder));
            _mergedCellHandler = mergedCellHandler ?? throw new ArgumentNullException(nameof(mergedCellHandler));
            _nodeBuilder = nodeBuilder ?? throw new ArgumentNullException(nameof(nodeBuilder));
        }

        /// <summary>
        /// 스키마 파싱 결과
        /// </summary>
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

        /// <summary>
        /// 워크시트에서 스키마를 파싱합니다.
        /// </summary>
        public SchemeParsingResult Parse()
        {
            _logger.Information($"스키마 파싱 시작: 시트={_worksheet.Name}");

            try
            {
                // 1. 파싱 컨텍스트 생성
                var context = CreateParsingContext();
                _logger.Debug($"파싱 컨텍스트 생성 완료: {context}");

                // 2. 스키마 노드 트리 구축
                var rootNode = BuildSchemeTree(context);
                
                // 3. 결과 생성
                var result = CreateParsingResult(rootNode, context);
                
                _logger.Information($"스키마 파싱 완료: 루트={rootNode.Key}, 타입={rootNode.NodeType}, " +
                                  $"데이터 시작행={result.ContentStartRowNum}, 끝행={result.EndRowNum}");
                
                return result;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "스키마 파싱 중 오류 발생");
                throw;
            }
        }

        private ParsingContext CreateParsingContext()
        {
            // 스키마 끝 마커 찾기
            var schemeEndRow = _endMarkerFinder.FindSchemeEndRow(_worksheet);
            if (schemeEndRow == SchemeConstants.RowNumbers.IllegalRow)
            {
                throw new InvalidOperationException(ErrorMessages.Schema.SchemeEndNotFound);
            }

            // 스키마 시작 행 가져오기
            var schemeStartRow = _worksheet.Row(SchemeConstants.Sheet.SchemaStartRow);
            if (schemeStartRow == null)
            {
                throw new InvalidOperationException(ErrorMessages.Schema.SchemeStartRowNotFound);
            }

            // 셀 범위 결정
            var firstCell = schemeStartRow.FirstCellUsed()?.Address.ColumnNumber ?? SchemeConstants.Position.FirstColumnIndex;
            var lastCell = schemeStartRow.LastCellUsed()?.Address.ColumnNumber ?? SchemeConstants.Position.FirstColumnIndex;

            return new ParsingContext(
                _worksheet,
                schemeStartRow,
                schemeEndRow,
                schemeEndRow + SchemeConstants.Position.DataRowOffset,
                firstCell,
                lastCell
            );
        }

        private SchemeNode BuildSchemeTree(ParsingContext context)
        {
            var parser = new SchemeTreeParser(context, _nodeBuilder, _mergedCellHandler, _logger);
            var rootNode = parser.ParseTree();

            if (rootNode == null)
            {
                _logger.Warning("루트 노드가 null입니다. 기본 ARRAY 노드를 생성합니다.");
                rootNode = SchemeNode.Create(
                    SchemeConstants.NodeTypes.Array,
                    context.SchemeStartRow.RowNumber(),
                    context.FirstCellNumber
                );
            }

            return rootNode;
        }

        private SchemeParsingResult CreateParsingResult(SchemeNode rootNode, ParsingContext context)
        {
            return new SchemeParsingResult
            {
                Root = rootNode,
                ContentStartRowNum = context.DataStartRowNumber,
                EndRowNum = context.GetDataEndRowNumber()
            };
        }

        /// <summary>
        /// 스키마 트리를 파싱하는 내부 클래스
        /// </summary>
        private class SchemeTreeParser
        {
            private readonly ParsingContext _context;
            private readonly ISchemeNodeBuilder _nodeBuilder;
            private readonly IMergedCellHandler _mergedCellHandler;
            private readonly ISimpleLogger _logger;

            public SchemeTreeParser(
                ParsingContext context,
                ISchemeNodeBuilder nodeBuilder,
                IMergedCellHandler mergedCellHandler,
                ISimpleLogger logger)
            {
                _context = context;
                _nodeBuilder = nodeBuilder;
                _mergedCellHandler = mergedCellHandler;
                _logger = logger;
            }

            public SchemeNode ParseTree()
            {
                return ParseRow(
                    null,
                    _context.SchemeStartRow.RowNumber(),
                    _context.FirstCellNumber,
                    _context.LastCellNumber
                );
            }

            private SchemeNode ParseRow(SchemeNode parent, int rowNumber, int startColumn, int endColumn)
            {
                _logger.Debug($"행 파싱: 행={rowNumber}, 열 범위={startColumn}-{endColumn}, 부모={parent?.Key ?? "없음"}");

                var currentRow = _context.Worksheet.Row(rowNumber);
                if (currentRow == null)
                    return parent;

                for (int columnNumber = startColumn; columnNumber <= endColumn; columnNumber++)
                {
                    var cell = currentRow.Cell(columnNumber);
                    var node = _nodeBuilder.BuildFromCell(cell);
                    
                    if (node == null)
                        continue;

                    _logger.Debug($"노드 생성됨: 키={node.Key}, 타입={node.NodeType}, 위치=({rowNumber},{columnNumber})");

                    if (parent == null)
                    {
                        parent = node;
                    }
                    else
                    {
                        parent.AddChild(node);
                        
                        // KEY 타입 노드는 다음 셀을 즉시 처리
                        if (node.NodeType == SchemeNodeType.Key)
                        {
                            columnNumber++;
                            ParseRow(node, rowNumber, columnNumber, columnNumber);
                            continue;
                        }
                    }

                    // 컨테이너 노드 처리
                    if (node.IsContainer)
                    {
                        columnNumber = ProcessContainerNode(node, rowNumber, columnNumber);
                    }
                }

                return parent;
            }

            private int ProcessContainerNode(SchemeNode containerNode, int rowNumber, int columnNumber)
            {
                var cell = _context.Worksheet.Row(rowNumber).Cell(columnNumber);
                var mergedRegions = _mergedCellHandler.GetMergedRegionsInRow(_context.Worksheet, rowNumber);
                var (startColumn, endColumn) = _mergedCellHandler.GetMergedCellRange(cell, mergedRegions);

                _logger.Debug($"컨테이너 노드 처리: 타입={containerNode.NodeType}, 병합 범위={startColumn}-{endColumn}");

                // 다음 행에서 자식 노드 파싱
                var nextRow = rowNumber + SchemeConstants.Position.NextRowOffset;
                if (nextRow < _context.SchemeEndRowNumber)
                {
                    ParseRow(containerNode, nextRow, startColumn, endColumn);
                }

                return endColumn; // 병합된 영역의 끝까지 이동
            }
        }
    }
}
