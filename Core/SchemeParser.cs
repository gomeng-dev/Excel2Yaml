using Microsoft.Extensions.Logging;
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
        private static readonly ILogger<SchemeParser> Logger = CreateLogger();

        private static ILogger<SchemeParser> CreateLogger()
        {
            // 간단한 로거 팩토리 생성
            var loggerFactory = LoggerFactory.Create(builder => 
            {
                // AddConsole 대신 디버그 로깅만 사용
                builder.SetMinimumLevel(LogLevel.Debug);
            });
            
            return loggerFactory.CreateLogger<SchemeParser>();
        }

        private const int ILLEGAL_ROW_NUM = -1;
        private const int COMMENT_ROW_NUM = 0;
        private const string SCHEME_END = "$scheme_end";

        private readonly IXLWorksheet sheet;
        private readonly IXLRow schemeStartRow;
        private readonly int firstCellNum;
        private readonly int lastCellNum;
        private int schemeEndRowNum;
        
        // parent 필드는 사용하지 않으므로 제거
        // private SchemeNode? parent;

        public SchemeParser(IXLWorksheet sheet)
        {
            this.sheet = sheet;
            Logger.LogInformation("SchemeParser 초기화: 시트명={SheetName}", sheet.Name);

            // ClosedXML에서는 행 인덱스가 1부터 시작함
            schemeEndRowNum = ILLEGAL_ROW_NUM;

            // 스키마 끝 마커 찾기
            foreach (var row in sheet.Rows())
            {
                Logger.LogDebug("행 검사: {RowNum}", row.RowNumber());

                if (row.RowNumber() == (COMMENT_ROW_NUM + 1) || !ContainsEndMarker(row))
                {
                    continue;
                }
                schemeEndRowNum = row.RowNumber();
                Logger.LogInformation("스키마 끝 마커 발견: 행={RowNum}", schemeEndRowNum);
                break;
            }

            if (schemeEndRowNum == ILLEGAL_ROW_NUM)
            {
                Logger.LogError("스키마 끝 마커를 찾을 수 없음");
                throw new InvalidOperationException("scheme end row marker not found.");
            }

            // ClosedXML에서는 행 번호가 1부터 시작하므로 2번 행이 스키마 시작 행임
            schemeStartRow = sheet.Row(2);
            if (schemeStartRow == null)
            {
                Logger.LogError("스키마 시작 행(2)이 없습니다.");
                throw new InvalidOperationException("Schema start row (2) not found.");
            }

            // ClosedXML에서는 첫 번째 셀과 마지막 셀을 다르게 찾음
            firstCellNum = schemeStartRow.FirstCellUsed()?.Address.ColumnNumber ?? 1;
            lastCellNum = schemeStartRow.LastCellUsed()?.Address.ColumnNumber ?? 1;
            
            Logger.LogInformation("스키마 범위: 시작={Start}, 끝={End}", firstCellNum, lastCellNum);
        }

        private List<IXLRange> GetMergedRegionsInRow(int rowNum)
        {
            List<IXLRange> regions = new List<IXLRange>();
            foreach (var range in sheet.MergedRanges)
            {
                if (range.FirstRow().RowNumber() <= rowNum && range.LastRow().RowNumber() >= rowNum)
                {
                    regions.Add(range);
                    Logger.LogDebug("병합 영역 발견: {Region}, 행={Row}", range.RangeAddress.ToString(), rowNum);
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
            Logger.LogInformation("스키마 파싱 시작");
            
            // 처음부터 끝까지 모든 셀을 검사하여 스키마 구조 파악
            SchemeNode rootNode = null;
            
            // 스키마 시작 행부터 스키마 끝 행까지 검사
            for (int rowNum = schemeStartRow.RowNumber(); rowNum < schemeEndRowNum; rowNum++)
            {
                IXLRow row = sheet.Row(rowNum);
                if (row == null || !row.CellsUsed().Any())
                {
                    continue;
                }
                
                // 첫 번째 유효한 행에서 루트 노드 생성 시도
                if (rootNode == null)
                {
                    // 첫 번째 셀 검사
                    for (int cellNum = firstCellNum; cellNum <= lastCellNum; cellNum++)
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
                        if (value.Contains("$"))
                        {
                            Logger.LogInformation("루트 노드 셀 발견: 값={Value}, 행={Row}, 열={Col}", value, rowNum, cellNum);
                            rootNode = new SchemeNode(sheet, rowNum, cellNum, value);
                            break;
                        }
                        else
                        {
                            // 기본적으로 MAP 타입으로 설정
                            Logger.LogInformation("기본 MAP 타입 루트 노드 생성: 값={Value}", value);
                            rootNode = new SchemeNode(sheet, rowNum, cellNum, value + "${}");
                            break;
                        }
                    }
                    
                    if (rootNode != null)
                    {
                        // 루트 노드를 찾았으므로 하위 구조 파싱
                        Parse(rootNode, rowNum, firstCellNum, lastCellNum);
                        break;
                    }
                }
            }
            
            // 루트 노드가 여전히 null이면 기본 ARRAY 노드 생성
            if (rootNode == null)
            {
                Logger.LogError("루트 노드가 null입니다. 기본 ARRAY 노드를 생성합니다.");
                rootNode = new SchemeNode(sheet, schemeStartRow.RowNumber(), firstCellNum, "$[]");
            }
            
            var result = new SchemeParsingResult
            {
                Root = rootNode,
                ContentStartRowNum = schemeEndRowNum + 1,
                EndRowNum = sheet.LastRowUsed()?.RowNumber() ?? schemeEndRowNum + 1
            };

            // 결과 로깅
            Logger.LogInformation("스키마 파싱 완료: 루트={Root}, 데이터 시작행={Start}, 끝행={End}", 
                rootNode.Key, result.ContentStartRowNum, result.EndRowNum);

            return result;
        }

        private SchemeNode Parse(SchemeNode parent, int rowNum, int startCellNum, int endCellNum)
        {
            Logger.LogDebug("Parse 호출: 행={Row}, 시작열={StartCell}, 끝열={EndCell}, 부모={Parent}", 
                rowNum, startCellNum, endCellNum, parent != null ? parent.Key : "null");
            
            for (int cellNum = startCellNum; cellNum <= endCellNum; cellNum++)
            {
                IXLCell cell = sheet.Cell(rowNum, cellNum);
                
                if (cell == null || cell.IsEmpty())
                {
                    continue;
                }
                
                string value = cell.GetString();
                
                if (string.IsNullOrEmpty(value) || value.Equals("^", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                
                Logger.LogDebug("셀 값 처리: 행={Row}, 열={Col}, 값={Value}", rowNum, cellNum, value);
                SchemeNode child = new SchemeNode(sheet, rowNum, cellNum, value);
                
                if (parent == null)
                {
                    parent = child;
                    Logger.LogDebug("부모 노드 설정: {Node}", parent.Key);
                }
                else
                {
                    parent.AddChild(child);
                    Logger.LogDebug("자식 노드 추가: 부모={Parent}, 자식={Child}", parent.Key, child.Key);
                    
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
                            Logger.LogDebug("병합 영역: {Region}, 첫열={First}, 마지막열={Last}", 
                                region.RangeAddress.ToString(), firstCellInRange, lastCellInRange);
                            break;
                        }
                    }
                    
                    // 다음 행에서 자식 노드 파싱
                    if (rowNum + 1 < schemeEndRowNum)
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
