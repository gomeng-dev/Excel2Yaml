using System;
using ExcelToYamlAddin.Infrastructure.Excel;
using ExcelToYamlAddin.Infrastructure.Excel.Parsing;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Tests.Common;
using ClosedXML.Excel;

namespace ExcelToYamlAddin.Tests.Infrastructure.Excel.Parsing
{
    /// <summary>
    /// 실제 Excel 환경에서 SchemeParser를 테스트하는 통합 테스트
    /// </summary>
    public class SchemeParserIntegrationTests
    {
        private readonly ISimpleLogger _logger = SimpleLoggerFactory.CreateLogger<SchemeParserIntegrationTests>();

        public void TestWithCurrentWorksheet()
        {
            try
            {
                // 현재 활성화된 워크시트 가져오기
                // 테스트용 ClosedXML 워크북 생성
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("테스트시트");
                    
                    // 테스트 스키마 설정
                    worksheet.Cell(1, 1).Value = "# 테스트 시트";
                    worksheet.Cell(2, 1).Value = SchemeConstants.NodeTypes.Map;
                    worksheet.Cell(3, 2).Value = "property1";
                    worksheet.Cell(3, 3).Value = "property2";
                    worksheet.Range(2, 1, 2, 3).Merge();
                    worksheet.Cell(4, 1).Value = SchemeConstants.Markers.SchemeEnd;
                    worksheet.Range(4, 1, 4, 3).Merge();
                    
                    // 테스트 데이터
                    worksheet.Cell(5, 2).Value = "value1";
                    worksheet.Cell(5, 3).Value = "value2";
                    
                    _logger.Information($"테스트 시트 생성 완료");

                    // SchemeParser 생성 및 파싱
                    var parser = SchemeParserFactory.Create(worksheet);
                var result = parser.Parse();

                // 결과 검증
                TestAssert.IsNotNull(result, "파싱 결과가 null입니다.");
                TestAssert.IsNotNull(result.Root, "루트 노드가 null입니다.");
                TestAssert.IsTrue(result.ContentStartRowNum > 0, "컨텐츠 시작 행이 유효하지 않습니다.");
                TestAssert.IsTrue(result.EndRowNum >= result.ContentStartRowNum, "끝 행이 시작 행보다 작습니다.");

                _logger.Information($"파싱 성공: 루트={result.Root.Key}, 타입={result.Root.NodeType}");
                _logger.Information($"데이터 범위: {result.ContentStartRowNum} ~ {result.EndRowNum}");
                _logger.Information($"자식 노드 수: {result.Root.ChildCount}");

                // 선형 노드 목록 테스트
                var linearNodes = result.GetLinearNodes();
                TestAssert.IsTrue(linearNodes.Count > 0, "선형 노드 목록이 비어있습니다.");
                _logger.Information($"전체 노드 수: {linearNodes.Count}");

                    _logger.Information("SchemeParser 통합 테스트 성공!");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "SchemeParser 통합 테스트 실패");
                throw;
            }
        }

        public void TestSchemeEndMarkerDetection()
        {
            try
            {
                _logger.Information("스키마 끝 마커 테스트 시작");

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("마커테스트");
                    
                    // 테스트를 위한 마커 추가
                    worksheet.Cell(5, 1).Value = SchemeConstants.Markers.SchemeEnd;
                    
                    var endMarkerFinder = new SchemeEndMarkerFinder(_logger);
                    var endRow = endMarkerFinder.FindSchemeEndRow(worksheet);
                    
                    TestAssert.AreEqual(5, endRow, "스키마 끝 마커를 찾지 못했습니다.");
                    _logger.Information($"스키마 끝 마커 발견: 행 {endRow}");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "스키마 끝 마커 테스트 실패");
                throw;
            }
        }

        public void TestMergedCellHandling()
        {
            try
            {
                _logger.Information("병합 셀 처리 테스트 시작");

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("병합셀테스트");
                    
                    // 병합 셀 생성
                    worksheet.Range(2, 1, 2, 3).Merge();
                    worksheet.Range(4, 2, 4, 5).Merge();
                    
                    var mergedCellHandler = new MergedCellHandler(_logger);

                    // 각 행에서 병합된 셀 확인
                    var mergedRegions2 = mergedCellHandler.GetMergedRegionsInRow(worksheet, 2);
                    var mergedRegions4 = mergedCellHandler.GetMergedRegionsInRow(worksheet, 4);
                    
                    TestAssert.AreEqual(1, mergedRegions2.Count, "행 2에 병합 영역이 있어야 합니다.");
                    TestAssert.AreEqual(1, mergedRegions4.Count, "행 4에 병합 영역이 있어야 합니다.");
                    
                    _logger.Information("병합 셀 처리 테스트 완료");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "병합 셀 처리 테스트 실패");
                throw;
            }
        }
    }
}