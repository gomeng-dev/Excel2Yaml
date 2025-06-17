using ClosedXML.Excel;
using System.Collections.Generic;

namespace ExcelToYamlAddin.Infrastructure.Excel.Parsing
{
    /// <summary>
    /// 병합된 셀을 처리하는 서비스의 인터페이스
    /// </summary>
    public interface IMergedCellHandler
    {
        /// <summary>
        /// 특정 행의 병합된 영역들을 가져옵니다.
        /// </summary>
        /// <param name="worksheet">워크시트</param>
        /// <param name="rowNumber">행 번호</param>
        /// <returns>병합된 영역 목록</returns>
        List<IXLRange> GetMergedRegionsInRow(IXLWorksheet worksheet, int rowNumber);

        /// <summary>
        /// 셀이 포함된 병합 영역의 범위를 가져옵니다.
        /// </summary>
        /// <param name="cell">확인할 셀</param>
        /// <param name="mergedRegions">병합된 영역 목록</param>
        /// <returns>시작 열 번호와 끝 열 번호를 포함한 튜플</returns>
        (int startColumn, int endColumn) GetMergedCellRange(IXLCell cell, List<IXLRange> mergedRegions);
    }
}