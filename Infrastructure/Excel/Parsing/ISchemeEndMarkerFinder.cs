using ClosedXML.Excel;

namespace ExcelToYamlAddin.Infrastructure.Excel.Parsing
{
    /// <summary>
    /// 스키마 끝 마커를 찾는 서비스의 인터페이스
    /// </summary>
    public interface ISchemeEndMarkerFinder
    {
        /// <summary>
        /// 워크시트에서 스키마 끝 마커를 찾습니다.
        /// </summary>
        /// <param name="worksheet">검색할 워크시트</param>
        /// <returns>스키마 끝 마커가 있는 행 번호. 찾지 못한 경우 -1</returns>
        int FindSchemeEndRow(IXLWorksheet worksheet);

        /// <summary>
        /// 특정 행이 스키마 끝 마커를 포함하는지 확인합니다.
        /// </summary>
        /// <param name="row">확인할 행</param>
        /// <returns>스키마 끝 마커를 포함하면 true</returns>
        bool ContainsEndMarker(IXLRow row);
    }
}