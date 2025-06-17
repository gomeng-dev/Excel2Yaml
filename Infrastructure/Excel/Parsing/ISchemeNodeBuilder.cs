using ExcelToYamlAddin.Domain.Entities;
using ClosedXML.Excel;

namespace ExcelToYamlAddin.Infrastructure.Excel.Parsing
{
    /// <summary>
    /// 스키마 노드를 생성하는 빌더의 인터페이스
    /// </summary>
    public interface ISchemeNodeBuilder
    {
        /// <summary>
        /// 셀 값으로부터 스키마 노드를 생성합니다.
        /// </summary>
        /// <param name="cell">Excel 셀</param>
        /// <returns>생성된 스키마 노드. 생성할 수 없는 경우 null</returns>
        SchemeNode BuildFromCell(IXLCell cell);

        /// <summary>
        /// 특정 값과 위치로부터 스키마 노드를 생성합니다.
        /// </summary>
        /// <param name="value">노드 값</param>
        /// <param name="row">행 번호</param>
        /// <param name="column">열 번호</param>
        /// <returns>생성된 스키마 노드</returns>
        SchemeNode Build(string value, int row, int column);

        /// <summary>
        /// 값이 무시해야 할 값인지 확인합니다.
        /// </summary>
        /// <param name="value">확인할 값</param>
        /// <returns>무시해야 하면 true</returns>
        bool ShouldIgnore(string value);
    }
}