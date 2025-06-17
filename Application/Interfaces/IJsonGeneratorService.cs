using ClosedXML.Excel;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;

namespace ExcelToYamlAddin.Application.Interfaces
{
    /// <summary>
    /// JSON 생성 서비스 인터페이스
    /// </summary>
    public interface IJsonGeneratorService
    {
        /// <summary>
        /// 스키마와 워크시트로부터 JSON을 생성합니다.
        /// </summary>
        /// <param name="scheme">스키마</param>
        /// <param name="worksheet">워크시트</param>
        /// <param name="includeEmptyFields">빈 필드 포함 여부</param>
        /// <returns>생성된 JSON 문자열</returns>
        string GenerateJson(Scheme scheme, IXLWorksheet worksheet, bool includeEmptyFields);

        /// <summary>
        /// 루트 노드를 처리하여 JSON 객체를 생성합니다.
        /// </summary>
        /// <param name="scheme">스키마</param>
        /// <param name="worksheet">워크시트</param>
        /// <returns>처리된 JSON 객체</returns>
        object ProcessRootNode(Scheme scheme, IXLWorksheet worksheet);
    }
}