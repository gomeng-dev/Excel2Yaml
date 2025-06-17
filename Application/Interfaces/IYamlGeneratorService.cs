using ClosedXML.Excel;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;

namespace ExcelToYamlAddin.Application.Interfaces
{
    /// <summary>
    /// YAML 생성 서비스 인터페이스
    /// </summary>
    public interface IYamlGeneratorService
    {
        /// <summary>
        /// 스키마와 워크시트로부터 YAML을 생성합니다.
        /// </summary>
        /// <param name="scheme">스키마</param>
        /// <param name="worksheet">워크시트</param>
        /// <param name="options">변환 옵션</param>
        /// <returns>생성된 YAML 문자열</returns>
        string Generate(Scheme scheme, IXLWorksheet worksheet, ConversionOptions options);

        /// <summary>
        /// 루트 노드를 처리하여 객체를 생성합니다.
        /// </summary>
        /// <param name="scheme">스키마</param>
        /// <param name="worksheet">워크시트</param>
        /// <returns>처리된 객체</returns>
        object ProcessRootNode(Scheme scheme, IXLWorksheet worksheet);
    }
}