using ExcelToYamlAddin.Domain.Entities;

namespace ExcelToYamlAddin.Domain.Interfaces
{
    /// <summary>
    /// 스키마 검증 서비스 인터페이스
    /// </summary>
    public interface ISchemeValidationService : IDomainService
    {
        /// <summary>
        /// 스키마의 유효성을 검증합니다.
        /// </summary>
        /// <param name="scheme">검증할 스키마</param>
        /// <returns>검증 결과</returns>
        SchemeValidationResult Validate(Scheme scheme);

        /// <summary>
        /// 스키마 노드의 유효성을 검증합니다.
        /// </summary>
        /// <param name="node">검증할 노드</param>
        /// <returns>검증 결과</returns>
        NodeValidationResult ValidateNode(SchemeNode node);

        /// <summary>
        /// 스키마 구조의 무결성을 검증합니다.
        /// </summary>
        /// <param name="scheme">검증할 스키마</param>
        /// <returns>무결성 검증 결과</returns>
        bool ValidateIntegrity(Scheme scheme);
    }
}