using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.Entities;

namespace ExcelToYamlAddin.Domain.Interfaces
{
    /// <summary>
    /// 스키마 리포지토리 인터페이스
    /// </summary>
    public interface ISchemeRepository : IRepository<Scheme, string>
    {
        /// <summary>
        /// 시트 이름으로 스키마를 조회합니다.
        /// </summary>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>스키마 또는 null</returns>
        Task<Scheme> GetBySheetNameAsync(string sheetName);

        /// <summary>
        /// 유효한 스키마만 조회합니다.
        /// </summary>
        /// <returns>유효한 스키마 목록</returns>
        Task<IEnumerable<Scheme>> GetValidSchemesAsync();

        /// <summary>
        /// 특정 노드 타입을 포함하는 스키마를 조회합니다.
        /// </summary>
        /// <param name="nodeType">노드 타입</param>
        /// <returns>스키마 목록</returns>
        Task<IEnumerable<Scheme>> GetByNodeTypeAsync(string nodeType);
    }
}