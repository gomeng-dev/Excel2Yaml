using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelToYamlAddin.Domain.Interfaces
{
    /// <summary>
    /// 리포지토리 패턴의 기본 인터페이스
    /// </summary>
    /// <typeparam name="T">엔티티 타입</typeparam>
    /// <typeparam name="TId">엔티티 식별자 타입</typeparam>
    public interface IRepository<T, TId> where T : class
    {
        /// <summary>
        /// ID로 엔티티를 조회합니다.
        /// </summary>
        /// <param name="id">엔티티 ID</param>
        /// <returns>엔티티 또는 null</returns>
        Task<T> GetByIdAsync(TId id);

        /// <summary>
        /// 모든 엔티티를 조회합니다.
        /// </summary>
        /// <returns>엔티티 목록</returns>
        Task<IEnumerable<T>> GetAllAsync();

        /// <summary>
        /// 새 엔티티를 추가합니다.
        /// </summary>
        /// <param name="entity">추가할 엔티티</param>
        /// <returns>추가된 엔티티</returns>
        Task<T> AddAsync(T entity);

        /// <summary>
        /// 엔티티를 업데이트합니다.
        /// </summary>
        /// <param name="entity">업데이트할 엔티티</param>
        /// <returns>업데이트된 엔티티</returns>
        Task<T> UpdateAsync(T entity);

        /// <summary>
        /// 엔티티를 삭제합니다.
        /// </summary>
        /// <param name="id">삭제할 엔티티 ID</param>
        /// <returns>삭제 성공 여부</returns>
        Task<bool> DeleteAsync(TId id);

        /// <summary>
        /// 엔티티 존재 여부를 확인합니다.
        /// </summary>
        /// <param name="id">확인할 엔티티 ID</param>
        /// <returns>존재 여부</returns>
        Task<bool> ExistsAsync(TId id);
    }
}