using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.Entities;

namespace ExcelToYamlAddin.Application.Services.Generation.Interfaces
{
    /// <summary>
    /// 스키마 노드를 순회하며 데이터를 수집하는 인터페이스
    /// </summary>
    public interface INodeTraverser
    {
        /// <summary>
        /// 컨텍스트를 순회합니다.
        /// </summary>
        /// <param name="context">생성 컨텍스트</param>
        void Traverse(GenerationContext context);

        /// <summary>
        /// 스키마 노드를 순회하며 데이터를 수집합니다.
        /// </summary>
        /// <param name="node">순회할 루트 노드</param>
        /// <param name="context">생성 컨텍스트</param>
        /// <param name="cancellationToken">취소 토큰</param>
        /// <returns>순회 결과</returns>
        Task<TraversalResult> TraverseAsync(
            SchemeNode node,
            GenerationContext context,
            CancellationToken cancellationToken = default);
    }

    /// <summary>
    /// 노드 순회 결과
    /// </summary>
    public class TraversalResult
    {
        public bool Success { get; private set; }
        public object Data { get; private set; }
        public string ErrorMessage { get; private set; }

        private TraversalResult() { }

        public static TraversalResult Ok(object data)
        {
            return new TraversalResult
            {
                Success = true,
                Data = data
            };
        }

        public static TraversalResult Error(string errorMessage)
        {
            return new TraversalResult
            {
                Success = false,
                ErrorMessage = errorMessage
            };
        }
    }
}