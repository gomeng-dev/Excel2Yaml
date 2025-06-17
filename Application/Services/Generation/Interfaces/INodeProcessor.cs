using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.Entities;

namespace ExcelToYamlAddin.Application.Services.Generation.Interfaces
{
    /// <summary>
    /// 특정 타입의 스키마 노드를 처리하는 인터페이스
    /// </summary>
    public interface INodeProcessor
    {
        /// <summary>
        /// 이 프로세서가 처리할 수 있는 노드 타입인지 확인
        /// </summary>
        /// <param name="node">확인할 노드</param>
        /// <returns>처리 가능 여부</returns>
        bool CanProcess(SchemeNode node);

        /// <summary>
        /// 노드를 처리하여 데이터를 생성합니다.
        /// </summary>
        /// <param name="node">처리할 노드</param>
        /// <param name="context">생성 컨텍스트</param>
        /// <returns>처리 결과</returns>
        object Process(SchemeNode node, GenerationContext context);

        /// <summary>
        /// 노드를 처리하여 데이터를 생성합니다. (비동기)
        /// </summary>
        /// <param name="node">처리할 노드</param>
        /// <param name="context">생성 컨텍스트</param>
        /// <param name="traverser">하위 노드 순회를 위한 traverser</param>
        /// <param name="cancellationToken">취소 토큰</param>
        /// <returns>처리 결과</returns>
        Task<NodeProcessResult> ProcessAsync(
            SchemeNode node,
            GenerationContext context,
            INodeTraverser traverser,
            CancellationToken cancellationToken = default);
    }

    /// <summary>
    /// 노드 처리 결과
    /// </summary>
    public class NodeProcessResult
    {
        /// <summary>
        /// 처리 성공 여부
        /// </summary>
        public bool Success { get; private set; }

        /// <summary>
        /// 처리된 데이터
        /// </summary>
        public object Data { get; private set; }

        /// <summary>
        /// 데이터를 건너뛸지 여부
        /// </summary>
        public bool ShouldSkip { get; private set; }

        /// <summary>
        /// 처리 후 행 이동 수
        /// </summary>
        public int RowsConsumed { get; private set; }

        /// <summary>
        /// 오류 메시지
        /// </summary>
        public string ErrorMessage { get; private set; }

        private NodeProcessResult() { }

        /// <summary>
        /// 성공 결과 생성
        /// </summary>
        public static NodeProcessResult Ok(object data, int rowsConsumed = 1)
        {
            return new NodeProcessResult
            {
                Success = true,
                Data = data,
                RowsConsumed = rowsConsumed
            };
        }

        /// <summary>
        /// 스킵 결과 생성
        /// </summary>
        public static NodeProcessResult Skip()
        {
            return new NodeProcessResult
            {
                Success = true,
                ShouldSkip = true,
                RowsConsumed = 1
            };
        }

        /// <summary>
        /// 오류 결과 생성
        /// </summary>
        public static NodeProcessResult Error(string errorMessage)
        {
            return new NodeProcessResult
            {
                Success = false,
                ErrorMessage = errorMessage
            };
        }
    }
}