using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation
{
    /// <summary>
    /// 스키마 노드를 순회하며 데이터를 수집하는 클래스
    /// </summary>
    public class NodeTraverser : INodeTraverser
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<NodeTraverser>();
        
        private readonly Dictionary<SchemeNodeType, INodeProcessor> _processors;

        public NodeTraverser(Dictionary<SchemeNodeType, INodeProcessor> processors)
        {
            _processors = processors ?? throw new ArgumentNullException(nameof(processors));
        }

        /// <summary>
        /// 컨텍스트를 순회합니다.
        /// </summary>
        public void Traverse(GenerationContext context)
        {
            if (context == null)
                throw new ArgumentNullException(nameof(context));

            var scheme = context.Scheme;
            if (scheme?.Root == null)
                return;

            // 루트 노드부터 시작하여 재귀적으로 순회
            TraverseNode(scheme.Root, context);
        }

        private void TraverseNode(SchemeNode node, GenerationContext context)
        {
            if (node == null || context.IsMaxDepthExceeded)
                return;

            // 프로세서 찾기
            if (_processors.TryGetValue(node.NodeType, out var processor))
            {
                // 노드 처리
                var result = processor.Process(node, context);
                
                // 처리 결과를 컨텍스트에 저장 (필요한 경우)
                if (result != null)
                {
                    context.AddMetadata($"node_{node.Key}", result);
                }
            }

            // 자식 노드들 순회
            if (node.Children != null)
            {
                foreach (var child in node.Children)
                {
                    TraverseNode(child, context);
                }
            }
        }

        /// <summary>
        /// 스키마 노드를 순회하며 데이터를 수집합니다.
        /// </summary>
        public async Task<TraversalResult> TraverseAsync(
            SchemeNode node,
            GenerationContext context,
            CancellationToken cancellationToken = default)
        {
            if (node == null)
            {
                return TraversalResult.Error("노드가 null입니다.");
            }

            if (context == null)
            {
                return TraversalResult.Error("컨텍스트가 null입니다.");
            }

            try
            {
                // 최대 깊이 확인
                if (context.IsMaxDepthExceeded)
                {
                    Logger.Warning($"최대 깊이 초과: {context.CurrentDepth}");
                    return TraversalResult.Error($"최대 깊이({context.Options.MaxDepth})를 초과했습니다.");
                }

                // 취소 요청 확인
                cancellationToken.ThrowIfCancellationRequested();

                // 적절한 프로세서 찾기
                if (!_processors.TryGetValue(node.NodeType, out var processor))
                {
                    return TraversalResult.Error($"노드 타입 '{node.NodeType}'에 대한 프로세서를 찾을 수 없습니다.");
                }

                Logger.Debug($"노드 처리 시작: {node.Key}, 타입: {node.NodeType}, 행: {context.CurrentRow}");

                // 노드 컨텍스트 추가
                context.PushNodeContext(node);

                try
                {
                    // 노드 처리
                    var result = await processor.ProcessAsync(node, context, this, cancellationToken);
                    
                    if (!result.Success)
                    {
                        return TraversalResult.Error(result.ErrorMessage);
                    }

                    if (result.ShouldSkip)
                    {
                        Logger.Debug($"노드 스킵: {node.Key}");
                        return TraversalResult.Ok(null);
                    }

                    Logger.Debug($"노드 처리 완료: {node.Key}, 소비된 행: {result.RowsConsumed}");
                    return TraversalResult.Ok(result.Data);
                }
                finally
                {
                    // 노드 컨텍스트 제거
                    context.PopNodeContext();
                }
            }
            catch (OperationCanceledException)
            {
                Logger.Warning("노드 순회가 취소되었습니다.");
                throw;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"노드 순회 중 오류 발생: {node.Key}");
                return TraversalResult.Error($"노드 순회 중 오류 발생: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// 노드 타입에 따라 적절한 프로세서를 선택하는 인터페이스
    /// </summary>
    public interface INodeProcessorResolver
    {
        /// <summary>
        /// 노드에 맞는 프로세서를 찾습니다.
        /// </summary>
        INodeProcessor Resolve(SchemeNode node);
    }

    /// <summary>
    /// 노드 프로세서 리졸버 구현
    /// </summary>
    public class NodeProcessorResolver : INodeProcessorResolver
    {
        private readonly IEnumerable<INodeProcessor> _processors;
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger(nameof(NodeProcessorResolver));

        public NodeProcessorResolver(IEnumerable<INodeProcessor> processors)
        {
            _processors = processors ?? throw new ArgumentNullException(nameof(processors));
        }

        public INodeProcessor Resolve(SchemeNode node)
        {
            foreach (var processor in _processors)
            {
                if (processor.CanProcess(node))
                {
                    Logger.Debug($"프로세서 선택: {processor.GetType().Name} -> 노드: {node.Key}");
                    return processor;
                }
            }

            Logger.Warning($"적절한 프로세서를 찾을 수 없음: 노드 타입 = {node.NodeType}");
            return null;
        }
    }
}