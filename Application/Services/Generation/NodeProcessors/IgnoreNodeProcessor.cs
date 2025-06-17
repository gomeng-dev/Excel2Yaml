using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation.NodeProcessors
{
    /// <summary>
    /// 무시(IGNORE) 노드를 처리하는 프로세서
    /// </summary>
    public class IgnoreNodeProcessor : NodeProcessorBase
    {
        protected override SchemeNodeType SupportedNodeType => SchemeNodeType.Ignore;

        public override object Process(SchemeNode node, GenerationContext context)
        {
            // IGNORE 노드는 항상 null 반환
            return null;
        }

        public override async Task<NodeProcessResult> ProcessAsync(
            SchemeNode node,
            GenerationContext context,
            INodeTraverser traverser,
            CancellationToken cancellationToken = default)
        {
            LogNodeProcessing("IGNORE 노드 스킵", node, context);
            return await Task.FromResult(NodeProcessResult.Skip());
        }
    }
}