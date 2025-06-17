using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.FileSystem;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation.NodeProcessors
{
    /// <summary>
    /// 맵(MAP) 노드를 처리하는 프로세서
    /// </summary>
    public class MapNodeProcessor : NodeProcessorBase
    {
        protected override SchemeNodeType SupportedNodeType => SchemeNodeType.Map;

        public override object Process(SchemeNode node, GenerationContext context)
        {
            // 맵 객체 생성
            var map = OrderedYamlFactory.CreateObject();
            
            // 현재 행이 비어있는지 확인
            if (IsRowEmpty(context, node))
            {
                return null;
            }

            return map;
        }

        public override async Task<NodeProcessResult> ProcessAsync(
            SchemeNode node,
            GenerationContext context,
            INodeTraverser traverser,
            CancellationToken cancellationToken = default)
        {
            LogNodeProcessing("맵 노드 처리 시작", node, context);

            try
            {
                var mapData = new MapData { Name = node.Key };
                var properties = new Dictionary<string, object>();

                // 현재 행이 비어있는지 확인
                if (IsRowEmpty(context, node))
                {
                    LogNodeProcessing("빈 행 스킵", node, context);
                    return NodeProcessResult.Skip();
                }

                // 자식 노드들 처리
                if (node.Children != null && node.Children.Any())
                {
                    foreach (var child in node.Children)
                    {
                        cancellationToken.ThrowIfCancellationRequested();

                        var result = await traverser.TraverseAsync(child, context, cancellationToken);
                        
                        if (!result.Success)
                        {
                            return NodeProcessResult.Error($"자식 노드 처리 실패: {result.ErrorMessage}");
                        }

                        if (result.Data != null)
                        {
                            // 자식 노드의 데이터를 맵에 추가
                            AddToMap(properties, result.Data);
                        }
                    }
                }

                // 빈 맵 처리
                if (properties.Count == 0 && !context.Options.ShowEmptyFields)
                {
                    LogNodeProcessing("빈 맵 스킵", node, context);
                    return NodeProcessResult.Skip();
                }

                mapData.Properties = properties;
                LogNodeProcessing($"맵 노드 처리 완료: 속성 수={properties.Count}", node, context);
                
                return NodeProcessResult.Ok(mapData, 1);
            }
            catch (System.Exception ex)
            {
                Logger.Error(ex, $"맵 노드 처리 중 오류: {node.Key}");
                return NodeProcessResult.Error($"맵 노드 처리 실패: {ex.Message}");
            }
        }

        private void AddToMap(Dictionary<string, object> map, object data)
        {
            switch (data)
            {
                case PropertyData property:
                    map[property.Name] = property.Value;
                    break;
                case MapData childMap:
                    if (!string.IsNullOrEmpty(childMap.Name))
                    {
                        map[childMap.Name] = childMap.Properties;
                    }
                    else
                    {
                        // 이름이 없는 맵의 경우 속성들을 직접 병합
                        foreach (var kvp in childMap.Properties)
                        {
                            map[kvp.Key] = kvp.Value;
                        }
                    }
                    break;
                case ArrayData array:
                    map[array.Name] = array.Items;
                    break;
                case KeyValueData kvData:
                    map[kvData.Key] = kvData.Value;
                    break;
            }
        }
    }

    /// <summary>
    /// 맵 데이터
    /// </summary>
    public class MapData
    {
        public string Name { get; set; }
        public Dictionary<string, object> Properties { get; set; } = new Dictionary<string, object>();
    }
}