using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Application.Interfaces;
using ExcelToYamlAddin.Application.Services.Generation;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;
using ExcelToYamlAddin.Application.Services.Generation.NodeProcessors;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.Logging;

namespace ExcelToYamlAddin.Application.Services
{
    /// <summary>
    /// 기존 YamlGenerator와 새로운 YamlGenerationService 간의 어댑터
    /// 점진적 마이그레이션을 위한 임시 클래스
    /// </summary>
    public class YamlGeneratorAdapter : IYamlGeneratorService
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<YamlGeneratorAdapter>();
        
        private readonly YamlGenerationService _newGenerator;
        private readonly bool _useNewGenerator;

        public YamlGeneratorAdapter(bool useNewGenerator = false)
        {
            _useNewGenerator = useNewGenerator;
            
            if (_useNewGenerator)
            {
                // 새로운 생성기 초기화
                var processors = CreateNodeProcessors();
                var traverser = new NodeTraverser(processors);
                var yamlBuilder = YamlBuilderFactory.Create();
                var optionsValidator = new GenerationOptionsValidator();
                _newGenerator = new YamlGenerationService(traverser, yamlBuilder, optionsValidator);
                Logger.Information("새로운 YAML 생성기를 사용합니다.");
            }
            else
            {
                Logger.Information("기존 YAML 생성기를 사용합니다.");
            }
        }

        /// <summary>
        /// YAML 생성
        /// </summary>
        public string Generate(Scheme scheme, IXLWorksheet worksheet, ConversionOptions options)
        {
            if (scheme == null)
                throw new ArgumentNullException(nameof(scheme));
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (options == null)
                throw new ArgumentNullException(nameof(options));

            try
            {
                if (_useNewGenerator)
                {
                    Logger.Debug("새로운 생성기로 YAML 생성 시작");
                    
                    // ConversionOptions를 YamlGenerationOptions로 변환
                    var yamlOptions = ConvertToYamlGenerationOptions(options);
                    
                    // 새로운 생성기 사용
                    return _newGenerator.Generate(scheme, worksheet, yamlOptions);
                }
                else
                {
                    Logger.Debug("기존 생성기로 YAML 생성 시작");
                    
                    // 기존 생성기 사용 (인스턴스 생성 후 호출)
                    var generator = new YamlGenerator(scheme, worksheet, options.IncludeEmptyFields);
                    return generator.Generate(
                        options.YamlStyle?.Style ?? YamlStyle.Block,
                        options.YamlStyle?.IndentSize ?? 2,
                        options.YamlStyle?.PreserveQuotes ?? false,
                        options.IncludeEmptyFields
                    );
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"YAML 생성 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// 특정 시트에 대한 YAML 생성 (옵션 오버라이드 가능)
        /// </summary>
        public string Generate(
            Scheme scheme, 
            IXLWorksheet worksheet, 
            ConversionOptions baseOptions,
            bool showEmptyFields = false,
            YamlStyleOptions yamlStyle = null)
        {
            // 옵션 오버라이드
            var builder = baseOptions.ToBuilder();
            
            if (showEmptyFields != baseOptions.IncludeEmptyFields)
                builder.WithIncludeEmptyFields(showEmptyFields);
            
            if (yamlStyle != null)
                builder.WithYamlStyle(yamlStyle);
            
            var options = builder.Build();

            return Generate(scheme, worksheet, options);
        }

        /// <summary>
        /// ConversionOptions를 YamlGenerationOptions로 변환
        /// </summary>
        private YamlGenerationOptions ConvertToYamlGenerationOptions(ConversionOptions options)
        {
            return new YamlGenerationOptions(
                showEmptyFields: options.IncludeEmptyFields,
                maxDepth: options.Validation?.MaxDepth ?? 100,
                indentSize: options.YamlStyle?.IndentSize ?? 2,
                style: options.YamlStyle?.Style ?? YamlStyle.Block,
                mergeKeyPaths: options.PostProcessing?.MergeKeyPaths,
                flowStylePaths: options.PostProcessing?.FlowStylePaths?.Keys.ToList(),
                outputPath: null, // 출력 경로는 별도로 관리
                postProcessing: Domain.ValueObjects.PostProcessingOptions.Create(
                    true, // enablePostProcessing
                    options.PostProcessing?.EnableMergeByKey ?? false,
                    options.PostProcessing?.ApplyFlowStyle ?? false,
                    options.PostProcessing?.MergeKeyPaths ?? new List<string>(),
                    options.PostProcessing?.FlowStylePaths ?? new Dictionary<string, string>()
                )
            );
        }

        /// <summary>
        /// 새로운 생성기 사용 여부 설정
        /// </summary>
        public void SetUseNewGenerator(bool useNew)
        {
            if (_useNewGenerator != useNew)
            {
                Logger.Information($"YAML 생성기 변경: {(_useNewGenerator ? "새로운" : "기존")} -> {(useNew ? "새로운" : "기존")}");
            }
        }

        /// <summary>
        /// 현재 사용 중인 생성기 정보
        /// </summary>
        public string GetGeneratorInfo()
        {
            return _useNewGenerator ? "새로운 컴포넌트 기반 생성기" : "기존 스택 기반 생성기";
        }

        /// <summary>
        /// 루트 노드를 처리하여 객체를 생성합니다.
        /// </summary>
        public object ProcessRootNode(Scheme scheme, IXLWorksheet worksheet)
        {
            if (scheme == null)
                throw new ArgumentNullException(nameof(scheme));
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));

            try
            {
                if (_useNewGenerator)
                {
                    // 새로운 생성기를 사용하여 루트 객체 생성
                    var options = YamlGenerationOptions.Default;
                    var context = new GenerationContext(worksheet, scheme, options);
                    var processors = CreateNodeProcessors();
                    var traverser = new NodeTraverser(processors);
                    
                    var result = traverser.TraverseAsync(scheme.Root, context).GetAwaiter().GetResult();
                    
                    if (!result.Success)
                    {
                        throw new InvalidOperationException($"루트 노드 처리 실패: {result.ErrorMessage}");
                    }
                    
                    return result.Data;
                }
                else
                {
                    // 기존 생성기 사용하여 루트 노드 처리
                    var generator = new YamlGenerator(scheme, worksheet, true);
                    return generator.ProcessRootNode();
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"루트 노드 처리 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// 노드 프로세서들을 생성합니다.
        /// </summary>
        private static Dictionary<SchemeNodeType, INodeProcessor> CreateNodeProcessors()
        {
            return new Dictionary<SchemeNodeType, INodeProcessor>
            {
                { SchemeNodeType.Map, new MapNodeProcessor() },
                { SchemeNodeType.Array, new ArrayNodeProcessor() },
                { SchemeNodeType.Property, new PropertyNodeProcessor() },
                { SchemeNodeType.Key, new KeyValueNodeProcessor() },
                { SchemeNodeType.Value, new ValueNodeProcessor() },
                { SchemeNodeType.Ignore, new IgnoreNodeProcessor() }
            };
        }
    }
}