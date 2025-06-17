using System;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using ExcelToYamlAddin.Application.Interfaces;
using ExcelToYamlAddin.Domain.Entities;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Domain.Constants;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation
{
    /// <summary>
    /// YAML 생성 서비스의 주요 오케스트레이션을 담당합니다.
    /// </summary>
    public class YamlGenerationService
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger(nameof(YamlGenerationService));
        
        private readonly INodeTraverser _traverser;
        private readonly IYamlBuilder _yamlBuilder;
        private readonly IGenerationOptionsValidator _optionsValidator;

        public YamlGenerationService(
            INodeTraverser traverser,
            IYamlBuilder yamlBuilder,
            IGenerationOptionsValidator optionsValidator)
        {
            _traverser = traverser ?? throw new ArgumentNullException(nameof(traverser));
            _yamlBuilder = yamlBuilder ?? throw new ArgumentNullException(nameof(yamlBuilder));
            _optionsValidator = optionsValidator ?? throw new ArgumentNullException(nameof(optionsValidator));
        }

        /// <summary>
        /// 스키마와 워크시트 데이터를 기반으로 YAML을 생성합니다.
        /// </summary>
        public async Task<GenerationResult> GenerateAsync(
            GenerationRequest request,
            CancellationToken cancellationToken = default)
        {
            if (request == null)
                throw new ArgumentNullException(nameof(request));

            try
            {
                Logger.Information($"YAML 생성 시작: {request.SheetName}");
                
                // 옵션 검증
                var validationResult = await _optionsValidator.ValidateAsync(request.Options);
                if (!validationResult.IsValid)
                {
                    return GenerationResult.CreateFailure(validationResult.ErrorMessage);
                }

                // 생성 컨텍스트 생성
                var context = new GenerationContext(
                    request.Worksheet,
                    request.Scheme,
                    request.Options);

                // 노드 순회 및 데이터 수집
                var traversalResult = await _traverser.TraverseAsync(
                    request.Scheme.Root,
                    context,
                    cancellationToken);

                if (!traversalResult.Success)
                {
                    return GenerationResult.CreateFailure(traversalResult.ErrorMessage);
                }

                // YAML 생성
                var yaml = await _yamlBuilder.BuildAsync(
                    traversalResult.Data,
                    request.Options,
                    cancellationToken);

                Logger.Information($"YAML 생성 완료: {request.SheetName}");
                
                return GenerationResult.CreateSuccess(yaml);
            }
            catch (OperationCanceledException)
            {
                Logger.Warning($"YAML 생성 취소됨: {request.SheetName}");
                return GenerationResult.CreateFailure("작업이 취소되었습니다.");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"YAML 생성 중 오류 발생: {request.SheetName}");
                return GenerationResult.CreateFailure($"YAML 생성 실패: {ex.Message}");
            }
        }

        /// <summary>
        /// 동기 방식의 YAML 생성 (기존 호환성)
        /// </summary>
        public string Generate(Scheme scheme, IXLWorksheet worksheet, YamlGenerationOptions options)
        {
            var request = new GenerationRequest
            {
                Scheme = scheme,
                Worksheet = worksheet,
                SheetName = worksheet.Name,
                Options = options
            };

            var result = GenerateAsync(request).GetAwaiter().GetResult();
            
            if (!result.Success)
            {
                throw new InvalidOperationException(result.ErrorMessage);
            }

            return result.Output;
        }
    }

    /// <summary>
    /// YAML 생성 요청 DTO
    /// </summary>
    public class GenerationRequest
    {
        public Scheme Scheme { get; set; }
        public IXLWorksheet Worksheet { get; set; }
        public string SheetName { get; set; }
        public YamlGenerationOptions Options { get; set; }
    }

    /// <summary>
    /// YAML 생성 결과 DTO
    /// </summary>
    public class GenerationResult
    {
        public bool Success { get; private set; }
        public string Output { get; private set; }
        public string ErrorMessage { get; private set; }

        private GenerationResult() { }

        public static GenerationResult CreateSuccess(string output)
        {
            return new GenerationResult
            {
                Success = true,
                Output = output
            };
        }

        public static GenerationResult CreateFailure(string errorMessage)
        {
            return new GenerationResult
            {
                Success = false,
                ErrorMessage = errorMessage
            };
        }
    }

}