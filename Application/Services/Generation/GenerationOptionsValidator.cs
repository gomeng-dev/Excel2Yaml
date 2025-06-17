using System;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation
{
    /// <summary>
    /// YAML 생성 옵션을 검증하는 클래스
    /// </summary>
    public class GenerationOptionsValidator : IGenerationOptionsValidator
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger(nameof(GenerationOptionsValidator));
        
        private const int MinMaxDepth = 1;
        private const int MaxMaxDepth = 1000;

        public async Task<ValidationResult> ValidateAsync(YamlGenerationOptions options)
        {
            if (options == null)
            {
                return ValidationResult.Invalid("생성 옵션이 null입니다.");
            }

            try
            {
                // MaxDepth 검증
                if (options.MaxDepth < MinMaxDepth || options.MaxDepth > MaxMaxDepth)
                {
                    return ValidationResult.Invalid(
                        $"MaxDepth는 {MinMaxDepth}와 {MaxMaxDepth} 사이여야 합니다. 현재 값: {options.MaxDepth}");
                }

                // IndentSize 검증
                if (options.IndentSize < 0 || options.IndentSize > 10)
                {
                    return ValidationResult.Invalid(
                        $"IndentSize는 0과 10 사이여야 합니다. 현재 값: {options.IndentSize}");
                }

                // OutputPath 검증 (선택적)
                if (!string.IsNullOrEmpty(options.OutputPath))
                {
                    try
                    {
                        // 경로 유효성 확인
                        var path = System.IO.Path.GetFullPath(options.OutputPath);
                    }
                    catch (Exception)
                    {
                        return ValidationResult.Invalid(
                            $"유효하지 않은 출력 경로입니다: {options.OutputPath}");
                    }
                }

                Logger.Debug("생성 옵션 검증 성공");
                return await Task.FromResult(ValidationResult.Valid());
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "생성 옵션 검증 중 오류 발생");
                return ValidationResult.Invalid($"옵션 검증 중 오류 발생: {ex.Message}");
            }
        }
    }
}