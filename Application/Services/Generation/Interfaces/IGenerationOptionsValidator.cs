using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.ValueObjects;

namespace ExcelToYamlAddin.Application.Services.Generation.Interfaces
{
    /// <summary>
    /// YAML 생성 옵션을 검증하는 인터페이스
    /// </summary>
    public interface IGenerationOptionsValidator
    {
        /// <summary>
        /// 생성 옵션의 유효성을 검증합니다.
        /// </summary>
        /// <param name="options">검증할 옵션</param>
        /// <returns>검증 결과</returns>
        Task<ValidationResult> ValidateAsync(YamlGenerationOptions options);
    }

    /// <summary>
    /// 검증 결과
    /// </summary>
    public class ValidationResult
    {
        public bool IsValid { get; private set; }
        public string ErrorMessage { get; private set; }

        private ValidationResult() { }

        public static ValidationResult Valid()
        {
            return new ValidationResult { IsValid = true };
        }

        public static ValidationResult Invalid(string errorMessage)
        {
            return new ValidationResult
            {
                IsValid = false,
                ErrorMessage = errorMessage
            };
        }
    }
}