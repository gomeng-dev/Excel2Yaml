using ExcelToYamlAddin.Domain.Entities;

namespace ExcelToYamlAddin.Application.DTOs
{
    /// <summary>
    /// 스키마 검증 요청 DTO
    /// </summary>
    public class SchemeValidationRequest
    {
        /// <summary>
        /// 검증할 스키마
        /// </summary>
        public Scheme Scheme { get; set; }

        /// <summary>
        /// 검증 옵션
        /// </summary>
        public ValidationOptions Options { get; set; }

        /// <summary>
        /// 기본 생성자
        /// </summary>
        public SchemeValidationRequest()
        {
            Options = new ValidationOptions();
        }

        /// <summary>
        /// 매개변수를 받는 생성자
        /// </summary>
        public SchemeValidationRequest(Scheme scheme, ValidationOptions options = null)
        {
            Scheme = scheme;
            Options = options ?? new ValidationOptions();
        }
    }

    /// <summary>
    /// 검증 옵션
    /// </summary>
    public class ValidationOptions
    {
        /// <summary>
        /// 구조 검증 활성화
        /// </summary>
        public bool ValidateStructure { get; set; } = true;

        /// <summary>
        /// 무결성 검증 활성화
        /// </summary>
        public bool ValidateIntegrity { get; set; } = true;

        /// <summary>
        /// 순환 참조 검증 활성화
        /// </summary>
        public bool ValidateCircularReferences { get; set; } = true;

        /// <summary>
        /// 중복 키 검증 활성화
        /// </summary>
        public bool ValidateDuplicateKeys { get; set; } = true;

        /// <summary>
        /// 빈 컨테이너 허용
        /// </summary>
        public bool AllowEmptyContainers { get; set; } = false;

        /// <summary>
        /// 최대 깊이 제한 (0 = 무제한)
        /// </summary>
        public int MaxDepth { get; set; } = 0;
    }
}