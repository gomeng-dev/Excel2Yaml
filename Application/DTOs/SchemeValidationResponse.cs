using System.Collections.Generic;
using System.Linq;

namespace ExcelToYamlAddin.Application.DTOs
{
    /// <summary>
    /// 스키마 검증 응답 DTO
    /// </summary>
    public class SchemeValidationResponse
    {
        /// <summary>
        /// 검증 성공 여부
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 검증 오류 목록
        /// </summary>
        public List<ValidationError> Errors { get; set; }

        /// <summary>
        /// 검증 경고 목록
        /// </summary>
        public List<ValidationWarning> Warnings { get; set; }

        /// <summary>
        /// 검증 통계
        /// </summary>
        public ValidationStatistics Statistics { get; set; }

        /// <summary>
        /// 기본 생성자
        /// </summary>
        public SchemeValidationResponse()
        {
            Errors = new List<ValidationError>();
            Warnings = new List<ValidationWarning>();
            Statistics = new ValidationStatistics();
        }

        /// <summary>
        /// 오류가 있는지 확인
        /// </summary>
        public bool HasErrors => Errors.Any();

        /// <summary>
        /// 경고가 있는지 확인
        /// </summary>
        public bool HasWarnings => Warnings.Any();

        /// <summary>
        /// 성공 응답 생성
        /// </summary>
        public static SchemeValidationResponse Success(ValidationStatistics statistics = null)
        {
            return new SchemeValidationResponse
            {
                IsValid = true,
                Statistics = statistics ?? new ValidationStatistics()
            };
        }

        /// <summary>
        /// 실패 응답 생성
        /// </summary>
        public static SchemeValidationResponse Failure(List<ValidationError> errors, ValidationStatistics statistics = null)
        {
            return new SchemeValidationResponse
            {
                IsValid = false,
                Errors = errors,
                Statistics = statistics ?? new ValidationStatistics()
            };
        }
    }

    /// <summary>
    /// 검증 오류
    /// </summary>
    public class ValidationError
    {
        /// <summary>
        /// 오류 코드
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 오류 메시지
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 오류 위치 (노드 경로)
        /// </summary>
        public string Location { get; set; }

        /// <summary>
        /// 오류 심각도
        /// </summary>
        public ErrorSeverity Severity { get; set; }

        /// <summary>
        /// 추가 컨텍스트 정보
        /// </summary>
        public Dictionary<string, object> Context { get; set; }

        public ValidationError()
        {
            Context = new Dictionary<string, object>();
        }
    }

    /// <summary>
    /// 검증 경고
    /// </summary>
    public class ValidationWarning
    {
        /// <summary>
        /// 경고 코드
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 경고 메시지
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 경고 위치 (노드 경로)
        /// </summary>
        public string Location { get; set; }

        /// <summary>
        /// 권장 조치
        /// </summary>
        public string RecommendedAction { get; set; }
    }

    /// <summary>
    /// 검증 통계
    /// </summary>
    public class ValidationStatistics
    {
        /// <summary>
        /// 총 노드 수
        /// </summary>
        public int TotalNodes { get; set; }

        /// <summary>
        /// 검증된 노드 수
        /// </summary>
        public int ValidatedNodes { get; set; }

        /// <summary>
        /// 최대 깊이
        /// </summary>
        public int MaxDepth { get; set; }

        /// <summary>
        /// 컨테이너 노드 수
        /// </summary>
        public int ContainerNodes { get; set; }

        /// <summary>
        /// 값 노드 수
        /// </summary>
        public int ValueNodes { get; set; }

        /// <summary>
        /// 빈 노드 수
        /// </summary>
        public int EmptyNodes { get; set; }
    }

    /// <summary>
    /// 오류 심각도
    /// </summary>
    public enum ErrorSeverity
    {
        /// <summary>
        /// 정보
        /// </summary>
        Info,

        /// <summary>
        /// 경고
        /// </summary>
        Warning,

        /// <summary>
        /// 오류
        /// </summary>
        Error,

        /// <summary>
        /// 치명적
        /// </summary>
        Critical
    }
}