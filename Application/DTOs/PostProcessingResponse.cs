using System;
using System.Collections.Generic;

namespace ExcelToYamlAddin.Application.DTOs
{
    /// <summary>
    /// 후처리 응답 DTO
    /// </summary>
    public class PostProcessingResponse
    {
        /// <summary>
        /// 처리 성공 여부
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 처리된 컨텐츠
        /// </summary>
        public string ProcessedContent { get; set; }

        /// <summary>
        /// 오류 메시지
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 처리 단계별 결과
        /// </summary>
        public List<ProcessingStepResult> StepResults { get; set; }

        /// <summary>
        /// 전체 처리 시간
        /// </summary>
        public TimeSpan ProcessingTime { get; set; }

        /// <summary>
        /// 처리 메타데이터
        /// </summary>
        public Dictionary<string, object> Metadata { get; set; }

        /// <summary>
        /// 기본 생성자
        /// </summary>
        public PostProcessingResponse()
        {
            StepResults = new List<ProcessingStepResult>();
            Metadata = new Dictionary<string, object>();
        }

        /// <summary>
        /// 성공 응답 생성
        /// </summary>
        public static PostProcessingResponse Success(string processedContent, List<ProcessingStepResult> stepResults = null)
        {
            return new PostProcessingResponse
            {
                IsSuccess = true,
                ProcessedContent = processedContent,
                StepResults = stepResults ?? new List<ProcessingStepResult>()
            };
        }

        /// <summary>
        /// 실패 응답 생성
        /// </summary>
        public static PostProcessingResponse Failure(string errorMessage, List<ProcessingStepResult> stepResults = null)
        {
            return new PostProcessingResponse
            {
                IsSuccess = false,
                ErrorMessage = errorMessage,
                StepResults = stepResults ?? new List<ProcessingStepResult>()
            };
        }
    }

    /// <summary>
    /// 처리 단계 결과
    /// </summary>
    public class ProcessingStepResult
    {
        /// <summary>
        /// 단계 이름
        /// </summary>
        public string StepName { get; set; }

        /// <summary>
        /// 성공 여부
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 처리 시간
        /// </summary>
        public TimeSpan Duration { get; set; }

        /// <summary>
        /// 변경 사항 수
        /// </summary>
        public int ChangesCount { get; set; }

        /// <summary>
        /// 단계별 메시지
        /// </summary>
        public string Message { get; set; }

        /// <summary>
        /// 단계별 세부 정보
        /// </summary>
        public Dictionary<string, object> Details { get; set; }

        public ProcessingStepResult()
        {
            Details = new Dictionary<string, object>();
        }

        public ProcessingStepResult(string stepName, bool isSuccess, TimeSpan duration)
        {
            StepName = stepName;
            IsSuccess = isSuccess;
            Duration = duration;
            Details = new Dictionary<string, object>();
        }
    }
}