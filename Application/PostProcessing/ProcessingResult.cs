using System;

namespace ExcelToYamlAddin.Application.PostProcessing
{
    /// <summary>
    /// 후처리 결과를 나타내는 클래스입니다.
    /// </summary>
    public class ProcessingResult
    {
        /// <summary>
        /// 처리된 출력 문자열
        /// </summary>
        public string Output { get; set; }

        /// <summary>
        /// 처리 성공 여부
        /// </summary>
        public bool Success { get; set; }

        /// <summary>
        /// 에러 메시지 (실패한 경우)
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 처리 소요 시간
        /// </summary>
        public TimeSpan ProcessingTime { get; set; }

        /// <summary>
        /// 처리기 이름
        /// </summary>
        public string ProcessorName { get; set; }

        /// <summary>
        /// 성공적인 결과를 생성합니다.
        /// </summary>
        public static ProcessingResult CreateSuccess(string output, string processorName = null)
        {
            return new ProcessingResult
            {
                Output = output,
                Success = true,
                ProcessorName = processorName
            };
        }

        /// <summary>
        /// 실패한 결과를 생성합니다.
        /// </summary>
        public static ProcessingResult CreateFailure(string errorMessage, string processorName = null)
        {
            return new ProcessingResult
            {
                Success = false,
                ErrorMessage = errorMessage,
                ProcessorName = processorName
            };
        }
    }
}