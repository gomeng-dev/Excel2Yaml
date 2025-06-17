using System.Collections.Generic;

namespace ExcelToYamlAddin.Application.DTOs
{
    /// <summary>
    /// 후처리 요청 DTO
    /// </summary>
    public class PostProcessingRequest
    {
        /// <summary>
        /// 처리할 컨텐츠
        /// </summary>
        public string Content { get; set; }

        /// <summary>
        /// 컨텐츠 타입 (yaml, json, xml 등)
        /// </summary>
        public string ContentType { get; set; }

        /// <summary>
        /// 적용할 후처리 목록
        /// </summary>
        public List<string> ProcessingSteps { get; set; }

        /// <summary>
        /// 후처리 옵션
        /// </summary>
        public Dictionary<string, object> Options { get; set; }

        /// <summary>
        /// 소스 파일 경로 (참조용)
        /// </summary>
        public string SourceFilePath { get; set; }

        /// <summary>
        /// 기본 생성자
        /// </summary>
        public PostProcessingRequest()
        {
            ProcessingSteps = new List<string>();
            Options = new Dictionary<string, object>();
        }

        /// <summary>
        /// 매개변수를 받는 생성자
        /// </summary>
        public PostProcessingRequest(string content, string contentType)
        {
            Content = content;
            ContentType = contentType;
            ProcessingSteps = new List<string>();
            Options = new Dictionary<string, object>();
        }

        /// <summary>
        /// 옵션 추가 헬퍼 메서드
        /// </summary>
        public PostProcessingRequest WithOption(string key, object value)
        {
            Options[key] = value;
            return this;
        }

        /// <summary>
        /// 처리 단계 추가 헬퍼 메서드
        /// </summary>
        public PostProcessingRequest WithProcessingStep(string step)
        {
            ProcessingSteps.Add(step);
            return this;
        }
    }
}