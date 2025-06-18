using ExcelToYamlAddin.Domain.ValueObjects;
using System.Collections.Generic;

namespace ExcelToYamlAddin.Application.PostProcessing
{
    /// <summary>
    /// 후처리 파이프라인에서 사용되는 컨텍스트 정보입니다.
    /// </summary>
    public class ProcessingContext
    {
        /// <summary>
        /// 처리 중인 파일 경로
        /// </summary>
        public string FilePath { get; set; }

        /// <summary>
        /// 출력 형식
        /// </summary>
        public OutputFormat OutputFormat { get; set; }

        /// <summary>
        /// 처리 옵션들
        /// </summary>
        public ProcessingOptions Options { get; set; }

        /// <summary>
        /// 추가 메타데이터
        /// </summary>
        public Dictionary<string, object> Metadata { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// 시트 이름
        /// </summary>
        public string SheetName { get; set; }
    }

    /// <summary>
    /// 후처리 옵션들
    /// </summary>
    public class ProcessingOptions
    {
        /// <summary>
        /// 병합 활성화 여부
        /// </summary>
        public bool EnableMerge { get; set; }

        /// <summary>
        /// Flow 스타일 적용 여부
        /// </summary>
        public bool ApplyFlowStyle { get; set; }

        /// <summary>
        /// 빈 필드 포함 여부
        /// </summary>
        public bool IncludeEmptyFields { get; set; }

        /// <summary>
        /// 병합 키 경로 설정
        /// </summary>
        public string MergeKeyPaths { get; set; }

        /// <summary>
        /// Flow 스타일 설정
        /// </summary>
        public string FlowStyleConfig { get; set; }
    }
}