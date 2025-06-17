using System.Collections.Generic;
using ExcelToYamlAddin.Domain.ValueObjects;

namespace ExcelToYamlAddin.Application.DTOs
{
    /// <summary>
    /// Excel 변환 요청 DTO
    /// </summary>
    public class ExcelConversionRequest
    {
        /// <summary>
        /// 입력 파일 경로
        /// </summary>
        public string InputFilePath { get; set; }

        /// <summary>
        /// 출력 파일 경로
        /// </summary>
        public string OutputFilePath { get; set; }

        /// <summary>
        /// 대상 시트 이름 (null이면 모든 시트)
        /// </summary>
        public string TargetSheetName { get; set; }

        /// <summary>
        /// 변환 옵션
        /// </summary>
        public ConversionOptions Options { get; set; }

        /// <summary>
        /// 후처리 옵션
        /// </summary>
        public Dictionary<string, object> PostProcessingOptions { get; set; }

        /// <summary>
        /// 기본 생성자
        /// </summary>
        public ExcelConversionRequest()
        {
            PostProcessingOptions = new Dictionary<string, object>();
        }

        /// <summary>
        /// 매개변수를 받는 생성자
        /// </summary>
        public ExcelConversionRequest(string inputFilePath, string outputFilePath, ConversionOptions options = null)
        {
            InputFilePath = inputFilePath;
            OutputFilePath = outputFilePath;
            Options = options ?? ConversionOptions.Default();
            PostProcessingOptions = new Dictionary<string, object>();
        }
    }
}