using System;
using System.Collections.Generic;

namespace ExcelToYamlAddin.Application.DTOs
{
    /// <summary>
    /// Excel 변환 응답 DTO
    /// </summary>
    public class ExcelConversionResponse
    {
        /// <summary>
        /// 성공 여부
        /// </summary>
        public bool IsSuccess { get; set; }

        /// <summary>
        /// 오류 메시지
        /// </summary>
        public string ErrorMessage { get; set; }

        /// <summary>
        /// 처리된 시트 정보
        /// </summary>
        public List<ProcessedSheetInfo> ProcessedSheets { get; set; }

        /// <summary>
        /// 전체 처리 시간
        /// </summary>
        public TimeSpan ProcessingTime { get; set; }

        /// <summary>
        /// 생성된 파일 경로 목록
        /// </summary>
        public List<string> GeneratedFiles { get; set; }

        /// <summary>
        /// 기본 생성자
        /// </summary>
        public ExcelConversionResponse()
        {
            ProcessedSheets = new List<ProcessedSheetInfo>();
            GeneratedFiles = new List<string>();
        }

        /// <summary>
        /// 성공 응답 생성
        /// </summary>
        public static ExcelConversionResponse Success(List<ProcessedSheetInfo> processedSheets, TimeSpan processingTime)
        {
            return new ExcelConversionResponse
            {
                IsSuccess = true,
                ProcessedSheets = processedSheets,
                ProcessingTime = processingTime
            };
        }

        /// <summary>
        /// 실패 응답 생성
        /// </summary>
        public static ExcelConversionResponse Failure(string errorMessage)
        {
            return new ExcelConversionResponse
            {
                IsSuccess = false,
                ErrorMessage = errorMessage
            };
        }
    }

    /// <summary>
    /// 처리된 시트 정보
    /// </summary>
    public class ProcessedSheetInfo
    {
        /// <summary>
        /// 시트 이름
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 출력 파일 경로
        /// </summary>
        public string OutputFilePath { get; set; }

        /// <summary>
        /// 처리된 행 수
        /// </summary>
        public int ProcessedRows { get; set; }

        /// <summary>
        /// 처리된 열 수
        /// </summary>
        public int ProcessedColumns { get; set; }

        /// <summary>
        /// 파일 크기 (바이트)
        /// </summary>
        public long FileSize { get; set; }

        /// <summary>
        /// MD5 해시 (생성된 경우)
        /// </summary>
        public string Md5Hash { get; set; }
    }
}