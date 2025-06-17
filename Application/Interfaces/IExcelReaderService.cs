using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.ValueObjects;

namespace ExcelToYamlAddin.Application.Interfaces
{
    /// <summary>
    /// Excel 파일 읽기 서비스 인터페이스
    /// </summary>
    public interface IExcelReaderService
    {
        /// <summary>
        /// Excel 파일을 처리합니다.
        /// </summary>
        /// <param name="inputPath">입력 파일 경로</param>
        /// <param name="outputPath">출력 파일 경로</param>
        /// <returns>비동기 작업</returns>
        Task ProcessExcelFileAsync(string inputPath, string outputPath);

        /// <summary>
        /// 특정 시트만 처리합니다.
        /// </summary>
        /// <param name="inputPath">입력 파일 경로</param>
        /// <param name="outputPath">출력 파일 경로</param>
        /// <param name="targetSheetName">대상 시트 이름</param>
        /// <returns>비동기 작업</returns>
        Task ProcessExcelFileAsync(string inputPath, string outputPath, string targetSheetName);

        /// <summary>
        /// Excel 파일을 읽고 문자열로 변환합니다.
        /// </summary>
        /// <param name="inputPath">입력 파일 경로</param>
        /// <param name="config">변환 설정</param>
        /// <returns>변환된 문자열</returns>
        Task<string> ConvertToStringAsync(string inputPath, ExcelToYamlConfig config);
    }
}