using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelToYamlAddin.Infrastructure.Interfaces
{
    /// <summary>
    /// 파일 시스템 서비스 인터페이스
    /// </summary>
    public interface IFileSystemService
    {
        /// <summary>
        /// 파일 존재 여부를 확인합니다.
        /// </summary>
        /// <param name="path">파일 경로</param>
        /// <returns>존재 여부</returns>
        bool FileExists(string path);

        /// <summary>
        /// 디렉토리 존재 여부를 확인합니다.
        /// </summary>
        /// <param name="path">디렉토리 경로</param>
        /// <returns>존재 여부</returns>
        bool DirectoryExists(string path);

        /// <summary>
        /// 파일을 읽습니다.
        /// </summary>
        /// <param name="path">파일 경로</param>
        /// <returns>파일 내용</returns>
        Task<string> ReadFileAsync(string path);

        /// <summary>
        /// 파일을 씁니다.
        /// </summary>
        /// <param name="path">파일 경로</param>
        /// <param name="content">파일 내용</param>
        /// <returns>비동기 작업</returns>
        Task WriteFileAsync(string path, string content);

        /// <summary>
        /// 파일을 복사합니다.
        /// </summary>
        /// <param name="sourcePath">원본 경로</param>
        /// <param name="destinationPath">대상 경로</param>
        /// <param name="overwrite">덮어쓰기 여부</param>
        /// <returns>비동기 작업</returns>
        Task CopyFileAsync(string sourcePath, string destinationPath, bool overwrite = false);

        /// <summary>
        /// 파일을 삭제합니다.
        /// </summary>
        /// <param name="path">파일 경로</param>
        /// <returns>비동기 작업</returns>
        Task DeleteFileAsync(string path);

        /// <summary>
        /// 디렉토리를 생성합니다.
        /// </summary>
        /// <param name="path">디렉토리 경로</param>
        /// <returns>생성된 디렉토리 정보</returns>
        DirectoryInfo CreateDirectory(string path);

        /// <summary>
        /// 디렉토리의 파일 목록을 가져옵니다.
        /// </summary>
        /// <param name="path">디렉토리 경로</param>
        /// <param name="searchPattern">검색 패턴</param>
        /// <param name="searchOption">검색 옵션</param>
        /// <returns>파일 목록</returns>
        IEnumerable<string> GetFiles(string path, string searchPattern = "*", SearchOption searchOption = SearchOption.TopDirectoryOnly);
    }
}