using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.FileSystem;

namespace ExcelToYamlAddin.Application.Services.Generation.Interfaces
{
    /// <summary>
    /// 수집된 데이터를 YAML 문자열로 변환하는 인터페이스
    /// </summary>
    public interface IYamlBuilder
    {
        /// <summary>
        /// 루트 객체를 생성합니다.
        /// </summary>
        object CreateRoot();

        /// <summary>
        /// 맵 객체를 생성합니다.
        /// </summary>
        object CreateMap();

        /// <summary>
        /// 배열 객체를 생성합니다.
        /// </summary>
        object CreateArray();

        /// <summary>
        /// 맵에 속성을 추가합니다.
        /// </summary>
        void AddProperty(object map, string key, object value);

        /// <summary>
        /// 배열에 요소를 추가합니다.
        /// </summary>
        void AddToArray(object array, object item);

        /// <summary>
        /// 객체를 YAML 문자열로 변환합니다.
        /// </summary>
        string BuildYaml(object data, YamlGenerationOptions options);

        /// <summary>
        /// 데이터를 YAML 문자열로 변환합니다. (비동기)
        /// </summary>
        /// <param name="data">변환할 데이터</param>
        /// <param name="options">생성 옵션</param>
        /// <param name="cancellationToken">취소 토큰</param>
        /// <returns>YAML 문자열</returns>
        Task<string> BuildAsync(
            object data,
            YamlGenerationOptions options,
            CancellationToken cancellationToken = default);
    }
}