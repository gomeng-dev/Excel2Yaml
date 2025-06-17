using System.Collections.Generic;
using System.Threading.Tasks;

namespace ExcelToYamlAddin.Infrastructure.Interfaces
{
    /// <summary>
    /// 구성 서비스 인터페이스
    /// </summary>
    public interface IConfigurationService
    {
        /// <summary>
        /// 구성 값을 가져옵니다.
        /// </summary>
        /// <typeparam name="T">값 타입</typeparam>
        /// <param name="key">구성 키</param>
        /// <returns>구성 값</returns>
        T GetValue<T>(string key);

        /// <summary>
        /// 구성 값을 가져옵니다. 없으면 기본값을 반환합니다.
        /// </summary>
        /// <typeparam name="T">값 타입</typeparam>
        /// <param name="key">구성 키</param>
        /// <param name="defaultValue">기본값</param>
        /// <returns>구성 값 또는 기본값</returns>
        T GetValue<T>(string key, T defaultValue);

        /// <summary>
        /// 구성 값을 설정합니다.
        /// </summary>
        /// <typeparam name="T">값 타입</typeparam>
        /// <param name="key">구성 키</param>
        /// <param name="value">값</param>
        void SetValue<T>(string key, T value);

        /// <summary>
        /// 구성 섹션을 가져옵니다.
        /// </summary>
        /// <param name="sectionName">섹션 이름</param>
        /// <returns>구성 섹션</returns>
        IConfigurationSection GetSection(string sectionName);

        /// <summary>
        /// 구성을 다시 로드합니다.
        /// </summary>
        /// <returns>비동기 작업</returns>
        Task ReloadAsync();

        /// <summary>
        /// 구성을 저장합니다.
        /// </summary>
        /// <returns>비동기 작업</returns>
        Task SaveAsync();
    }

    /// <summary>
    /// 구성 섹션 인터페이스
    /// </summary>
    public interface IConfigurationSection
    {
        /// <summary>
        /// 섹션 이름
        /// </summary>
        string Name { get; }

        /// <summary>
        /// 섹션 경로
        /// </summary>
        string Path { get; }

        /// <summary>
        /// 섹션 값
        /// </summary>
        string Value { get; set; }

        /// <summary>
        /// 하위 섹션들
        /// </summary>
        IEnumerable<IConfigurationSection> Children { get; }

        /// <summary>
        /// 키-값 쌍으로 가져옵니다.
        /// </summary>
        /// <returns>키-값 쌍</returns>
        IEnumerable<KeyValuePair<string, string>> AsEnumerable();
    }
}