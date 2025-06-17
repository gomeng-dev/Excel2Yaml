using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.FileSystem;
using ExcelToYamlAddin.Infrastructure.Logging;
using ExcelToYamlAddin.Application.Services;
using ExcelToYamlAddin.Application.Services.Generation.Interfaces;

namespace ExcelToYamlAddin.Application.Services.Generation
{
    /// <summary>
    /// 수집된 데이터를 YAML 문자열로 변환하는 클래스
    /// </summary>
    public class YamlBuilder : IYamlBuilder
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<YamlBuilder>();

        public YamlBuilder()
        {
        }

        /// <summary>
        /// 루트 객체를 생성합니다.
        /// </summary>
        public object CreateRoot()
        {
            return OrderedYamlFactory.CreateObject();
        }

        /// <summary>
        /// 맵 객체를 생성합니다.
        /// </summary>
        public object CreateMap()
        {
            return OrderedYamlFactory.CreateObject();
        }

        /// <summary>
        /// 배열 객체를 생성합니다.
        /// </summary>
        public object CreateArray()
        {
            return OrderedYamlFactory.CreateArray();
        }

        /// <summary>
        /// 맵에 속성을 추가합니다.
        /// </summary>
        public void AddProperty(object map, string key, object value)
        {
            if (map is YamlObject yamlObj)
            {
                yamlObj.Add(key, value);
            }
            else
            {
                throw new ArgumentException("map은 YamlObject 타입이어야 합니다.", nameof(map));
            }
        }

        /// <summary>
        /// 배열에 요소를 추가합니다.
        /// </summary>
        public void AddToArray(object array, object item)
        {
            if (array is YamlArray yamlArray)
            {
                yamlArray.Add(item);
            }
            else
            {
                throw new ArgumentException("array는 YamlArray 타입이어야 합니다.", nameof(array));
            }
        }

        /// <summary>
        /// 객체를 YAML 문자열로 변환합니다.
        /// </summary>
        public string BuildYaml(object data, YamlGenerationOptions options)
        {
            if (data == null)
            {
                Logger.Warning("빌드할 데이터가 null입니다.");
                return string.Empty;
            }

            try
            {
                // 빈 필드 제거 (옵션에 따라)
                if (!options.ShowEmptyFields && data != null)
                {
                    OrderedYamlFactory.RemoveEmptyProperties(data);
                }

                // OrderedYamlFactory를 사용하여 YAML 생성
                return OrderedYamlFactory.SerializeToYaml(
                    data,
                    options.IndentSize,
                    options.Style,
                    false, // preserveQuotes
                    options.ShowEmptyFields
                );
            }
            catch (Exception ex)
            {
                Logger.Error($"YAML 빌드 중 오류 발생: {ex.Message}", ex);
                throw new InvalidOperationException($"YAML 빌드 실패: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 데이터를 YAML 문자열로 변환합니다. (비동기)
        /// </summary>
        public async Task<string> BuildAsync(
            object data,
            YamlGenerationOptions options,
            CancellationToken cancellationToken = default)
        {
            return await Task.Run(() => BuildYaml(data, options), cancellationToken);
        }

    }

    /// <summary>
    /// YAML 빌더 팩토리
    /// </summary>
    public class YamlBuilderFactory
    {
        public static IYamlBuilder Create()
        {
            return new YamlBuilder();
        }
    }
}