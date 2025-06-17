using System.Collections.Generic;
using ExcelToYamlAddin.Domain.Common;

namespace ExcelToYamlAddin.Domain.ValueObjects
{
    /// <summary>
    /// 출력 형식을 나타내는 값 객체
    /// </summary>
    public class OutputFormat : ValueObject
    {
        /// <summary>
        /// YAML 형식
        /// </summary>
        public static readonly OutputFormat Yaml = new OutputFormat("YAML", ".yaml");

        /// <summary>
        /// JSON 형식
        /// </summary>
        public static readonly OutputFormat Json = new OutputFormat("JSON", ".json");

        /// <summary>
        /// XML 형식
        /// </summary>
        public static readonly OutputFormat Xml = new OutputFormat("XML", ".xml");

        /// <summary>
        /// HTML 형식
        /// </summary>
        public static readonly OutputFormat Html = new OutputFormat("HTML", ".html");

        /// <summary>
        /// 형식 이름
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// 파일 확장자
        /// </summary>
        public string Extension { get; }

        private OutputFormat(string name, string extension)
        {
            Name = name;
            Extension = extension;
        }

        /// <summary>
        /// 문자열로부터 OutputFormat 생성
        /// </summary>
        public static OutputFormat FromString(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return Yaml;

            switch (value.ToUpperInvariant())
            {
                case "YAML":
                case "YML":
                    return Yaml;
                case "JSON":
                    return Json;
                case "XML":
                    return Xml;
                case "HTML":
                    return Html;
                default:
                    return Yaml;
            }
        }

        /// <summary>
        /// 확장자로부터 OutputFormat 생성
        /// </summary>
        public static OutputFormat FromExtension(string extension)
        {
            if (string.IsNullOrWhiteSpace(extension))
                return Yaml;

            var ext = extension.TrimStart('.');
            switch (ext.ToLowerInvariant())
            {
                case "yaml":
                case "yml":
                    return Yaml;
                case "json":
                    return Json;
                case "xml":
                    return Xml;
                case "html":
                case "htm":
                    return Html;
                default:
                    return Yaml;
            }
        }

        /// <summary>
        /// 모든 정의된 출력 형식 반환
        /// </summary>
        public static IEnumerable<OutputFormat> GetAll()
        {
            yield return Yaml;
            yield return Json;
            yield return Xml;
            yield return Html;
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return Name;
        }

        public override string ToString()
        {
            return Name;
        }

        /// <summary>
        /// 암시적 문자열 변환
        /// </summary>
        public static implicit operator string(OutputFormat format)
        {
            return format?.Name;
        }
    }
}