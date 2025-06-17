using System.Collections.Generic;
using ExcelToYamlAddin.Domain.Common;

namespace ExcelToYamlAddin.Domain.ValueObjects
{
    /// <summary>
    /// YAML 스타일을 나타내는 값 객체
    /// </summary>
    public class YamlStyle : ValueObject
    {
        /// <summary>
        /// 표준 스타일
        /// </summary>
        public static readonly YamlStyle Canonical = new YamlStyle("CANONICAL", "표준");

        /// <summary>
        /// 플로우 스타일
        /// </summary>
        public static readonly YamlStyle Flow = new YamlStyle("FLOW", "플로우");

        /// <summary>
        /// 블록 스타일
        /// </summary>
        public static readonly YamlStyle Block = new YamlStyle("BLOCK", "블록");

        /// <summary>
        /// 자동 스타일
        /// </summary>
        public static readonly YamlStyle Auto = new YamlStyle("AUTO", "자동");

        /// <summary>
        /// 스타일 값
        /// </summary>
        public string Value { get; }

        /// <summary>
        /// 스타일 설명
        /// </summary>
        public string Description { get; }

        private YamlStyle(string value, string description)
        {
            Value = value;
            Description = description;
        }

        /// <summary>
        /// 문자열로부터 YamlStyle 생성
        /// </summary>
        public static YamlStyle FromString(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return Canonical;

            switch (value.ToUpperInvariant())
            {
                case "CANONICAL":
                    return Canonical;
                case "FLOW":
                    return Flow;
                case "BLOCK":
                    return Block;
                case "AUTO":
                    return Auto;
                default:
                    return Canonical;
            }
        }

        /// <summary>
        /// 모든 정의된 YAML 스타일 반환
        /// </summary>
        public static IEnumerable<YamlStyle> GetAll()
        {
            yield return Canonical;
            yield return Flow;
            yield return Block;
            yield return Auto;
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return Value;
        }

        public override string ToString()
        {
            return $"{Value} ({Description})";
        }

        /// <summary>
        /// 암시적 문자열 변환
        /// </summary>
        public static implicit operator string(YamlStyle style)
        {
            return style?.Value;
        }
    }
}