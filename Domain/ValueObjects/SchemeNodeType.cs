using System;
using System.Collections.Generic;
using ExcelToYamlAddin.Domain.Common;
using ExcelToYamlAddin.Domain.Constants;

namespace ExcelToYamlAddin.Domain.ValueObjects
{
    /// <summary>
    /// 스키마 노드의 타입을 나타내는 값 객체
    /// </summary>
#pragma warning disable CS0660, CS0661 // ValueObject 기본 클래스에서 Equals와 GetHashCode를 구현함
    public class SchemeNodeType : ValueObject
#pragma warning restore CS0660, CS0661
    {
        /// <summary>
        /// 속성 노드 (일반 필드)
        /// </summary>
        public static readonly SchemeNodeType Property = new SchemeNodeType("PROPERTY", "속성");

        /// <summary>
        /// 맵/객체 노드
        /// </summary>
        public static readonly SchemeNodeType Map = new SchemeNodeType("MAP", "맵");

        /// <summary>
        /// 배열 노드
        /// </summary>
        public static readonly SchemeNodeType Array = new SchemeNodeType("ARRAY", "배열");

        /// <summary>
        /// 동적 키 노드
        /// </summary>
        public static readonly SchemeNodeType Key = new SchemeNodeType("KEY", "키");

        /// <summary>
        /// 동적 값 노드
        /// </summary>
        public static readonly SchemeNodeType Value = new SchemeNodeType("VALUE", "값");

        /// <summary>
        /// 무시 노드
        /// </summary>
        public static readonly SchemeNodeType Ignore = new SchemeNodeType("IGNORE", "무시");

        /// <summary>
        /// 노드 타입 코드
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// 노드 타입 설명
        /// </summary>
        public string Description { get; }

        /// <summary>
        /// 컨테이너 타입인지 여부 (MAP, ARRAY)
        /// </summary>
        public bool IsContainer => this == Map || this == Array;

        /// <summary>
        /// 동적 타입인지 여부 (KEY, VALUE)
        /// </summary>
        public bool IsDynamic => this == Key || this == Value;

        /// <summary>
        /// 데이터를 가질 수 있는 타입인지 여부
        /// </summary>
        public bool CanHaveData => this != Ignore;

        private SchemeNodeType(string value, string description)
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException(ErrorMessages.Validation.NodeTypeValueIsEmpty, nameof(value));

            Code = value;
            Description = description ?? value;
        }

        /// <summary>
        /// 스키마 이름으로부터 노드 타입 판별
        /// </summary>
        public static SchemeNodeType FromSchemeName(string schemeName)
        {
            if (string.IsNullOrWhiteSpace(schemeName))
                return Property;

            // 무시 마커 확인
            if (schemeName.Equals(SchemeConstants.Markers.Ignore, StringComparison.OrdinalIgnoreCase))
                return Ignore;

            // $ 마커가 없으면 일반 속성
            if (!schemeName.Contains(SchemeConstants.Markers.MarkerPrefix))
                return Property;

            // 특정 마커 확인
            if (schemeName.Contains(SchemeConstants.Markers.ArrayStart))
                return Array;

            if (schemeName.Contains(SchemeConstants.Markers.MapStart))
                return Map;

            if (schemeName.Contains(SchemeConstants.Markers.DynamicKey))
                return Key;

            if (schemeName.Contains(SchemeConstants.Markers.DynamicValue))
                return Value;

            // 마커 문자열로 타입 판별
            var parts = schemeName.Split(new[] { SchemeConstants.Markers.MarkerPrefix[0] }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length > 0)
            {
                var typeString = parts[parts.Length - 1];
                return FromTypeString(typeString);
            }

            return Property;
        }

        /// <summary>
        /// 타입 문자열로부터 노드 타입 생성
        /// </summary>
        private static SchemeNodeType FromTypeString(string typeString)
        {
            switch (typeString.ToUpperInvariant())
            {
                case "MAP":
                case SchemeConstants.NodeTypes.Map:
                    return Map;
                case "ARRAY":
                case SchemeConstants.NodeTypes.Array:
                    return Array;
                case "KEY":
                    return Key;
                case "VALUE":
                    return Value;
                case "IGNORE":
                case SchemeConstants.NodeTypes.Ignore:
                    return Ignore;
                default:
                    return Property;
            }
        }

        /// <summary>
        /// 모든 정의된 노드 타입 반환
        /// </summary>
        public static IEnumerable<SchemeNodeType> GetAll()
        {
            yield return Property;
            yield return Map;
            yield return Array;
            yield return Key;
            yield return Value;
            yield return Ignore;
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return Code;
        }

        public override string ToString()
        {
            return $"{Code} ({Description})";
        }

        public static bool operator ==(SchemeNodeType left, SchemeNodeType right)
        {
            return EqualOperator(left, right);
        }

        public static bool operator !=(SchemeNodeType left, SchemeNodeType right)
        {
            return NotEqualOperator(left, right);
        }

        /// <summary>
        /// 암시적 문자열 변환
        /// </summary>
        public static implicit operator string(SchemeNodeType nodeType)
        {
            return nodeType?.Code;
        }
    }
}