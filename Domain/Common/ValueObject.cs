using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelToYamlAddin.Domain.Common
{
    /// <summary>
    /// 값 객체의 기본 추상 클래스
    /// </summary>
    public abstract class ValueObject
    {
        /// <summary>
        /// 두 값 객체가 같은지 비교
        /// </summary>
        protected static bool EqualOperator(ValueObject left, ValueObject right)
        {
            if (ReferenceEquals(left, null) ^ ReferenceEquals(right, null))
            {
                return false;
            }
            return ReferenceEquals(left, null) || left.Equals(right);
        }

        /// <summary>
        /// 두 값 객체가 다른지 비교
        /// </summary>
        protected static bool NotEqualOperator(ValueObject left, ValueObject right)
        {
            return !EqualOperator(left, right);
        }

        /// <summary>
        /// 동등성 비교를 위한 구성 요소들을 반환
        /// </summary>
        protected abstract IEnumerable<object> GetEqualityComponents();

        /// <summary>
        /// 값 객체의 동등성 비교
        /// </summary>
        public override bool Equals(object obj)
        {
            if (obj == null || obj.GetType() != GetType())
            {
                return false;
            }

            var other = (ValueObject)obj;

            return GetEqualityComponents().SequenceEqual(other.GetEqualityComponents());
        }

        /// <summary>
        /// 해시 코드 생성
        /// </summary>
        public override int GetHashCode()
        {
            return GetEqualityComponents()
                .Select(x => x != null ? x.GetHashCode() : 0)
                .Aggregate((x, y) => x ^ y);
        }

        /// <summary>
        /// 값 객체의 복사본 생성
        /// </summary>
        public ValueObject GetCopy()
        {
            return MemberwiseClone() as ValueObject;
        }
    }
}