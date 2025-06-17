using System;
using System.Collections.Generic;
using ExcelToYamlAddin.Domain.Common;
using ExcelToYamlAddin.Domain.Constants;

namespace ExcelToYamlAddin.Domain.ValueObjects
{
    /// <summary>
    /// Excel 셀의 위치를 나타내는 값 객체
    /// </summary>
#pragma warning disable CS0660, CS0661 // ValueObject 기본 클래스에서 Equals와 GetHashCode를 구현함
    public class CellPosition : ValueObject
#pragma warning restore CS0660, CS0661
    {
        /// <summary>
        /// 행 번호 (1부터 시작)
        /// </summary>
        public int Row { get; }

        /// <summary>
        /// 열 번호 (1부터 시작)
        /// </summary>
        public int Column { get; }

        /// <summary>
        /// 열 문자 (A, B, C...)
        /// </summary>
        public string ColumnLetter { get; }

        /// <summary>
        /// 셀 주소 (예: A1, B2)
        /// </summary>
        public string Address { get; }

        /// <summary>
        /// CellPosition 생성자
        /// </summary>
        /// <param name="row">행 번호 (1부터 시작)</param>
        /// <param name="column">열 번호 (1부터 시작)</param>
        public CellPosition(int row, int column)
        {
            if (row < 1)
                throw new ArgumentException(ErrorMessages.Validation.RowLessThanOne, nameof(row));
            if (column < 1)
                throw new ArgumentException(ErrorMessages.Validation.ColumnLessThanOne, nameof(column));

            Row = row;
            Column = column;
            ColumnLetter = GetColumnLetter(column);
            Address = $"{ColumnLetter}{Row}";
        }

        /// <summary>
        /// 셀 주소 문자열로부터 CellPosition 생성
        /// </summary>
        /// <param name="address">셀 주소 (예: A1, B2)</param>
        public static CellPosition FromAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                throw new ArgumentException(ErrorMessages.Validation.CellAddressIsEmpty, nameof(address));

            var columnLetters = "";
            var rowNumbers = "";
            var isReadingNumbers = false;

            foreach (var ch in address)
            {
                if (char.IsLetter(ch) && !isReadingNumbers)
                {
                    columnLetters += ch;
                }
                else if (char.IsDigit(ch))
                {
                    isReadingNumbers = true;
                    rowNumbers += ch;
                }
                else
                {
                    throw new ArgumentException(string.Format(ErrorMessages.Validation.InvalidCellAddressFormat, address), nameof(address));
                }
            }

            if (string.IsNullOrEmpty(columnLetters) || string.IsNullOrEmpty(rowNumbers))
            {
                throw new ArgumentException(string.Format(ErrorMessages.Validation.InvalidCellAddressFormat, address), nameof(address));
            }

            var column = GetColumnNumber(columnLetters);
            var row = int.Parse(rowNumbers);

            return new CellPosition(row, column);
        }

        /// <summary>
        /// 열 번호를 문자로 변환 (1 -> A, 2 -> B, 27 -> AA)
        /// </summary>
        private static string GetColumnLetter(int columnNumber)
        {
            string columnLetter = "";
            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnLetter = Convert.ToChar(65 + modulo) + columnLetter;
                columnNumber = (columnNumber - modulo) / 26;
            }
            return columnLetter;
        }

        /// <summary>
        /// 열 문자를 번호로 변환 (A -> 1, B -> 2, AA -> 27)
        /// </summary>
        private static int GetColumnNumber(string columnLetter)
        {
            int columnNumber = 0;
            for (int i = 0; i < columnLetter.Length; i++)
            {
                columnNumber = columnNumber * 26 + (columnLetter[i] - 'A' + 1);
            }
            return columnNumber;
        }

        /// <summary>
        /// 현재 위치에서 오프셋만큼 이동한 새 위치 생성
        /// </summary>
        public CellPosition Offset(int rowOffset, int columnOffset)
        {
            return new CellPosition(Row + rowOffset, Column + columnOffset);
        }

        /// <summary>
        /// 다음 행의 동일한 열 위치
        /// </summary>
        public CellPosition NextRow()
        {
            return Offset(1, 0);
        }

        /// <summary>
        /// 다음 열의 동일한 행 위치
        /// </summary>
        public CellPosition NextColumn()
        {
            return Offset(0, 1);
        }

        /// <summary>
        /// 두 위치 사이의 거리 계산
        /// </summary>
        public int DistanceTo(CellPosition other)
        {
            if (other == null)
                throw new ArgumentNullException(nameof(other));

            var rowDiff = Math.Abs(Row - other.Row);
            var colDiff = Math.Abs(Column - other.Column);
            return Math.Max(rowDiff, colDiff);
        }

        /// <summary>
        /// 특정 범위 내에 있는지 확인
        /// </summary>
        public bool IsInRange(CellPosition topLeft, CellPosition bottomRight)
        {
            if (topLeft == null || bottomRight == null)
                return false;

            return Row >= topLeft.Row && Row <= bottomRight.Row &&
                   Column >= topLeft.Column && Column <= bottomRight.Column;
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return Row;
            yield return Column;
        }

        public override string ToString()
        {
            return $"[{Row}, {Column}]";
        }

        public static bool operator ==(CellPosition left, CellPosition right)
        {
            return EqualOperator(left, right);
        }

        public static bool operator !=(CellPosition left, CellPosition right)
        {
            return NotEqualOperator(left, right);
        }
    }
}