using System;

namespace ExcelToYamlAddin.Domain.Constants
{
    /// <summary>
    /// HTML 및 CSS 스타일 관련 상수를 정의하는 클래스
    /// </summary>
    public static class HtmlStyles
    {
        /// <summary>
        /// 테이블 스타일
        /// </summary>
        public static class Table
        {
            /// <summary>
            /// 기본 테이블 스타일
            /// </summary>
            public const string Base = "border-collapse: collapse; margin: 20px;";

            /// <summary>
            /// 테이블 셀 스타일
            /// </summary>
            public const string Cell = "border: 1px solid #999; padding: 8px; min-width: 80px;";

            /// <summary>
            /// 테이블 헤더 스타일
            /// </summary>
            public const string Header = "background-color: #f2f2f2; font-weight: bold;";
        }

        /// <summary>
        /// 셀 배경색 스타일
        /// </summary>
        public static class CellBackground
        {
            /// <summary>
            /// 병합된 셀
            /// </summary>
            public const string Merged = "background-color: #e8f4fc;";

            /// <summary>
            /// 스키마 종료 마커
            /// </summary>
            public const string SchemeEnd = "background-color: #ff0000; color: white; text-align: center;";

            /// <summary>
            /// 배열 마커
            /// </summary>
            public const string ArrayMarker = "background-color: #00CC00;";

            /// <summary>
            /// 객체 마커
            /// </summary>
            public const string ObjectMarker = "background-color: #CCFFCC;";

            /// <summary>
            /// 빈 셀
            /// </summary>
            public const string Empty = "background-color: #f9f9f9;";

            /// <summary>
            /// 행 헤더
            /// </summary>
            public const string RowHeader = "background-color: #ddd; font-weight: bold; width: 50px;";

            /// <summary>
            /// 열 헤더
            /// </summary>
            public const string ColHeader = "background-color: #ddd; font-weight: bold; height: 30px; text-align: center;";
        }

        /// <summary>
        /// 색상 값
        /// </summary>
        public static class Colors
        {
            /// <summary>
            /// 빨간색 임계값 (RGB에서 R 값)
            /// </summary>
            public const int RedThreshold = 200;

            /// <summary>
            /// 초록색 임계값 (RGB에서 G 값)
            /// </summary>
            public const int GreenThreshold = 200;

            /// <summary>
            /// 낮은 색상 임계값
            /// </summary>
            public const int LowColorThreshold = 100;
        }

        /// <summary>
        /// CSS 클래스명
        /// </summary>
        public static class CssClasses
        {
            /// <summary>
            /// 병합된 셀
            /// </summary>
            public const string Merged = "merged";

            /// <summary>
            /// 스키마 종료
            /// </summary>
            public const string SchemeEnd = "scheme-end";

            /// <summary>
            /// 배열 마커
            /// </summary>
            public const string ArrayMarker = "array-marker";

            /// <summary>
            /// 객체 마커
            /// </summary>
            public const string ObjectMarker = "object-marker";

            /// <summary>
            /// 빈 셀
            /// </summary>
            public const string Empty = "empty";

            /// <summary>
            /// 행 헤더
            /// </summary>
            public const string RowHeader = "row-header";

            /// <summary>
            /// 열 헤더
            /// </summary>
            public const string ColHeader = "col-header";
        }

        /// <summary>
        /// HTML 태그 및 속성
        /// </summary>
        public static class HtmlTags
        {
            /// <summary>
            /// DOCTYPE 선언
            /// </summary>
            public const string DocType = "<!DOCTYPE html>";

            /// <summary>
            /// UTF-8 메타 태그
            /// </summary>
            public const string Utf8Meta = "<meta charset='UTF-8'>";

            /// <summary>
            /// 범례 섹션 스타일
            /// </summary>
            public const string LegendSectionStyle = "margin: 20px;";

            /// <summary>
            /// 범례 항목 스타일
            /// </summary>
            public const string LegendItemStyle = "padding: 5px;";
        }

        /// <summary>
        /// 레이아웃 관련 상수
        /// </summary>
        public static class Layout
        {
            /// <summary>
            /// 기본 여백
            /// </summary>
            public const int DefaultMargin = 20;

            /// <summary>
            /// 셀 패딩
            /// </summary>
            public const int CellPadding = 8;

            /// <summary>
            /// 최소 셀 너비
            /// </summary>
            public const int MinCellWidth = 80;

            /// <summary>
            /// 행 헤더 너비
            /// </summary>
            public const int RowHeaderWidth = 50;

            /// <summary>
            /// 열 헤더 높이
            /// </summary>
            public const int ColHeaderHeight = 30;

            /// <summary>
            /// 범례 항목 패딩
            /// </summary>
            public const int LegendItemPadding = 5;
        }
    }
}