using System;

namespace ExcelToYamlAddin.Domain.Constants
{
    /// <summary>
    /// Excel 스키마 관련 상수를 정의하는 클래스
    /// </summary>
    public static class SchemeConstants
    {
        /// <summary>
        /// 스키마 마커 관련 상수
        /// </summary>
        public static class Markers
        {
            /// <summary>
            /// 스키마 종료를 나타내는 마커
            /// </summary>
            public const string SchemeEnd = "$scheme_end";

            /// <summary>
            /// 배열 시작을 나타내는 마커
            /// </summary>
            public const string ArrayStart = "$[]";

            /// <summary>
            /// 맵/객체 시작을 나타내는 마커
            /// </summary>
            public const string MapStart = "${}";

            /// <summary>
            /// 동적 키를 나타내는 마커
            /// </summary>
            public const string DynamicKey = "$key";

            /// <summary>
            /// 동적 값을 나타내는 마커
            /// </summary>
            public const string DynamicValue = "$value";

            /// <summary>
            /// 무시할 셀을 나타내는 마커
            /// </summary>
            public const string Ignore = "^";

            /// <summary>
            /// 마커 접두사
            /// </summary>
            public const string MarkerPrefix = "$";
        }

        /// <summary>
        /// 시트 관련 상수
        /// </summary>
        public static class Sheet
        {
            /// <summary>
            /// 변환 대상 시트를 나타내는 접두사
            /// </summary>
            public const string ConversionPrefix = "!";

            /// <summary>
            /// 설정 시트 이름
            /// </summary>
            public const string ConfigurationName = "excel2yamlconfig";

            /// <summary>
            /// 스키마 시작 행 번호 (1-based)
            /// </summary>
            public const int SchemaStartRow = 2;

            /// <summary>
            /// 헤더 행 번호 (1-based)
            /// </summary>
            public const int HeaderRow = 1;

            /// <summary>
            /// 데이터 시작 행 번호 (1-based)
            /// </summary>
            public const int DataStartRow = 2;
        }

        /// <summary>
        /// 노드 타입 관련 상수
        /// </summary>
        public static class NodeTypes
        {
            /// <summary>
            /// 맵/객체 타입 식별자
            /// </summary>
            public const string Map = "{}";

            /// <summary>
            /// 배열 타입 식별자
            /// </summary>
            public const string Array = "[]";

            /// <summary>
            /// 키 타입 식별자
            /// </summary>
            public const string Key = "key";

            /// <summary>
            /// 값 타입 식별자
            /// </summary>
            public const string Value = "value";

            /// <summary>
            /// 무시 타입 식별자
            /// </summary>
            public const string Ignore = "^";
        }

        /// <summary>
        /// 특수 행 번호 관련 상수
        /// </summary>
        public static class RowNumbers
        {
            /// <summary>
            /// 잘못된 행 번호
            /// </summary>
            public const int IllegalRow = -1;

            /// <summary>
            /// 주석 행 번호
            /// </summary>
            public const int CommentRow = 0;
        }

        /// <summary>
        /// 설정 관련 상수
        /// </summary>
        public static class Configuration
        {
            /// <summary>
            /// 시트명 열 번호 (1-based)
            /// </summary>
            public const int SheetNameColumn = 1;

            /// <summary>
            /// 설정 키 열 번호 (1-based)
            /// </summary>
            public const int ConfigKeyColumn = 2;

            /// <summary>
            /// 설정 값 열 번호 (1-based)
            /// </summary>
            public const int ConfigValueColumn = 3;

            /// <summary>
            /// YAML 빈 필드 설정 열 번호 (1-based)
            /// </summary>
            public const int YamlEmptyFieldsColumn = 4;

            /// <summary>
            /// 빈 배열 필드 설정 열 번호 (1-based)
            /// </summary>
            public const int EmptyArrayFieldsColumn = 5;

            /// <summary>
            /// 설정 시트 업데이트 대기 시간 (초)
            /// </summary>
            public const int UpdateWaitTimeSeconds = 5;
        }

        /// <summary>
        /// 설정 키 이름 관련 상수
        /// </summary>
        public static class ConfigKeys
        {
            /// <summary>
            /// 시트명 키
            /// </summary>
            public const string SheetName = "SheetName";

            /// <summary>
            /// 설정 키
            /// </summary>
            public const string ConfigKey = "ConfigKey";

            /// <summary>
            /// 설정 값
            /// </summary>
            public const string ConfigValue = "ConfigValue";

            /// <summary>
            /// YAML 빈 필드 포함 여부
            /// </summary>
            public const string YamlEmptyFields = "YamlEmptyFields";

            /// <summary>
            /// 빈 배열 필드 포함 여부
            /// </summary>
            public const string EmptyArrayFields = "EmptyArrayFields";

            /// <summary>
            /// 머지 키 경로
            /// </summary>
            public const string MergeKeyPaths = "MergeKeyPaths";

            /// <summary>
            /// 플로우 스타일 설정
            /// </summary>
            public const string FlowStyle = "FlowStyle";
        }

        /// <summary>
        /// 파일 확장자 관련 상수
        /// </summary>
        public static class FileExtensions
        {
            /// <summary>
            /// JSON 파일 확장자
            /// </summary>
            public const string Json = ".json";

            /// <summary>
            /// YAML 파일 확장자
            /// </summary>
            public const string Yaml = ".yaml";

            /// <summary>
            /// MD5 체크섬 파일 확장자
            /// </summary>
            public const string Md5 = ".md5";

            /// <summary>
            /// Excel 파일 확장자
            /// </summary>
            public const string Excel = ".xlsx";

            /// <summary>
            /// XML 파일 확장자
            /// </summary>
            public const string Xml = ".xml";
        }

        /// <summary>
        /// 기본값 관련 상수
        /// </summary>
        public static class Defaults
        {
            /// <summary>
            /// 최대 파일 표시 개수
            /// </summary>
            public const int MaxFileDisplayCount = 5;

            /// <summary>
            /// 기본 타임아웃 (밀리초)
            /// </summary>
            public const int DefaultTimeout = 120000;

            /// <summary>
            /// 기본 날짜 형식
            /// </summary>
            public const string DefaultDateFormat = "yyyy-MM-dd";

            /// <summary>
            /// 기본 날짜/시간 형식
            /// </summary>
            public const string DefaultDateTimeFormat = "yyyy-MM-dd HH:mm:ss";
        }

        /// <summary>
        /// 특수 문자 관련 상수
        /// </summary>
        public static class SpecialCharacters
        {
            /// <summary>
            /// 줄바꿈 문자 (LF)
            /// </summary>
            public const string LineFeed = "\n";

            /// <summary>
            /// 캐리지 리턴 문자 (CR)
            /// </summary>
            public const string CarriageReturn = "\r";

            /// <summary>
            /// 줄바꿈 이스케이프 시퀀스
            /// </summary>
            public const string LineFeedEscape = "\\n";

            /// <summary>
            /// 캐리지 리턴 이스케이프 시퀀스
            /// </summary>
            public const string CarriageReturnEscape = "\\r";
        }
    }
}