using System;
using System.Text.RegularExpressions;

namespace ExcelToYamlAddin.Domain.Constants
{
    /// <summary>
    /// 정규식 패턴 상수를 정의하는 클래스
    /// </summary>
    public static class RegexPatterns
    {
        /// <summary>
        /// 시트명 검증 패턴
        /// </summary>
        public static class SheetName
        {
            /// <summary>
            /// 변환 대상 시트 패턴 (! 접두사로 시작)
            /// </summary>
            public const string ConversionSheetPattern = @"^!.*";

            /// <summary>
            /// 유효한 시트명 패턴 (특수문자 제외)
            /// </summary>
            public const string ValidSheetNamePattern = @"^[a-zA-Z0-9가-힣_\-]+$";

            /// <summary>
            /// 설정 시트 패턴
            /// </summary>
            public const string ConfigSheetPattern = @"^excel2yamlconfig$";
        }

        /// <summary>
        /// 파일 경로 관련 패턴
        /// </summary>
        public static class FilePath
        {
            /// <summary>
            /// Windows 파일 경로 패턴
            /// </summary>
            public const string WindowsPathPattern = @"^[a-zA-Z]:\\(?:[^\\/:*?""<>|\r\n]+\\)*[^\\/:*?""<>|\r\n]*$";

            /// <summary>
            /// 상대 경로 패턴
            /// </summary>
            public const string RelativePathPattern = @"^\.\.?[\\/].*$";

            /// <summary>
            /// 파일 확장자 추출 패턴
            /// </summary>
            public const string FileExtensionPattern = @"\.([^.]+)$";
        }

        /// <summary>
        /// 스키마 관련 패턴
        /// </summary>
        public static class Schema
        {
            /// <summary>
            /// 스키마 마커 패턴 ($로 시작하는 마커)
            /// </summary>
            public const string SchemaMarkerPattern = @"^\$.*";

            /// <summary>
            /// 배열 마커 패턴
            /// </summary>
            public const string ArrayMarkerPattern = @"\$\[\]";

            /// <summary>
            /// 객체 마커 패턴
            /// </summary>
            public const string ObjectMarkerPattern = @"\$\{\}";

            /// <summary>
            /// 동적 키 패턴
            /// </summary>
            public const string DynamicKeyPattern = @"\$key";

            /// <summary>
            /// 동적 값 패턴
            /// </summary>
            public const string DynamicValuePattern = @"\$value";

            /// <summary>
            /// 무시 마커 패턴
            /// </summary>
            public const string IgnoreMarkerPattern = @"^\^$";
        }

        /// <summary>
        /// 데이터 유형 검증 패턴
        /// </summary>
        public static class DataTypes
        {
            /// <summary>
            /// 정수 패턴
            /// </summary>
            public const string IntegerPattern = @"^-?\d+$";

            /// <summary>
            /// 실수 패턴
            /// </summary>
            public const string FloatPattern = @"^-?\d+\.\d+$";

            /// <summary>
            /// 부울 값 패턴
            /// </summary>
            public const string BooleanPattern = @"^(true|false|TRUE|FALSE|True|False)$";

            /// <summary>
            /// 날짜 패턴 (yyyy-MM-dd)
            /// </summary>
            public const string DatePattern = @"^\d{4}-\d{2}-\d{2}$";

            /// <summary>
            /// 날짜/시간 패턴 (yyyy-MM-dd HH:mm:ss)
            /// </summary>
            public const string DateTimePattern = @"^\d{4}-\d{2}-\d{2}\s\d{2}:\d{2}:\d{2}$";

            /// <summary>
            /// ISO 8601 날짜/시간 패턴
            /// </summary>
            public const string Iso8601Pattern = @"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{3})?(?:Z|[+-]\d{2}:\d{2})?$";
        }

        /// <summary>
        /// YAML 관련 패턴
        /// </summary>
        public static class Yaml
        {
            /// <summary>
            /// YAML 주석 패턴
            /// </summary>
            public const string CommentPattern = @"^\s*#.*$";

            /// <summary>
            /// YAML 키-값 패턴
            /// </summary>
            public const string KeyValuePattern = @"^(\s*)([^:]+):\s*(.*)$";

            /// <summary>
            /// YAML 배열 항목 패턴
            /// </summary>
            public const string ArrayItemPattern = @"^(\s*)-\s*(.*)$";

            /// <summary>
            /// YAML 멀티라인 문자열 시작 패턴 (|, >)
            /// </summary>
            public const string MultilineStringPattern = @"^\s*[|>]\s*$";

            /// <summary>
            /// YAML 플로우 스타일 감지 패턴
            /// </summary>
            public const string FlowStylePattern = @"[\[\]{}]";
        }

        /// <summary>
        /// XML 관련 패턴
        /// </summary>
        public static class Xml
        {
            /// <summary>
            /// XML 태그 패턴
            /// </summary>
            public const string TagPattern = @"<([^>]+)>([^<]*)</\1>";

            /// <summary>
            /// XML 속성 패턴
            /// </summary>
            public const string AttributePattern = @"(\w+)\s*=\s*[""']([^""']*)[""']";

            /// <summary>
            /// XML 특수 문자 패턴
            /// </summary>
            public const string SpecialCharPattern = @"[&<>""']";

            /// <summary>
            /// XML 네임스페이스 패턴
            /// </summary>
            public const string NamespacePattern = @"xmlns(?::\w+)?=";
        }

        /// <summary>
        /// 문자열 처리 관련 패턴
        /// </summary>
        public static class String
        {
            /// <summary>
            /// 앞뒤 공백 패턴
            /// </summary>
            public const string TrimPattern = @"^\s+|\s+$";

            /// <summary>
            /// 연속된 공백 패턴
            /// </summary>
            public const string MultipleSpacesPattern = @"\s{2,}";

            /// <summary>
            /// 줄바꿈 문자 패턴
            /// </summary>
            public const string NewlinePattern = @"\r\n|\r|\n";

            /// <summary>
            /// 이스케이프 시퀀스 패턴
            /// </summary>
            public const string EscapeSequencePattern = @"\\[\\""'nrtbf]";

            /// <summary>
            /// 유니코드 이스케이프 패턴
            /// </summary>
            public const string UnicodeEscapePattern = @"\\u[0-9a-fA-F]{4}";
        }

        /// <summary>
        /// 설정 값 관련 패턴
        /// </summary>
        public static class Configuration
        {
            /// <summary>
            /// 부울 설정 값 패턴 (Yes/No, True/False 등)
            /// </summary>
            public const string BooleanConfigPattern = @"^(yes|no|true|false|on|off|1|0)$";

            /// <summary>
            /// 경로 구분자 패턴 (쉼표, 세미콜론 등)
            /// </summary>
            public const string PathSeparatorPattern = @"[,;|]";

            /// <summary>
            /// 설정 키 패턴 (영문, 숫자, 언더스코어)
            /// </summary>
            public const string ConfigKeyPattern = @"^[a-zA-Z_]\w*$";
        }

        /// <summary>
        /// 미리 컴파일된 정규식 객체들
        /// </summary>
        public static class Compiled
        {
            /// <summary>
            /// 변환 대상 시트 정규식
            /// </summary>
            public static readonly Regex ConversionSheet = new Regex(SheetName.ConversionSheetPattern, RegexOptions.Compiled);

            /// <summary>
            /// 스키마 마커 정규식
            /// </summary>
            public static readonly Regex SchemaMarker = new Regex(Schema.SchemaMarkerPattern, RegexOptions.Compiled);

            /// <summary>
            /// 정수 검증 정규식
            /// </summary>
            public static readonly Regex Integer = new Regex(DataTypes.IntegerPattern, RegexOptions.Compiled);

            /// <summary>
            /// 실수 검증 정규식
            /// </summary>
            public static readonly Regex Float = new Regex(DataTypes.FloatPattern, RegexOptions.Compiled);

            /// <summary>
            /// 부울 값 검증 정규식
            /// </summary>
            public static readonly Regex Boolean = new Regex(DataTypes.BooleanPattern, RegexOptions.Compiled | RegexOptions.IgnoreCase);

            /// <summary>
            /// YAML 키-값 정규식
            /// </summary>
            public static readonly Regex YamlKeyValue = new Regex(Yaml.KeyValuePattern, RegexOptions.Compiled);

            /// <summary>
            /// 줄바꿈 문자 정규식
            /// </summary>
            public static readonly Regex Newline = new Regex(String.NewlinePattern, RegexOptions.Compiled);
        }
    }
}