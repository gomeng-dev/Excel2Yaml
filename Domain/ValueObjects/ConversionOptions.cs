using System;
using System.Collections.Generic;
using System.Linq;
using ExcelToYamlAddin.Domain.Common;

namespace ExcelToYamlAddin.Domain.ValueObjects
{
    /// <summary>
    /// 변환 옵션을 나타내는 값 객체
    /// </summary>
    public class ConversionOptions : ValueObject
    {
        /// <summary>
        /// 출력 형식
        /// </summary>
        public OutputFormat OutputFormat { get; }

        /// <summary>
        /// 빈 필드 포함 여부
        /// </summary>
        public bool IncludeEmptyFields { get; }

        /// <summary>
        /// 빈 배열 필드 포함 여부
        /// </summary>
        public bool IncludeEmptyArrayFields { get; }

        /// <summary>
        /// 속성 순서 유지 여부
        /// </summary>
        public bool PreservePropertyOrder { get; }

        /// <summary>
        /// 날짜 형식
        /// </summary>
        public string DateFormat { get; }

        /// <summary>
        /// 날짜/시간 형식
        /// </summary>
        public string DateTimeFormat { get; }

        /// <summary>
        /// 숫자 형식
        /// </summary>
        public string NumberFormat { get; }

        /// <summary>
        /// YAML 스타일 설정
        /// </summary>
        public YamlStyleOptions YamlStyle { get; }

        /// <summary>
        /// 후처리 옵션
        /// </summary>
        public PostProcessingOptions PostProcessing { get; }

        /// <summary>
        /// 검증 옵션
        /// </summary>
        public ValidationOptions Validation { get; }

        private ConversionOptions()
        {
            // 기본값 설정
            OutputFormat = OutputFormat.Yaml;
            IncludeEmptyFields = false;
            IncludeEmptyArrayFields = false;
            PreservePropertyOrder = true;
            DateFormat = Constants.SchemeConstants.Defaults.DefaultDateFormat;
            DateTimeFormat = Constants.SchemeConstants.Defaults.DefaultDateTimeFormat;
            NumberFormat = "G";
            YamlStyle = YamlStyleOptions.Default();
            PostProcessing = PostProcessingOptions.Default();
            Validation = ValidationOptions.Default();
        }

        private ConversionOptions(
            OutputFormat outputFormat,
            bool includeEmptyFields,
            bool includeEmptyArrayFields,
            bool preservePropertyOrder,
            string dateFormat,
            string dateTimeFormat,
            string numberFormat,
            YamlStyleOptions yamlStyle,
            PostProcessingOptions postProcessing,
            ValidationOptions validation)
        {
            OutputFormat = outputFormat;
            IncludeEmptyFields = includeEmptyFields;
            IncludeEmptyArrayFields = includeEmptyArrayFields;
            PreservePropertyOrder = preservePropertyOrder;
            DateFormat = dateFormat ?? Constants.SchemeConstants.Defaults.DefaultDateFormat;
            DateTimeFormat = dateTimeFormat ?? Constants.SchemeConstants.Defaults.DefaultDateTimeFormat;
            NumberFormat = numberFormat ?? "G";
            YamlStyle = yamlStyle ?? YamlStyleOptions.Default();
            PostProcessing = postProcessing ?? PostProcessingOptions.Default();
            Validation = validation ?? ValidationOptions.Default();
        }

        /// <summary>
        /// 기본 옵션 생성
        /// </summary>
        public static ConversionOptions Default()
        {
            return new ConversionOptions();
        }

        /// <summary>
        /// 빌더를 통한 옵션 생성
        /// </summary>
        public static ConversionOptionsBuilder Builder()
        {
            return new ConversionOptionsBuilder();
        }

        /// <summary>
        /// 현재 옵션을 기반으로 빌더 생성
        /// </summary>
        public ConversionOptionsBuilder ToBuilder()
        {
            return new ConversionOptionsBuilder(this);
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return OutputFormat;
            yield return IncludeEmptyFields;
            yield return IncludeEmptyArrayFields;
            yield return PreservePropertyOrder;
            yield return DateFormat;
            yield return DateTimeFormat;
            yield return NumberFormat;
            yield return YamlStyle;
            yield return PostProcessing;
            yield return Validation;
        }

        /// <summary>
        /// ConversionOptions 빌더 클래스
        /// </summary>
        public class ConversionOptionsBuilder
        {
            private OutputFormat _outputFormat = OutputFormat.Yaml;
            private bool _includeEmptyFields = false;
            private bool _includeEmptyArrayFields = false;
            private bool _preservePropertyOrder = true;
            private string _dateFormat = Constants.SchemeConstants.Defaults.DefaultDateFormat;
            private string _dateTimeFormat = Constants.SchemeConstants.Defaults.DefaultDateTimeFormat;
            private string _numberFormat = "G";
            private YamlStyleOptions _yamlStyle = YamlStyleOptions.Default();
            private PostProcessingOptions _postProcessing = PostProcessingOptions.Default();
            private ValidationOptions _validation = ValidationOptions.Default();

            public ConversionOptionsBuilder()
            {
            }

            public ConversionOptionsBuilder(ConversionOptions options)
            {
                _outputFormat = options.OutputFormat;
                _includeEmptyFields = options.IncludeEmptyFields;
                _includeEmptyArrayFields = options.IncludeEmptyArrayFields;
                _preservePropertyOrder = options.PreservePropertyOrder;
                _dateFormat = options.DateFormat;
                _dateTimeFormat = options.DateTimeFormat;
                _numberFormat = options.NumberFormat;
                _yamlStyle = options.YamlStyle;
                _postProcessing = options.PostProcessing;
                _validation = options.Validation;
            }

            public ConversionOptionsBuilder WithOutputFormat(OutputFormat format)
            {
                _outputFormat = format;
                return this;
            }

            public ConversionOptionsBuilder WithIncludeEmptyFields(bool include)
            {
                _includeEmptyFields = include;
                return this;
            }

            public ConversionOptionsBuilder WithIncludeEmptyArrayFields(bool include)
            {
                _includeEmptyArrayFields = include;
                return this;
            }

            public ConversionOptionsBuilder WithPreservePropertyOrder(bool preserve)
            {
                _preservePropertyOrder = preserve;
                return this;
            }

            public ConversionOptionsBuilder WithDateFormat(string format)
            {
                _dateFormat = format;
                return this;
            }

            public ConversionOptionsBuilder WithDateTimeFormat(string format)
            {
                _dateTimeFormat = format;
                return this;
            }

            public ConversionOptionsBuilder WithNumberFormat(string format)
            {
                _numberFormat = format;
                return this;
            }

            public ConversionOptionsBuilder WithYamlStyle(YamlStyleOptions style)
            {
                _yamlStyle = style;
                return this;
            }

            public ConversionOptionsBuilder WithPostProcessing(PostProcessingOptions postProcessing)
            {
                _postProcessing = postProcessing;
                return this;
            }

            public ConversionOptionsBuilder WithValidation(ValidationOptions validation)
            {
                _validation = validation;
                return this;
            }

            public ConversionOptions Build()
            {
                return new ConversionOptions(
                    _outputFormat,
                    _includeEmptyFields,
                    _includeEmptyArrayFields,
                    _preservePropertyOrder,
                    _dateFormat,
                    _dateTimeFormat,
                    _numberFormat,
                    _yamlStyle,
                    _postProcessing,
                    _validation);
            }
        }
    }

    /// <summary>
    /// YAML 스타일 옵션
    /// </summary>
    public class YamlStyleOptions : ValueObject
    {
        public YamlStyle Style { get; }
        public int IndentSize { get; }
        public bool PreserveQuotes { get; }
        public bool UseFlowStyle { get; }

        private YamlStyleOptions()
        {
            Style = YamlStyle.Canonical;
            IndentSize = 2;
            PreserveQuotes = false;
            UseFlowStyle = false;
        }

        private YamlStyleOptions(YamlStyle style, int indentSize, bool preserveQuotes, bool useFlowStyle)
        {
            Style = style;
            IndentSize = indentSize > 0 ? indentSize : 2;
            PreserveQuotes = preserveQuotes;
            UseFlowStyle = useFlowStyle;
        }

        public static YamlStyleOptions Default() => new YamlStyleOptions();

        public static YamlStyleOptions Create(YamlStyle style, int indentSize, bool preserveQuotes, bool useFlowStyle)
        {
            return new YamlStyleOptions(style, indentSize, preserveQuotes, useFlowStyle);
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return Style;
            yield return IndentSize;
            yield return PreserveQuotes;
            yield return UseFlowStyle;
        }
    }

    /// <summary>
    /// 후처리 옵션
    /// </summary>
    public class PostProcessingOptions : ValueObject
    {
        public bool EnablePostProcessing { get; }
        public bool EnableMergeByKey { get; }
        public bool ApplyFlowStyle { get; }
        public List<string> MergeKeyPaths { get; }
        public Dictionary<string, string> FlowStylePaths { get; }

        private PostProcessingOptions()
        {
            EnablePostProcessing = true;
            EnableMergeByKey = false;
            ApplyFlowStyle = false;
            MergeKeyPaths = new List<string>();
            FlowStylePaths = new Dictionary<string, string>();
        }

        private PostProcessingOptions(
            bool enablePostProcessing,
            bool enableMergeByKey,
            bool applyFlowStyle,
            List<string> mergeKeyPaths,
            Dictionary<string, string> flowStylePaths)
        {
            EnablePostProcessing = enablePostProcessing;
            EnableMergeByKey = enableMergeByKey;
            ApplyFlowStyle = applyFlowStyle;
            MergeKeyPaths = mergeKeyPaths ?? new List<string>();
            FlowStylePaths = flowStylePaths ?? new Dictionary<string, string>();
        }

        public static PostProcessingOptions Default() => new PostProcessingOptions();

        public static PostProcessingOptions Create(
            bool enablePostProcessing,
            bool enableMergeByKey,
            bool applyFlowStyle,
            List<string> mergeKeyPaths,
            Dictionary<string, string> flowStylePaths)
        {
            return new PostProcessingOptions(
                enablePostProcessing,
                enableMergeByKey,
                applyFlowStyle,
                mergeKeyPaths,
                flowStylePaths);
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return EnablePostProcessing;
            yield return EnableMergeByKey;
            yield return ApplyFlowStyle;
            foreach (var path in MergeKeyPaths ?? Enumerable.Empty<string>())
                yield return path;
            foreach (var kvp in FlowStylePaths ?? new Dictionary<string, string>())
            {
                yield return kvp.Key;
                yield return kvp.Value;
            }
        }
    }

    /// <summary>
    /// 검증 옵션
    /// </summary>
    public class ValidationOptions : ValueObject
    {
        public bool ValidateOutput { get; }
        public bool StrictMode { get; }
        public int MaxDepth { get; }
        public int MaxArraySize { get; }
        public bool AllowDuplicateKeys { get; }

        private ValidationOptions()
        {
            ValidateOutput = true;
            StrictMode = false;
            MaxDepth = 100;
            MaxArraySize = 10000;
            AllowDuplicateKeys = false;
        }

        private ValidationOptions(
            bool validateOutput,
            bool strictMode,
            int maxDepth,
            int maxArraySize,
            bool allowDuplicateKeys)
        {
            ValidateOutput = validateOutput;
            StrictMode = strictMode;
            MaxDepth = maxDepth > 0 ? maxDepth : 100;
            MaxArraySize = maxArraySize > 0 ? maxArraySize : 10000;
            AllowDuplicateKeys = allowDuplicateKeys;
        }

        public static ValidationOptions Default() => new ValidationOptions();

        public static ValidationOptions Create(
            bool validateOutput,
            bool strictMode,
            int maxDepth,
            int maxArraySize,
            bool allowDuplicateKeys)
        {
            return new ValidationOptions(
                validateOutput,
                strictMode,
                maxDepth,
                maxArraySize,
                allowDuplicateKeys);
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return ValidateOutput;
            yield return StrictMode;
            yield return MaxDepth;
            yield return MaxArraySize;
            yield return AllowDuplicateKeys;
        }
    }
}