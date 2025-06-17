using System;
using System.Collections.Generic;
using System.Linq;
using ExcelToYamlAddin.Domain.Common;

namespace ExcelToYamlAddin.Domain.ValueObjects
{
    /// <summary>
    /// YAML 생성을 위한 옵션을 정의하는 값 객체입니다.
    /// </summary>
    public class YamlGenerationOptions : ValueObject
    {
        /// <summary>
        /// 빈 필드를 출력할지 여부
        /// </summary>
        public bool ShowEmptyFields { get; private set; }

        /// <summary>
        /// 최대 처리 깊이
        /// </summary>
        public int MaxDepth { get; private set; }

        /// <summary>
        /// 들여쓰기 크기
        /// </summary>
        public int IndentSize { get; private set; }

        /// <summary>
        /// YAML 스타일 (Flow 또는 Block)
        /// </summary>
        public YamlStyle Style { get; private set; }

        /// <summary>
        /// 병합할 키 경로 목록
        /// </summary>
        public IReadOnlyList<string> MergeKeyPaths { get; private set; }

        /// <summary>
        /// 플로우 스타일을 적용할 경로 목록
        /// </summary>
        public IReadOnlyList<string> FlowStylePaths { get; private set; }

        /// <summary>
        /// 출력 파일 경로
        /// </summary>
        public string OutputPath { get; private set; }

        /// <summary>
        /// 사용자 정의 메타데이터
        /// </summary>
        public IReadOnlyDictionary<string, object> Metadata { get; private set; }

        /// <summary>
        /// 컬럼 필터링 옵션
        /// </summary>
        public ColumnFilterOptions ColumnFilter { get; private set; }

        /// <summary>
        /// 후처리 옵션
        /// </summary>
        public PostProcessingOptions PostProcessing { get; private set; }

        private YamlGenerationOptions()
        {
            // EF Core를 위한 빈 생성자
        }

        public YamlGenerationOptions(
            bool showEmptyFields = true,
            int maxDepth = 100,
            int indentSize = 2,
            YamlStyle style = null,
            IEnumerable<string> mergeKeyPaths = null,
            IEnumerable<string> flowStylePaths = null,
            string outputPath = null,
            IDictionary<string, object> metadata = null,
            ColumnFilterOptions columnFilter = null,
            PostProcessingOptions postProcessing = null)
        {
            if (maxDepth <= 0)
                throw new ArgumentException("최대 깊이는 0보다 커야 합니다.", nameof(maxDepth));
            
            if (indentSize <= 0)
                throw new ArgumentException("들여쓰기 크기는 0보다 커야 합니다.", nameof(indentSize));

            ShowEmptyFields = showEmptyFields;
            MaxDepth = maxDepth;
            IndentSize = indentSize;
            Style = style ?? YamlStyle.Block;
            MergeKeyPaths = mergeKeyPaths?.ToList() ?? new List<string>();
            FlowStylePaths = flowStylePaths?.ToList() ?? new List<string>();
            OutputPath = outputPath;
            if (metadata != null)
            {
                Metadata = new Dictionary<string, object>(metadata);
            }
            else
            {
                Metadata = new Dictionary<string, object>();
            }
            ColumnFilter = columnFilter ?? new ColumnFilterOptions();
            PostProcessing = postProcessing ?? PostProcessingOptions.Default();
        }

        /// <summary>
        /// 빌더 패턴을 위한 복사 메서드
        /// </summary>
        public YamlGenerationOptions With(
            bool? showEmptyFields = null,
            int? maxDepth = null,
            int? indentSize = null,
            YamlStyle style = null,
            IEnumerable<string> mergeKeyPaths = null,
            IEnumerable<string> flowStylePaths = null,
            string outputPath = null,
            IDictionary<string, object> metadata = null,
            ColumnFilterOptions columnFilter = null,
            PostProcessingOptions postProcessing = null)
        {
            return new YamlGenerationOptions(
                showEmptyFields ?? ShowEmptyFields,
                maxDepth ?? MaxDepth,
                indentSize ?? IndentSize,
                style ?? Style,
                mergeKeyPaths ?? MergeKeyPaths,
                flowStylePaths ?? FlowStylePaths,
                outputPath ?? OutputPath,
                metadata != null ? new Dictionary<string, object>(metadata) : Metadata.ToDictionary(kvp => kvp.Key, kvp => kvp.Value),
                columnFilter ?? ColumnFilter,
                postProcessing ?? PostProcessing
            );
        }

        /// <summary>
        /// 기본 옵션 생성
        /// </summary>
        public static YamlGenerationOptions Default => new YamlGenerationOptions();

        /// <summary>
        /// 설정 파일에서 옵션 생성
        /// </summary>
        public static YamlGenerationOptions FromConfig(ConversionOptions conversionOptions)
        {
            if (conversionOptions == null)
                return Default;

            return new YamlGenerationOptions(
                showEmptyFields: conversionOptions.IncludeEmptyFields,
                maxDepth: 100, // 기본값
                indentSize: 2, // 기본값
                style: conversionOptions.YamlStyle?.Style ?? YamlStyle.Block,
                mergeKeyPaths: conversionOptions.PostProcessing?.MergeKeyPaths,
                flowStylePaths: conversionOptions.PostProcessing?.FlowStylePaths?.Keys.ToList(),
                outputPath: null, // 출력 경로는 별도로 관리
                postProcessing: PostProcessingOptions.Create(
                    true, // enablePostProcessing
                    conversionOptions.PostProcessing?.EnableMergeByKey ?? false,
                    conversionOptions.PostProcessing?.ApplyFlowStyle ?? false,
                    conversionOptions.PostProcessing?.MergeKeyPaths ?? new List<string>(),
                    conversionOptions.PostProcessing?.FlowStylePaths ?? new Dictionary<string, string>()
                )
            );
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return ShowEmptyFields;
            yield return MaxDepth;
            yield return IndentSize;
            yield return Style;
            yield return string.Join(",", MergeKeyPaths);
            yield return string.Join(",", FlowStylePaths);
            yield return OutputPath;
            yield return ColumnFilter;
            yield return PostProcessing;
        }
    }

    /// <summary>
    /// 컬럼 필터링 옵션
    /// </summary>
    public class ColumnFilterOptions : ValueObject
    {
        /// <summary>
        /// 포함할 컬럼 목록 (null이면 모든 컬럼 포함)
        /// </summary>
        public IReadOnlyList<int> IncludedColumns { get; private set; }

        /// <summary>
        /// 제외할 컬럼 목록
        /// </summary>
        public IReadOnlyList<int> ExcludedColumns { get; private set; }

        public ColumnFilterOptions()
        {
            IncludedColumns = new List<int>();
            ExcludedColumns = new List<int>();
        }

        public ColumnFilterOptions(
            IEnumerable<int> includedColumns = null,
            IEnumerable<int> excludedColumns = null)
        {
            IncludedColumns = includedColumns?.ToList() ?? new List<int>();
            ExcludedColumns = excludedColumns?.ToList() ?? new List<int>();
        }

        /// <summary>
        /// 특정 컬럼이 포함되어야 하는지 확인
        /// </summary>
        public bool ShouldIncludeColumn(int column)
        {
            // 제외 목록에 있으면 제외
            if (ExcludedColumns.Contains(column))
                return false;

            // 포함 목록이 비어있으면 모든 컬럼 포함
            if (!IncludedColumns.Any())
                return true;

            // 포함 목록에 있는지 확인
            return IncludedColumns.Contains(column);
        }

        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return string.Join(",", IncludedColumns);
            yield return string.Join(",", ExcludedColumns);
        }
    }
}