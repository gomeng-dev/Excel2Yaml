using System;
using System.Collections.Generic;
using System.Linq;
using YamlDotNet.RepresentationModel;
using ExcelToYamlAddin.Logging;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    /// <summary>
    /// 중복 요소 분석 및 스키마 최적화를 담당하는 매니저
    /// </summary>
    public class DuplicateElementManager
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<DuplicateElementManager>();

        /// <summary>
        /// 배열의 중복 요소를 분석하여 최대 출현 횟수를 계산
        /// </summary>
        /// <param name="array">분석할 YAML 배열</param>
        /// <returns>요소 경로별 최대 출현 횟수</returns>
        public Dictionary<string, int> AnalyzeDuplicateElements(YamlSequenceNode array)
        {
            var duplicateCounts = new Dictionary<string, int>();

            // 각 배열 요소 분석
            foreach (var element in array.Children)
            {
                if (element is YamlMappingNode mapping)
                {
                    AnalyzeMappingDuplicates(mapping, "", duplicateCounts);
                }
            }

            // 결과 로깅
            foreach (var kvp in duplicateCounts)
            {
                Logger.Debug($"중복 요소 감지: {kvp.Key} = {kvp.Value}개");
            }

            return duplicateCounts;
        }

        /// <summary>
        /// 매핑 노드의 중복 요소를 재귀적으로 분석
        /// </summary>
        private void AnalyzeMappingDuplicates(YamlMappingNode mapping, string basePath, Dictionary<string, int> duplicateCounts)
        {
            foreach (var kvp in mapping.Children)
            {
                var key = kvp.Key.ToString();
                var path = string.IsNullOrEmpty(basePath) ? key : $"{basePath}.{key}";

                if (kvp.Value is YamlSequenceNode sequence)
                {
                    // 배열의 경우 요소 수 카운트
                    if (!duplicateCounts.ContainsKey(path))
                        duplicateCounts[path] = 0;

                    duplicateCounts[path] = Math.Max(duplicateCounts[path], sequence.Children.Count);

                    // 배열 요소가 매핑인 경우 재귀 분석
                    for (int i = 0; i < sequence.Children.Count; i++)
                    {
                        if (sequence.Children[i] is YamlMappingNode childMapping)
                        {
                            AnalyzeMappingDuplicates(childMapping, $"{path}[{i}]", duplicateCounts);
                        }
                    }
                }
                else if (kvp.Value is YamlMappingNode childMapping)
                {
                    // 중첩 객체의 경우 재귀 분석
                    AnalyzeMappingDuplicates(childMapping, path, duplicateCounts);
                }
            }
        }

        /// <summary>
        /// 실제 사용된 컬럼 수에 따라 스키마 병합 정보를 업데이트
        /// </summary>
        /// <param name="scheme">업데이트할 Excel 스키마</param>
        /// <param name="actualUsedColumns">실제 사용된 컬럼 수</param>
        /// <param name="lastSchemaRow">스키마의 마지막 행 번호</param>
        public void UpdateSchemaMerging(DynamicSchemaBuilder.ExcelScheme scheme, int actualUsedColumns, int lastSchemaRow)
        {
            Logger.Information($"스키마 병합 업데이트: 실제 사용 컬럼={actualUsedColumns}, 마지막 스키마 행={lastSchemaRow}");

            // $scheme_end 마커 업데이트
            scheme.UpdateSchemeEndMarker(actualUsedColumns);

            // 병합된 셀들의 범위 조정
            UpdateMergedCellRanges(scheme, actualUsedColumns);
        }

        /// <summary>
        /// 병합된 셀들의 범위를 실제 사용된 컬럼에 맞게 조정
        /// </summary>
        private void UpdateMergedCellRanges(DynamicSchemaBuilder.ExcelScheme scheme, int actualUsedColumns)
        {
            // 각 행의 병합 셀 정보 업데이트
            for (int row = 2; row <= scheme.LastSchemaRow; row++)
            {
                var mergedCells = scheme.GetMergedCellsInRow(row);
                
                foreach (var merged in mergedCells)
                {
                    // 루트 배열 마커나 객체 마커인 경우 전체 컬럼으로 확장
                    if (merged.StartColumn == 1 || merged.StartColumn == 2)
                    {
                        scheme.UpdateMergedCell(row, merged.StartColumn, actualUsedColumns);
                        Logger.Debug($"병합 셀 업데이트: 행={row}, 시작={merged.StartColumn}, 끝={actualUsedColumns}");
                    }
                }
            }
        }

        /// <summary>
        /// 가변 속성을 가진 배열 요소들을 분석
        /// </summary>
        public Dictionary<string, List<string>> AnalyzeVariableProperties(YamlNode root)
        {
            var variableProperties = new Dictionary<string, List<string>>();

            if (root is YamlSequenceNode rootArray)
            {
                foreach (var element in rootArray.Children)
                {
                    if (element is YamlMappingNode mapping)
                    {
                        AnalyzeVariablePropertiesInMapping(mapping, "", variableProperties);
                    }
                }
            }
            else if (root is YamlMappingNode rootMapping)
            {
                AnalyzeVariablePropertiesInMapping(rootMapping, "", variableProperties);
            }

            return variableProperties;
        }

        /// <summary>
        /// 매핑 내의 가변 속성 분석
        /// </summary>
        private void AnalyzeVariablePropertiesInMapping(
            YamlMappingNode mapping, 
            string basePath, 
            Dictionary<string, List<string>> variableProperties)
        {
            foreach (var kvp in mapping.Children)
            {
                var key = kvp.Key.ToString();
                var path = string.IsNullOrEmpty(basePath) ? key : $"{basePath}.{key}";

                if (kvp.Value is YamlSequenceNode sequence)
                {
                    // 배열의 각 요소가 가진 속성들을 수집
                    var allPropertiesInArray = new HashSet<string>();
                    var propertyOccurrences = new Dictionary<string, int>();

                    foreach (var element in sequence.Children)
                    {
                        if (element is YamlMappingNode elementMapping)
                        {
                            foreach (var prop in elementMapping.Children)
                            {
                                var propKey = prop.Key.ToString();
                                allPropertiesInArray.Add(propKey);
                                
                                if (!propertyOccurrences.ContainsKey(propKey))
                                    propertyOccurrences[propKey] = 0;
                                propertyOccurrences[propKey]++;
                            }
                        }
                    }

                    // 모든 요소에 나타나지 않는 속성이 있으면 가변 속성
                    var variableProps = allPropertiesInArray
                        .Where(p => propertyOccurrences[p] < sequence.Children.Count)
                        .ToList();

                    if (variableProps.Any())
                    {
                        variableProperties[path] = variableProps;
                        Logger.Debug($"가변 속성 감지: {path} - {string.Join(", ", variableProps)}");
                    }
                }
            }
        }

        /// <summary>
        /// 배열 요소의 최적 컬럼 배치를 계산
        /// </summary>
        public int CalculateOptimalColumnCount(
            string arrayPath,
            Dictionary<string, int> duplicateCounts,
            Dictionary<string, List<string>> variableProperties)
        {
            // 기본 컬럼 수
            int baseColumns = 4; // level, material, mineral, damage/addDamage

            // 가변 속성이 있는 경우 추가 컬럼 필요
            if (variableProperties.ContainsKey(arrayPath))
            {
                var varProps = variableProperties[arrayPath];
                // damage와 addDamage가 가변인 경우
                if (varProps.Contains("addDamage"))
                {
                    // 일부 요소는 4컬럼, 일부는 5컬럼 필요
                    return baseColumns + 1; // 최대 컬럼 수
                }
            }

            return baseColumns;
        }
    }
}