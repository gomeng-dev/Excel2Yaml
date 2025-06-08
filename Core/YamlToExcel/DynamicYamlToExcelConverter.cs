using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Logging;
using YamlDotNet.RepresentationModel;

namespace ExcelToYamlAddin.Core.YamlToExcel
{
    /// <summary>
    /// YAML 파일을 Excel 파일로 변환하는 메인 컨버터
    /// </summary>
    public class DynamicYamlToExcelConverter
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<DynamicYamlToExcelConverter>();

        private readonly DynamicStructureAnalyzer _analyzer;
        private readonly DynamicPatternRecognizer _recognizer;
        private readonly DynamicSchemaBuilder _schemaBuilder;
        private readonly DynamicDataMapper _dataMapper;
        private readonly DuplicateElementManager _duplicateManager;

        public DynamicYamlToExcelConverter()
        {
            _analyzer = new DynamicStructureAnalyzer();
            _recognizer = new DynamicPatternRecognizer();
            _schemaBuilder = new DynamicSchemaBuilder();
            _dataMapper = new DynamicDataMapper();
            _duplicateManager = new DuplicateElementManager();
        }

        /// <summary>
        /// YAML 파일을 Excel 파일로 변환
        /// </summary>
        /// <param name="yamlPath">입력 YAML 파일 경로</param>
        /// <param name="excelPath">출력 Excel 파일 경로</param>
        public void Convert(string yamlPath, string excelPath)
        {
            try
            {
                Logger.Information($"YAML to Excel 변환 시작: {yamlPath} -> {excelPath}");

                // 1. YAML 로드
                Logger.Debug("YAML 파일 로드 중...");
                var yamlContent = File.ReadAllText(yamlPath);
                var yaml = new YamlStream();
                yaml.Load(new StringReader(yamlContent));

                if (yaml.Documents.Count == 0)
                {
                    throw new InvalidOperationException("YAML 파일에 문서가 없습니다.");
                }

                var rootNode = yaml.Documents[0].RootNode;

                // 2. 구조 분석 (완전 동적)
                Logger.Debug("YAML 구조 분석 중...");
                var pattern = _analyzer.AnalyzeStructure(rootNode);
                Logger.Information($"구조 분석 완료: 타입={pattern.Type}, 속성 수={pattern.Properties.Count}, 배열 수={pattern.Arrays.Count}");

                // 3. 최적 전략 결정 (패턴 기반)
                Logger.Debug("레이아웃 전략 결정 중...");
                var strategy = _recognizer.DetermineStrategy(pattern);
                Logger.Information($"레이아웃 전략 결정: {strategy}");

                // 4. 레이아웃 정보 생성
                Logger.Debug("레이아웃 정보 생성 중...");
                var layoutInfo = GenerateLayoutInfo(rootNode, pattern, strategy);

                // 5. Excel 스키마 생성
                Logger.Debug("Excel 스키마 생성 중...");
                var scheme = _schemaBuilder.BuildScheme(pattern, strategy, layoutInfo);
                
                // 스키마 매핑 상황 디버깅
                scheme.DebugAllMappings();

                // 6. 중복 요소 분석 및 스키마 최적화
                if (rootNode is YamlSequenceNode rootSequence)
                {
                    Logger.Debug("중복 요소 분석 중...");
                    var duplicateCounts = _duplicateManager.AnalyzeDuplicateElements(rootSequence);
                    scheme.OptimizeForDuplicates(duplicateCounts);
                }

                // 7. 데이터 매핑
                Logger.Debug("데이터 매핑 중...");
                var rows = _dataMapper.MapToExcelRows(rootNode, scheme, pattern);
                Logger.Information($"데이터 매핑 완료: {rows.Count}개 행");

                // 8. 실제 사용 컬럼 계산 및 스키마 업데이트
                var actualUsedColumns = scheme.CalculateActualUsedColumns(rows);
                Logger.Debug($"실제 사용 컬럼 수: {actualUsedColumns}");
                _duplicateManager.UpdateSchemaMerging(scheme, actualUsedColumns, scheme.LastSchemaRow);

                // 9. Excel 파일 작성
                Logger.Debug("Excel 파일 작성 중...");
                WriteExcel(scheme, rows, excelPath);

                Logger.Information($"변환 완료: {excelPath}");
            }
            catch (Exception ex)
            {
                Logger.Error($"변환 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// YAML 내용을 직접 Excel로 변환
        /// </summary>
        /// <param name="yamlContent">YAML 내용</param>
        /// <param name="excelPath">출력 Excel 파일 경로</param>
        public void ConvertFromContent(string yamlContent, string excelPath)
        {
            try
            {
                Logger.Information($"YAML 내용을 Excel로 변환: -> {excelPath}");

                // 1. YAML 로드
                var yaml = new YamlStream();
                yaml.Load(new StringReader(yamlContent));

                if (yaml.Documents.Count == 0)
                {
                    throw new InvalidOperationException("YAML 내용에 문서가 없습니다.");
                }

                var rootNode = yaml.Documents[0].RootNode;

                // 이후 과정은 Convert 메서드와 동일
                var pattern = _analyzer.AnalyzeStructure(rootNode);
                var strategy = _recognizer.DetermineStrategy(pattern);
                var layoutInfo = GenerateLayoutInfo(rootNode, pattern, strategy);
                var scheme = _schemaBuilder.BuildScheme(pattern, strategy, layoutInfo);
                
                // 중복 요소 분석
                if (rootNode is YamlSequenceNode rootSequence)
                {
                    var duplicateCounts = _duplicateManager.AnalyzeDuplicateElements(rootSequence);
                    scheme.OptimizeForDuplicates(duplicateCounts);
                }
                
                var rows = _dataMapper.MapToExcelRows(rootNode, scheme, pattern);
                
                // rows가 null이거나 비어있는지 확인
                if (rows == null)
                {
                    Logger.Warning("MapToExcelRows가 null을 반환했습니다.");
                    rows = new List<DynamicDataMapper.ExcelRow>();
                }
                else if (rows.Count == 0)
                {
                    Logger.Warning("MapToExcelRows가 빈 리스트를 반환했습니다.");
                }
                else
                {
                    Logger.Information($"데이터 매핑 완료: {rows.Count}개 행");
                }
                
                // 실제 사용 컬럼 계산 및 스키마 업데이트
                var actualUsedColumns = scheme.CalculateActualUsedColumns(rows);
                _duplicateManager.UpdateSchemaMerging(scheme, actualUsedColumns, scheme.LastSchemaRow);
                
                WriteExcel(scheme, rows, excelPath);

                Logger.Information($"변환 완료: {excelPath}");
            }
            catch (Exception ex)
            {
                Logger.Error($"변환 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        /// <summary>
        /// YAML을 Excel 워크북으로 변환 (파일로 저장하지 않고 반환)
        /// </summary>
        /// <param name="yamlContent">YAML 내용</param>
        /// <param name="sheetName">시트 이름</param>
        /// <returns>생성된 워크북</returns>
        public IXLWorkbook ConvertToWorkbook(string yamlContent, string sheetName = "Sheet1")
        {
            try
            {
                Logger.Information($"YAML을 워크북으로 변환");

                // 1. YAML 로드
                var yaml = new YamlStream();
                yaml.Load(new StringReader(yamlContent));

                if (yaml.Documents.Count == 0)
                {
                    throw new InvalidOperationException("YAML 내용에 문서가 없습니다.");
                }

                var rootNode = yaml.Documents[0].RootNode;

                // 구조 분석 및 변환
                var pattern = _analyzer.AnalyzeStructure(rootNode);
                var strategy = _recognizer.DetermineStrategy(pattern);
                var layoutInfo = GenerateLayoutInfo(rootNode, pattern, strategy);
                var scheme = _schemaBuilder.BuildScheme(pattern, strategy, layoutInfo);
                
                // 중복 요소 분석
                if (rootNode is YamlSequenceNode rootSequence)
                {
                    var duplicateCounts = _duplicateManager.AnalyzeDuplicateElements(rootSequence);
                    scheme.OptimizeForDuplicates(duplicateCounts);
                }
                
                var rows = _dataMapper.MapToExcelRows(rootNode, scheme, pattern);
                
                // rows가 null이거나 비어있는지 확인
                if (rows == null)
                {
                    Logger.Warning("MapToExcelRows가 null을 반환했습니다.");
                    rows = new List<DynamicDataMapper.ExcelRow>();
                }
                else if (rows.Count == 0)
                {
                    Logger.Warning("MapToExcelRows가 빈 리스트를 반환했습니다.");
                }
                else
                {
                    Logger.Information($"데이터 매핑 완료: {rows.Count}개 행");
                }
                
                // 실제 사용 컬럼 계산 및 스키마 업데이트
                var actualUsedColumns = scheme.CalculateActualUsedColumns(rows);
                _duplicateManager.UpdateSchemaMerging(scheme, actualUsedColumns, scheme.LastSchemaRow);

                // 워크북 생성
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add(sheetName);

                // 스키마 작성
                scheme.WriteToWorksheet(worksheet);

                // 데이터 작성
                int dataStartRow = scheme.LastSchemaRow + 1;
                foreach (var row in rows)
                {
                    row.WriteToWorksheet(worksheet, dataStartRow++);
                }

                // 자동 너비 조정
                worksheet.Columns().AdjustToContents();

                return workbook;
            }
            catch (Exception ex)
            {
                Logger.Error($"워크북 변환 중 오류 발생: {ex.Message}", ex);
                throw;
            }
        }

        private dynamic GenerateLayoutInfo(
            YamlNode root,
            DynamicStructureAnalyzer.StructurePattern pattern,
            DynamicPatternRecognizer.LayoutStrategy strategy)
        {
            switch (strategy)
            {
                case DynamicPatternRecognizer.LayoutStrategy.HorizontalExpansion:
                    return GenerateHorizontalLayout(root, pattern);
                case DynamicPatternRecognizer.LayoutStrategy.VerticalNesting:
                    return GenerateVerticalLayout(root, pattern);
                case DynamicPatternRecognizer.LayoutStrategy.Mixed:
                    return GenerateMixedLayout(root, pattern);
                default:
                    return null;
            }
        }

        private DynamicHorizontalExpander.HorizontalLayout GenerateHorizontalLayout(
            YamlNode root,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            var expander = new DynamicHorizontalExpander();
            
            if (root is YamlSequenceNode rootSequence)
            {
                // 루트가 배열인 경우, 전체 루트 요소를 분석
                var layout = expander.AnalyzeHorizontalLayout(rootSequence, pattern);
                
                // 배열 속성들에 대해 통합 스키마 전달
                foreach (var arrayName in pattern.Arrays.Keys)
                {
                    if (layout.ArrayLayouts.ContainsKey(arrayName))
                    {
                        // 이미 AnalyzeHorizontalLayout에서 통합 분석됨
                        Logger.Debug($"배열 '{arrayName}'은 이미 통합 분석 완료");
                    }
                }
                
                return layout;
            }
            else if (root is YamlMappingNode rootMapping)
            {
                // 루트가 객체인 경우, 배열 속성들을 찾아서 처리
                var layout = new DynamicHorizontalExpander.HorizontalLayout();
                
                foreach (var kvp in rootMapping.Children)
                {
                    var key = kvp.Key.ToString();
                    if (kvp.Value is YamlSequenceNode arrayNode && pattern.Arrays.ContainsKey(key))
                    {
                        var arrayLayout = expander.CalculateArrayLayout(key, arrayNode, pattern.Arrays[key].ElementProperties);
                        layout.ArrayLayouts[key] = arrayLayout;
                    }
                }
                
                return layout;
            }

            return new DynamicHorizontalExpander.HorizontalLayout();
        }

        private DynamicVerticalNester.VerticalLayout GenerateVerticalLayout(
            YamlNode root,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            var nester = new DynamicVerticalNester();
            
            if (root is YamlSequenceNode rootSequence)
            {
                var items = rootSequence.Children.ToList();
                return nester.GenerateVerticalLayout(pattern, items);
            }

            return new DynamicVerticalNester.VerticalLayout();
        }

        private dynamic GenerateMixedLayout(
            YamlNode root,
            DynamicStructureAnalyzer.StructurePattern pattern)
        {
            // 혼합 레이아웃은 주로 수평 확장을 기본으로 사용
            return GenerateHorizontalLayout(root, pattern);
        }

        private void WriteExcel(
            DynamicSchemaBuilder.ExcelScheme scheme,
            List<DynamicDataMapper.ExcelRow> rows,
            string path)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // 스키마 작성
                scheme.WriteToWorksheet(worksheet);

                // 데이터 작성
                int dataStartRow = scheme.LastSchemaRow + 1;
                foreach (var row in rows)
                {
                    row.WriteToWorksheet(worksheet, dataStartRow++);
                }

                // 자동 너비 조정
                worksheet.Columns().AdjustToContents();

                // 스타일 적용
                ApplyStyles(worksheet, scheme.LastSchemaRow);

                // 디렉토리 생성
                var directory = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                workbook.SaveAs(path);
            }
        }

        private void ApplyStyles(IXLWorksheet worksheet, int schemaEndRow)
        {
            // 스키마 영역 스타일
            var schemaRange = worksheet.Range(1, 1, schemaEndRow, worksheet.LastColumnUsed().ColumnNumber());
            schemaRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            schemaRange.Style.Font.Bold = true;
            schemaRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            schemaRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

            // 데이터 영역 테두리
            if (worksheet.LastRowUsed().RowNumber() > schemaEndRow + 1)
            {
                var dataRange = worksheet.Range(
                    schemaEndRow + 2, 1,
                    worksheet.LastRowUsed().RowNumber(),
                    worksheet.LastColumnUsed().ColumnNumber());
                dataRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                dataRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }
        }
    }
}