# YAML to Excel 동적 변환 설계서

## 1. 개요

### 1.1 목적
본 문서는 YAML 파일을 Excel 스키마로 자동 변환하는 완전히 동적인 알고리즘을 설계합니다. 모든 하드코딩을 제거하고, 어떤 YAML 구조든 자동으로 분석하여 최적의 Excel 레이아웃을 생성합니다.

### 1.2 핵심 원칙
- **No Hardcoding**: 속성명, 파일 패턴, 배열 크기 등 어떤 하드코딩도 없음
- **Dynamic Analysis**: YAML 구조를 동적으로 분석하여 패턴 인식
- **Adaptive Layout**: 구조에 따라 최적의 레이아웃 전략 자동 선택
- **Pattern Recognition**: 파일명이 아닌 구조 패턴으로 처리 방식 결정

## 2. 시스템 아키텍처

```
YAML 파일 → DynamicParser → StructureAnalyzer → PatternRecognizer → LayoutOptimizer → ExcelWriter
                                    ↓                    ↓
                            TypeInferenceEngine    StrategySelector
```

### 2.1 주요 컴포넌트

#### 2.1.1 DynamicParser
- YamlDotNet을 사용한 YAML 파싱
- 구조를 손실 없이 보존하는 노드 트리 생성

#### 2.1.2 StructureAnalyzer
- 재귀적 구조 분석
- 깊이, 너비, 패턴 자동 감지
- 동적 스키마 추출

#### 2.1.3 PatternRecognizer
- 구조 패턴 인식 (단순/중첩/가변/고정)
- 반복 패턴 감지
- 속성 출현 빈도 분석

#### 2.1.4 LayoutOptimizer
- 분석된 패턴에 따른 최적 레이아웃 결정
- 수평/수직 확장 전략 자동 선택
- 컬럼 공간 최적화

## 3. 동적 구조 분석 알고리즘

### 3.1 구조 분석기

```csharp
public class DynamicStructureAnalyzer
{
    public class StructurePattern
    {
        public PatternType Type { get; set; }
        public Dictionary<string, PropertyPattern> Properties { get; set; }
        public Dictionary<string, ArrayPattern> Arrays { get; set; }
        public int MaxDepth { get; set; }
        public double ConsistencyScore { get; set; }
    }

    public StructurePattern AnalyzeStructure(YamlNode root)
    {
        var pattern = new StructurePattern
        {
            Properties = new Dictionary<string, PropertyPattern>(),
            Arrays = new Dictionary<string, ArrayPattern>()
        };

        // 동적 분석 - 하드코딩 없음
        if (root is YamlSequenceNode sequence)
        {
            pattern.Type = PatternType.RootArray;
            AnalyzeArrayElements(sequence, pattern);
        }
        else if (root is YamlMappingNode mapping)
        {
            pattern.Type = PatternType.RootObject;
            AnalyzeObjectProperties(mapping, pattern);
        }

        pattern.ConsistencyScore = CalculateConsistencyScore(pattern);
        pattern.MaxDepth = CalculateMaxDepth(root);

        return pattern;
    }

    private void AnalyzeArrayElements(YamlSequenceNode array, StructurePattern pattern)
    {
        var elementSchemas = new List<Dictionary<string, object>>();
        
        // 모든 배열 요소 분석
        foreach (var element in array.Children)
        {
            var schema = ExtractElementSchema(element);
            elementSchemas.Add(schema);
        }

        // 동적 패턴 인식
        pattern.Properties = UnifySchemas(elementSchemas);
        pattern.Arrays = DetectNestedArrays(elementSchemas);
    }

    private Dictionary<string, PropertyPattern> UnifySchemas(List<Dictionary<string, object>> schemas)
    {
        var unified = new Dictionary<string, PropertyPattern>();
        
        // 모든 속성 수집 및 분석
        foreach (var schema in schemas)
        {
            foreach (var prop in schema)
            {
                if (!unified.ContainsKey(prop.Key))
                {
                    unified[prop.Key] = new PropertyPattern
                    {
                        Name = prop.Key,
                        OccurrenceCount = 0,
                        Types = new HashSet<Type>(),
                        FirstAppearanceIndex = schemas.IndexOf(schema)
                    };
                }
                
                unified[prop.Key].OccurrenceCount++;
                unified[prop.Key].Types.Add(prop.Value?.GetType() ?? typeof(object));
            }
        }

        // 출현 비율 계산
        foreach (var prop in unified.Values)
        {
            prop.OccurrenceRatio = (double)prop.OccurrenceCount / schemas.Count;
            prop.IsRequired = prop.OccurrenceRatio > 0.8; // 80% 이상 출현시 필수
        }

        return unified;
    }
}
```

### 3.2 패턴 인식기

```csharp
public class DynamicPatternRecognizer
{
    public enum LayoutStrategy
    {
        Simple,              // 단순 구조
        VerticalNesting,     // 수직 중첩 (각 요소가 여러 행)
        HorizontalExpansion, // 수평 확장 (배열을 컬럼으로)
        Mixed               // 혼합 전략
    }

    public LayoutStrategy DetermineStrategy(StructurePattern pattern)
    {
        // 동적 전략 결정 - 하드코딩 없음
        var metrics = CalculateMetrics(pattern);
        
        if (metrics.IsSimpleStructure)
        {
            return LayoutStrategy.Simple;
        }
        
        if (metrics.HasLargeNestedArrays && metrics.ArrayElementCount > 5)
        {
            return LayoutStrategy.HorizontalExpansion;
        }
        
        if (metrics.HasVariableDepth || metrics.HasOptionalNesting)
        {
            return LayoutStrategy.VerticalNesting;
        }
        
        return LayoutStrategy.Mixed;
    }

    private StructureMetrics CalculateMetrics(StructurePattern pattern)
    {
        return new StructureMetrics
        {
            IsSimpleStructure = pattern.MaxDepth <= 2 && !pattern.Arrays.Any(),
            HasLargeNestedArrays = pattern.Arrays.Any(a => a.Value.MaxSize > 3),
            ArrayElementCount = pattern.Arrays.Sum(a => a.Value.MaxSize),
            HasVariableDepth = pattern.Properties.Any(p => p.Value.OccurrenceRatio < 0.5),
            HasOptionalNesting = pattern.Arrays.Any(a => a.Value.OccurrenceRatio < 1.0)
        };
    }
}
```

### 3.3 동적 속성 순서 결정

```csharp
public class DynamicPropertyOrderer
{
    public List<string> DeterminePropertyOrder(Dictionary<string, PropertyPattern> properties)
    {
        // 동적 우선순위 계산 - 하드코딩 없음
        return properties
            .OrderByDescending(p => p.Value.OccurrenceRatio)  // 출현 빈도
            .ThenBy(p => p.Value.FirstAppearanceIndex)        // 첫 등장 순서
            .ThenBy(p => p.Key.Length)                        // 이름 길이 (짧은 것 우선)
            .ThenBy(p => p.Key)                               // 알파벳 순
            .Select(p => p.Key)
            .ToList();
    }

    public List<string> OptimizeForHorizontalLayout(
        Dictionary<string, PropertyPattern> properties,
        List<Dictionary<string, object>> samples)
    {
        // 샘플 데이터 분석을 통한 최적화
        var groupings = AnalyzePropertyGroupings(samples);
        var ordered = new List<string>();

        // 함께 나타나는 속성들을 그룹화
        foreach (var group in groupings.OrderByDescending(g => g.Strength))
        {
            var groupProperties = group.Properties
                .OrderByDescending(p => properties[p].OccurrenceRatio)
                .ToList();
            ordered.AddRange(groupProperties);
        }

        // 그룹에 속하지 않은 속성들 추가
        var ungrouped = properties.Keys.Except(ordered);
        ordered.AddRange(DeterminePropertyOrder(
            properties.Where(p => ungrouped.Contains(p.Key))
                     .ToDictionary(p => p.Key, p => p.Value)));

        return ordered;
    }

    private List<PropertyGrouping> AnalyzePropertyGroupings(
        List<Dictionary<string, object>> samples)
    {
        var cooccurrence = new Dictionary<(string, string), int>();
        
        // 속성 동시 출현 분석
        foreach (var sample in samples)
        {
            var props = sample.Keys.ToList();
            for (int i = 0; i < props.Count; i++)
            {
                for (int j = i + 1; j < props.Count; j++)
                {
                    var pair = (props[i], props[j]);
                    if (!cooccurrence.ContainsKey(pair))
                        cooccurrence[pair] = 0;
                    cooccurrence[pair]++;
                }
            }
        }

        // 강한 연관성을 가진 속성 그룹 생성
        return CreatePropertyGroups(cooccurrence, samples.Count);
    }
}
```

## 4. 동적 레이아웃 생성

### 4.1 수평 확장 알고리즘

```csharp
public class DynamicHorizontalExpander
{
    public class DynamicArrayLayout
    {
        public string ArrayPath { get; set; }
        public int ElementCount { get; set; }
        public List<ElementLayout> Elements { get; set; }
        public int TotalColumns { get; set; }
    }

    public DynamicArrayLayout CalculateArrayLayout(
        string arrayPath, 
        YamlSequenceNode array,
        Dictionary<string, PropertyPattern> unifiedSchema)
    {
        var layout = new DynamicArrayLayout
        {
            ArrayPath = arrayPath,
            ElementCount = array.Children.Count,
            Elements = new List<ElementLayout>()
        };

        // 각 요소별 레이아웃 계산
        for (int i = 0; i < array.Children.Count; i++)
        {
            if (array.Children[i] is YamlMappingNode element)
            {
                var elementLayout = new ElementLayout
                {
                    Index = i,
                    Properties = element.Children.Keys
                        .Select(k => k.ToString())
                        .ToList(),
                    RequiredColumns = 0
                };

                // 동적으로 필요한 컬럼 수 계산
                elementLayout.RequiredColumns = CalculateRequiredColumns(
                    elementLayout.Properties, 
                    unifiedSchema);
                
                layout.Elements.Add(elementLayout);
            }
        }

        // 전체 컬럼 수 계산
        layout.TotalColumns = layout.Elements.Sum(e => e.RequiredColumns);

        return layout;
    }

    private int CalculateRequiredColumns(
        List<string> properties, 
        Dictionary<string, PropertyPattern> schema)
    {
        // 스키마에 있는 모든 속성 중 이 요소가 가진 속성 수
        var schemaProperties = schema.Keys.Where(properties.Contains).Count();
        
        // 스키마에 없는 추가 속성
        var extraProperties = properties.Except(schema.Keys).Count();
        
        return schemaProperties + extraProperties;
    }
}
```

### 4.2 수직 중첩 알고리즘

```csharp
public class DynamicVerticalNester
{
    public class VerticalLayout
    {
        public List<RowGroup> RowGroups { get; set; }
        public Dictionary<string, int> ColumnMapping { get; set; }
        public bool RequiresMerging { get; set; }
        public string MergeKey { get; set; }
    }

    public VerticalLayout GenerateVerticalLayout(
        StructurePattern pattern,
        List<YamlNode> items)
    {
        var layout = new VerticalLayout
        {
            RowGroups = new List<RowGroup>(),
            ColumnMapping = new Dictionary<string, int>()
        };

        // 병합 키 자동 감지
        layout.MergeKey = DetectMergeKey(items);
        layout.RequiresMerging = !string.IsNullOrEmpty(layout.MergeKey);

        // 컬럼 매핑 생성
        int currentColumn = 1;
        foreach (var prop in pattern.Properties.Keys)
        {
            layout.ColumnMapping[prop] = currentColumn++;
        }

        // 행 그룹 생성
        if (layout.RequiresMerging)
        {
            CreateMergedRowGroups(items, layout);
        }
        else
        {
            CreateSimpleRowGroups(items, layout);
        }

        return layout;
    }

    private string DetectMergeKey(List<YamlNode> items)
    {
        // 중복 값을 가진 속성 찾기
        var propertyValues = new Dictionary<string, List<object>>();
        
        foreach (var item in items.OfType<YamlMappingNode>())
        {
            foreach (var kvp in item.Children)
            {
                var key = kvp.Key.ToString();
                var value = kvp.Value.ToString();
                
                if (!propertyValues.ContainsKey(key))
                    propertyValues[key] = new List<object>();
                
                propertyValues[key].Add(value);
            }
        }

        // 중복 값이 있는 속성 찾기
        foreach (var prop in propertyValues)
        {
            var uniqueValues = prop.Value.Distinct().Count();
            if (uniqueValues < prop.Value.Count && uniqueValues < items.Count * 0.5)
            {
                return prop.Key; // 병합 키 후보
            }
        }

        return null;
    }
}
```

## 5. Excel 스키마 생성

### 5.1 동적 스키마 빌더

```csharp
public class DynamicSchemaBuilder
{
    public ExcelScheme BuildScheme(
        StructurePattern pattern,
        LayoutStrategy strategy,
        dynamic layoutInfo)
    {
        var scheme = new ExcelScheme();
        int currentRow = 2; // 2행부터 시작

        switch (strategy)
        {
            case LayoutStrategy.Simple:
                BuildSimpleScheme(scheme, pattern, currentRow);
                break;
                
            case LayoutStrategy.HorizontalExpansion:
                BuildHorizontalScheme(scheme, pattern, layoutInfo, currentRow);
                break;
                
            case LayoutStrategy.VerticalNesting:
                BuildVerticalScheme(scheme, pattern, layoutInfo, currentRow);
                break;
                
            case LayoutStrategy.Mixed:
                BuildMixedScheme(scheme, pattern, layoutInfo, currentRow);
                break;
        }

        // $scheme_end 추가
        scheme.AddSchemeEndRow(scheme.LastSchemaRow + 1);

        return scheme;
    }

    private void BuildHorizontalScheme(
        ExcelScheme scheme,
        StructurePattern pattern,
        DynamicArrayLayout arrayLayout,
        int startRow)
    {
        // 루트 배열 마커
        scheme.AddCell(startRow, 1, "$[]");
        
        // ^ 마커와 기본 속성들
        int row = startRow + 1;
        scheme.AddCell(row, 1, "^");
        
        int col = 2;
        var orderer = new DynamicPropertyOrderer();
        
        // 단순 속성들
        var simpleProps = pattern.Properties
            .Where(p => !p.Value.IsArray)
            .ToDictionary(p => p.Key, p => p.Value);
        
        foreach (var prop in orderer.DeterminePropertyOrder(simpleProps))
        {
            scheme.AddCell(row, col++, prop);
        }

        // 배열 속성 처리
        foreach (var array in pattern.Arrays)
        {
            var arrayStartCol = col;
            var arrayEndCol = col + arrayLayout.TotalColumns - 1;
            
            // 배열 마커 (병합)
            scheme.AddMergedCell(row, arrayStartCol, arrayEndCol, $"{array.Key}$[]");
            
            // 각 요소별 처리
            BuildArrayElementScheme(scheme, arrayLayout, row + 1, arrayStartCol);
            
            col = arrayEndCol + 1;
        }
    }

    private void BuildArrayElementScheme(
        ExcelScheme scheme,
        DynamicArrayLayout layout,
        int startRow,
        int startCol)
    {
        int currentCol = startCol;
        
        // 각 배열 요소에 대한 스키마
        foreach (var element in layout.Elements)
        {
            // ${} 마커
            scheme.AddMergedCell(startRow, currentCol, 
                currentCol + element.RequiredColumns - 1, "${}");
            
            // 속성들
            int propCol = currentCol;
            foreach (var prop in element.Properties)
            {
                scheme.AddCell(startRow + 1, propCol++, prop);
            }
            
            currentCol += element.RequiredColumns;
        }
    }
}
```

## 6. 데이터 변환

### 6.1 동적 데이터 매퍼

```csharp
public class DynamicDataMapper
{
    public List<ExcelRow> MapToExcelRows(
        YamlNode data,
        ExcelScheme scheme,
        StructurePattern pattern)
    {
        var rows = new List<ExcelRow>();
        
        if (data is YamlSequenceNode sequence)
        {
            foreach (var item in sequence.Children)
            {
                var mappedRows = MapItem(item, scheme, pattern);
                rows.AddRange(mappedRows);
            }
        }
        else if (data is YamlMappingNode mapping)
        {
            var row = MapSingleItem(mapping, scheme, pattern);
            rows.Add(row);
        }

        return rows;
    }

    private List<ExcelRow> MapItem(
        YamlNode item,
        ExcelScheme scheme,
        StructurePattern pattern)
    {
        var rows = new List<ExcelRow>();
        
        if (pattern.Arrays.Any(a => a.Value.RequiresMultipleRows))
        {
            // 수직 확장 필요
            rows.AddRange(ExpandVertically(item, scheme, pattern));
        }
        else
        {
            // 단일 행 매핑
            rows.Add(MapHorizontally(item, scheme, pattern));
        }

        return rows;
    }

    private ExcelRow MapHorizontally(
        YamlNode item,
        ExcelScheme scheme,
        StructurePattern pattern)
    {
        var row = new ExcelRow();
        
        if (item is YamlMappingNode mapping)
        {
            // ^ 마커
            row.SetCell(1, "^");
            
            // 속성 매핑
            foreach (var prop in mapping.Children)
            {
                var key = prop.Key.ToString();
                var columnIndex = scheme.GetColumnIndex(key);
                
                if (columnIndex > 0)
                {
                    var value = ConvertValue(prop.Value);
                    row.SetCell(columnIndex, value);
                }
                
                // 중첩 배열 처리
                if (prop.Value is YamlSequenceNode nestedArray)
                {
                    MapNestedArray(row, nestedArray, key, scheme, pattern);
                }
            }
        }

        return row;
    }

    private void MapNestedArray(
        ExcelRow row,
        YamlSequenceNode array,
        string arrayName,
        ExcelScheme scheme,
        StructurePattern pattern)
    {
        var arrayPattern = pattern.Arrays[arrayName];
        var startColumn = scheme.GetArrayStartColumn(arrayName);
        
        int currentCol = startColumn;
        foreach (var element in array.Children)
        {
            if (element is YamlMappingNode elementMapping)
            {
                foreach (var prop in elementMapping.Children)
                {
                    var value = ConvertValue(prop.Value);
                    row.SetCell(currentCol++, value);
                }
            }
        }
    }

    private object ConvertValue(YamlNode node)
    {
        return node switch
        {
            YamlScalarNode scalar => ConvertScalar(scalar),
            YamlSequenceNode => "[Array]",
            YamlMappingNode => "[Object]",
            _ => null
        };
    }

    private object ConvertScalar(YamlScalarNode scalar)
    {
        var value = scalar.Value;
        
        // 동적 타입 추론
        if (bool.TryParse(value, out bool boolResult))
            return boolResult;
            
        if (int.TryParse(value, out int intResult))
            return intResult;
            
        if (double.TryParse(value, NumberStyles.Any, 
            CultureInfo.InvariantCulture, out double doubleResult))
            return doubleResult;
            
        if (DateTime.TryParse(value, out DateTime dateResult))
            return dateResult;
            
        return value;
    }
}
```

## 7. 통합 실행 엔진

```csharp
public class DynamicYamlToExcelConverter
{
    private readonly DynamicStructureAnalyzer _analyzer;
    private readonly DynamicPatternRecognizer _recognizer;
    private readonly DynamicSchemaBuilder _schemaBuilder;
    private readonly DynamicDataMapper _dataMapper;

    public void Convert(string yamlPath, string excelPath)
    {
        // 1. YAML 로드
        var yamlContent = File.ReadAllText(yamlPath);
        var yaml = new YamlStream();
        yaml.Load(new StringReader(yamlContent));
        
        var rootNode = yaml.Documents[0].RootNode;

        // 2. 구조 분석 (완전 동적)
        var pattern = _analyzer.AnalyzeStructure(rootNode);

        // 3. 최적 전략 결정 (패턴 기반)
        var strategy = _recognizer.DetermineStrategy(pattern);

        // 4. 레이아웃 정보 생성
        var layoutInfo = GenerateLayoutInfo(rootNode, pattern, strategy);

        // 5. Excel 스키마 생성
        var scheme = _schemaBuilder.BuildScheme(pattern, strategy, layoutInfo);

        // 6. 데이터 매핑
        var rows = _dataMapper.MapToExcelRows(rootNode, scheme, pattern);

        // 7. Excel 파일 작성
        WriteExcel(scheme, rows, excelPath);
    }

    private dynamic GenerateLayoutInfo(
        YamlNode root,
        StructurePattern pattern,
        LayoutStrategy strategy)
    {
        return strategy switch
        {
            LayoutStrategy.HorizontalExpansion => 
                GenerateHorizontalLayout(root, pattern),
            LayoutStrategy.VerticalNesting => 
                GenerateVerticalLayout(root, pattern),
            LayoutStrategy.Mixed => 
                GenerateMixedLayout(root, pattern),
            _ => null
        };
    }

    private void WriteExcel(ExcelScheme scheme, List<ExcelRow> rows, string path)
    {
        using (var workbook = new XLWorkbook())
        {
            var worksheet = workbook.Worksheets.Add("Sheet1");

            // 스키마 작성
            scheme.WriteToWorksheet(worksheet);

            // 데이터 작성
            int dataStartRow = scheme.LastSchemaRow + 2;
            foreach (var row in rows)
            {
                row.WriteToWorksheet(worksheet, dataStartRow++);
            }

            workbook.SaveAs(path);
        }
    }
}
```

## 8. 주요 특징

### 8.1 완전 동적 처리
- 속성명, 배열 크기, 파일 패턴 등 어떤 하드코딩도 없음
- 구조 분석을 통한 자동 패턴 인식
- 데이터 기반 우선순위 결정

### 8.2 적응형 레이아웃
- 구조 복잡도에 따른 자동 전략 선택
- 수평/수직 확장 자동 결정
- 공간 최적화

### 8.3 패턴 기반 처리
- 파일명이 아닌 구조 패턴으로 처리 방식 결정
- 유사 구조 자동 그룹화
- 재사용 가능한 패턴 라이브러리

### 8.4 확장성
- 새로운 YAML 구조 자동 지원
- 커스텀 패턴 인식기 추가 가능
- 플러그인 아키텍처

## 9. 사용 예시

```csharp
// 사용법은 매우 간단
var converter = new DynamicYamlToExcelConverter();

// 어떤 YAML 파일이든 자동 처리
converter.Convert("any_structure.yaml", "output.xlsx");
converter.Convert("complex_nested.yaml", "output2.xlsx");
converter.Convert("variable_arrays.yaml", "output3.xlsx");

// 모든 처리가 자동으로 이루어짐
```

이 설계는 완전히 동적이며, 어떤 YAML 구조든 자동으로 분석하여 최적의 Excel 레이아웃을 생성합니다.