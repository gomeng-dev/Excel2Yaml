# 📚 Excel2Yaml 프로젝트 상세 리팩토링 계획서

## 🎯 프로젝트 개요

Excel2Yaml는 Excel 스프레드시트를 YAML, JSON, XML 등의 구조화된 데이터 형식으로 변환하는 VSTO 애드인입니다. 현재 코드베이스는 기능적으로는 완성도가 높지만, 유지보수성과 확장성 측면에서 개선이 필요한 상태입니다.

## 🔍 현황 분석 (As-Is)

### 주요 문제점

1. **아키텍처 문제**
   - 단일 책임 원칙(SRP) 위반: 한 클래스가 너무 많은 책임을 가짐
   - 의존성 역전 원칙(DIP) 위반: 구체적인 구현에 직접 의존
   - 테스트 불가능한 구조: 정적 메서드와 싱글톤 과다 사용

2. **코드 품질 문제**
   - 높은 순환 복잡도: 일부 메서드가 20 이상의 복잡도를 가짐
   - 코드 중복: DRY 원칙 위반
   - 매직 값: 하드코딩된 문자열과 숫자

3. **유지보수성 문제**
   - 변경 시 영향 범위가 넓음
   - 새로운 기능 추가가 어려움
   - 디버깅과 문제 추적이 복잡함

## 🚀 목표 아키텍처 (To-Be)

### 핵심 설계 원칙

1. **클린 아키텍처**: 계층 간 명확한 책임 분리
2. **SOLID 원칙**: 객체지향 설계 원칙 준수
3. **DDD(Domain-Driven Design)**: 도메인 중심 설계
4. **테스트 가능한 구조**: 모든 비즈니스 로직의 단위 테스트 가능

### 아키텍처 다이어그램

```
┌─────────────────────────────────────────────────────────────┐
│                      Presentation Layer                      │
│  ┌─────────────┐  ┌──────────────┐  ┌─────────────────┐   │
│  │   Ribbon    │  │    Forms     │  │   ViewModels    │   │
│  └─────────────┘  └──────────────┘  └─────────────────┘   │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                     Application Layer                        │
│  ┌──────────────┐  ┌──────────────┐  ┌───────────────┐    │
│  │   Services   │  │   Commands   │  │    Queries    │    │
│  └──────────────┘  └──────────────┘  └───────────────┘    │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                       Domain Layer                           │
│  ┌──────────────┐  ┌──────────────┐  ┌───────────────┐    │
│  │   Entities   │  │ Value Objects│  │  Domain Svcs  │    │
│  └──────────────┘  └──────────────┘  └───────────────┘    │
└─────────────────────────────────────────────────────────────┘
                              │
┌─────────────────────────────────────────────────────────────┐
│                   Infrastructure Layer                       │
│  ┌──────────────┐  ┌──────────────┐  ┌───────────────┐    │
│  │ Repositories │  │   External   │  │ Configuration │    │
│  └──────────────┘  └──────────────┘  └───────────────┘    │
└─────────────────────────────────────────────────────────────┘
```

## 📋 상세 리팩토링 계획

### Phase 1: 기반 구조 구축 (1-2주)

#### 1.1 프로젝트 구조 재구성

**목표**: 클린 아키텍처에 맞는 폴더 구조 확립

```
ExcelToYaml/
├── Domain/
│   ├── Entities/
│   ├── ValueObjects/
│   ├── Interfaces/
│   └── Services/
├── Application/
│   ├── Commands/
│   ├── Queries/
│   ├── Services/
│   └── DTOs/
├── Infrastructure/
│   ├── Excel/
│   ├── FileSystem/
│   ├── Configuration/
│   └── Logging/
├── Presentation/
│   ├── Ribbon/
│   ├── Forms/
│   └── ViewModels/
└── Tests/
    ├── Unit/
    ├── Integration/
    └── TestUtilities/
```

**To-Do List**:
- [ ] 새로운 폴더 구조 생성
- [ ] 기존 파일들을 적절한 레이어로 이동
- [ ] 네임스페이스 정리 및 업데이트
- [ ] 프로젝트 참조 관계 재설정
- [ ] 빌드 확인 및 컴파일 오류 수정

#### 1.2 상수 및 설정 중앙화

**목표**: 모든 매직 값을 상수로 추출하여 중앙 관리

**구현 예시**:

```csharp
// Domain/Constants/SchemeConstants.cs
namespace ExcelToYaml.Domain.Constants
{
    public static class SchemeConstants
    {
        public static class Markers
        {
            public const string SchemeEnd = "$scheme_end";
            public const string ArrayStart = "$[]";
            public const string MapStart = "${}";
            public const string DynamicKey = "$key";
            public const string DynamicValue = "$value";
            public const string Ignore = "^";
        }

        public static class Sheet
        {
            public const string ConversionPrefix = "!";
            public const string ConfigurationName = "_ExcelToYamlConfig";
            public const int SchemaStartRow = 2;
        }

        public static class Defaults
        {
            public const int MaxFileDisplayCount = 5;
            public const int DefaultTimeout = 120000;
            public const string DefaultDateFormat = "yyyy-MM-dd";
        }
    }
}

// Domain/Constants/ErrorMessages.cs
namespace ExcelToYaml.Domain.Constants
{
    public static class ErrorMessages
    {
        public static class Schema
        {
            public const string EndMarkerNotFound = "스키마 종료 마커($scheme_end)를 찾을 수 없습니다.";
            public const string InvalidStructure = "잘못된 스키마 구조입니다.";
            public const string MissingRequiredColumn = "필수 열이 누락되었습니다: {0}";
        }

        public static class Conversion
        {
            public const string NoSheetsFound = "변환할 시트를 찾을 수 없습니다.";
            public const string ConversionFailed = "변환 중 오류가 발생했습니다: {0}";
            public const string InvalidSheetName = "시트 이름은 '!'로 시작해야 합니다.";
        }

        public static class File
        {
            public const string SaveFailed = "파일 저장에 실패했습니다: {0}";
            public const string InvalidPath = "잘못된 경로입니다: {0}";
            public const string AccessDenied = "파일에 접근할 수 없습니다: {0}";
        }
    }
}
```

**To-Do List**:
- [ ] SchemeConstants 클래스 생성
- [ ] ErrorMessages 클래스 생성
- [ ] RegexPatterns 클래스 생성
- [ ] 전체 코드베이스에서 하드코딩된 값 검색
- [ ] 하드코딩된 값을 상수로 교체
- [ ] 상수 사용 부분 테스트

#### 1.3 도메인 모델 정의

**목표**: 비즈니스 로직의 핵심이 되는 도메인 모델 확립

**구현 예시**:

```csharp
// Domain/Entities/Scheme.cs
namespace ExcelToYaml.Domain.Entities
{
    public class Scheme
    {
        public SchemeNode Root { get; private set; }
        public string SheetName { get; private set; }
        public int EndRow { get; private set; }
        
        private Scheme() { }
        
        public static Scheme Create(string sheetName, SchemeNode root, int endRow)
        {
            if (string.IsNullOrWhiteSpace(sheetName))
                throw new ArgumentException("Sheet name cannot be empty");
            
            if (root == null)
                throw new ArgumentNullException(nameof(root));
            
            return new Scheme
            {
                SheetName = sheetName,
                Root = root,
                EndRow = endRow
            };
        }
    }
}

// Domain/ValueObjects/CellPosition.cs
namespace ExcelToYaml.Domain.ValueObjects
{
    public class CellPosition : ValueObject
    {
        public int Row { get; }
        public int Column { get; }
        
        public CellPosition(int row, int column)
        {
            if (row < 1) throw new ArgumentException("Row must be positive");
            if (column < 1) throw new ArgumentException("Column must be positive");
            
            Row = row;
            Column = column;
        }
        
        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return Row;
            yield return Column;
        }
    }
}

// Domain/ValueObjects/SchemeNodeType.cs
namespace ExcelToYaml.Domain.ValueObjects
{
    public class SchemeNodeType : ValueObject
    {
        public static readonly SchemeNodeType Property = new("PROPERTY");
        public static readonly SchemeNodeType Map = new("MAP");
        public static readonly SchemeNodeType Array = new("ARRAY");
        public static readonly SchemeNodeType Key = new("KEY");
        public static readonly SchemeNodeType Value = new("VALUE");
        
        public string Value { get; }
        
        private SchemeNodeType(string value)
        {
            Value = value;
        }
        
        protected override IEnumerable<object> GetEqualityComponents()
        {
            yield return Value;
        }
    }
}
```

**To-Do List**:
- [ ] Scheme 엔티티 생성
- [ ] SchemeNode 엔티티 리팩토링
- [ ] CellPosition 값 객체 생성
- [ ] SchemeNodeType 값 객체 생성
- [ ] ConversionOptions 값 객체 생성
- [ ] 도메인 모델 단위 테스트 작성

#### 1.4 인터페이스 및 추상화 정의

**목표**: 의존성 역전을 위한 인터페이스 계층 구축

**구현 예시**:

```csharp
// Domain/Interfaces/ISchemeParser.cs
namespace ExcelToYaml.Domain.Interfaces
{
    public interface ISchemeParser
    {
        /// <summary>
        /// Excel 워크시트에서 스키마 구조를 파싱합니다.
        /// </summary>
        Scheme Parse(IWorksheet worksheet);
        
        /// <summary>
        /// 스키마 유효성을 검증합니다.
        /// </summary>
        ValidationResult ValidateSchema(Scheme scheme);
    }
}

// Domain/Interfaces/IDataGenerator.cs
namespace ExcelToYaml.Domain.Interfaces
{
    public interface IDataGenerator<TOutput>
    {
        /// <summary>
        /// 스키마와 워크시트 데이터를 기반으로 출력을 생성합니다.
        /// </summary>
        TOutput Generate(Scheme scheme, IWorksheet worksheet, GenerationOptions options);
        
        /// <summary>
        /// 생성 가능 여부를 확인합니다.
        /// </summary>
        bool CanGenerate(Scheme scheme);
    }
}

// Domain/Interfaces/IPostProcessor.cs
namespace ExcelToYaml.Domain.Interfaces
{
    public interface IPostProcessor
    {
        /// <summary>
        /// 처리 우선순위 (낮을수록 먼저 실행)
        /// </summary>
        int Priority { get; }
        
        /// <summary>
        /// 후처리를 수행합니다.
        /// </summary>
        Task<ProcessingResult> ProcessAsync(string input, ProcessingContext context);
        
        /// <summary>
        /// 이 프로세서가 처리 가능한지 확인합니다.
        /// </summary>
        bool CanProcess(ProcessingContext context);
    }
}

// Application/Interfaces/IConversionService.cs
namespace ExcelToYaml.Application.Interfaces
{
    public interface IConversionService
    {
        Task<ConversionResult> ConvertAsync(
            ConversionRequest request, 
            CancellationToken cancellationToken = default);
        
        Task<IEnumerable<string>> GetConvertibleSheetsAsync(
            string workbookPath,
            CancellationToken cancellationToken = default);
    }
}
```

**To-Do List**:
- [ ] 도메인 레이어 인터페이스 정의
- [ ] 애플리케이션 레이어 인터페이스 정의
- [ ] 인프라 레이어 인터페이스 정의
- [ ] DTO 및 요청/응답 모델 정의
- [ ] 인터페이스 문서화 (XML 주석)

### Phase 2: 핵심 컴포넌트 리팩토링 (2-3주)

#### 2.1 SchemeParser 리팩토링

**목표**: 복잡도를 낮추고 테스트 가능한 구조로 개선

**현재 문제점**:
- Parse 메서드의 복잡도가 너무 높음 (15+)
- 재귀 호출과 복잡한 조건문이 혼재
- 병합 셀 처리 로직이 파싱 로직과 섞여 있음

**개선 방안**:

```csharp
// Infrastructure/Excel/SchemeParser.cs
namespace ExcelToYaml.Infrastructure.Excel
{
    public class SchemeParser : ISchemeParser
    {
        private readonly ISchemeValidator _validator;
        private readonly ISchemeNodeFactory _nodeFactory;
        private readonly ILogger<SchemeParser> _logger;
        
        public SchemeParser(
            ISchemeValidator validator,
            ISchemeNodeFactory nodeFactory,
            ILogger<SchemeParser> logger)
        {
            _validator = validator;
            _nodeFactory = nodeFactory;
            _logger = logger;
        }
        
        public Scheme Parse(IWorksheet worksheet)
        {
            _logger.LogInformation("스키마 파싱 시작: {SheetName}", worksheet.Name);
            
            var endRow = FindSchemeEndRow(worksheet);
            var rootNode = ParseRootNode(worksheet, endRow);
            var scheme = Scheme.Create(worksheet.Name, rootNode, endRow);
            
            var validationResult = _validator.Validate(scheme);
            if (!validationResult.IsValid)
            {
                throw new SchemeParsingException(validationResult.Errors);
            }
            
            return scheme;
        }
        
        private int FindSchemeEndRow(IWorksheet worksheet)
        {
            // 단일 책임: 스키마 종료 행 찾기
            for (int row = 1; row <= worksheet.RowCount; row++)
            {
                var firstCell = worksheet.GetCell(row, 1);
                if (firstCell?.Value?.ToString() == SchemeConstants.Markers.SchemeEnd)
                {
                    return row;
                }
            }
            
            throw new SchemeParsingException("스키마 종료 마커를 찾을 수 없습니다.");
        }
        
        private SchemeNode ParseRootNode(IWorksheet worksheet, int endRow)
        {
            var context = new ParsingContext(worksheet, endRow);
            var builder = new SchemeNodeBuilder(_nodeFactory);
            
            // 각 열을 순회하며 노드 구성
            for (int col = 1; col <= worksheet.ColumnCount; col++)
            {
                var columnNodes = ParseColumn(context, col);
                builder.AddColumn(columnNodes);
            }
            
            return builder.Build();
        }
    }
}

// Infrastructure/Excel/SchemeNodeBuilder.cs
namespace ExcelToYaml.Infrastructure.Excel
{
    public class SchemeNodeBuilder
    {
        private readonly ISchemeNodeFactory _factory;
        private readonly Dictionary<int, List<SchemeNode>> _columnNodes;
        
        public void AddColumn(IEnumerable<SchemeNode> nodes)
        {
            // 열별로 노드를 수집하여 계층 구조 구성
        }
        
        public SchemeNode Build()
        {
            // 수집된 노드들을 바탕으로 루트 노드 구성
            return _factory.CreateRootNode(_columnNodes);
        }
    }
}
```

**To-Do List**:
- [ ] SchemeParser를 작은 단위로 분해
- [ ] SchemeNodeFactory 구현
- [ ] SchemeValidator 구현
- [ ] ParsingContext 클래스 생성
- [ ] 병합 셀 처리 로직 분리
- [ ] 단위 테스트 작성 (최소 80% 커버리지)

#### 2.2 YamlGenerator 리팩토링

**목표**: 600줄이 넘는 거대한 클래스를 책임별로 분리

**현재 문제점**:
- 노드 순회, 데이터 생성, 포맷팅이 한 클래스에 혼재
- 스택 관리와 비즈니스 로직이 섞여 있음
- 테스트하기 어려운 구조

**개선 방안**:

```csharp
// Application/Services/Generation/YamlGenerationService.cs
namespace ExcelToYaml.Application.Services.Generation
{
    public class YamlGenerationService : IDataGenerator<string>
    {
        private readonly INodeTraverser _traverser;
        private readonly IYamlBuilder _yamlBuilder;
        private readonly IGenerationOptionsValidator _optionsValidator;
        private readonly ILogger<YamlGenerationService> _logger;
        
        public string Generate(
            Scheme scheme, 
            IWorksheet worksheet, 
            GenerationOptions options)
        {
            _optionsValidator.Validate(options);
            
            var context = new GenerationContext(worksheet, options);
            var data = _traverser.Traverse(scheme.Root, context);
            var yaml = _yamlBuilder.Build(data, options);
            
            return yaml;
        }
    }
}

// Domain/Services/NodeTraverser.cs
namespace ExcelToYaml.Domain.Services
{
    public class NodeTraverser : INodeTraverser
    {
        private readonly INodeProcessorResolver _processorResolver;
        
        public object Traverse(SchemeNode node, GenerationContext context)
        {
            var processor = _processorResolver.Resolve(node.Type);
            return processor.Process(node, context, this);
        }
    }
}

// Domain/Services/NodeProcessors/PropertyNodeProcessor.cs
namespace ExcelToYaml.Domain.Services.NodeProcessors
{
    public class PropertyNodeProcessor : INodeProcessor
    {
        public object Process(
            SchemeNode node, 
            GenerationContext context, 
            INodeTraverser traverser)
        {
            // PROPERTY 노드만 처리하는 단일 책임
            var cellValue = context.Worksheet.GetCell(context.CurrentRow, node.Column);
            
            if (ShouldSkipEmpty(cellValue, context.Options))
            {
                return NodeProcessResult.Skip;
            }
            
            return new PropertyData
            {
                Name = node.Name,
                Value = FormatValue(cellValue, node.Format)
            };
        }
        
        private bool ShouldSkipEmpty(object value, GenerationOptions options)
        {
            return value == null && options.SkipEmptyFields;
        }
    }
}

// Domain/Services/NodeProcessors/ArrayNodeProcessor.cs
namespace ExcelToYaml.Domain.Services.NodeProcessors
{
    public class ArrayNodeProcessor : INodeProcessor
    {
        public object Process(
            SchemeNode node, 
            GenerationContext context, 
            INodeTraverser traverser)
        {
            var items = new List<object>();
            var startRow = context.CurrentRow;
            
            // 배열 항목 처리
            while (context.CurrentRow <= context.EndRow)
            {
                var item = ProcessArrayItem(node, context, traverser);
                if (item != NodeProcessResult.Skip)
                {
                    items.Add(item);
                }
                context.MoveToNextRow();
            }
            
            return new ArrayData
            {
                Name = node.Name,
                Items = items
            };
        }
    }
}
```

**To-Do List**:
- [ ] YamlGenerationService 생성 (오케스트레이션)
- [ ] NodeTraverser 구현 (순회 로직)
- [ ] NodeProcessor 인터페이스 및 구현체들
- [ ] YamlBuilder 구현 (YAML 생성)
- [ ] GenerationContext 구현
- [ ] 스택 관리 로직 제거 및 개선
- [ ] 각 컴포넌트별 단위 테스트

#### 2.3 Ribbon UI 분리

**목표**: UI 로직과 비즈니스 로직 분리 (MVP 패턴 적용)

**현재 문제점**:
- 1000줄이 넘는 거대한 Ribbon 클래스
- UI 이벤트 핸들러에 비즈니스 로직이 혼재
- 테스트 불가능한 구조

**개선 방안**:

```csharp
// Presentation/ViewModels/ConversionViewModel.cs
namespace ExcelToYaml.Presentation.ViewModels
{
    public class ConversionViewModel : ViewModelBase
    {
        private readonly IConversionService _conversionService;
        private readonly ISheetSelectionService _sheetSelection;
        private readonly IProgressReporter _progressReporter;
        private readonly IDialogService _dialogService;
        
        public ICommand ConvertToYamlCommand { get; }
        public ICommand ConvertToJsonCommand { get; }
        public ICommand ConfigureSettingsCommand { get; }
        
        public ObservableCollection<SheetInfo> AvailableSheets { get; }
        public bool IsProcessing { get; private set; }
        
        public ConversionViewModel(
            IConversionService conversionService,
            ISheetSelectionService sheetSelection,
            IProgressReporter progressReporter,
            IDialogService dialogService)
        {
            _conversionService = conversionService;
            _sheetSelection = sheetSelection;
            _progressReporter = progressReporter;
            _dialogService = dialogService;
            
            ConvertToYamlCommand = new AsyncCommand(ConvertToYamlAsync);
            ConvertToJsonCommand = new AsyncCommand(ConvertToJsonAsync);
            ConfigureSettingsCommand = new Command(ConfigureSettings);
        }
        
        private async Task ConvertToYamlAsync()
        {
            try
            {
                IsProcessing = true;
                
                var sheets = await _sheetSelection.GetSelectedSheetsAsync();
                if (!sheets.Any())
                {
                    await _dialogService.ShowWarningAsync("변환할 시트를 선택해주세요.");
                    return;
                }
                
                var request = new ConversionRequest
                {
                    Sheets = sheets,
                    OutputFormat = OutputFormat.Yaml,
                    Options = await GetConversionOptionsAsync()
                };
                
                var progress = new Progress<ConversionProgress>(OnProgressUpdate);
                var result = await _conversionService.ConvertAsync(request, progress);
                
                await ShowResultAsync(result);
            }
            catch (Exception ex)
            {
                await _dialogService.ShowErrorAsync($"변환 실패: {ex.Message}");
            }
            finally
            {
                IsProcessing = false;
            }
        }
    }
}

// Presentation/Ribbon/RibbonPresenter.cs
namespace ExcelToYaml.Presentation.Ribbon
{
    public class RibbonPresenter
    {
        private readonly ConversionViewModel _viewModel;
        private readonly Ribbon _view;
        
        public RibbonPresenter(Ribbon view, ConversionViewModel viewModel)
        {
            _view = view;
            _viewModel = viewModel;
            
            BindCommands();
            SubscribeToEvents();
        }
        
        private void BindCommands()
        {
            _view.ConvertToYamlButton.Click += (s, e) => 
                _viewModel.ConvertToYamlCommand.Execute(null);
            
            _view.ConvertToJsonButton.Click += (s, e) => 
                _viewModel.ConvertToJsonCommand.Execute(null);
        }
    }
}
```

**To-Do List**:
- [ ] ConversionViewModel 생성
- [ ] Command 패턴 구현 (ICommand)
- [ ] DialogService 구현
- [ ] ProgressReporter 구현
- [ ] SheetSelectionService 구현
- [ ] RibbonPresenter 구현
- [ ] 기존 Ribbon.cs 리팩토링
- [ ] ViewModel 단위 테스트

### Phase 3: 후처리 시스템 현대화 (1주)

#### 3.1 후처리 파이프라인 구축

**목표**: 확장 가능하고 테스트 가능한 후처리 시스템

**구현 예시**:

```csharp
// Application/PostProcessing/ProcessingPipeline.cs
namespace ExcelToYaml.Application.PostProcessing
{
    public class ProcessingPipeline : IProcessingPipeline
    {
        private readonly IEnumerable<IPostProcessor> _processors;
        private readonly ILogger<ProcessingPipeline> _logger;
        
        public ProcessingPipeline(
            IEnumerable<IPostProcessor> processors,
            ILogger<ProcessingPipeline> logger)
        {
            _processors = processors.OrderBy(p => p.Priority);
            _logger = logger;
        }
        
        public async Task<ProcessingResult> ProcessAsync(
            string input, 
            ProcessingContext context,
            CancellationToken cancellationToken = default)
        {
            var result = new ProcessingResult(input);
            
            foreach (var processor in _processors)
            {
                if (!processor.CanProcess(context))
                {
                    _logger.LogDebug("Skipping processor: {Processor}", 
                        processor.GetType().Name);
                    continue;
                }
                
                try
                {
                    _logger.LogInformation("Applying processor: {Processor}", 
                        processor.GetType().Name);
                    
                    result = await processor.ProcessAsync(
                        result.Output, 
                        context, 
                        cancellationToken);
                    
                    if (!result.Success)
                    {
                        _logger.LogWarning("Processor failed: {Processor}, {Error}", 
                            processor.GetType().Name, result.Error);
                        break;
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex, "Error in processor: {Processor}", 
                        processor.GetType().Name);
                    throw;
                }
            }
            
            return result;
        }
    }
}

// Application/PostProcessing/Processors/YamlMergeProcessor.cs
namespace ExcelToYaml.Application.PostProcessing.Processors
{
    public class YamlMergeProcessor : PostProcessorBase
    {
        public override int Priority => 10;
        
        private readonly IMergeStrategy _mergeStrategy;
        
        public YamlMergeProcessor(IMergeStrategy mergeStrategy)
        {
            _mergeStrategy = mergeStrategy;
        }
        
        public override bool CanProcess(ProcessingContext context)
        {
            return context.Options.EnableMerge && 
                   context.OutputFormat == OutputFormat.Yaml;
        }
        
        protected override async Task<string> ProcessCoreAsync(
            string input, 
            ProcessingContext context)
        {
            var yaml = ParseYaml(input);
            var merged = await _mergeStrategy.MergeAsync(yaml, context.MergeOptions);
            return SerializeYaml(merged);
        }
    }
}

// Application/PostProcessing/Processors/YamlFlowStyleProcessor.cs
namespace ExcelToYaml.Application.PostProcessing.Processors
{
    public class YamlFlowStyleProcessor : PostProcessorBase
    {
        public override int Priority => 20;
        
        private readonly IFlowStyleAnalyzer _analyzer;
        
        public override bool CanProcess(ProcessingContext context)
        {
            return context.Options.ApplyFlowStyle && 
                   context.OutputFormat == OutputFormat.Yaml;
        }
        
        protected override async Task<string> ProcessCoreAsync(
            string input, 
            ProcessingContext context)
        {
            var flowStyleRules = await _analyzer.AnalyzeAsync(input);
            return ApplyFlowStyle(input, flowStyleRules);
        }
    }
}
```

**To-Do List**:
- [ ] ProcessingPipeline 구현
- [ ] PostProcessorBase 추상 클래스
- [ ] YamlMergeProcessor 리팩토링
- [ ] YamlFlowStyleProcessor 리팩토링
- [ ] JsonFormatterProcessor 구현
- [ ] XmlFormatterProcessor 구현
- [ ] 처리 순서 및 우선순위 시스템
- [ ] 각 프로세서 단위 테스트

### Phase 4: 설정 관리 시스템 (1주)

#### 4.1 설정 관리 현대화

**목표**: 유연하고 확장 가능한 설정 시스템

**구현 예시**:

```csharp
// Application/Configuration/ConfigurationService.cs
namespace ExcelToYaml.Application.Configuration
{
    public class ConfigurationService : IConfigurationService
    {
        private readonly IConfigurationRepository _repository;
        private readonly IConfigurationValidator _validator;
        private readonly IEventBus _eventBus;
        
        public async Task<TConfig> GetConfigurationAsync<TConfig>() 
            where TConfig : class, IConfiguration, new()
        {
            var config = await _repository.LoadAsync<TConfig>();
            if (config == null)
            {
                config = new TConfig();
                await SaveConfigurationAsync(config);
            }
            
            return config;
        }
        
        public async Task SaveConfigurationAsync<TConfig>(TConfig configuration) 
            where TConfig : class, IConfiguration
        {
            var validationResult = await _validator.ValidateAsync(configuration);
            if (!validationResult.IsValid)
            {
                throw new ConfigurationException(validationResult.Errors);
            }
            
            await _repository.SaveAsync(configuration);
            await _eventBus.PublishAsync(new ConfigurationChangedEvent(configuration));
        }
    }
}

// Domain/Configuration/ConversionConfiguration.cs
namespace ExcelToYaml.Domain.Configuration
{
    public class ConversionConfiguration : IConfiguration
    {
        public string ConfigurationId => "ConversionSettings";
        
        public OutputSettings Output { get; set; } = new();
        public ProcessingSettings Processing { get; set; } = new();
        public AdvancedSettings Advanced { get; set; } = new();
        
        public class OutputSettings
        {
            public bool SkipEmptyFields { get; set; } = true;
            public bool PreservePropertyOrder { get; set; } = true;
            public string DateTimeFormat { get; set; } = "yyyy-MM-dd HH:mm:ss";
            public string NumberFormat { get; set; } = "G";
        }
        
        public class ProcessingSettings
        {
            public bool EnablePostProcessing { get; set; } = true;
            public bool EnableMergeByKey { get; set; } = false;
            public bool ApplyFlowStyle { get; set; } = false;
            public List<string> MergeKeyPaths { get; set; } = new();
        }
        
        public class AdvancedSettings
        {
            public int MaxDepth { get; set; } = 100;
            public int MaxArraySize { get; set; } = 10000;
            public bool ValidateOutput { get; set; } = true;
        }
    }
}

// Infrastructure/Configuration/ExcelConfigurationRepository.cs
namespace ExcelToYaml.Infrastructure.Configuration
{
    public class ExcelConfigurationRepository : IConfigurationRepository
    {
        private readonly IExcelWorkbook _workbook;
        private readonly ISerializer _serializer;
        
        public async Task<T> LoadAsync<T>() where T : class, IConfiguration
        {
            var configSheet = GetOrCreateConfigSheet();
            var configData = ReadConfigurationData(configSheet, typeof(T).Name);
            
            if (string.IsNullOrEmpty(configData))
                return null;
            
            return _serializer.Deserialize<T>(configData);
        }
        
        public async Task SaveAsync<T>(T configuration) where T : class, IConfiguration
        {
            var configSheet = GetOrCreateConfigSheet();
            var serialized = _serializer.Serialize(configuration);
            
            WriteConfigurationData(configSheet, configuration.ConfigurationId, serialized);
            await Task.CompletedTask;
        }
    }
}
```

**To-Do List**:
- [ ] ConfigurationService 구현
- [ ] IConfiguration 인터페이스 정의
- [ ] ConversionConfiguration 클래스
- [ ] SheetPathConfiguration 클래스
- [ ] ConfigurationValidator 구현
- [ ] ExcelConfigurationRepository 구현
- [ ] JsonConfigurationRepository 구현 (대안)
- [ ] 설정 마이그레이션 도구

### Phase 5: 에러 처리 및 로깅 (1주)

#### 5.1 구조화된 에러 처리

**목표**: 일관되고 유용한 에러 처리 시스템

**구현 예시**:

```csharp
// Domain/Exceptions/ExcelConversionException.cs
namespace ExcelToYaml.Domain.Exceptions
{
    public abstract class ExcelConversionException : Exception
    {
        public string ErrorCode { get; }
        public Dictionary<string, object> Context { get; }
        
        protected ExcelConversionException(
            string errorCode, 
            string message, 
            Exception innerException = null) 
            : base(message, innerException)
        {
            ErrorCode = errorCode;
            Context = new Dictionary<string, object>();
        }
        
        public ExcelConversionException WithContext(string key, object value)
        {
            Context[key] = value;
            return this;
        }
    }
}

// Domain/Exceptions/SchemeParsingException.cs
namespace ExcelToYaml.Domain.Exceptions
{
    public class SchemeParsingException : ExcelConversionException
    {
        public string SheetName { get; }
        public int? Row { get; }
        public int? Column { get; }
        
        public SchemeParsingException(
            string message, 
            string sheetName = null, 
            int? row = null, 
            int? column = null) 
            : base("SCHEME_PARSE_ERROR", message)
        {
            SheetName = sheetName;
            Row = row;
            Column = column;
            
            if (!string.IsNullOrEmpty(sheetName))
                WithContext("SheetName", sheetName);
            if (row.HasValue)
                WithContext("Row", row.Value);
            if (column.HasValue)
                WithContext("Column", column.Value);
        }
    }
}

// Application/ErrorHandling/GlobalErrorHandler.cs
namespace ExcelToYaml.Application.ErrorHandling
{
    public class GlobalErrorHandler : IGlobalErrorHandler
    {
        private readonly ILogger<GlobalErrorHandler> _logger;
        private readonly IUserNotificationService _notificationService;
        
        public async Task<ErrorHandlingResult> HandleAsync(Exception exception)
        {
            switch (exception)
            {
                case SchemeParsingException spe:
                    return await HandleSchemeParsingError(spe);
                    
                case DataConversionException dce:
                    return await HandleDataConversionError(dce);
                    
                case ConfigurationException ce:
                    return await HandleConfigurationError(ce);
                    
                default:
                    return await HandleUnknownError(exception);
            }
        }
        
        private async Task<ErrorHandlingResult> HandleSchemeParsingError(
            SchemeParsingException exception)
        {
            _logger.LogError(exception, 
                "스키마 파싱 오류 발생 - Sheet: {Sheet}, Row: {Row}, Column: {Col}",
                exception.SheetName, exception.Row, exception.Column);
            
            var userMessage = BuildUserFriendlyMessage(exception);
            await _notificationService.ShowErrorAsync(userMessage);
            
            return new ErrorHandlingResult
            {
                Handled = true,
                ShouldRetry = false,
                UserAction = UserAction.FixSchemaAndRetry
            };
        }
    }
}
```

**To-Do List**:
- [ ] 예외 계층 구조 설계
- [ ] 도메인별 예외 클래스 생성
- [ ] GlobalErrorHandler 구현
- [ ] 에러 복구 전략 구현
- [ ] 사용자 친화적 에러 메시지
- [ ] 에러 로깅 및 추적
- [ ] 예외 처리 단위 테스트

#### 5.2 구조화된 로깅

**구현 예시**:

```csharp
// Infrastructure/Logging/StructuredLogger.cs
namespace ExcelToYaml.Infrastructure.Logging
{
    public class StructuredLogger : ILogger<T>
    {
        private readonly ILoggerFactory _loggerFactory;
        private readonly ILogContext _context;
        
        public void LogInformation(string message, params object[] args)
        {
            using (_context.Push("CorrelationId", Guid.NewGuid()))
            using (_context.Push("Timestamp", DateTime.UtcNow))
            {
                _innerLogger.LogInformation(message, args);
            }
        }
        
        public IDisposable BeginScope<TState>(TState state)
        {
            return _context.Push("Scope", state);
        }
    }
}
```

**To-Do List**:
- [ ] 구조화된 로깅 구현
- [ ] 로그 컨텍스트 관리
- [ ] 성능 메트릭 로깅
- [ ] 감사(Audit) 로깅
- [ ] 로그 필터링 및 레벨 관리

## 📊 성공 지표

### 코드 품질 메트릭
- **순환 복잡도**: 최대 10 이하
- **메서드 길이**: 최대 30줄
- **클래스 크기**: 최대 300줄
- **코드 중복**: 5% 이하

### 아키텍처 품질
- **레이어 간 의존성**: 단방향 유지
- **인터페이스 분리**: 모든 주요 컴포넌트
- **테스트 커버리지**: 핵심 로직 80% 이상

### 개발 생산성
- **새 기능 추가**: 기존 대비 50% 시간 단축
- **버그 수정**: 기존 대비 70% 시간 단축
- **코드 리뷰**: 평균 리뷰 시간 50% 단축

## 🚀 실행 계획

### Week 1-2: 기반 구조
- [ ] 프로젝트 구조 재구성
- [ ] 상수 및 설정 중앙화
- [ ] 도메인 모델 정의
- [ ] 인터페이스 계층 구축

### Week 3-5: 핵심 리팩토링
- [ ] SchemeParser 개선
- [ ] YamlGenerator 분해
- [ ] Ribbon UI 분리
- [ ] 단위 테스트 작성

### Week 6: 후처리 시스템
- [ ] 파이프라인 구축
- [ ] 프로세서 리팩토링
- [ ] 통합 테스트

### Week 7: 설정 및 에러 처리
- [ ] 설정 시스템 구현
- [ ] 에러 처리 개선
- [ ] 로깅 시스템 구축

### Week 8: 마무리
- [ ] 통합 테스트
- [ ] 성능 최적화
- [ ] 문서화
- [ ] 코드 리뷰

## 📝 위험 관리

### 주요 위험 요소
1. **기존 기능 손상**: 점진적 리팩토링으로 최소화
2. **일정 지연**: 우선순위 기반 접근
3. **팀 저항**: 명확한 이익 제시 및 교육

### 완화 전략
1. **기능 테스트**: 각 단계마다 회귀 테스트
2. **점진적 접근**: 작은 단위로 나누어 진행
3. **문서화**: 변경사항 상세 기록

이 리팩토링 계획을 통해 Excel2Yaml 프로젝트는 더욱 견고하고 유지보수가 용이한 구조로 발전할 것입니다.