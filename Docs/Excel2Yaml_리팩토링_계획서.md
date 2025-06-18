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

### 최근 완료된 리팩토링 (2025-06-18)

#### Presentation Layer 리팩토링 ✅
1. **Ribbon.cs 대규모 리팩토링**
   - 1821줄 → 약 1150줄 (37% 감소)
   - 단일 책임 원칙(SRP) 적용
   - 서비스 분리 완료

2. **새로운 서비스 구조**
   ```
   Presentation/
   ├── Services/
   │   ├── ConversionService.cs     - Excel 변환 로직
   │   ├── ImportExportService.cs   - Import/Export 기능
   │   └── PostProcessingService.cs - YAML 후처리
   └── Helpers/
       └── RibbonHelpers.cs         - 공통 유틸리티
   ```

3. **의존성 주입 패턴 적용**
   - 서비스 필드를 통한 의존성 관리
   - 테스트 가능한 구조로 개선

4. **추가 개선사항** (2025-06-18 오후)
   - Import 함수 통합: OnImportXmlClick, OnImportYamlClick, OnImportJsonClick → HandleImport(fileType)
   - 중복 코드 제거로 약 100줄 추가 감소
   - 진행률 표시 개선: 상세한 단계별 프로그레스 바 적용
   - ConversionService와 PostProcessingService에 세밀한 진행률 보고 추가

5. **Phase 2.3.3 완료** (2025-06-18 저녁) ✅
   - Convert 함수 통합: OnConvertToYamlClick, OnConvertToXmlClick, OnConvertYamlToJsonClick → HandleConvert(targetFormat)
   - 중복 코드 제거로 약 560줄 추가 감소 (Ribbon.cs: 1821줄 → 880줄, 총 52% 감소)
   - 6개의 개별 함수를 2개의 통합 함수로 리팩토링
   - C# 7.3 호환성 유지 (switch 표현식을 switch 문으로 변환)

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

### Phase 1: 기반 구조 구축 (1-2주) ✅

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
- [x] 새로운 폴더 구조 생성 ✅
- [x] 기존 파일들을 적절한 레이어로 이동 ✅
- [x] 네임스페이스 정리 및 업데이트 ✅
- [x] 프로젝트 참조 관계 재설정 ✅
- [x] 빌드 확인 및 컴파일 오류 수정 ✅

**참고사항**:
- Application/DTOs, Domain/Interfaces, Infrastructure/Interfaces 디렉토리의 파일들은 향후 구현을 위한 준비 단계로 생성되었으나 아직 사용되지 않음
- 실제 구현 시 이들을 활용하여 더 깔끔한 아키텍처로 전환 예정

**추가 수정사항**:
- Generator.cs: Config.ExcelToYamlConfig → ExcelToYamlConfig로 변경
- Generator.cs: Config.OutputFormat → OutputFormat로 변경
- YamlGenerator.cs: using 문 업데이트 (Domain.ValueObjects, Domain.Entities 추가)
- OrderedYamlFactory.cs: using ExcelToYamlAddin.Config → Domain.ValueObjects로 변경
- ExcelReader.cs: 불필요한 using 제거

**완료된 작업 상세**:
1. 클린 아키텍처 폴더 구조 생성 완료
   - Domain, Application, Infrastructure, Presentation 레이어 생성
2. 파일 이동 완료
   - Domain Layer: Scheme, SchemeNode, ExcelToYamlConfig
   - Application Layer: YamlGenerator, Generator, ExcelReader, PostProcessors, XmlToYamlConverter, XmlToExcelViaYamlConverter, YamlToExcel 관련 파일들
   - Infrastructure Layer: SchemeParser, OrderedFactories, ConfigManagers, Logging, ExcelToHtmlExporter, SheetAnalyzer
   - Presentation Layer: Ribbon, Forms
3. Core 폴더 완전 제거
4. 프로젝트 파일(.csproj) 업데이트 완료
5. 모든 파일의 네임스페이스 업데이트 (Domain.Entities)
6. Using 문 정리
7. 빌드오류 해결

**남은 작업**:

#### 1.2 상수 및 설정 중앙화 ✅

**목표**: 모든 매직 값을 상수로 추출하여 중앙 관리

**완료된 작업**:

1. **Domain/Constants 폴더 생성 및 상수 클래스 구현**
   - SchemeConstants.cs: Excel 스키마 관련 모든 상수
   - ErrorMessages.cs: 모든 에러 메시지 상수
   - RegexPatterns.cs: 정규식 패턴 상수
   - HtmlStyles.cs: HTML/CSS 스타일 관련 상수

2. **하드코딩된 값 교체 완료**
   - SchemeParser.cs: 모든 매직 값을 SchemeConstants로 교체
   - SchemeNode.cs: 노드 타입 및 마커를 상수로 교체
   - SheetAnalyzer.cs: 시트 접두사를 상수로 교체
   - ExcelConfigManager.cs: 설정 관련 상수 교체
   - YamlGenerator.cs: 에러 메시지 및 특수 문자 상수 교체
   - ExcelReader.cs: 파일 확장자 및 에러 메시지 상수 교체
   - ExcelToHtmlExporter.cs: HTML 스타일을 HtmlStyles 상수로 교체

3. **네임스페이스 정리**
   - SheetAnalyzer.cs: Core → Infrastructure.Excel
   - ExcelToHtmlExporter.cs: Core → Infrastructure.Excel

4. **프로젝트 파일 업데이트**
   - ExcelToYamlAddin.csproj에 모든 상수 클래스 추가

**구현된 상수 클래스 구조**:

```csharp
// Domain/Constants/SchemeConstants.cs
namespace ExcelToYamlAddin.Domain.Constants
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
            public const string MarkerPrefix = "$";
        }

        public static class Sheet
        {
            public const string ConversionPrefix = "!";
            public const string ConfigurationName = "excel2yamlconfig";
            public const int SchemaStartRow = 2;
            public const int HeaderRow = 1;
            public const int DataStartRow = 2;
        }

        public static class NodeTypes
        {
            public const string Map = "{}";
            public const string Array = "[]";
            public const string Key = "key";
            public const string Value = "value";
            public const string Ignore = "^";
        }

        public static class RowNumbers
        {
            public const int IllegalRow = -1;
            public const int CommentRow = 0;
        }

        public static class Configuration
        {
            public const int SheetNameColumn = 1;
            public const int ConfigKeyColumn = 2;
            public const int ConfigValueColumn = 3;
            public const int YamlEmptyFieldsColumn = 4;
            public const int EmptyArrayFieldsColumn = 5;
            public const int UpdateWaitTimeSeconds = 5;
        }

        public static class ConfigKeys
        {
            public const string SheetName = "SheetName";
            public const string ConfigKey = "ConfigKey";
            public const string ConfigValue = "ConfigValue";
            public const string YamlEmptyFields = "YamlEmptyFields";
            public const string EmptyArrayFields = "EmptyArrayFields";
            public const string MergeKeyPaths = "MergeKeyPaths";
            public const string FlowStyle = "FlowStyle";
        }

        public static class FileExtensions
        {
            public const string Json = ".json";
            public const string Yaml = ".yaml";
            public const string Md5 = ".md5";
            public const string Excel = ".xlsx";
            public const string Xml = ".xml";
        }

        public static class Defaults
        {
            public const int MaxFileDisplayCount = 5;
            public const int DefaultTimeout = 120000;
            public const string DefaultDateFormat = "yyyy-MM-dd";
            public const string DefaultDateTimeFormat = "yyyy-MM-dd HH:mm:ss";
        }

        public static class SpecialCharacters
        {
            public const string LineFeed = "\n";
            public const string CarriageReturn = "\r";
            public const string LineFeedEscape = "\\n";
            public const string CarriageReturnEscape = "\\r";
        }
    }
}
```

**To-Do List**:
- [x] SchemeConstants 클래스 생성
- [x] ErrorMessages 클래스 생성
- [x] RegexPatterns 클래스 생성
- [x] HtmlStyles 클래스 생성 (추가)
- [x] 전체 코드베이스에서 하드코딩된 값 검색
- [x] 하드코딩된 값을 상수로 교체
- [x] 프로젝트 파일에 상수 클래스 추가
- [x] 상수 사용 부분 테스트

#### 1.3 도메인 모델 정의 ✅

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
- [x] Scheme 엔티티 생성
- [x] SchemeNode 엔티티 리팩토링
- [x] CellPosition 값 객체 생성
- [x] SchemeNodeType 값 객체 생성
- [x] ConversionOptions 값 객체 생성
- [x] 도메인 모델 단위 테스트 작성

**구현 완료 사항**:
1. **ValueObject 기본 클래스** - 모든 값 객체의 기반이 되는 추상 클래스 구현
2. **CellPosition** - Excel 셀 위치를 나타내는 값 객체 (행/열 변환, 네비게이션 메서드 포함)
3. **SchemeNodeType** - 노드 타입을 나타내는 열거형 값 객체 (Property, Map, Array, Key, Value, Ignore)
4. **ConversionOptions** - 변환 옵션을 나타내는 복합 값 객체 (Builder 패턴 적용)
5. **OutputFormat** - 출력 형식을 나타내는 값 객체 (YAML, JSON, XML, HTML)
6. **YamlStyle** - YAML 스타일을 나타내는 값 객체
7. **Scheme 엔티티** - DDD 원칙에 따른 리치 도메인 모델로 리팩토링 (팩토리 메서드, 검증, 메타데이터 지원)
8. **SchemeNode 엔티티** - 불변성과 검증 로직을 갖춘 도메인 엔티티로 리팩토링
9. **도메인 상수 적용** - 모든 도메인 모델에 ErrorMessages와 SchemeConstants 적용
10. **ReverseSchemeBuilder 리팩토링** - 새로운 도메인 구조에 맞춰 업데이트
11. **도메인 모델 단위 테스트** - CellPosition, SchemeNodeType, SchemeNode, Scheme에 대한 테스트 작성

#### 1.4 인터페이스 및 추상화 정의 ✅

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
- [x] 도메인 레이어 인터페이스 정의 ✅
- [x] 애플리케이션 레이어 인터페이스 정의 ✅
- [x] 인프라 레이어 인터페이스 정의 ✅
- [x] DTO 및 요청/응답 모델 정의 ✅
- [x] 인터페이스 문서화 (XML 주석) ✅

**완료된 작업 상세**:

1. **도메인 레이어 인터페이스 (Domain/Interfaces/)**
   - IRepository.cs: 리포지토리 패턴의 기본 인터페이스 (CRUD 작업)
   - ISchemeRepository.cs: 스키마 리포지토리 인터페이스
   - IDomainService.cs: 도메인 서비스의 기본 인터페이스
   - ISchemeValidationService.cs: 스키마 검증 서비스 인터페이스

2. **애플리케이션 레이어 인터페이스 (Application/Interfaces/)**
   - IExcelReaderService.cs: Excel 파일 읽기 서비스
   - IYamlGeneratorService.cs: YAML 생성 서비스
   - IJsonGeneratorService.cs: JSON 생성 서비스
   - IPostProcessingService.cs: 후처리 서비스
   - ISchemeBuilderService.cs: 스키마 빌더 서비스 (관련 클래스 포함)

3. **인프라 레이어 인터페이스 (Infrastructure/Interfaces/)**
   - IFileSystemService.cs: 파일 시스템 서비스
   - IExcelService.cs: Excel 처리 서비스
   - IConfigurationService.cs: 구성 서비스 (관련 인터페이스 포함)
   - ILoggingService.cs: 로깅 서비스

4. **DTO 및 요청/응답 모델 (Application/DTOs/)**
   - ExcelConversionRequest.cs: Excel 변환 요청 DTO
   - ExcelConversionResponse.cs: Excel 변환 응답 DTO (관련 클래스 포함)
   - SchemeValidationRequest.cs: 스키마 검증 요청 DTO
   - SchemeValidationResponse.cs: 스키마 검증 응답 DTO (관련 클래스 포함)
   - PostProcessingRequest.cs: 후처리 요청 DTO
   - PostProcessingResponse.cs: 후처리 응답 DTO (관련 클래스 포함)

5. **문서화**
   - 모든 인터페이스와 DTO에 XML 주석 추가 완료
   - 메서드, 속성, 클래스에 대한 상세 설명 포함
   - 빌더 패턴, 팩토리 메서드 등 사용 예시 포함

6. **프로젝트 파일 업데이트**
   - ExcelToYamlAddin.csproj에 모든 새 파일 추가 완료

### Phase 2: 핵심 컴포넌트 리팩토링 (2-3주) ✅

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
- [x] SchemeParser를 작은 단위로 분해 ✅
- [x] SchemeNodeFactory 구현 (SchemeNodeBuilder로 구현) ✅
- [x] SchemeValidator 구현 (ParsingContext 내 검증 로직으로 구현) ✅
- [x] ParsingContext 클래스 생성 ✅
- [x] 병합 셀 처리 로직 분리 ✅
- [x] 단위 테스트 작성 ✅

**완료된 작업 상세**:

1. **책임 분리 완료**
   - ISchemeEndMarkerFinder: 스키마 끝 마커 찾기 전담
   - IMergedCellHandler: 병합 셀 처리 전담
   - ISchemeNodeBuilder: 노드 생성 전담
   - ParsingContext: 파싱 컨텍스트 관리

2. **구현된 클래스**
   - SchemeEndMarkerFinder: 스키마 끝 마커 검색 구현
   - MergedCellHandler: 병합 셀 범위 계산 구현
   - SchemeNodeBuilder: 셀에서 노드 생성 구현
   - SchemeParserV2: 리팩토링된 파서 구현
   - SchemeTreeParser: 내부 트리 파싱 로직 분리
   - SchemeParserFactory: 의존성 주입을 위한 팩토리

3. **기존 호환성 유지**
   - 기존 SchemeParser를 래퍼로 변경하여 API 호환성 유지
   - SchemeParsingResult 구조 유지

4. **단위 테스트 작성 완료**
   - SchemeEndMarkerFinderTests: 마커 찾기 테스트
   - MergedCellHandlerTests: 병합 셀 처리 테스트
   - SchemeNodeBuilderTests: 노드 빌더 테스트
   - ParsingContextTests: 컨텍스트 검증 테스트

5. **개선된 점**
   - 순환 복잡도 감소: Parse 메서드가 여러 작은 메서드로 분리됨
   - 테스트 가능성 향상: 모든 컴포넌트가 인터페이스를 통해 목(mock) 가능
   - 유지보수성 향상: 각 클래스가 단일 책임만 가짐
   - 의존성 주입 지원: 팩토리를 통한 유연한 인스턴스 생성

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
- [x] YamlGenerationService 생성 (오케스트레이션) ✅
- [x] NodeTraverser 구현 (순회 로직) ✅
- [x] NodeProcessor 인터페이스 및 구현체들 ✅
- [x] YamlBuilder 구현 (YAML 생성) ✅
- [x] GenerationContext 구현 ✅
- [x] YamlGenerationOptions ValueObject 생성 ✅
- [x] 스택 관리 로직 제거 및 개선 ✅
- [x] 기존 YamlGenerator와 통합 (YamlGeneratorAdapter 구현) ✅
- [x] 각 컴포넌트별 단위 테스트 ✅
  - NodeTraverserTests
  - YamlBuilderTests
  - NodeProcessorTests
  - YamlGeneratorIntegrationTests

### 2.2.1 기존 YamlGenerator 통합 방안

**현재 상황**:
- 기존 YamlGenerator는 복잡한 스택 기반 처리 로직 사용
- 새로운 Generation 컴포넌트들은 명확한 책임 분리와 의존성 주입 기반
- 점진적 마이그레이션이 필요

**완료된 구현**:
- ✅ YamlGenerationService: YAML 생성 오케스트레이션
- ✅ NodeTraverser: 노드 순회 로직 구현
- ✅ NodeProcessor 인터페이스 및 구현체들 (Array, Map, Property, KeyValue, Ignore)
- ✅ YamlBuilder: YAML 구조 생성
- ✅ GenerationContext: 상태 관리 (스택 기반에서 컨텍스트 기반으로 개선)
- ✅ YamlGenerationOptions: 생성 옵션 값 객체
- ✅ YamlGeneratorAdapter: 기존 시스템과의 통합 어댑터
- ✅ 포괄적인 단위 테스트 및 통합 테스트 작성

**주요 개선사항**:
1. **스택 관리 로직 제거**: 복잡한 스택 기반 처리를 GenerationContext를 통한 명확한 상태 관리로 대체
2. **책임 분리**: 각 컴포넌트가 단일 책임 원칙을 따르도록 설계
3. **테스트 가능성**: 모든 컴포넌트에 대한 단위 테스트 작성
4. **점진적 마이그레이션**: YamlGeneratorAdapter를 통한 안전한 전환 경로 제공

**통합 전략**:

1. **Adapter 패턴 적용**
```csharp
// Application/Services/YamlGeneratorAdapter.cs
public class YamlGeneratorAdapter : IYamlGeneratorService
{
    private readonly IYamlGenerationService _newGenerator;
    private readonly bool _useNewGenerator;
    
    public YamlGeneratorAdapter(
        IYamlGenerationService newGenerator,
        IConfiguration configuration)
    {
        _newGenerator = newGenerator;
        _useNewGenerator = configuration.GetValue<bool>("UseNewYamlGenerator", false);
    }
    
    public string Generate(Scheme scheme, IXLWorksheet sheet, ConversionOptions options)
    {
        if (_useNewGenerator)
        {
            // 새로운 생성기 사용
            var yamlOptions = YamlGenerationOptions.FromConfig(options);
            return _newGenerator.GenerateYaml(scheme, sheet, yamlOptions);
        }
        else
        {
            // 기존 생성기 사용 (임시)
            return YamlGenerator.Generate(
                scheme, 
                sheet, 
                options.YamlStyle, 
                2, 
                false, 
                options.ShowEmptyFields);
        }
    }
}
```

2. **기존 코드 점진적 제거**
- Phase 1: Adapter를 통한 새 구현 테스트
- Phase 2: 기존 YamlGenerator의 정적 메서드를 인스턴스 메서드로 변경
- Phase 3: 기존 코드 완전 제거 및 새 구현으로 교체

3. **호환성 레이어**
```csharp
// Application/Services/Generation/LegacyCompatibility.cs
public static class LegacyCompatibility
{
    public static ConversionOptions ToLegacyOptions(YamlGenerationOptions options)
    {
        return new ConversionOptions
        {
            ShowEmptyFields = options.ShowEmptyFields,
            YamlStyle = options.Style,
            MergeKeyPaths = options.MergeKeyPaths.ToList(),
            FlowStylePaths = options.FlowStylePaths.ToList(),
            OutputPath = options.OutputPath,
            OutputFormat = options.PostProcessing.OutputFormat
        };
    }
    
    public static YamlGenerationOptions FromLegacyOptions(ConversionOptions options)
    {
        return YamlGenerationOptions.FromConfig(options);
    }
}

#### 2.3 Ribbon UI 분리 🔄 재설계 필요

**현재 문제점**:
- 2018줄이 넘는 거대한 Ribbon_legacy.cs 클래스 분석 완료
- **⚠️ 현재 MVP 구현 완성도 부족**: 기존 기능의 5% 정도만 구현됨
- **핵심 기능 누락**: 실제 변환 로직, 후처리, 설정 관리 등 대부분 미구현
- **기능적 호환성 부족**: 기존 사용자 워크플로우와 완전히 단절

**Ribbon_legacy.cs 분석 결과**:

1. **핵심 변환 메서드들** (모두 미구현):
   ```csharp
   - PrepareAndValidateSheets(): 시트 검증 및 준비
   - ConvertExcelFile(): 실제 Excel 파일 변환
   - ConvertExcelFileToTemp(): 임시 파일 변환
   - ApplyYamlPostProcessing(): YAML 후처리 파이프라인
   ```

2. **상세 이벤트 핸들러들** (대부분 placeholder):
   ```csharp
   - OnConvertToYamlClick(): 복잡한 YAML 변환 로직 (450줄)
   - OnConvertToXmlClick(): XML 변환 파이프라인 (190줄)
   - OnConvertYamlToJsonClick(): YAML→JSON 변환 (190줄)
   - OnImportYamlClick(): YAML 가져오기 (180줄)
   - OnImportXmlClick(): XML 가져오기 (75줄)
   - OnImportJsonClick(): JSON 가져오기 (55줄)
   ```

3. **설정 관리 시스템**:
   ```csharp
   - Ribbon_Load(): 복잡한 초기화 로직
   - OnSheetPathSettingsClick(): 시트별 경로 설정
   - 체크박스 상태 관리 (EmptyFields, HashGen, AddEmptyYaml)
   - SheetPathManager와 ExcelConfigManager 통합
   ```

4. **진행 상황 관리**:
   ```csharp
   - ProgressForm과 통합된 복잡한 진행률 보고
   - CancellationToken 지원
   - 단계별 상세 메시지 표시
   ```

**재설계 전략**:

### Phase 2.3.1: Legacy 기능 분석 및 매핑 📋 ✅ (완료)

**목표**: 기존 기능을 신규 MVP 아키텍처로 완전 이관

1. **기능 매핑 테이블 작성**:
   ```markdown
   | Legacy 메서드 | 기능 설명 | MVP 위치 | 구현 상태 |
   |--------------|-----------|----------|-----------|
   | PrepareAndValidateSheets | 시트 검증 | RibbonHelpers | ✅ 완료 |
   | ConvertExcelFile | Excel 변환 | ConversionService | ✅ 완료 |
   | ApplyYamlPostProcessing | 후처리 | PostProcessingService | ✅ 완료 |
   | OnImportXmlClick | XML Import | ImportExportService | ✅ 완료 |
   | OnImportYamlClick | YAML Import | ImportExportService | ✅ 완료 |
   | OnImportJsonClick | JSON Import | ImportExportService | ✅ 완료 |
   ```

2. **상태 관리 분석**:
   ```csharp
   // Legacy에서 사용하는 상태들
   private bool includeEmptyFields = false;
   private bool enableHashGen = false;
   private bool addEmptyYamlFields = false;
   private readonly ExcelToYamlConfig config = new ExcelToYamlConfig();
   private Forms.SheetPathSettingsForm settingsForm = null;
   ```

### Phase 2.3.2: 핵심 서비스 레이어 구축 🏗️ ✅ (완료)

**목표**: Legacy 로직을 서비스로 분리

**완료된 작업**:
1. **Presentation/Services 폴더 구조 생성**
   - ConversionService.cs - Excel 변환 관련 로직
   - ImportExportService.cs - Import/Export 기능
   - PostProcessingService.cs - YAML 후처리 기능

2. **Presentation/Helpers 폴더 구조 생성**
   - RibbonHelpers.cs - 공통 유틸리티 메서드들

3. **Ribbon.cs 리팩토링**
   - 1821줄에서 약 1150줄로 감소 (37% 감소)
   - 서비스 의존성 주입 패턴 적용
   - 단일 책임 원칙(SRP) 준수

1. **ConversionOrchestrationService** 구현:
   ```csharp
   public class ConversionOrchestrationService : IConversionOrchestrationService
   {
       public async Task<ConversionResult> ExecuteYamlConversionAsync(ConversionRequest request)
       {
           // PrepareAndValidateSheets 로직 이관
           var sheets = await ValidateAndPrepareSheets(request);
           
           // ConvertExcelFile 로직 이관
           var convertedFiles = await ConvertToYaml(sheets, request.Config);
           
           // ApplyYamlPostProcessing 로직 이관
           var postProcessed = await ApplyPostProcessing(convertedFiles, sheets);
           
           return new ConversionResult { Files = postProcessed };
       }
   }
   ```

2. **SheetValidationService** 구현:
   ```csharp
   public class SheetValidationService : ISheetValidationService
   {
       public async Task<SheetValidationResult> ValidateSheets(IWorkbook workbook)
       {
           // PrepareAndValidateSheets의 검증 로직
           var convertibleSheets = GetConvertibleSheets(workbook);
           var enabledSheets = FilterEnabledSheets(convertibleSheets);
           return new SheetValidationResult { Sheets = enabledSheets };
       }
   }
   ```

3. **PostProcessingOrchestrator** 구현:
   ```csharp
   public class PostProcessingOrchestrator : IPostProcessingOrchestrator
   {
       public async Task<PostProcessingResult> ApplyYamlPostProcessing(
           List<string> yamlFiles, 
           List<Sheet> sheets,
           PostProcessingOptions options)
       {
           // ApplyYamlPostProcessing 로직 완전 이관
           var mergeResult = await ApplyMergeKeyPaths(yamlFiles, sheets);
           var flowResult = await ApplyFlowStyle(yamlFiles, sheets);
           return new PostProcessingResult { MergeCount = mergeResult, FlowCount = flowResult };
       }
   }
   ```

### Phase 2.3.3: 이벤트 핸들러 완전 구현 🎯 ✅ (2025-06-18 완료)

**목표**: 모든 버튼 클릭 이벤트의 완전한 기능 구현

**완료된 작업**:

1. **HandleConvert 함수 통합 구현** ✅
   ```csharp
   private void HandleConvert(string targetFormat)
   {
       // OnConvertToYamlClick, OnConvertToXmlClick, OnConvertYamlToJsonClick 통합
       // - 약 560줄의 중복 코드 제거
       // - 파라미터 기반 분기 처리
       // - 세밀한 단계별 진행률 보고
   }
   ```

2. **HandleImport 함수 통합 구현** ✅
   ```csharp
   private void HandleImport(string fileType)
   {
       // OnImportXmlClick, OnImportYamlClick, OnImportJsonClick 통합
       // - 약 100줄의 중복 코드 제거
       // - 통합된 Import 로직
       // - 파일 타입별 설정 분기
   }
   ```

3. **중복 제거 및 코드 개선** ✅
   - Ribbon.cs: 1821줄 → 약 880줄 (52% 감소)
   - 6개의 개별 함수를 2개의 통합 함수로 리팩토링
   - C# 7.3 호환성 유지 (switch 표현식 → switch 문)
   - 진행률 표시를 ConversionService와 PostProcessingService로 이동

4. **Ribbon.Designer.cs 업데이트** ✅
   - Import 버튼들의 이벤트 핸들러를 wrapper 메서드로 연결
   - VS 디자이너 자동 생성 사용하지 않고 직접 수정


### Phase 3: 후처리 시스템 현대화 (1주)

#### 3.1 후처리 파이프라인 구축 ✅ (완료)

**목표**: 확장 가능한 후처리 시스템

**완료된 작업**:
1. **IPostProcessor 인터페이스 구현**
   - Priority 속성으로 실행 순서 제어
   - CanProcess 메서드로 조건부 실행
   - ProcessAsync 메서드로 비동기 처리

2. **ProcessingPipeline 클래스 구현**
   - 우선순위 기반 프로세서 실행
   - IProgress<T> 지원으로 진행률 보고
   - CancellationToken 지원
   - 포괄적인 에러 처리

3. **PostProcessorBase 추상 클래스 구현**
   - 공통 에러 처리 로직
   - 처리 시간 측정
   - Template Method 패턴 적용

4. **기존 프로세서 리팩토링**
   - YamlMergeProcessor (Priority: 10)
   - YamlFlowStyleProcessor (Priority: 20)
   - JsonFormatterProcessor (Priority: 30)
   - XmlFormatterProcessor (Priority: 30)

5. **PostProcessingServiceV2 구현**
   - 파이프라인 기반 처리
   - 기존 서비스와의 호환성 유지
   - 비동기 처리 지원

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
- [x] ProcessingPipeline 구현 ✅
- [x] PostProcessorBase 추상 클래스 ✅
- [x] YamlMergeProcessor 리팩토링 ✅
- [x] YamlFlowStyleProcessor 리팩토링 ✅
- [x] JsonFormatterProcessor 구현 ✅
- [x] XmlFormatterProcessor 구현 ✅
- [x] 처리 순서 및 우선순위 시스템 ✅

#### 3.2 기존 시스템과의 통합 ✅ (완료)

**목표**: 새로운 파이프라인을 기존 시스템에 통합

**완료된 작업**:
1. **Ribbon.cs 업데이트**
   - PostProcessingService → PostProcessingServiceV2로 전환
   - ApplyYamlPostProcessing → ApplyYamlPostProcessingAsync로 변경
   - async/await 패턴 적용 (Task.Wait 사용)

2. **호환성 유지**
   - 기존 메서드 시그니처와 호환되는 래퍼 메서드 구현
   - 기존 로직과 동일한 동작 보장
   - 진행률 보고 기능 유지

3. **프로젝트 파일 업데이트**
   - 모든 새로운 파일을 ExcelToYamlAddin.csproj에 추가
   - 올바른 컴파일 순서 보장

**To-Do List**:
- [x] ConversionService에서 PostProcessingServiceV2 사용하도록 수정 ✅
- [x] 기존 PostProcessingService와의 호환성 확인 ✅
- [x] 통합 테스트 및 검증 ✅

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

### Week 1-2: 기반 구조 ✅
- [x] 프로젝트 구조 재구성 ✅
- [x] 상수 및 설정 중앙화 ✅
- [x] 도메인 모델 정의 ✅
- [x] 인터페이스 계층 구축 ✅

### Week 3-5: 핵심 리팩토링 ✅
- [x] SchemeParser 개선 ✅
- [x] YamlGenerator 분해 ✅
- [x] Ribbon UI 분리 ✅
- [x] 단위 테스트 작성 ✅

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