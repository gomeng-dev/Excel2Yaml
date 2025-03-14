# SheetAnalyzer 클래스 가이드

## 개요

`SheetAnalyzer` 클래스는 Excel2YAML 애드인에서 엑셀 워크북 내의 시트들을 분석하여 YAML로 변환 가능한 시트를 식별하는 역할을 담당합니다. 이 클래스는 시트 이름의 특정 접두사를 통해 자동 변환 대상을 식별하는 핵심 로직을 제공합니다.

## 주요 기능

### 1. 변환 가능한 시트 식별

`SheetAnalyzer` 클래스는 엑셀 워크북 내의 모든 시트를 검사하여 YAML로 자동 변환할 대상 시트를 찾아냅니다. 이를 통해 사용자는 특정 시트만 선택적으로 변환할 수 있습니다.

### 2. 자동 변환 규칙 적용

시트 이름이 `!` 접두사로 시작하는 경우 해당 시트는 자동 변환 대상으로 간주됩니다. 이 접두사는 시트를 YAML로 변환할 때 사용하는 명시적인 마커 역할을 합니다.

## 구현 상세

### `GetConvertibleSheets` 메서드

```csharp
public static List<Worksheet> GetConvertibleSheets(Workbook workbook)
```

이 메서드는 워크북 내의 모든 시트를 순회하면서 변환 가능한 시트들을 식별하여 목록으로 반환합니다.

- **입력**: 엑셀 `Workbook` 객체
- **출력**: 변환 가능한 `Worksheet` 객체들의 목록
- **동작 과정**:
  1. 워크북 내의 모든 시트를 순회
  2. 각 시트에 대해 `IsSheetConvertible` 메서드를 통해 변환 가능 여부 확인
  3. 변환 가능한 시트들만 결과 목록에 추가

### `IsSheetConvertible` 메서드

```csharp
private static bool IsSheetConvertible(Worksheet sheet)
```

이 메서드는 개별 시트가 변환 가능한지 판별합니다.

- **입력**: 엑셀 `Worksheet` 객체
- **출력**: 변환 가능 여부를 나타내는 불리언 값
- **판별 기준**: 시트 이름이 `!` 접두사로 시작하는지 여부

## 사용 예시

### 자동 변환 대상 시트 지정하기

엑셀 파일에서 특정 시트를 자동 변환 대상으로 지정하려면:

1. 시트 이름 앞에 `!` 기호를 붙입니다.
   - 예: `Data` → `!Data`
2. Excel2YAML 도구를 실행하면 `!` 접두사가 있는 시트들만 자동으로 처리됩니다.

### 코드에서의 활용 예시

```csharp
// 워크북에서 변환 가능한 시트 목록 가져오기
var workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
var convertibleSheets = SheetAnalyzer.GetConvertibleSheets(workbook);

// 변환 가능한 시트 목록 출력
foreach (var sheet in convertibleSheets)
{
    Console.WriteLine($"변환 대상 시트: {sheet.Name}");
}
```

## 주의사항

- 시트 이름에 `!` 접두사를 추가할 때는 엑셀의 시트 이름 규칙을 준수해야 합니다.
- 자동 변환을 원하지 않는 시트에는 `!` 접두사를 사용하지 않도록 주의하세요.
- 시트 이름의 `!` 접두사는 YAML 파일 이름에는 포함되지 않습니다. 변환 결과 파일 이름은 접두사를 제외한 시트 이름으로 생성됩니다.

## 활용 시나리오

1. **다중 시트 엑셀 파일 처리**: 여러 시트가 있는 엑셀 파일에서 특정 시트만 변환하고 싶을 때 유용합니다.
2. **자동화된 데이터 파이프라인**: 시트 이름 규칙을 통해 자동 변환 대상을 지정함으로써 데이터 처리 과정을 자동화할 수 있습니다.
3. **테스트 및 프로덕션 데이터 구분**: 동일한 엑셀 파일 내에서 테스트용 시트와 실제 사용할 시트를 구분하여 관리할 수 있습니다. 