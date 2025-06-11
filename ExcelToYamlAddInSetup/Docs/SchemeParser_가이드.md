# SchemeParser 클래스 가이드

## 개요

`SchemeParser` 클래스는 Excel2YAML 애드인에서 엑셀 시트의 스키마 구조를 분석하고 파싱하는 핵심 클래스입니다. 이 클래스는 엑셀 시트의 첫 행을 스키마 정의로 해석하여 계층적 데이터 구조를 생성하고, 이를 YAML 변환 과정에서 활용할 수 있는 트리 구조로 변환합니다.

## 주요 기능

### 1. 스키마 영역 식별

`SchemeParser`는 엑셀 시트에서 스키마 정의 영역과 실제 데이터 영역을 구분합니다. 스키마의 끝을 표시하는 특수 마커(`$scheme_end`)를 사용하여 스키마 영역의 경계를 식별합니다.

### 2. 계층적 스키마 구조 파싱

엑셀 시트의 첫 행에 정의된 필드명을 기반으로 계층적인 스키마 구조를 구성합니다. 점(.) 표기법, 배열 표시([]), 맵 표시({}) 등의 특수 표기를 해석하여 중첩된 객체와 배열을 포함하는 트리 구조를 생성합니다.

### 3. 병합된 셀 처리

스키마 정의 영역에서 병합된 셀이 있는 경우, 이를 적절히 처리하여 올바른 계층 구조를 파싱합니다. 병합된 셀은 보통 상위 컨테이너 노드를 나타내는 데 사용됩니다.

### 4. 데이터 영역 경계 확인

스키마 끝 마커를 찾아 실제 데이터가 시작되는 행과 끝나는 행을 정확히 파악합니다. 이 정보는 이후 데이터 처리 과정에서 중요하게 활용됩니다.

## 구현 상세

### `SchemeParser` 생성자

```csharp
public SchemeParser(IXLWorksheet sheet)
```

이 생성자는 엑셀 시트를 입력받아 SchemeParser 인스턴스를 초기화합니다.

- **입력**: ClosedXML 라이브러리의 `IXLWorksheet` 객체
- **동작 과정**:
  1. 엑셀 시트에서 스키마 끝 마커(`$scheme_end`)를 찾습니다.
  2. 스키마 정의 시작 행(기본적으로 2번째 행)을 설정합니다.
  3. 스키마 영역의 첫 번째 셀과 마지막 셀을 식별합니다.

### `SchemeParsingResult` 클래스

스키마 파싱 결과를 담는 내부 클래스로, 다음 속성들을 포함합니다:

- **Root**: 파싱된 스키마의 루트 노드
- **ContentStartRowNum**: 실제 데이터가 시작되는 행 번호(스키마 끝 마커 다음 행)
- **EndRowNum**: 데이터가 끝나는 행 번호
- **LinearNodes**: 트리 구조의 노드들을 선형 목록으로 변환한 컬렉션

### `Parse` 메서드

```csharp
public SchemeParsingResult Parse()
```

이 메서드는 스키마 파싱 프로세스를 시작하고 결과를 반환합니다.

- **출력**: `SchemeParsingResult` 객체
- **동작 과정**:
  1. 재귀적 파싱 메서드를 호출하여 스키마 트리 구조를 생성합니다.
  2. 루트 노드가 없을 경우 기본 배열 타입 노드를 생성합니다.
  3. 콘텐츠 시작 행과 종료 행 정보를 포함한 결과 객체를 반환합니다.

### 재귀적 `Parse` 메서드

```csharp
private SchemeNode Parse(SchemeNode parent, int rowNum, int startCellNum, int endCellNum)
```

이 메서드는 지정된 행과 열 범위를 처리하여 스키마 노드를 재귀적으로 파싱합니다.

- **입력**: 부모 노드, 현재 행 번호, 시작 열 번호, 종료 열 번호
- **출력**: 파싱된 부모 노드
- **동작 과정**:
  1. 지정된 범위의 각 셀을 순회합니다.
  2. 셀 값을 기반으로 `SchemeNode` 객체를 생성합니다.
  3. 컨테이너 타입(MAP, ARRAY) 노드일 경우 다음 행을 재귀적으로 파싱합니다.
  4. 병합된 셀 영역을 처리하여 올바른 계층 구조를 유지합니다.

### `ContainsEndMarker` 메서드

```csharp
private bool ContainsEndMarker(IXLRow row)
```

이 메서드는 특정 행이 스키마 끝 마커(`$scheme_end`)를 포함하는지 확인합니다.

- **입력**: 엑셀 행 객체
- **출력**: 마커 포함 여부를 나타내는 불리언 값
- **동작 과정**: 행의 첫 번째 셀이 특수 마커 문자열과 일치하는지 검사합니다.

## 사용 예시

### 스키마 파싱 구현

```csharp
// 엑셀 시트 객체 획득
IXLWorksheet sheet = workbook.Worksheet("DataSheet");

// SchemeParser 인스턴스 생성
SchemeParser parser = new SchemeParser(sheet);

// 스키마 파싱 수행
SchemeParser.SchemeParsingResult result = parser.Parse();

// 파싱 결과 활용
SchemeNode rootNode = result.Root;
int dataStartRow = result.ContentStartRowNum;
int dataEndRow = result.EndRowNum;

// 파싱된 스키마 트리 활용하기
foreach (var node in result.GetLinearNodes())
{
    Console.WriteLine($"Node: {node.Key}, Type: {node.NodeType}");
}
```

### 스키마 끝 마커 사용하기

엑셀 시트에서 스키마 영역의 끝을 표시하려면:

```
| name | age | job |
|------|-----|-----|
| $scheme_end |  |  |
| 홍길동 | 30 | 개발자 |
| 김철수 | 25 | 디자이너 |
```

위 예시에서 `$scheme_end`가 포함된 행은 스키마의 끝을 표시하며, 그 다음 행부터 실제 데이터가 시작됩니다.

## 주의사항

- 스키마 끝 마커(`$scheme_end`)는 반드시 첫 번째 열에 위치해야 합니다.
- 스키마 끝 마커가 없는 경우 예외가 발생합니다.
- 스키마 정의는 기본적으로 엑셀 시트의 두 번째 행(인덱스 2)에서 시작합니다.
- 병합된 셀을 사용할 때는 계층 구조가 올바르게 표현되었는지 확인해야 합니다.

## 활용 시나리오

1. **복잡한 데이터 구조 변환**: 중첩된 객체와 배열이 포함된 복잡한 데이터 구조를 정의하고 변환할 때 유용합니다.
2. **스키마와 데이터 분리**: 엑셀 시트에서 스키마 정의와 실제 데이터를 명확하게 구분하여 관리할 수 있습니다.
3. **계층적 구조 시각화**: 엑셀에서 계층적 데이터 구조를 시각적으로 표현하고 이를 프로그래밍 가능한 형태로 변환합니다. 