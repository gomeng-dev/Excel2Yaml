# Excel2Yaml 스키마 구조 분석

## 전체 시스템 아키텍처

### 1. 스키마 파싱 시스템 (SchemeParser.cs)

**기본 규칙:**
- **1행**: 주석 행 (COMMENT_ROW_NUM = 0, 비워둠)
- **2행부터**: 스키마 시작 (_schemeStartRow)
- **$scheme_end 마커**: 스키마 종료 표시
- **그 이후**: 실제 데이터 영역

**파싱 알고리즘:**
```
Parse(parent, rowNum, startCol, endCol):
  for cellNum in range(startCol, endCol):
    - 빈 셀이나 '^' 마커는 건너뜀
    - 셀 값으로 SchemeNode 생성
    - 컨테이너 노드(MAP, ARRAY)인 경우:
      * 병합된 셀 영역 확인
      * 다음 행(rowNum + 1)에서 자식 노드 재귀 파싱
      * 병합 영역이 자식 파싱 범위가 됨
```

### 2. 스키마 노드 타입 (SchemeNode.cs)

**노드 타입:**
- `PROPERTY`: 단순 속성 (예: Name, Type)
- `KEY`: 동적 키 ($key)
- `VALUE`: 동적 값 ($value)
- `MAP`: 객체 (${}마커)
- `ARRAY`: 배열 ($[]마커)
- `IGNORE`: 무시 마커 (^)

**타입 결정 규칙:**
```
- "$[]" 포함 → ARRAY
- "$key" 포함 → KEY  
- "$value" 포함 → VALUE
- "${}" 포함 → MAP
- "$" 없음 → PROPERTY
```

**컨테이너 노드:** MAP, ARRAY, KEY (다음 행에 자식 요소)

### 3. 스키마 구조 예시

```
행2: ${}                          // 루트 MAP (전체 열 병합)
행3: ^    Item$[]                 // 배열 마커 (B열부터 병합)
행4: ^    ^      ${}              // 배열 항목 객체 (C열부터 병합)
행5: ^    ^      _ID  Name  ShortCut${}  Condition${}  ...
행6: ^    ^      ^    ^     _Type  _Value  _Type  _Value  ...
행7: $scheme_end                  // 스키마 종료 (전체 열 병합)
```

### 4. 속성 처리 규칙 (YamlToXmlConverter.cs 기준)

**속성 명명 규칙:**
- XML 속성은 `_속성명` 형태로 표현
- 예: `<ShortCut Type="value">` → `_Type` 헤더

**구조 변환:**
```xml
<ShortCut Type="GameStart" Value="1"/>
```
↓ Excel 스키마
```
ShortCut${}
_Type  _Value
```

### 5. 병합 셀 규칙

**병합 대상:**
1. **루트 객체 마커**: 전체 MaxColumns까지
2. **배열 마커**: 시작 컬럼부터 MaxColumns까지  
3. **중첩 객체**: 해당 객체의 속성/자식 컬럼 수만큼
4. **스키마 종료**: 전체 MaxColumns까지

**병합 계산:**
- `GetNestedColumnCount()`: 속성 수 + 자식 요소 수 + 텍스트 내용(있다면 +1)
- 최소 1개 컬럼은 보장

### 6. 무한 행 확장 지원

**동적 구조 처리:**
- SchemeParser는 `$scheme_end` 마커를 찾을 때까지 무한히 행을 검사
- 컨테이너 노드는 재귀적으로 자식을 파싱 (`rowNum + 1 < _schemeEndRowNum`)
- 복잡한 중첩 구조도 자동으로 처리

**행 제한 없음:**
```csharp
// SchemeParser.Parse() - 무한 확장 가능
if (rowNum + 1 < _schemeEndRowNum) {
    Parse(child, rowNum + 1, firstCellInRange, lastCellInRange);
}
```

### 7. 데이터 생성 시스템 (Generator.cs)

**처리 방식:**
- 루트 노드 타입에 따라 MAP/ARRAY 처리 분기
- 각 데이터 행에 대해 스키마 노드 순회
- 노드 타입별 값 추출 및 JSON 구조 생성

**키 결정 로직:**
1. KEY 노드: 실제 셀 값 사용
2. PROPERTY 노드: 스키마에 정의된 키 사용
3. 동적 키: `GetKey(row)` 메서드로 런타임 결정

### 8. 개발 시 주의사항

**XmlToExcelConverter 개선 포인트:**
1. **속성 처리**: `_속성명` 형태로 정확히 표현
2. **병합 계산**: 모든 가능한 속성과 자식 요소 고려
3. **무한 확장**: 행 수 제한 없이 동적 구조 생성
4. **순서 보장**: XML 태그 순서대로 컬럼 배치
5. **타입 정확성**: 속성 있는 요소는 복잡한 객체로만 분류

**구조 병합 규칙:**
- 같은 이름 요소의 모든 구조 수집
- 속성과 자식 요소 통합
- 네임스페이스 속성 제외
- 순서 보장 위해 List 사용

이 분석을 바탕으로 XML→Excel 변환 시 정확한 스키마 구조를 생성할 수 있습니다.