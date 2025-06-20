# Excel2YAML 사용설명서
![Example](https://github.com/user-attachments/assets/2fef96dc-63f2-4284-8be2-e6143d24c9fc)

## 목차
- [1. 개요](#1-개요)
- [2. 이 프로그램의 용도](#2-이-프로그램의-용도)
- [3. 사용 가능한 기능들](#3-사용-가능한-기능들)
- [4. 엑셀 작성 방법](#4-엑셀-작성-방법)

## 1. 개요

Excel2YAML은 Microsoft Excel 애드인으로, 엑셀 시트의 데이터를 구조화된 YAML 형식으로 변환하는 도구입니다. 이 도구는 특히 게임 개발, 설정 관리, 데이터 모델링 등 계층적 데이터 구조가 필요한 분야에서 매우 유용합니다.

Excel2YAML은 엑셀의 사용 편의성과 YAML의 가독성 및 계층적 구조화 능력을 결합하여, 복잡한 데이터도 직관적으로 관리하고 효율적으로 변환할 수 있게 해줍니다. 특히 일반적인 YAML 변환과 달리 **속성의 순서를 보존**하는 고유한 기능을 제공합니다.

### 주요 특징
- **직관적인 데이터 구조 설계**: 엑셀 시트에서 데이터 구조를 쉽게 설계할 수 있습니다.
- **계층적 데이터 지원**: 중첩된 객체, 목록, 키-값 쌍 등 복잡한 데이터 구조를 표현할 수 있습니다.
- **선택적 시트 변환**: 시트 이름 앞에 `!` 기호를 추가하여 변환할 시트를 지정할 수 있습니다.
- **사용하기 쉬운 UI**: Excel 내에서 바로 접근할 수 있는 버튼과 메뉴를 제공합니다.
- **다양한 출력 옵션**: 빈 필드 포함 여부, 출력 형식 등 여러 설정을 제공합니다.
- **데이터 병합 기능**: 동일한 ID를 가진 항목을 자동으로 병합할 수 있습니다.

> **참고:** Excel2YAML은 Microsoft Excel 2010 이상 버전에서 작동하며, 처음 설치 후에는 Excel에서 애드인을 활성화해야 합니다.

## 2. 이 프로그램의 용도

Excel2YAML은 엑셀에서 작성한 데이터를 구조화된 YAML 형식으로 변환하여 다양한 실무 환경에서 활용할 수 있게 해주는 도구입니다. 이 프로그램은 다음과 같은 상황에서 특히 유용합니다.

### 주요 활용 분야

#### 1. 게임 개발 데이터 관리
게임 개발 시 필요한 다양한 게임 요소 데이터(캐릭터, 아이템, 스킬, 퀘스트 등)를 엑셀에서 편집하고 게임 엔진에서 사용할 수 있는 YAML 파일로 변환할 수 있습니다.

**예시:** 게임 아이템 정보, 캐릭터 능력치, 레벨업 테이블 등을 엑셀로 관리하고 YAML로 변환

#### 2. 시스템 설정 파일 생성
애플리케이션 설정, 환경 구성, 시스템 파라미터 등을 엑셀에서 쉽게 편집하고 YAML 설정 파일로 변환할 수 있습니다.

**예시:** 웹 애플리케이션 설정, 서버 환경 변수, 사용자 권한 설정 등

#### 3. 다국어 리소스 관리
여러 언어의 텍스트 리소스를 엑셀에서 관리하고 각 언어별 YAML 파일로 내보낼 수 있습니다.

**예시:** 앱이나 웹사이트의 다국어 지원 텍스트, 오류 메시지, 도움말 텍스트 등

#### 4. 데이터 모델 설계 및 변환
복잡한 데이터 모델을 엑셀에서 시각적으로 설계하고 이를 YAML 형식으로 변환할 수 있습니다.

**예시:** API 응답 형식 모델링, 데이터베이스 초기 데이터 생성, 테스트 데이터 준비 등

#### 5. 문서 자동화
구조화된 문서나 보고서 템플릿을 엑셀에서 관리하고 YAML로 변환하여 자동 문서 생성 시스템에 활용할 수 있습니다.

**예시:** 프로젝트 명세서, 작업 지시서, 제품 카탈로그 등

### Excel2YAML을 사용하면 좋은 경우
- **비개발자가 데이터를 관리해야 할 때**: 익숙한 엑셀 환경에서 데이터를 편집하고 구조화된 출력을 얻을 수 있습니다.
- **복잡한 계층 구조의 데이터가 필요할 때**: 중첩된 객체와 목록을 쉽게 표현할 수 있습니다.
- **데이터 형식이 자주 변경될 때**: 엑셀에서 스키마를 쉽게 수정하고 즉시 변환할 수 있습니다.
- **대량의 구조화된 데이터를 다룰 때**: 엑셀의 필터링, 정렬 등 기능을 활용해 대량 데이터를 효율적으로 관리할 수 있습니다.
- **특정 형식으로 데이터를 내보내야 할 때**: 엑셀 데이터를 시스템에서 사용할 수 있는 YAML 형식으로 변환할 수 있습니다.

> **활용 팁:** Excel2YAML은 데이터 작성을 엑셀에서 하고 결과물은 YAML로 얻고자 하는 모든 상황에서 유용합니다. 특히 팀 내 기술 지식이 다양한 구성원들이 함께 데이터를 관리할 때 효과적입니다. 엑셀에 익숙한 비개발자도 복잡한 데이터 구조를 쉽게 작성할 수 있습니다.

## 3. 사용 가능한 기능들

Excel2YAML은 다양한 기능을 제공하여 엑셀 데이터를 YAML로 변환하는 과정을 유연하고 효율적으로 만들어 줍니다. 여기서는 사용자가 실제로 활용할 수 있는 주요 기능들을 소개합니다.

### 3.1 엑셀 시트 YAML 변환
엑셀 시트의 데이터를 YAML 파일로 변환하는 기본 기능입니다.

#### 사용 방법:
1. Excel에서 데이터가 포함된 엑셀 파일을 엽니다.
2. Excel2YAML 탭을 클릭합니다.
3. "YAML로 변환" 버튼을 클릭합니다.
4. YAML 파일을 저장할 위치를 선택합니다.
5. 변환이 완료되면 결과 메시지가 표시됩니다.

이 기능은 엑셀 시트에 정의된 스키마(구조)에 따라 데이터를 YAML 형식으로 변환합니다. 스키마 정의 방법은 [4. 엑셀 작성 방법](#4-엑셀-작성-방법) 섹션에서 자세히 설명합니다.

### 3.2 자동 변환 시트 지정
Excel2YAML은 `!` 기호가 접두사로 붙은 시트만 변환할 수 있습니다. 이 기호는 어떤 시트를 YAML로 변환할지 지정하는 필수 표시입니다.

#### 사용 방법:
1. YAML로 변환하려는 시트 이름 앞에 느낌표(`!`) 기호를 추가합니다.
2. 예: "데이터" → "!데이터"
3. "YAML로 변환" 버튼을 클릭하면 `!` 기호가 붙은 시트만 변환됩니다.
4. 느낌표가 없는 시트는 변환되지 않습니다.

> **참고:** YAML 파일 이름은 시트 이름에서 `!` 기호를 제외한 이름으로 생성됩니다. 예를 들어 "!캐릭터데이터" 시트는 "캐릭터데이터.yaml" 파일로 저장됩니다.

> **주의:** 변환하고자 하는 모든 시트에는 반드시 `!` 기호를 추가해야 합니다. 이 표시가 없는 시트는 변환 대상에서 제외됩니다.

### 3.3 시트별 출력 경로 설정
각 시트별로 다른 저장 경로를 지정할 수 있어, 변환된 YAML 파일을 체계적으로 관리할 수 있습니다.

#### 사용 방법:
1. Excel2YAML 탭에서 "시트 경로 설정" 버튼을 클릭합니다.
2. 대화상자에서 시트 목록을 확인합니다.
3. 각 시트별로 저장 경로를 지정합니다.
4. "저장" 버튼을 클릭하여 설정을 저장합니다.

이 기능은 프로젝트에서 여러 종류의 데이터를 관리할 때 유용합니다. 예를 들어, 아이템 데이터는 "Items" 폴더에, 캐릭터 데이터는 "Characters" 폴더에 자동으로 저장되도록 설정할 수 있습니다.

### 3.4 변환 옵션 설정
YAML 변환 과정에서 적용할 수 있는 옵션을 제공합니다.

#### 제공 옵션:
- **빈 YAML 필드 추가**: YAML에 빈 필드를 추가할지 여부를 설정합니다. 이 옵션은 빈 객체나 배열을 출력에 포함할지 결정합니다.

#### 사용 방법:
1. Excel2YAML 탭의 "설정" 그룹에서 옵션을 토글(클릭)합니다.
2. 활성화된 옵션은 체크 표시되며, 다음 변환 작업부터 적용됩니다.

> **팁:** 데이터 구조의 일관성을 유지하기 위해 "빈 필드 포함" 옵션을 활성화하는 것이 도움이 될 수 있습니다. 특히 다른 시스템에서 특정 필드가 항상 존재해야 하는 경우에 유용합니다.

### 3.5 YAML 항목 병합 기능
동일한 식별자(ID)를 가진 항목들을 자동으로 하나로 합치는 기능입니다. 이는 여러 행의 데이터를 하나의 구조로 통합할 때 유용합니다.

#### 설정 방법:
1. Excel2YAML 탭에서 "시트 경로 설정" 버튼을 클릭합니다.
2. 시트 설정 대화상자가 열리면 표 형태의 설정 화면이 나타납니다.
3. 설정하려는 시트의 행에서 다음 네 열에 값을 입력합니다:
   - **ID 경로**: 항목을 구분하는 고유 ID 위치 (예: `id`)
   - **병합 경로**: 합칠 데이터 위치들, 쉼표로 구분 (예: `items,stats`)
   - **키 경로 전략**: 특정 필드의 병합 방식, 콜론으로 구분 (예: `level:merge`)
   - **배열 필드 경로**: 배열 항목의 병합 방식 (예: `skills:append`)
4. "저장" 버튼을 클릭하여 설정을 저장합니다.

#### 시트 설정 화면 예시:
| 시트 이름 | 경로 | 빈 필드 포함 | ID 경로 | 병합 경로 | 키 경로 전략 | 배열 필드 경로 | 기타 설정... |
|-----------|------|-------------|----------|-----------|--------------|---------------|-------------|
| !캐릭터데이터 | Data/Characters/ | ✓ | id | items,stats | level:merge | skills:append | ... |

#### 예시:
위 설정대로 입력했을 때 (ID 경로: `id`, 병합 경로: `items,stats`, 키 경로 전략: `level:merge`, 배열 필드 경로: `skills:append`):

병합 전 데이터:
```yaml
- id: 1
  name: "전사"
  level: 5
  stats:
    strength: 10
    agility: 5
  skills:
    - name: "베기"
      damage: 20
- id: 1
  level: 6
  stats:
    defense: 8
    health: 100
  skills:
    - name: "방패 방어"
      block: 15
  items:
    - name: "검"
      power: 25
```

병합 후 결과:
```yaml
- id: 1
  name: "전사"
  level: 6
  stats:
    strength: 10
    agility: 5
    defense: 8
    health: 100
  skills:
    - name: "베기"
      damage: 20
    - name: "방패 방어"
      block: 15
  items:
    - name: "검"
      power: 25
```

> **설명:**
> - **ID 경로**: `id`가 동일한 두 항목이 병합되었습니다.
> - **병합 경로**: `items,stats`가 지정되어 두 항목의 내용이 합쳐졌습니다.
> - **키 경로 전략**: `level:merge`로 지정되어 level 필드가 나중 값(6)으로 대체되었습니다.
> - **배열 필드 경로**: `skills:append`로 지정되어 skills 배열의 모든 항목이 유지되었습니다.
> - 지정되지 않은 `name` 필드는 자동으로 유지됩니다.

> **참고:** 병합 기능에 대한 자세한 설명과 추가 예시는 [YAML 항목 병합 설정](YAML_후처리_가이드.md#yaml-merge) 섹션을 참조하세요.

### 3.6 YAML 표시 형식 설정
YAML 출력의 표시 형식을 사용자가 설정할 수 있습니다. 특히 특정 항목을 한 줄로 간결하게 표시하는 옵션이 유용합니다.

#### 설정 방법:
1. Excel2YAML 탭에서 "시트 경로 설정" 버튼을 클릭합니다.
2. 시트 설정 대화상자가 열리면 표 형태의 설정 화면이 나타납니다.
3. 설정하려는 시트의 행에서 다음 두 열에 값을 입력합니다:
   - **흐름 필드**: 한 줄로 표시할 항목 이름들(쉼표로 구분)
   - **흐름 스타일 항목 필드**: 목록 내 각 항목을 한 줄로 표시할 항목 이름들(쉼표로 구분)
4. "저장" 버튼을 클릭하여 설정을 저장합니다.

#### 시트 설정 화면 예시:
| 시트 이름 | 경로 | 빈 필드 포함 | 흐름 필드 | 흐름 스타일 항목 필드 | 기타 설정... |
|-----------|------|-------------|-----------|----------------------|-------------|
| !캐릭터데이터 | Data/Characters/ | ✓ | stats,position | abilities,inventory | ... |

#### 예시:
위 설정대로 입력했을 때 (흐름 필드: `stats,position`, 흐름 스타일 항목 필드: `abilities,inventory`):

적용 전:
```yaml
character:
  name: 플레이어
  stats:
    strength: 10
    dexterity: 8
  position:
    x: 100
    y: 200
  abilities:
    - name: 화염구
      damage: 50
    - name: 회복
      healing: 40
  inventory:
    - id: 101
      name: 검
    - id: 102
      name: 방패
```

적용 후:
```yaml
character:
  name: 플레이어
  stats: { strength: 10, dexterity: 8 }
  position: { x: 100, y: 200 }
  abilities:
    - { name: 화염구, damage: 50 }
    - { name: 회복, healing: 40 }
  inventory:
    - { id: 101, name: 검 }
    - { id: 102, name: 방패 }
```

> **설명:**
> - **흐름 필드**에 `stats,position`을 입력하면 해당 객체들이 한 줄로 표시됩니다.
> - **흐름 스타일 항목 필드**에 `abilities,inventory`를 입력하면 해당 배열 내 각 항목이 한 줄로 표시됩니다.
> - 두 설정을 동시에 적용하면 위와 같이 여러 요소가 간결하게 표시됩니다.

> **참고:** 표시 형식 설정에 대한 자세한 설명과 예시는 [YAML 항목 표시 형식 설정](YAML_후처리_가이드.md#yaml-flow) 섹션을 참조하세요.

> **주의사항:** YAML 변환 전에 항상 데이터와 설정을 확인하세요. 특히 대용량 데이터를 처리할 때는 변환에 시간이 걸릴 수 있으니 인내심을 가지고 기다려 주세요.

## 4. 엑셀 작성 방법

Excel2YAML을 효과적으로 사용하기 위해서는 엑셀 시트에 데이터를 특정 구조로 작성해야 합니다. 이 섹션에서는 변환 가능한 엑셀 시트를 작성하는 방법을 단계별로 설명합니다.

### 4.1 스키마 정의 기본 원칙

Excel2YAML에서는 스키마를 정의하여 엑셀 데이터의 구조를 YAML로 변환할 수 있습니다. 스키마는 데이터의 계층 구조와 형식을 결정합니다.

#### 기본 원칙:
- **상단 행들에 스키마 정의**: 시트의 상단 1~3개 행에 스키마 구조를 정의합니다.
- **중간 부분은 실제 데이터**: 스키마 정의 이후부터 실제 데이터를 입력합니다.
- **맨 아래 행에 스키마 종료 표시**: 마지막 행에는 반드시 `$scheme_end` 표시를 전체 열을 병합하여 포함해야 합니다.
- **특수 마커 사용**: 객체, 배열 등을 표현하기 위한 특수 마커를 사용합니다.

> **중요:** 스키마의 시작과 끝을 명확히 구분해야 합니다. 상단에 스키마 구조를 정의하고, 마지막 행은 반드시 `$scheme_end`로 전체 열을 병합하여 붉은색으로 표시해야 합니다.

> **중요:** 시트 이름 앞에 `!` 기호를 붙여야 해당 시트가 변환 대상으로 인식됩니다.

### 4.2 스키마 마커

스키마를 정의할 때 다음과 같은 특수 마커를 사용하여 데이터 구조를 표현합니다:

| 마커 | 의미 | 설명 |
|------|------|------|
| `${}` | 객체(Object) | 중괄호로 표시된 객체를 의미합니다. 여러 속성을 포함할 수 있습니다. |
| `$[]` | 배열(Array) | 대괄호로 표시된 배열을 의미합니다. 여러 항목을 순서대로 포함할 수 있습니다. |
| `$key` | 키(Key) | YAML의 키 부분을 정의합니다. 값과 쌍을 이룹니다. 부모가 객체일 때는 객체의 키를, 배열일 때는 각 배열 항목의 키를 나타냅니다. |
| `$value` | 값(Value) | YAML의 값 부분을 정의합니다. 키와 쌍을 이룹니다. 부모가 객체일 때는 객체의 값을, 배열일 때는 각 배열 항목의 값을 나타냅니다. |
| `^` | 무시(Ignore) | 변환 시 해당 셀을 무시합니다. |
| `$scheme_end` | 스키마 종료 | 스키마의 끝을 표시합니다. 반드시 마지막 행에 전체 열이 병합된 형태로 포함되어야 합니다. |

> **주의:** 마커는 반드시 셀의 맨 뒤에 위치해야 하며, 마커와 이름 사이에 공백이 없어야 합니다. 예: `character${}`

> **키와 값 처리에 대한 중요 정보:**
> - **부모가 객체(${})**일 때:
>   - `$key`는 객체 내 속성의 이름으로 사용됩니다.
>   - `$value`는 해당 키에 대응하는 값으로 사용됩니다.
>   - 결과 YAML 구조: `key: value`
> - **부모가 배열($[])**일 때:
>   - `$key`와 `$value`는 배열 항목 내 단일 속성을 나타냅니다.
>   - 결과 YAML 구조: `- key: value` 형태의 배열 항목이 됩니다.
> - **중첩 구조**의 경우, 각 계층별 부모-자식 관계에 따라 키-값 쌍이 적절히 구성됩니다.

### 4.3 기본 스키마 작성법

첨부된 이미지를 기반으로 실제 스키마 작성법을 설명합니다. 모든 스키마는 다음과 같은 공통 구조를 가집니다:

#### 1. 기본 스키마 구조

<table class="excel-grid" style="width: 100%;">
    <thead>
        <tr class="excel-row">
            <th class="excel-header" colspan="4" style="background-color: #00CC00;">$[]</th>
        </tr>
    </thead>
    <tbody>
        <tr class="excel-row">
            <td class="excel-cell">^</td>
            <td class="excel-cell" colspan="3" style="background-color: #CCFFCC;">${}</td>
        </tr>
        <tr class="excel-row">
            <td class="excel-cell" style="width: 10%;">^</td>
            <td class="excel-cell" style="width: 30%;">Id</td>
            <td class="excel-cell" style="width: 30%;">Name</td>
            <td class="excel-cell" style="width: 30%;">Value</td>
        </tr>
        <tr class="excel-row" style="background-color: #ff0000; color: white;">
            <td class="excel-cell" colspan="4" style="text-align: center;">$scheme_end</td>
        </tr>
        <tr class="excel-row">
            <td class="excel-cell">^</td>
            <td class="excel-cell">1</td>
            <td class="excel-cell">아이템1</td>
            <td class="excel-cell">100</td>
        </tr>
        <tr class="excel-row">
            <td class="excel-cell">^</td>
            <td class="excel-cell">2</td>
            <td class="excel-cell">아이템2</td>
            <td class="excel-cell">200</td>
        </tr>
    </tbody>
</table>

**YAML 결과:**
```yaml
- Id: 1
  Name: 아이템1
  Value: 100
- Id: 2
  Name: 아이템2
  Value: 200
```

> **핵심 특징:**
> - 첫 번째 행(녹색 배경)에 `$[]` 마커로 전체가 배열임을 정의합니다.
> - 두 번째 행(연한 녹색 배경)에 각 열의 속성명을 정의합니다.
> - 세 번째 행부터는 실제 데이터입니다.
> - 맨 마지막 행은 **붉은색 배경**으로 `$scheme_end`가 표시되며, 전체 열이 병합되어 있습니다.
> - 데이터 행의 첫 열에는 `^` 표시가 있어 해당 셀을 무시하도록 합니다.

#### 2. 중첩 배열 구조

<table class="excel-grid" style="width: 100%;">
    <thead>
        <tr class="excel-row">
            <th class="excel-header" colspan="8" style="background-color: #00CC00;">$[]</th>
        </tr>
    </thead>
    <tbody>
        <tr class="excel-row">
            <td class="excel-cell">^</td>
            <td class="excel-cell" colspan="7" style="background-color: #CCFFCC;">${}</td>
        </tr>
        <tr class="excel-row">
            <td class="excel-cell" style="width: 5%;">^</td>
            <td class="excel-cell" style="width: 15%;">Rarity</td>
            <td class="excel-cell" colspan="6" style="width: 80%;">Materials$[]</td>
        </tr>
        <tr class="excel-row">
            <td class="excel-cell">^</td>
            <td class="excel-cell">^</td>
            <td class="excel-cell" colspan="2" style="width: 26%;">${}</td>
            <td class="excel-cell" colspan="2" style="width: 26%;">${}</td>
            <td class="excel-cell" colspan="2" style="width: 26%;">${}</td>
        </tr>
        <tr class="excel-row">
            <td class="excel-cell">^</td>
            <td class="excel-cell">^</td>
            <td class="excel-cell">Id</td>
            <td class="excel-cell">Count</td>
            <td class="excel-cell">Id</td>
            <td class="excel-cell">Count</td>
            <td class="excel-cell">Id</td>
            <td class="excel-cell">Count</td>
        </tr>
        <tr class="excel-row" style="background-color: #ff0000; color: white;">
            <td class="excel-cell" colspan="8" style="text-align: center;">$scheme_end</td>
        </tr>
        <tr class="excel-row">
            <td class="excel-cell"></td>
            <td class="excel-cell">Rare</td>
            <td class="excel-cell">iron</td>
            <td class="excel-cell">10</td>
            <td class="excel-cell">gold</td>
            <td class="excel-cell">5</td>
            <td class="excel-cell">gem</td>
            <td class="excel-cell">2</td>
        </tr>
        <tr class="excel-row">
            <td class="excel-cell"></td>
            <td class="excel-cell">Epic</td>
            <td class="excel-cell">steel</td>
            <td class="excel-cell">15</td>
            <td class="excel-cell">diamond</td>
            <td class="excel-cell">3</td>
            <td class="excel-cell">magic</td>
            <td class="excel-cell">7</td>
        </tr>
    </tbody>
</table>

**YAML 결과:**
```yaml
- Rarity: Rare
  Materials:
    - Id: iron
      Count: 10
    - Id: gold
      Count: 5
    - Id: gem
      Count: 2
- Rarity: Epic
  Materials:
    - Id: steel
      Count: 15
    - Id: diamond
      Count: 3
    - Id: magic
      Count: 7
```

### 4.4 복합 스키마 작성법

실전에서는 객체, 배열, 키-값이 복합적으로 사용됩니다. 아래 예시를 통해 복합 구조를 작성하는 방법을 알아보겠습니다.

<table class="excel-grid" style="width: 100%;">
  <thead>
      <tr class="excel-row">
          <th class="excel-header" colspan="8" style="background-color: #00CC00;">$[]</th>
      </tr>
  </thead>
  <tbody>
      <tr class="excel-row">
          <td class="excel-cell">^</td>
          <td class="excel-cell" colspan="7" style="background-color: #CCFFCC;">${}</td>
      </tr>
      <tr class="excel-row">
          <td class="excel-cell" style="width: 5%;">^</td>
          <td class="excel-cell" style="width: 15%;">Rarity</td>
          <td class="excel-cell" colspan="6" style="width: 80%;">Materials$[]</td>
      </tr>
      <tr class="excel-row">
          <td class="excel-cell">^</td>
          <td class="excel-cell">^</td>
          <td class="excel-cell" colspan="2" style="width: 26%;">${}</td>
          <td class="excel-cell" colspan="2" style="width: 26%;">${}</td>
          <td class="excel-cell" colspan="2" style="width: 26%;">${}</td>
      </tr>
      <tr class="excel-row">
          <td class="excel-cell">^</td>
          <td class="excel-cell">^</td>
          <td class="excel-cell">Id</td>
          <td class="excel-cell">Count</td>
          <td class="excel-cell">Id</td>
          <td class="excel-cell">Count</td>
          <td class="excel-cell">Id</td>
          <td class="excel-cell">Count</td>
      </tr>
      <tr class="excel-row" style="background-color: #ff0000; color: white;">
          <td class="excel-cell" colspan="8" style="text-align: center;">$scheme_end</td>
      </tr>
      <tr class="excel-row">
          <td class="excel-cell"></td>
          <td class="excel-cell">Rare</td>
          <td class="excel-cell">iron</td>
          <td class="excel-cell">10</td>
          <td class="excel-cell">gold</td>
          <td class="excel-cell">5</td>
          <td class="excel-cell">gem</td>
          <td class="excel-cell">2</td>
      </tr>
      <tr class="excel-row">
          <td class="excel-cell"></td>
          <td class="excel-cell">Epic</td>
          <td class="excel-cell">steel</td>
          <td class="excel-cell">15</td>
          <td class="excel-cell">diamond</td>
          <td class="excel-cell">3</td>
          <td class="excel-cell">magic</td>
          <td class="excel-cell">7</td>
      </tr>
  </tbody>
</table>

**YAML 결과:**
```yaml
- id: C001
  name: 전사
  type: 근접
  inventory:
    - hp: 200
      atk: 35
      def: 40
  skills:
    - id: shieldBash
      name: 방패 강타
      damage: 15
      cost: 15
- id: C002
  name: 마법사
  type: 원거리
  inventory:
    - hp: 150
      atk: 80
      def: 20
  skills:
    - id: fireball
      name: 화염구
      damage: 60
      cost: 35
```

> **복합 구조 작성 팁:** 중첩된 배열과 객체를 함께 사용할 때는 각 필드의 계층 구조를 명확하게 표현하는 것이 중요합니다. 이미지와 같이 스키마의 마지막 행에 $scheme_end를 전체 열을 병합하여 표시하고, 실제 데이터는 그 아래에 작성합니다.

### 4.5 최종 체크리스트

엑셀 시트를 YAML로 변환하기 전에 다음 사항을 확인하세요:

1. 시트 이름 앞에 `!` 기호가 있는지 확인 (예: "!캐릭터데이터")
2. 스키마 정의가 상단에 올바르게 작성되었는지 확인
3. 스키마 마지막 행에 `$scheme_end` 마커가 있고, 모든 열이 병합되었는지 확인
4. 배열(`$[]`)과 객체(`${}`) 마커 뒤에 이름이 올바르게 지정되었는지 확인
5. 중첩 구조의 경우 계층 관계가 올바르게 표현되었는지 확인
6. 데이터 행에 필요한 모든 값이 입력되었는지 확인
7. 무시해야 할 셀에 `^` 마커가 있는지 확인

> **팁:** 복잡한 구조를 작성하기 전에 간단한 예시로 테스트하여 변환 결과를 확인하는 것이 좋습니다. 이를 통해 스키마 구조의 오류를 빠르게 파악하고 수정할 수 있습니다. 
