# Ribbon 클래스 가이드

## 개요

`Ribbon` 클래스는 Excel2YAML 애드인의, 엑셀 내에서 사용자 인터페이스(UI)를 제공하는 핵심 컴포넌트입니다. 이 클래스는 엑셀 리본(Ribbon)에 탭, 그룹, 버튼 등의 컨트롤을 추가하여 사용자가 쉽게 Excel2YAML 기능에 접근할 수 있도록 합니다. 리본 UI는 사용자가 엑셀 시트를 YAML 형식으로 변환하고, 관련 설정을 구성할 수 있는 직관적인 방법을 제공합니다.

## 구성 요소

Ribbon 클래스는 두 개의 파일로 구성되어 있습니다:

1. **Ribbon.Designer.cs**: 리본 UI의 시각적 요소와 디자인을 정의합니다. 이 파일은 주로 Visual Studio 디자이너에 의해 자동 생성됩니다.

2. **Ribbon.cs**: 리본 컨트롤의 이벤트 핸들러와 비즈니스 로직을 구현합니다. 이 파일은 사용자 상호작용에 대한 응답과 실제 기능 수행을 담당합니다.

## 주요 UI 요소

### 탭과 그룹

```csharp
internal Microsoft.Office.Tools.Ribbon.RibbonTab tabExcelToJson;
internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupConvert;
internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSettings;
```

- **tabExcelToJson**: 엑셀 리본에 추가되는 주 탭으로, 모든 Excel2YAML 기능을 포함합니다.
- **groupConvert**: 변환 관련 버튼들을 그룹화합니다.
- **groupSettings**: 설정 관련 버튼들을 그룹화합니다.

### 주요 버튼

```csharp
internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertToYaml;
internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSheetPathSettings;
```

- **btnConvertToYaml**: 현재 선택된 엑셀 시트를 YAML 형식으로 변환합니다.
- **btnSheetPathSettings**: 시트별 출력 경로 설정 대화상자를 엽니다.

## 주요 기능

### 초기화 및 로드

```csharp
public Ribbon() : base(Globals.Factory.GetRibbonFactory())
{
    InitializeComponent();
    // 설정 로드 로직...
}

private void Ribbon_Load(object sender, RibbonUIEventArgs e)
{
    // 리본 UI 로드 시 실행되는 로직...
}
```

- **생성자**: 리본 컴포넌트를 초기화하고 사용자 설정을 로드합니다.
- **Ribbon_Load**: 리본 UI가 로드될 때 호출되며, 초기 상태 설정 및 설정 불러오기를 담당합니다.

### YAML 변환 기능

```csharp
public void OnConvertToYamlClick(object sender, RibbonControlEventArgs e)
{
    // YAML 변환 로직...
}
```

이 메서드는 사용자가 'Convert to YAML' 버튼을 클릭할 때 호출되며 다음 작업을 수행합니다:

1. 현재 워크북과 시트 설정 초기화
2. 변환 대상 시트 확인
3. 출력 경로 설정 확인
4. 시트 데이터를 YAML 형식으로 변환
5. 후처리 옵션 적용(항목 병합, 흐름 스타일 등)
6. YAML 파일 저장

### 설정 관리

```csharp
public void OnSheetPathSettingsClick(object sender, RibbonControlEventArgs e)
{
    // 시트 경로 설정 대화상자 표시...
}

public bool GetEmptyFieldsState(IRibbonControl control)
{
    return includeEmptyFields;
}

public void OnEmptyFieldsClicked(IRibbonControl control, bool pressed)
{
    includeEmptyFields = pressed;
}

// 기타 설정 관련 메서드...
```

- **OnSheetPathSettingsClick**: 시트별 출력 경로를 설정하는 대화상자를 표시합니다.
- **GetXXXState/OnXXXClicked**: 각종 설정 옵션(빈 필드 포함, 해시 생성, 빈 YAML 필드 추가 등)의 상태를 관리합니다.

## 설정 옵션

Ribbon 클래스는 YAML 변환 시 사용되는 여러 설정 옵션을 관리합니다:

```csharp
private bool includeEmptyFields = false;
private bool enableHashGen = false;
private bool addEmptyYamlFields = false;
```

- **includeEmptyFields**: 빈 셀의 데이터도 출력에 포함할지 여부를 결정합니다.
- **enableHashGen**: 해시 생성 기능을 활성화할지 여부를 결정합니다.
- **addEmptyYamlFields**: 빈 YAML 필드를 출력에 포함할지 여부를 결정합니다.

## 변환 프로세스

```csharp
private List<string> ConvertExcelFile(ExcelToJsonConfig config)
{
    // 엑셀 파일 변환 로직...
}
```

이 메서드는 Excel2YAML의 핵심 기능을 수행합니다:

1. 엑셀 파일 열기 및 시트 분석
2. 시트별 스키마 파싱
3. 스키마에 따라 데이터 추출
4. 추출된 데이터를 YAML 형식으로 변환
5. 변환된 YAML 데이터 후처리
6. 결과 파일 저장

## 사용자 정의 설정 폼

```csharp
private Forms.SheetPathSettingsForm settingsForm = null;
```

이 필드는 시트별 출력 경로를 설정하는 대화상자 폼을 참조합니다. `OnSheetPathSettingsClick` 메서드에서 이 폼을 생성하고 표시합니다.

## 사용 예시

### 1. YAML 변환 실행

1. 엑셀 파일을 엽니다.
2. Excel2YAML 탭을 클릭합니다.
3. "Convert to YAML" 버튼을 클릭합니다.
4. 출력 경로를 선택합니다(필요한 경우).
5. 변환이 완료되면 완료 메시지가 표시됩니다.

### 2. 시트별 출력 경로 설정

1. Excel2YAML 탭을 클릭합니다.
2. "Sheet Path Settings" 버튼을 클릭합니다.
3. 대화상자에서 시트별 출력 경로를 설정합니다.
4. "Save" 버튼을 클릭하여 설정을 저장합니다.

### 3. 변환 옵션 설정

1. Excel2YAML 탭을 클릭합니다.
2. "Settings" 그룹에서 원하는 옵션을 토글합니다:
   - "Include Empty Fields": 빈 셀 데이터를 출력에 포함할지 여부
   - "Generate Hash": 해시 생성 활성화 여부
   - "Add Empty YAML Fields": 빈 YAML 필드를 출력에 포함할지 여부

## 주의사항

1. **시트 이름 관련**: 자동 변환을 원하는 시트 이름 앞에 '!' 기호를 붙여야 합니다.

2. **설정 저장**: 시트별 출력 경로 설정은 애플리케이션 설정에 저장되므로 엑셀을 다시 시작해도 유지됩니다.

3. **대용량 파일 처리**: 큰 엑셀 파일을 변환할 때는 시간이 오래 걸릴 수 있으므로 인내심을 가지고 기다려야 합니다.

4. **YAML 후처리 옵션**: 항목 병합, 흐름 스타일 등의 후처리 옵션은 시트 설정에서 별도로 구성해야 합니다.

## 개발자를 위한 정보

Ribbon 클래스를 확장하거나 수정하려면 다음 사항을 고려하세요:

1. **Ribbon.Designer.cs**: 시각적 요소를 수정하려면 Visual Studio 디자이너를 사용하는 것이 좋습니다.

2. **Ribbon.cs**: 기능 로직을 수정하려면 이 파일의 이벤트 핸들러를 수정하세요.

3. **RibbonUI.xml**: 리본 UI의 XML 정의는 별도의 XML 파일로 관리됩니다. 복잡한 UI 변경은 이 파일을 수정해야 할 수 있습니다.

4. **Forms 연동**: 설정 대화상자는 Forms 네임스페이스의 별도 클래스로 구현되어 있습니다. UI와 로직이 분리되어 있음을 기억하세요. 