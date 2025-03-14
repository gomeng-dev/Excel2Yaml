# OrderedYamlFactory 클래스 가이드

## 개요

`OrderedYamlFactory` 클래스는 Excel2YAML 애드인에서 YAML 데이터의 생성과 직렬화를 담당하는 핵심 컴포넌트입니다. 이 클래스는 일반적인 YAML 라이브러리와 달리 **속성의 순서를 보존**하는 특별한 기능을 제공합니다. 또한 `YamlObject`와 `YamlArray` 클래스를 통해 계층적 데이터 구조를 효과적으로 구성하고 처리할 수 있습니다.

## 주요 클래스

### YamlObject 클래스

속성 순서가 보존되는 객체를 표현하는 클래스로, 이름-값 쌍의 컬렉션을 관리합니다.

```csharp
public class YamlObject
{
    private readonly Dictionary<string, object> properties = new Dictionary<string, object>();
    private readonly List<string> propertyOrder = new List<string>();
    
    // 메서드들...
}
```

#### 주요 속성 및 메서드

- **Add(string name, object value)**: 객체에 속성을 추가합니다.
- **ContainsKey(string key)**: 지정된 키가 객체에 존재하는지 확인합니다.
- **Remove(string key)**: 지정된 키와 그에 해당하는 값을 제거합니다.
- **HasValues**: 객체에 속성이 하나 이상 있는지 확인합니다.
- **Properties**: 객체의 모든 속성을 반환합니다. 중요한 점은 이 속성이 `propertyOrder` 리스트에 따라 **순서대로** 속성을 반환한다는 점입니다.

### YamlArray 클래스

순서가 있는 값의 컬렉션을 표현하는 클래스입니다.

```csharp
public class YamlArray
{
    private readonly List<object> items = new List<object>();
    
    // 메서드들...
}
```

#### 주요 속성 및 메서드

- **Add(object value)**: 배열에 항목을 추가합니다.
- **RemoveAt(int index)**: 지정된 인덱스의 항목을 제거합니다.
- **Count**: 배열 내 항목의 수를 반환합니다.
- **HasValues**: 배열에 항목이 하나 이상 있는지 확인합니다.
- **Items**: 배열의 모든 항목을 반환합니다.

### OrderedYamlFactory 클래스

YAML 객체와 배열을 생성하고 직렬화하는 정적 메서드들을 제공하는 클래스입니다.

```csharp
public static class OrderedYamlFactory
{
    // 정적 메서드들...
}
```

#### 주요 메서드

- **CreateObject()**: 새로운 `YamlObject` 인스턴스를 생성합니다.
- **CreateArray()**: 새로운 `YamlArray` 인스턴스를 생성합니다.
- **RemoveEmptyProperties(object token)**: 객체나 배열에서 빈 속성이나 항목을 제거합니다.
- **SerializeToYaml(object obj, ...)**: 객체를 YAML 문자열로 직렬화합니다.
- **SaveToYaml(object obj, string filePath, ...)**: 객체를 YAML 파일로 저장합니다.

## 주요 기능

### 1. 속성 순서 보존

일반적인 YAML 라이브러리와 달리, `OrderedYamlFactory`는 객체의 속성 순서를 정확하게 보존합니다. 이는 데이터 표현에서 순서가 중요한 경우에 특히 유용합니다.

```csharp
var obj = OrderedYamlFactory.CreateObject();
obj.Add("name", "홍길동");
obj.Add("age", 30);
obj.Add("job", "개발자");

// YAML 출력:
// name: 홍길동
// age: 30
// job: 개발자
// (입력 순서 그대로 유지됨)
```

### 2. 다양한 직렬화 옵션

YAML 출력 형식을 세밀하게 제어할 수 있는 다양한 옵션을 제공합니다:

- **들여쓰기 크기**: 계층 구조의 들여쓰기 공백 수 지정
- **스타일**: Block 스타일(기본값) 또는 Flow 스타일(한 줄 표시)
- **따옴표 보존**: 문자열 값에 따옴표를 유지할지 여부
- **빈 필드 포함**: 빈 속성이나 컬렉션을 출력에 포함할지 여부

```csharp
string yaml = OrderedYamlFactory.SerializeToYaml(obj, 
    indentSize: 4,                  // 4칸 들여쓰기
    style: YamlStyle.Flow,         // 흐름 스타일(한 줄) 사용
    preserveQuotes: true,          // 문자열 따옴표 유지
    includeEmptyFields: false);    // 빈 필드 제외
```

### 3. 빈 속성 관리

객체나 배열에서 빈 속성을 자동으로 감지하고 처리(제거 또는 유지)할 수 있습니다.

```csharp
var obj = OrderedYamlFactory.CreateObject();
obj.Add("name", "홍길동");
obj.Add("description", "");
obj.Add("tags", new YamlArray()); // 빈 배열

// 빈 속성 제거
OrderedYamlFactory.RemoveEmptyProperties(obj);

// 결과: description과 tags 속성이 제거됨
```

### 4. 계층적 데이터 구조 지원

복잡한 중첩 객체와 배열을 쉽게 구성하고 직렬화할 수 있습니다.

```csharp
var character = OrderedYamlFactory.CreateObject();
character.Add("name", "전사");
character.Add("level", 30);

var stats = OrderedYamlFactory.CreateObject();
stats.Add("strength", 20);
stats.Add("agility", 15);
character.Add("stats", stats);

var items = OrderedYamlFactory.CreateArray();
var item1 = OrderedYamlFactory.CreateObject();
item1.Add("name", "검");
item1.Add("damage", 10);
items.Add(item1);

character.Add("items", items);

// 직렬화하여 YAML 출력
string yaml = OrderedYamlFactory.SerializeToYaml(character);
```

### 5. 자동 문자열 인코딩

특수 문자나 여러 줄 텍스트를 포함하는 문자열을 자동으로 적절히 인코딩합니다:

- 특수 문자가 포함된 문자열: 따옴표로 묶음
- 여러 줄 텍스트: `|` 또는 `>` 문자 사용
- 숫자 형태의 문자열: 따옴표로 묶어 문자열로 처리

## 사용 예시

### 기본 객체 생성 및 직렬화

```csharp
// 객체 생성
var config = OrderedYamlFactory.CreateObject();
config.Add("appName", "MyApp");
config.Add("version", "1.0.0");
config.Add("debugMode", true);

// 하위 객체 추가
var database = OrderedYamlFactory.CreateObject();
database.Add("host", "localhost");
database.Add("port", 3306);
database.Add("username", "admin");
config.Add("database", database);

// YAML로 직렬화
string yaml = OrderedYamlFactory.SerializeToYaml(config);
Console.WriteLine(yaml);
```

출력 결과:
```yaml
appName: MyApp
version: 1.0.0
debugMode: true
database:
  host: localhost
  port: 3306
  username: admin
```

### 배열 처리

```csharp
// 배열 생성
var users = OrderedYamlFactory.CreateArray();

// 첫 번째 사용자
var user1 = OrderedYamlFactory.CreateObject();
user1.Add("id", 1);
user1.Add("name", "홍길동");
users.Add(user1);

// 두 번째 사용자
var user2 = OrderedYamlFactory.CreateObject();
user2.Add("id", 2);
user2.Add("name", "김철수");
users.Add(user2);

// YAML로 직렬화 및 저장
OrderedYamlFactory.SaveToYaml(users, "users.yaml");
```

출력 파일 (users.yaml):
```yaml
- id: 1
  name: 홍길동
- id: 2
  name: 김철수
```

### 다양한 스타일 적용

```csharp
var person = OrderedYamlFactory.CreateObject();
person.Add("name", "홍길동");

var contact = OrderedYamlFactory.CreateObject();
contact.Add("email", "hong@example.com");
contact.Add("phone", "010-1234-5678");
person.Add("contact", contact);

// 블록 스타일(기본)
string blockStyle = OrderedYamlFactory.SerializeToYaml(person);
// 결과:
// name: 홍길동
// contact:
//   email: hong@example.com
//   phone: 010-1234-5678

// 흐름 스타일
string flowStyle = OrderedYamlFactory.SerializeToYaml(person, style: YamlStyle.Flow);
// 결과:
// { name: 홍길동, contact: { email: hong@example.com, phone: 010-1234-5678 } }
```

## 주의사항

1. **값 타입 처리**: 기본적으로 문자열, 숫자, 불리언과 같은 타입은 자동으로 인식되어 적절히 처리됩니다. 문자열로 처리되어야 하는 숫자 형태의 데이터는 문자열 객체로 명시적으로 지정해야 합니다.

2. **빈 값 처리**: `includeEmptyFields` 옵션을 `false`로 설정하면 빈 속성이나 배열이 출력에서 제외됩니다. 빈 값도 유지해야 하는 경우 `true`로 설정하세요.

3. **큰 데이터 처리**: 대용량 데이터를 처리할 때는 메모리 사용량에 주의하세요. 필요하다면 데이터를 분할하여 처리하는 것이 좋습니다.

4. **특수 문자 처리**: YAML 문법에서 특별한 의미를 가지는 문자(`:`, `-`, `[`, `]` 등)가 포함된 문자열은 자동으로 따옴표로 처리됩니다.

## 활용 시나리오

- **설정 파일 생성**: 애플리케이션 설정을 YAML 파일로 저장
- **데이터 직렬화**: 메모리 내 데이터 구조를 YAML 형식으로 변환
- **API 응답 포맷팅**: API 응답 데이터를 YAML로 포맷팅하여 제공
- **게임 데이터 관리**: 게임 내 아이템, 캐릭터 등의 데이터를 YAML로 관리
- **엑셀에서 변환된 데이터 출력**: `SchemeNode`에서 파싱된 엑셀 데이터를 YAML로 출력 