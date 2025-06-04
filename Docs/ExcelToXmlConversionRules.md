# Excel 데이터를 XML로 변환 후 JSON으로 변환하는 규칙 (Excel2Yaml 확장)

Excel 데이터를 중간 형식인 XML로 변환한 후, 최종적으로 특정 규칙에 따라 JSON으로 변환하는 기능입니다.

## 변환 규칙 (XML -> JSON)

제공된 예시 XML 데이터를 기준으로 JSON으로 변환하는 규칙은 다음과 같습니다.

1.  **루트 요소 (Root Element)**:
    *   XML 문서의 루트 요소는 JSON 객체의 최상위 키가 됩니다.
    *   예시: `<ArenaSettings>`는 JSON에서 `"ArenaSettings": {}`의 형태로 시작됩니다.

2.  **자식 요소 (Child Elements)**:
    *   **단일 자식 요소**: 부모 요소의 속성(키-값 쌍)으로 변환됩니다. 요소의 이름이 키가 되고, 요소의 텍스트 내용이 값이 됩니다.
        *   예시: `<DailyTrainingCount>5</DailyTrainingCount>`는 `"DailyTrainingCount": "5"`로 변환됩니다.
    *   **동일한 이름의 여러 자식 요소**: 해당 요소의 이름을 키로 하는 JSON 배열로 변환됩니다. 각 자식 XML 요소는 배열 내의 JSON 객체가 됩니다.
        *   예시: `WinRewards` 요소 내의 여러 `WinReward` 요소들은 `"WinReward": [{}, {}, ...]` 형태로 변환됩니다.

3.  **속성 (Attributes)**:
    *   XML 요소의 속성들은 해당 요소가 변환된 JSON 객체 내에서 밑줄(`_`) 접두사가 붙은 키로 변환됩니다. 속성의 이름이 밑줄 뒤에 오고, 속성의 값이 JSON 값이 됩니다.
    *   예시: `<WinReward Step="0" WinCount="3">`는 `"_Step": "0"`, `"_WinCount": "3"`을 포함하는 JSON 객체로 변환됩니다.

4.  **중첩 구조 (Nested Structure)**:
    *   XML 요소 내에 다른 자식 요소나 속성이 중첩된 경우, 위의 규칙들이 재귀적으로 적용됩니다.
    *   예시: `WinReward` 요소 내의 `Reward` 요소는 `WinReward`에 해당하는 JSON 객체 내에 `"Reward": {}` 형태로 중첩되어 표현됩니다. `Reward` 요소의 속성들(`Type`, `Count`) 또한 밑줄 접두사 규칙에 따라 변환됩니다.
        ```xml
        <WinReward Step="0" WinCount="3">
            <Reward Type="Token_ARENA" Count="50"/>
        </WinReward>
        ```
        위 XML은 아래 JSON 구조의 일부가 됩니다:
        ```json
        {
            "_Step": "0",
            "_WinCount": "3",
            "Reward": {
                "_Type": "Token_ARENA",
                "_Count": "50"
            }
        }
        ```

## 전체 변환 예시

**입력 XML 데이터:**
```xml
<ArenaSettings>
    <WinRewards>
        <WinReward Step="0" WinCount="3">
            <Reward Type="Token_ARENA" Count="50"/>
        </WinReward>
        <WinReward Step="1" WinCount="8">
            <Reward Type="Token_ARENA" Count="150"/>
        </WinReward>
        <WinReward Step="2" WinCount="18">
            <Reward Type="Token_ARENA" Count="300"/>
        </WinReward>
        <WinReward Step="3" WinCount="30">
            <Reward Type="Token_ARENA" Count="500"/>
        </WinReward>
        <WinReward Step="4" WinCount="45">
            <Reward Type="Token_ARENA" Count="1000"/>
        </WinReward>
    </WinRewards>
    <DailyTrainingCount>5</DailyTrainingCount>
</ArenaSettings>
```

**출력 JSON 데이터:**
```json
{
	"ArenaSettings": {
		"WinRewards": {
			"WinReward": [
				{
					"Reward": {
						"_Type": "Token_ARENA",
						"_Count": "50"
					},
					"_Step": "0",
					"_WinCount": "3"
				},
				{
					"Reward": {
						"_Type": "Token_ARENA",
						"_Count": "150"
					},
					"_Step": "1",
					"_WinCount": "8"
				},
				{
					"Reward": {
						"_Type": "Token_ARENA",
						"_Count": "300"
					},
					"_Step": "2",
					"_WinCount": "18"
				},
				{
					"Reward": {
						"_Type": "Token_ARENA",
						"_Count": "500"
					},
					"_Step": "3",
					"_WinCount": "30"
				},
				{
					"Reward": {
						"_Type": "Token_ARENA",
						"_Count": "1000"
					},
					"_Step": "4",
					"_WinCount": "45"
				}
			]
		},
		"DailyTrainingCount": "5"
	}
}
```
