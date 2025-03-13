using ExcelToYamlAddin.Logging;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelToYamlAddin.Config;
using System.Diagnostics;
using System.Windows.Forms;

namespace ExcelToYamlAddin.Core
{
    public class YamlGenerator
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<YamlGenerator>();

        private readonly Scheme _scheme;
        private readonly IXLWorksheet _sheet;
        private readonly Stack<object> _stack = new Stack<object>();  // Java와 같은 스택 기반 접근법 적용

        public YamlGenerator(Scheme scheme)
        {
            _scheme = scheme;
            _sheet = scheme.Sheet;
            Logger.Debug("YamlGenerator 초기화: 스키마 노드 타입={0}", scheme.Root.NodeType);
        }

        // 스택 관리 메서드
        private void Push(object obj)
        {
            _stack.Push(obj);
            Logger.Debug($"스택에 객체 추가: 현재 크기={_stack.Count}, 타입={obj.GetType().Name}");
        }

        private object Pop()
        {
            object popped = _stack.Pop();
            Logger.Debug($"스택에서 객체 제거: 현재 크기={_stack.Count}, 타입={popped.GetType().Name}");
            return popped;
        }

        private int GetStackSize()
        {
            return _stack.Count;
        }

        // 외부에서 호출할 정적 메서드
        public static string Generate(Scheme scheme, YamlStyle style = YamlStyle.Block, int indentSize = 2, bool preserveQuotes = false, bool includeEmptyFields = false)
        {
            try 
            {
                Debug.WriteLine($"[YamlGenerator] Generate 호출됨: style={style}, includeEmptyFields={includeEmptyFields}");
                Debug.WriteLine($"[YamlGenerator] Generate 호출 스택: {Environment.StackTrace}");
                var generator = new YamlGenerator(scheme);
                object rootObj = generator.ProcessRootNode();
                
                // 디버그 로그 추가
                Debug.WriteLine($"[YamlGenerator] 정적 Generate 메서드: includeEmptyFields={includeEmptyFields}, 이 값이 SerializeToYaml에 전달됩니다");
                
                // SerializeToYaml에서 includeEmptyFields 매개변수를 통해 빈 속성 처리를 수행하므로
                // 여기서는 RemoveEmptyProperties 호출을 제거합니다.
                Debug.WriteLine($"[YamlGenerator] OrderedYamlFactory.SerializeToYaml 호출 전: includeEmptyFields={includeEmptyFields}");
                var result = OrderedYamlFactory.SerializeToYaml(rootObj, indentSize, style, preserveQuotes, includeEmptyFields);
                Debug.WriteLine($"[YamlGenerator] OrderedYamlFactory.SerializeToYaml 호출 후 결과 길이: {result?.Length ?? 0}");
                return result;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"YAML 생성 중 오류 발생: {ex.Message}");
                throw;
            }
        }

        public string Generate(YamlStyle style = YamlStyle.Block, int indentSize = 2, bool preserveQuotes = false, bool includeEmptyFields = false)
        {
            try
            {
                Logger.Debug("YAML 생성 시작: 스타일={0}, 들여쓰기={1}, 빈 필드 포함={2}", style, indentSize, includeEmptyFields);
                
                object rootObj = ProcessRootNode();
                
                return OrderedYamlFactory.SerializeToYaml(rootObj, indentSize, style, preserveQuotes, includeEmptyFields);
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "YAML 생성 중 오류 발생");
                throw;
            }
        }
        
        // Java 버전과 유사한 스택 기반 루트 노드 처리 메서드
        public object ProcessRootNode()
        {
            SchemeNode rootNode = _scheme.Root;
            Logger.Debug($"ProcessRootNode 시작: 루트 타입={rootNode.NodeType}, 키={rootNode.Key}, 자식 수={rootNode.ChildCount}");
            
            // 루트 노드에 따른 객체 생성
            object rootObj;
            if (rootNode.NodeType == SchemeNode.SchemeNodeType.MAP)
            {
                rootObj = OrderedYamlFactory.CreateObject();
                Logger.Debug("MAP 루트 노드 객체 생성");
            }
            else if (rootNode.NodeType == SchemeNode.SchemeNodeType.ARRAY)
            {
                rootObj = OrderedYamlFactory.CreateArray();
                Logger.Debug("ARRAY 루트 노드 객체 생성");
            }
            else
            {
                Logger.Warning($"지원되지 않는 루트 노드 타입: {rootNode.NodeType}, MAP으로 처리합니다.");
                rootObj = OrderedYamlFactory.CreateObject();
            }
            
            // 스택에 루트 객체 추가
            Push(rootObj);
            
            // 모든 데이터 행 처리
            for (int rowNum = _scheme.ContentStartRowNum; rowNum <= _scheme.EndRowNum; rowNum++)
            {
                IXLRow row = _sheet.Row(rowNum);
                if (row == null) continue;
                
                Logger.Debug($"행 {rowNum} 처리 중");
                
                // 현재 객체는 항상 스택의 맨 위에 있는 객체
                object currentObject = rootObj;
                
                // 스키마 노드를 순회하며 처리
                var linearNodes = _scheme.GetLinearNodes().ToList();
                
                for (int i = 0; i < linearNodes.Count; i++)
                {
                    SchemeNode currentNode = linearNodes[i];
                    Logger.Debug($"노드 처리: 키={currentNode.Key}, 타입={currentNode.NodeType}, 스키마 행={currentNode.SchemeRowNum}");
                    
                    // 노드 깊이와 스택 크기를 비교하여 스택 관리
                    int stackSize = GetStackSize();
                    if (currentNode.SchemeRowNum < stackSize)
                    {
                        int popCount = stackSize - currentNode.SchemeRowNum;
                        for (int j = 0; j < popCount; j++)
                        {
                            currentObject = Pop();
                        }
                    }
                    
                    // 컨테이너 타입 노드 처리 (MAP, ARRAY)
                    if (currentNode.NodeType == SchemeNode.SchemeNodeType.MAP || 
                        currentNode.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                    {
                        ProcessContainerNode(currentNode, currentObject, row);
                    }
                    // 값 타입 노드 처리 (PROPERTY, VALUE)
                    else if (currentNode.NodeType == SchemeNode.SchemeNodeType.PROPERTY ||
                             currentNode.NodeType == SchemeNode.SchemeNodeType.VALUE)
                    { 
                        ProcessValueNode(currentNode, currentObject, row);
                    }
                    // KEY-VALUE 쌍 처리
                    else if (currentNode.NodeType == SchemeNode.SchemeNodeType.KEY)
                    {
                        ProcessKeyNode(currentNode, currentObject, row);
                    }
                }
                
                // 행 처리 후 스택 초기화 - 루트 객체만 남기기
                while (GetStackSize() > 1)
                {
                    Pop();
                }
                currentObject = _stack.Peek();
            }
            
            // 마지막 객체 반환
            return Pop();
        }
        
        // 컨테이너 노드 처리 (MAP, ARRAY)
        private void ProcessContainerNode(SchemeNode node, object parentObject, IXLRow row)
        {
            string key = GetNodeKey(node, row);
            
            // 부모가 객체이고 키가 비어있는 경우 (이는 JSON/YAML 표준에 맞지 않음)
            if (string.IsNullOrEmpty(key) && parentObject is YamlObject)
            {
                string errorMessage = "오류: JSON/YAML 표준에서는 객체 내에 이름 없는 속성을 가질 수 없습니다. 키 값이 비어있습니다.";
                Logger.Error(errorMessage);
                
                // 오류 창 표시
                DialogResult result = MessageBox.Show(
                    errorMessage,
                    "JSON/YAML 표준 오류",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.DefaultDesktopOnly);
                
                // 확인 버튼을 누른 경우 변환 과정 중단
                if (result == DialogResult.OK)
                {
                    string abortMessage = "사용자 요청에 의해 변환 작업이 중단되었습니다.";
                    Logger.Error(abortMessage);
                    throw new InvalidOperationException(abortMessage);
                }
                
                Logger.Debug($"컨테이너 노드의 키가 비어 있어 무시: 타입={node.NodeType}");
                return;
            }
            
            // 부모가 객체인 경우
            if (parentObject is YamlObject parentMap)
            {
                // 이미 키가 있는 경우, 기존 객체 사용
                if (parentMap.ContainsKey(key))
                {
                    object existingObj = parentMap.Properties.First(p => p.Key == key).Value;
                    Push(parentObject);
                    Push(existingObj);
                    Logger.Debug($"기존 객체 사용: 키={key}, 타입={existingObj.GetType().Name}");
                    return;
                }
                
                // 새 객체 생성
                if (node.NodeType == SchemeNode.SchemeNodeType.MAP)
                {
                    YamlObject newMap = OrderedYamlFactory.CreateObject();
                    parentMap.Add(key, newMap);
                    Push(parentObject);
                    Push(newMap);
                    Logger.Debug($"새 MAP 객체 생성: 키={key}");
                }
                else if (node.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                {
                    // 객체 안에 배열 추가 - 이는 표준에 맞는 정상적인 동작임
                    YamlArray newArray = OrderedYamlFactory.CreateArray();
                    parentMap.Add(key, newArray);
                    Push(parentObject);
                    Push(newArray);
                    Logger.Debug($"객체에 새 ARRAY 추가: 키={key}");
                }
            }
            // 부모가 배열인 경우
            else if (parentObject is YamlArray parentArray)
            {
                if (node.NodeType == SchemeNode.SchemeNodeType.MAP)
                {
                    YamlObject newMap = OrderedYamlFactory.CreateObject();
                    parentArray.Add(newMap);
                    Push(parentObject);
                    Push(newMap);
                    Logger.Debug($"배열에 새 MAP 객체 추가");
                }
                else if (node.NodeType == SchemeNode.SchemeNodeType.ARRAY)
                {
                    // JSON/YAML 표준: 배열 안에 직접 배열을 넣는 것은 권장되지 않지만 처리 가능
                    // 경고나 중단 없이 처리함
                    string warningMessage = "배열 안에 직접 배열을 추가하는 것은 일부 파서에서 문제가 될 수 있습니다. 가능하면 이름 있는 객체로 감싸는 것이 좋습니다.";
                    Logger.Warning(warningMessage); // 로그에만 기록하고 사용자에게 표시하지 않음
                    
                    YamlArray newArray = OrderedYamlFactory.CreateArray();
                    parentArray.Add(newArray);
                    Push(parentObject);
                    Push(newArray);
                    Logger.Debug($"배열에 새 ARRAY 객체 추가 (비권장)");
                }
            }
        }
        
        // 값 노드 처리 (PROPERTY, VALUE)
        private void ProcessValueNode(SchemeNode node, object parentObject, IXLRow row)
        {
            // KEY 노드의 VALUE 자식인 경우, 이미 KEY 노드에서 처리됐으므로 중복 추가 방지
            if (node.NodeType == SchemeNode.SchemeNodeType.VALUE && 
                node.Parent != null && 
                node.Parent.NodeType == SchemeNode.SchemeNodeType.KEY)
            {
                Logger.Debug($"VALUE 노드가 KEY 노드의 자식이므로 중복 추가 방지를 위해 스킵: 부모={node.Parent.GetKey(row)}");
                return;
            }
            
            // 값 노드의 키 가져오기 (여러 곳에서 사용하므로 먼저 선언)
            string key = GetNodeKey(node, row);
            
            // 부모가 객체인 경우, 키 확인
            if (parentObject is YamlObject)
            {
                if (string.IsNullOrEmpty(key))
                {
                    string errorMessage = $"오류: JSON/YAML 표준에서는 객체 내에 이름 없는 값을 가질 수 없습니다. 노드 타입: {node.NodeType}";
                    Logger.Error(errorMessage);
                    
                    // 오류 창 표시
                    DialogResult result = MessageBox.Show(
                        errorMessage,
                        "JSON/YAML 표준 오류",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error,
                        MessageBoxDefaultButton.Button1,
                        MessageBoxOptions.DefaultDesktopOnly);
                    
                    // 확인 버튼을 누른 경우 변환 과정 중단
                    if (result == DialogResult.OK)
                    {
                        string abortMessage = "사용자 요청에 의해 변환 작업이 중단되었습니다.";
                        Logger.Error(abortMessage);
                        throw new InvalidOperationException(abortMessage);
                    }
                    
                    Logger.Debug($"값 노드의 키가 비어 있어 무시: 타입={node.NodeType}");
                    return;
                }
            }
            
            // 값 가져오기
            object value = node.GetValue(row);
            
            if (value == null || string.IsNullOrEmpty(value.ToString()))
            {
                Logger.Debug($"값이 비어 있어 무시: 노드={node.SchemeName}");
                return;
            }
            
            value = FormatStringValue(value);
            
            // VALUE 노드가 직접 ARRAY의 자식인 경우 (독립 $value) - 직접 값 추가
            if (node.NodeType == SchemeNode.SchemeNodeType.VALUE && 
                node.Parent != null && 
                (node.Parent.NodeType == SchemeNode.SchemeNodeType.ARRAY || 
                 node.Parent.NodeType == SchemeNode.SchemeNodeType.MAP))
            {
                if (parentObject is YamlArray parentArray)
                {
                    parentArray.Add(value);
                    Logger.Debug($"배열에 값 직접 추가 (독립 $value): 값={value}");
                    return;
                }
            }
            
            // 부모가 객체인 경우 - 이미 키를 확인했으므로 중복 선언 제거
            if (parentObject is YamlObject parentMap)
            {
                if (!string.IsNullOrEmpty(key))
                {
                    parentMap.Add(key, value);
                    Logger.Debug($"객체에 값 추가: 키={key}, 값={value}");
                }
            }
            // 부모가 배열인 경우
            else if (parentObject is YamlArray parentArray)
            {
                // $value만 있고 $key가 없는 경우 - 값을 직접 추가
                if (node.NodeType == SchemeNode.SchemeNodeType.VALUE && string.IsNullOrEmpty(key))
                {
                    parentArray.Add(value);
                    Logger.Debug($"배열에 값 직접 추가: 값={value}");
                }
                // 키가 있는 경우 - 객체로 추가
                else if (!string.IsNullOrEmpty(key))
                {
                    // 키가 있는 경우 단일 속성을 가진 객체 추가
                    YamlObject singlePropObj = OrderedYamlFactory.CreateObject();
                    singlePropObj.Add(key, value);
                    parentArray.Add(singlePropObj);
                    Logger.Debug($"배열에 객체로 값 추가: 키={key}, 값={value}");
                }
                else
                {
                    // 안전장치: 키가 없지만 VALUE 노드가 아닌 경우
                    parentArray.Add(value);
                    Logger.Debug($"배열에 값 직접 추가 (기타 케이스): 값={value}");
                }
            }
        }
        
        // KEY 노드 처리
        private void ProcessKeyNode(SchemeNode node, object parentObject, IXLRow row)
        {
            // 셀 값을 키로 사용 (스키마에서 $key 컬럼의 값)
            string cellKey = string.Empty;
            if (node.NodeType == SchemeNode.SchemeNodeType.KEY)
            {
                object cellValue = node.GetValue(row);
                if (cellValue != null && !string.IsNullOrEmpty(cellValue.ToString()))
                {
                    cellKey = cellValue.ToString();
                    Logger.Debug($"KEY 노드에서 셀 값을 키로 사용: {cellKey}");
                }
            }
            
            // 기존 코드는 GetKey를 호출하여 고정된 키("value")를 사용함
            string dynamicKey = node.GetKey(row);
            if (string.IsNullOrEmpty(dynamicKey))
            {
                Logger.Debug("KEY 노드에서 동적 키를 가져올 수 없음");
                return;
            }
            
            // KEY 노드의 VALUE 자식 찾기
            SchemeNode valueNode = node.Children.FirstOrDefault(c => c.NodeType == SchemeNode.SchemeNodeType.VALUE);
            if (valueNode == null)
            {
                Logger.Debug($"KEY 노드에 대응하는 VALUE 노드가 없음: 키={dynamicKey}");
                return;
            }
            
            object value = valueNode.GetValue(row);
            if (value == null || string.IsNullOrEmpty(value.ToString()))
            {
                Logger.Debug($"VALUE 값이 비어 있어 무시: 키={dynamicKey}");
                return;
            }
            
            value = FormatStringValue(value);
            
            // 부모가 객체인 경우
            if (parentObject is YamlObject parentMap)
            {
                parentMap.Add(dynamicKey, value);
                Logger.Debug($"KEY-VALUE 쌍 추가: 키={dynamicKey}, 값={value}");
            }
            // 부모가 배열인 경우 - 실제 셀 값을 키로 사용
            else if (parentObject is YamlArray parentArray)
            {
                // cellKey가 있으면 사용, 없으면 dynamicKey 사용
                string actualKey = !string.IsNullOrEmpty(cellKey) ? cellKey : dynamicKey;
                
                YamlObject keyValueObj = OrderedYamlFactory.CreateObject();
                keyValueObj.Add(actualKey, value);
                parentArray.Add(keyValueObj);
                Logger.Debug($"배열에 동적 키-값 쌍 추가: 키={actualKey}, 값={value}");
            }
        }
        
        // 노드에서 키 가져오기
        private string GetNodeKey(SchemeNode node, IXLRow row)
        {
            // 노드 자체에 키가 있는 경우 직접 사용
            if (!string.IsNullOrEmpty(node.Key))
            {
                return node.Key;
            }
            
            // KEY 타입 노드는 GetKey 메서드로 키를 가져옴
            if (node.NodeType == SchemeNode.SchemeNodeType.KEY)
            {
                return node.GetKey(row);
            }
            
            // 노드에 키가 지정된 자식이 있는 경우 (MAP, ARRAY 타입에서 사용)
            var keyChild = node.Children.FirstOrDefault(c => c.NodeType == SchemeNode.SchemeNodeType.KEY);
            if (keyChild != null)
            {
                return keyChild.GetKey(row);
            }
            
            // 스키마 이름이 있으면 키로 사용
            if (!string.IsNullOrEmpty(node.SchemeName))
            {
                return node.SchemeName;
            }
            
            return string.Empty;
        }
        
        // 문자열 값 포맷
        private object FormatStringValue(object value)
        {
            if (value == null)
                return null;
                
            string strValue = value.ToString();
            if (string.IsNullOrEmpty(strValue))
                return strValue;
                
            // 개행 문자 포함 여부 확인 및 처리
            if (strValue.Contains('\n') || strValue.Contains('\r'))
            {
                // 개행 문자가 이미 이스케이프되어 있는지 확인
                if (!strValue.Contains("\\n") && !strValue.Contains("\\r"))
                {
                    // 새 줄 문자를 이스케이프 처리하지 않고 보존하면
                    // YAML에서 자동으로 블록 스타일을 적용하므로 그대로 반환
                    return strValue;
                }
            }
            
            return strValue;
        }
        
        // 빈 속성 제거
        private bool RemoveEmptyAttributes(object arg)
        {
            bool valueExist = false;
            
            if (arg is string str)
            {
                valueExist = !string.IsNullOrEmpty(str);
            }
            else if (arg is int || arg is long || arg is float || arg is double || arg is decimal)
            {
                valueExist = true;
            }
            else if (arg is bool)
            {
                valueExist = true;
            }
            else if (arg is YamlObject yamlObject)
            {
                var keysToRemove = new List<string>();
                
                foreach (var property in yamlObject.Properties)
                {
                    if (!RemoveEmptyAttributes(property.Value))
                    {
                        keysToRemove.Add(property.Key);
                    }
                    else
                    {
                        valueExist = true;
                    }
                }
                
                foreach (var key in keysToRemove)
                {
                    yamlObject.Remove(key);
                }
            }
            else if (arg is YamlArray yamlArray)
            {
                for (int i = 0; i < yamlArray.Count; i++)
                {
                    var item = yamlArray[i];
                    if (!RemoveEmptyAttributes(item))
                    {
                        yamlArray.RemoveAt(i);
                        i--;
                    }
                    else
                    {
                        valueExist = true;
                    }
                }
            }
            
            return valueExist;
        }
    }
} 