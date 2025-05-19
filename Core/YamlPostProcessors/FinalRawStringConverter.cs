using System;
using System.Diagnostics;
using System.IO;
using System.Text; // For StringBuilder
// YamlDotNet을 사용하여 파싱/재직렬화 하는 대신 텍스트 직접 처리

namespace ExcelToYamlAddin.Core.YamlPostProcessors
{
    public class FinalRawStringConverter
    {
        public bool ProcessYamlFile(string yamlPath)
        {
            try
            {
                Debug.WriteLine($"[FinalRawStringConverter] YAML 파일 처리 시작: {yamlPath}");

                if (!File.Exists(yamlPath))
                {
                    Debug.WriteLine($"[FinalRawStringConverter] 오류: YAML 파일을 찾을 수 없습니다: {yamlPath}");
                    return false;
                }

                string originalContent = File.ReadAllText(yamlPath);
                if (string.IsNullOrWhiteSpace(originalContent))
                {
                    Debug.WriteLine($"[FinalRawStringConverter] YAML 파일이 비어있거나 공백만 포함합니다: {yamlPath}");
                    return true; // 빈 파일은 처리할 내용 없음
                }

                // YAML 문자열 내에서만 이스케이프 시퀀스를 변환해야 하므로,
                // 정규 표현식을 사용하여 큰따옴표로 묶인 문자열을 찾는 것이 더 안전할 수 있습니다.
                // 하지만 단순화를 위해 전체 텍스트에 대해 치환을 시도합니다.
                // 이는 주석이나 다른 부분에 의도치 않은 변경을 유발할 수 있으므로 주의해야 합니다.
                // 더 정교한 방법은 YAML 파서를 사용하여 문자열 노드만 대상으로 하는 것입니다.
                // 하지만 현재 "들여쓰기 오류" 때문에 파서 사용이 어려우므로 텍스트 기반으로 갑니다.

                StringBuilder sb = new StringBuilder(originalContent.Length);
                bool inDoubleQuotedString = false;
                char prevChar = '\0';

                for (int i = 0; i < originalContent.Length; i++)
                {
                    char currentChar = originalContent[i];

                    if (currentChar == '"' && prevChar != '\\') // 이스케이프되지 않은 따옴표
                    {
                        inDoubleQuotedString = !inDoubleQuotedString;
                        sb.Append(currentChar);
                    }
                    else if (inDoubleQuotedString && currentChar == '\\' && i + 1 < originalContent.Length)
                    {
                        char nextChar = originalContent[i + 1];
                        switch (nextChar)
                        {
                            case 'n':
                                sb.Append('\n'); // \\n -> 실제 줄바꿈 문자로 변경
                                i++; // 다음 문자까지 처리했으므로 인덱스 증가
                                break;
                            case 'r':
                                sb.Append('\r'); // \\r -> 실제 캐리지리턴 문자로 변경
                                i++;
                                break;
                            case '"':
                                sb.Append('"');  // \\" -> " 로 변경
                                i++;
                                break;
                            case '\\':
                                sb.Append('\\'); // \\\\ -> \ 로 변경
                                i++;
                                break;
                            // 필요한 다른 이스케이프 시퀀스 추가 (예: \t)
                            case 't':
                                sb.Append('\t'); // \\t -> 실제 탭 문자로 변경
                                i++;
                                break;
                            default:
                                // 인식할 수 없는 이스케이프 시퀀스는 그대로 둠 (예: \a)
                                sb.Append(currentChar);
                                sb.Append(nextChar);
                                i++;
                                break;
                        }
                    }
                    else
                    {
                        sb.Append(currentChar);
                    }
                    prevChar = currentChar;
                }

                string modifiedContent = sb.ToString();

                // 변경 사항이 있을 경우에만 파일 쓰기
                if (originalContent != modifiedContent)
                {
                    File.WriteAllText(yamlPath, modifiedContent);
                    Debug.WriteLine($"[FinalRawStringConverter] YAML 파일 내용 변환 완료: {yamlPath}");
                }
                else
                {
                    Debug.WriteLine($"[FinalRawStringConverter] 변경 사항 없음, 파일 쓰기 건너뜀: {yamlPath}");
                }

                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[FinalRawStringConverter] 처리 중 오류 발생: {ex.Message}");
                Debug.WriteLine($"[FinalRawStringConverter] 스택 추적: {ex.StackTrace}");
                return false;
            }
        }

        // 이 후처리기를 Ribbon.cs 등에서 호출하는 로직 필요
        // 예를 들어, 모든 다른 후처리가 끝난 후 마지막으로 호출
        public static bool Process(string filePath) // 정적 메서드로 만들어 호출 용이하게
        {
            var processor = new FinalRawStringConverter();
            return processor.ProcessYamlFile(filePath);
        }
    }
}