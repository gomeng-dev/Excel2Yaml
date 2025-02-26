using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Diagnostics;
using ExcelToJsonAddin.Core.YamlPostProcessors;

namespace ExcelToJsonAddin.Core
{
    /// <summary>
    /// YAML 파일의 스키마를 분석하고, 선택적 필드에 빈 값을 추가하는 후처리 클래스
    /// 새로운 YamlOptionalFieldsProcessor를 사용하는 간소화된 버전입니다.
    /// </summary>
    public class YamlSchemaProcessor
    {
        /// <summary>
        /// YAML 파일을 처리하고 선택적 필드에 빈 값을 추가합니다.
        /// </summary>
        /// <param name="yamlPath">YAML 파일 경로</param>
        /// <param name="addEmpty">선택적 필드에 빈 값 추가 여부</param>
        /// <returns>처리 성공 여부</returns>
        public bool ProcessYamlWithEmptyOptionals(string yamlPath, bool addEmpty = true)
        {
            try
            {
                Debug.WriteLine($"[YamlSchemaProcessor] YAML 파일 처리 (간소화된 버전): {yamlPath}");
                
                // addEmpty가 true인 경우에만 처리
                if (addEmpty)
                {
                    var processor = new YamlOptionalFieldsProcessor();
                    return processor.ProcessYamlFile(yamlPath);
                }
                
                // addEmpty가 false인 경우 아무 작업도 수행하지 않고 성공 반환
                Debug.WriteLine($"[YamlSchemaProcessor] 선택적 필드 처리 건너뜀 (추가 옵션이 false): {yamlPath}");
                return true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[YamlSchemaProcessor] 처리 중 오류 발생: {ex.Message}");
                Debug.WriteLine($"[YamlSchemaProcessor] 스택 추적: {ex.StackTrace}");
                return false;
            }
        }
    }
} 