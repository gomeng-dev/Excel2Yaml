using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace ExcelToYamlAddin.Application.PostProcessing.Processors
{
    /// <summary>
    /// JSON 출력을 포맷팅하는 후처리기입니다.
    /// </summary>
    public class JsonFormatterProcessor : PostProcessorBase
    {
        private readonly ISimpleLogger _logger;

        public JsonFormatterProcessor()
        {
            _logger = SimpleLoggerFactory.CreateLogger<JsonFormatterProcessor>();
        }

        /// <summary>
        /// 처리 우선순위
        /// </summary>
        public override int Priority => 30;

        /// <summary>
        /// 이 프로세서가 처리할 수 있는지 확인합니다.
        /// </summary>
        public override bool CanProcess(ProcessingContext context)
        {
            return context.OutputFormat == OutputFormat.Json;
        }

        /// <summary>
        /// JSON 포맷팅을 수행합니다.
        /// </summary>
        protected override async Task<string> ProcessCoreAsync(string input, ProcessingContext context, CancellationToken cancellationToken)
        {
            _logger.Information("JSON 포맷팅 시작");

            try
            {
                // JSON 파싱 및 재포맷팅
                var jsonObject = JToken.Parse(input);
                
                // 들여쓰기와 포맷팅 적용
                var formatted = jsonObject.ToString(Formatting.Indented);
                
                _logger.Information("JSON 포맷팅 완료");
                return await Task.FromResult(formatted);
            }
            catch (JsonException ex)
            {
                _logger.Error($"JSON 파싱 오류: {ex.Message}", ex);
                // JSON이 유효하지 않은 경우 원본 반환
                return await Task.FromResult(input);
            }
            catch (Exception ex)
            {
                _logger.Error($"JSON 포맷팅 중 오류: {ex.Message}", ex);
                throw;
            }
        }
    }
}