using ExcelToYamlAddin.Domain.ValueObjects;
using ExcelToYamlAddin.Infrastructure.Logging;
using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace ExcelToYamlAddin.Application.PostProcessing.Processors
{
    /// <summary>
    /// XML 출력을 포맷팅하는 후처리기입니다.
    /// </summary>
    public class XmlFormatterProcessor : PostProcessorBase
    {
        private readonly ISimpleLogger _logger;

        public XmlFormatterProcessor()
        {
            _logger = SimpleLoggerFactory.CreateLogger<XmlFormatterProcessor>();
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
            return context.OutputFormat == OutputFormat.Xml;
        }

        /// <summary>
        /// XML 포맷팅을 수행합니다.
        /// </summary>
        protected override async Task<string> ProcessCoreAsync(string input, ProcessingContext context, CancellationToken cancellationToken)
        {
            _logger.Information("XML 포맷팅 시작");

            try
            {
                // XML 파싱 및 재포맷팅
                var doc = XDocument.Parse(input);
                
                // 들여쓰기 설정
                var settings = new XmlWriterSettings
                {
                    Indent = true,
                    IndentChars = "  ",
                    NewLineChars = "\r\n",
                    NewLineHandling = NewLineHandling.Replace,
                    Encoding = new UTF8Encoding(false),
                    OmitXmlDeclaration = false
                };

                // 포맷팅된 XML 생성
                using (var stringWriter = new StringWriter())
                {
                    using (var xmlWriter = XmlWriter.Create(stringWriter, settings))
                    {
                        doc.Save(xmlWriter);
                    }
                    
                    var formatted = stringWriter.ToString();
                    _logger.Information("XML 포맷팅 완료");
                    return await Task.FromResult(formatted);
                }
            }
            catch (XmlException ex)
            {
                _logger.Error($"XML 파싱 오류: {ex.Message}", ex);
                // XML이 유효하지 않은 경우 원본 반환
                return await Task.FromResult(input);
            }
            catch (Exception ex)
            {
                _logger.Error($"XML 포맷팅 중 오류: {ex.Message}", ex);
                throw;
            }
        }
    }
}