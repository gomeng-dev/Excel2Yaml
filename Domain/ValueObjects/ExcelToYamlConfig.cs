namespace ExcelToYamlAddin.Domain.ValueObjects
{
    /// <summary>
    /// Excel to YAML 변환 설정을 나타내는 값 객체
    /// </summary>
    public class ExcelToYamlConfig
    {
        public bool EnableHashGen { get; set; }
        public string WorkingDirectory { get; set; }
        public OutputFormat OutputFormat { get; set; }
        public int YamlIndentSize { get; set; }
        public bool YamlPreserveQuotes { get; set; }
        public YamlStyle YamlStyle { get; set; }
        public bool IncludeEmptyFields { get; set; }

        public ExcelToYamlConfig()
        {
            EnableHashGen = false;
            WorkingDirectory = System.IO.Directory.GetCurrentDirectory();
            OutputFormat = OutputFormat.Json;  // 이미 정적 필드로 정의됨
            YamlIndentSize = 2;
            YamlPreserveQuotes = false;
            YamlStyle = YamlStyle.Block;  // 이미 정적 필드로 정의됨
            IncludeEmptyFields = false;
        }
    }
}
