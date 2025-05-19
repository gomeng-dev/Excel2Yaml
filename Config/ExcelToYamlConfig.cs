namespace ExcelToYamlAddin.Config
{
    public enum OutputFormat
    {
        Json,
        Yaml,
        Xml
    }

    public enum YamlStyle
    {
        Block,     // 블록 스타일 (기본값)
        Flow       // 플로우 스타일 (한 줄로 컴팩트하게)
    }

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
            OutputFormat = OutputFormat.Json;
            YamlIndentSize = 2;
            YamlPreserveQuotes = false;
            YamlStyle = YamlStyle.Block;
            IncludeEmptyFields = false;
        }
    }
}
