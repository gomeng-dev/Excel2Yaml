namespace ExcelToJsonAddin.Config
{
    public enum OutputFormat
    {
        JSON,
        YAML
    }
    
    public enum YamlStyle
    {
        Block,     // 블록 스타일 (기본값)
        Flow       // 플로우 스타일 (한 줄로 컴팩트하게)
    }

    public class ExcelToJsonConfig
    {
        public bool EnableHashFileGen { get; set; }
        public string WorkingDirectory { get; set; }
        public OutputFormat Format { get; set; }
        public int YamlIndentSize { get; set; }
        public bool YamlPreserveQuotes { get; set; }
        public YamlStyle YamlStyle { get; set; }
        public bool IncludeEmptyOptionals { get; set; }

        public ExcelToJsonConfig()
        {
            EnableHashFileGen = false;
            WorkingDirectory = System.IO.Directory.GetCurrentDirectory();
            Format = OutputFormat.JSON;
            YamlIndentSize = 2;
            YamlPreserveQuotes = false;
            YamlStyle = YamlStyle.Block;
            IncludeEmptyOptionals = false;
        }
    }
}
