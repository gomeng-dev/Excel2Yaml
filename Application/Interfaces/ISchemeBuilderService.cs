using System.Collections.Generic;
using ExcelToYamlAddin.Domain.Entities;
using YamlDotNet.RepresentationModel;

namespace ExcelToYamlAddin.Application.Interfaces
{
    /// <summary>
    /// 스키마 빌더 서비스 인터페이스
    /// </summary>
    public interface ISchemeBuilderService
    {
        /// <summary>
        /// YAML 노드로부터 스키마 트리를 빌드합니다.
        /// </summary>
        /// <param name="yamlRoot">YAML 루트 노드</param>
        /// <returns>빌드 결과</returns>
        SchemeBuildResult BuildSchemaTree(YamlNode yamlRoot);

        /// <summary>
        /// 스키마 노드로부터 Excel 구조를 생성합니다.
        /// </summary>
        /// <param name="rootNode">루트 스키마 노드</param>
        /// <returns>Excel 구조 정보</returns>
        ExcelStructureInfo BuildExcelStructure(SchemeNode rootNode);
    }

    /// <summary>
    /// 스키마 빌드 결과
    /// </summary>
    public class SchemeBuildResult
    {
        public SchemeNode RootNode { get; set; }
        public int TotalRows { get; set; }
        public int TotalColumns { get; set; }
        public Dictionary<int, List<SchemeNode>> RowNodes { get; set; }
        public Dictionary<string, int> ColumnMappings { get; set; }

        public SchemeBuildResult()
        {
            RowNodes = new Dictionary<int, List<SchemeNode>>();
            ColumnMappings = new Dictionary<string, int>();
        }
    }

    /// <summary>
    /// Excel 구조 정보
    /// </summary>
    public class ExcelStructureInfo
    {
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int StartColumn { get; set; }
        public int EndColumn { get; set; }
        public List<MergedCellInfo> MergedCells { get; set; }

        public ExcelStructureInfo()
        {
            MergedCells = new List<MergedCellInfo>();
        }
    }

    /// <summary>
    /// 병합 셀 정보
    /// </summary>
    public class MergedCellInfo
    {
        public int StartRow { get; set; }
        public int EndRow { get; set; }
        public int StartColumn { get; set; }
        public int EndColumn { get; set; }
    }
}