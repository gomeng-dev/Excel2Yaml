using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Xml.Linq;
using ClosedXML.Excel;
using ExcelToYamlAddin.Logging;

namespace ExcelToYamlAddin.Core
{
    /// <summary>
    /// XML 파일을 Excel 시트로 변환하는 클래스
    /// Excel2Yaml의 역변환 기능을 제공합니다.
    /// </summary>
    public class XmlToExcelConverter
    {
        private static readonly ISimpleLogger Logger = SimpleLoggerFactory.CreateLogger<XmlToExcelConverter>();
        
        // XML 속성을 나타내는 접두사
        private const string AttributePrefix = "_";
        
        // 텍스트 내용을 나타내는 특수 키
        private const string TextContentKey = "__text";
        
        // 스키마 마커들
        private const string ArrayMarker = "$[]";
        private const string ObjectMarker = "${}";
        private const string SchemeEndMarker = "$scheme_end";
        private const string IgnoreMarker = "^";

        /// <summary>
        /// XML 파일을 Excel 워크북으로 변환합니다.
        /// </summary>
        /// <param name="xmlContent">XML 내용</param>
        /// <param name="sheetName">생성할 시트 이름</param>
        /// <returns>변환된 Excel 워크북</returns>
        public XLWorkbook ConvertToExcel(string xmlContent, string sheetName = "Sheet1")
        {
            try
            {
                Logger.Information($"XML to Excel 변환 시작: 시트명 = {sheetName}");
                
                // XML 파싱
                var doc = XDocument.Parse(xmlContent);
                var root = doc.Root;
                
                if (root == null)
                {
                    throw new InvalidOperationException("XML 루트 요소를 찾을 수 없습니다.");
                }
                
                // 새 워크북 생성
                var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add($"!{root.Name.LocalName}"); // 루트 요소 이름을 시트명으로 사용
                
                // XML 구조 분석
                var structure = AnalyzeXmlStructure(root);
                
                // 스키마 생성 (2행부터 시작)
                int currentRow = 2;
                currentRow = WriteSchema(worksheet, structure, currentRow);
                
                // 스키마 종료 마커 추가
                worksheet.Cell(currentRow, 1).Value = SchemeEndMarker;
                worksheet.Range(currentRow, 1, currentRow, structure.MaxColumns).Merge();
                worksheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.Red;
                worksheet.Row(currentRow).Style.Font.FontColor = XLColor.White;
                worksheet.Row(currentRow).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                currentRow++;
                
                // 데이터 작성
                WriteData(worksheet, root, structure, currentRow);
                
                // 열 너비 자동 조정
                worksheet.Columns().AdjustToContents();
                
                Logger.Information("XML to Excel 변환 완료");
                return workbook;
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "XML to Excel 변환 중 오류 발생");
                throw;
            }
        }

        /// <summary>
        /// XML 구조를 분석하여 Excel 스키마 정보를 생성합니다.
        /// </summary>
        private XmlStructureInfo AnalyzeXmlStructure(XElement element)
        {
            var info = new XmlStructureInfo
            {
                ElementName = element.Name.LocalName,
                IsArray = false,
                Columns = new List<ColumnInfo>()
            };
            
            // 동일한 이름의 자식 요소가 여러 개 있는지 확인
            var childGroups = element.Elements()
                .GroupBy(e => e.Name.LocalName)
                .ToDictionary(g => g.Key, g => g.Count());
            
            // 루트가 배열인지 확인 (자식 요소가 모두 동일한 이름인 경우)
            // Items -> Item, Item, Item 같은 구조 처리
            if (childGroups.Count == 1 && childGroups.First().Value > 1)
            {
                info.IsArray = true;
                info.ArrayItemName = childGroups.First().Key; // "Item"
                var childElements = element.Elements().ToList();
                if (childElements.Any())
                {
                    // 모든 Item 요소를 분석해서 전체 스키마 생성
                    info.ItemStructure = AnalyzeAllElementsStructure(childElements);
                }
            }
            else
            {
                // 단일 객체 구조 분석
                info.Columns.Add(new ColumnInfo { Name = IgnoreMarker, Level = 0 });
                AnalyzeObjectStructure(element, info.Columns, 1);
            }
            
            info.MaxColumns = CalculateMaxColumns(info);
            return info;
        }

        /// <summary>
        /// 요소의 구조를 분석합니다.
        /// </summary>
        private ElementStructure AnalyzeElementStructure(XElement element)
        {
            var structure = new ElementStructure
            {
                Attributes = element.Attributes().Select(a => a.Name.LocalName).ToList(),
                ChildElements = new Dictionary<string, ElementStructure>(),
                HasTextContent = !string.IsNullOrWhiteSpace(element.Value) && !element.HasElements
            };
            
            // 자식 요소 분석
            var childGroups = element.Elements()
                .GroupBy(e => e.Name.LocalName);
            
            foreach (var group in childGroups)
            {
                var firstChild = group.First();
                structure.ChildElements[group.Key] = AnalyzeElementStructure(firstChild);
                
                // 배열인지 확인
                if (group.Count() > 1)
                {
                    structure.ChildElements[group.Key].IsArray = true;
                }
            }
            
            return structure;
        }

        /// <summary>
        /// 여러 요소를 분석해서 통합된 구조를 생성합니다.
        /// </summary>
        private ElementStructure AnalyzeAllElementsStructure(List<XElement> elements)
        {
            var structure = new ElementStructure
            {
                Attributes = new List<string>(),
                ChildElements = new Dictionary<string, ElementStructure>(),
                HasTextContent = false
            };
            
            // 모든 단순 자식 요소 이름을 수집하여 순서 보장
            var allChildNames = new HashSet<string>();
            var complexChildNames = new HashSet<string>();
            
            // 모든 요소에서 속성과 자식 요소 수집
            foreach (var element in elements)
            {
                // 속성 수집
                foreach (var attr in element.Attributes())
                {
                    if (!structure.Attributes.Contains(attr.Name.LocalName))
                    {
                        structure.Attributes.Add(attr.Name.LocalName);
                    }
                }
                
                // 자식 요소 수집
                foreach (var child in element.Elements())
                {
                    string childName = child.Name.LocalName;
                    allChildNames.Add(childName);
                    
                    if (child.HasElements)
                    {
                        complexChildNames.Add(childName);
                    }
                    
                    if (!structure.ChildElements.ContainsKey(childName))
                    {
                        structure.ChildElements[childName] = AnalyzeElementStructure(child);
                    }
                }
                
                // 텍스트 내용 확인
                if (!string.IsNullOrWhiteSpace(element.Value) && !element.HasElements)
                {
                    structure.HasTextContent = true;
                }
            }
            
            return structure;
        }

        /// <summary>
        /// 객체 구조를 분석하여 컬럼 정보를 생성합니다.
        /// </summary>
        private void AnalyzeObjectStructure(XElement element, List<ColumnInfo> columns, int level)
        {
            // 속성 추가
            foreach (var attr in element.Attributes())
            {
                columns.Add(new ColumnInfo 
                { 
                    Name = AttributePrefix + attr.Name.LocalName,
                    Level = level
                });
            }
            
            // 텍스트 내용이 있으면 추가
            if (!string.IsNullOrWhiteSpace(element.Value) && !element.HasElements)
            {
                columns.Add(new ColumnInfo 
                { 
                    Name = TextContentKey,
                    Level = level
                });
            }
            
            // 자식 요소 처리
            var childGroups = element.Elements()
                .GroupBy(e => e.Name.LocalName);
            
            foreach (var group in childGroups)
            {
                if (group.Count() > 1)
                {
                    // 배열인 경우
                    columns.Add(new ColumnInfo 
                    { 
                        Name = group.Key + ArrayMarker,
                        Level = level,
                        IsArray = true
                    });
                    
                    // 배열 항목의 구조 분석
                    var firstChild = group.First();
                    AnalyzeObjectStructure(firstChild, columns, level + 1);
                }
                else
                {
                    // 단일 요소
                    var child = group.First();
                    if (child.HasElements || child.HasAttributes)
                    {
                        // 복잡한 객체
                        columns.Add(new ColumnInfo 
                        { 
                            Name = group.Key,
                            Level = level,
                            IsObject = true
                        });
                        AnalyzeObjectStructure(child, columns, level + 1);
                    }
                    else
                    {
                        // 단순 값
                        columns.Add(new ColumnInfo 
                        { 
                            Name = group.Key,
                            Level = level
                        });
                    }
                }
            }
        }

        /// <summary>
        /// 최대 컬럼 수를 계산합니다.
        /// </summary>
        private int CalculateMaxColumns(XmlStructureInfo info)
        {
            if (info.IsArray && info.ItemStructure != null)
            {
                return CalculateElementColumns(info.ItemStructure) + 2; // +2 for A, B ignore columns
            }
            return info.Columns.Count;
        }

        /// <summary>
        /// 요소의 컬럼 수를 계산합니다.
        /// </summary>
        private int CalculateElementColumns(ElementStructure structure)
        {
            int count = 0;
            
            // 속성 수
            count += structure.Attributes.Count;
            
            // 텍스트 내용
            if (structure.HasTextContent) count++;
            
            // 자식 요소들
            foreach (var child in structure.ChildElements.Values)
            {
                if (child.IsArray)
                {
                    // 배열의 경우 최대 항목 수를 고려해야 함
                    count += CalculateElementColumns(child) * 3; // 예시로 3개 항목 가정
                }
                else if (child.ChildElements.Any())
                {
                    // 중첩 객체인 경우 하위 요소 수
                    count += child.ChildElements.Count;
                }
                else
                {
                    // 단순 요소
                    count += 1;
                }
            }
            
            return Math.Max(count, 1);
        }

        /// <summary>
        /// Excel에 스키마를 작성합니다.
        /// </summary>
        private int WriteSchema(IXLWorksheet worksheet, XmlStructureInfo structure, int startRow)
        {
            int currentRow = startRow;
            
            if (structure.IsArray)
            {
                // 루트 객체 마커 (시트명이 루트이므로 ${} 마커만)
                worksheet.Cell(currentRow, 1).Value = ObjectMarker;
                worksheet.Range(currentRow, 1, currentRow, structure.MaxColumns).Merge();
                worksheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.LightGreen;
                currentRow++;
                
                // 배열 항목 마커 (예: Item$[])
                worksheet.Cell(currentRow, 1).Value = IgnoreMarker;
                worksheet.Cell(currentRow, 2).Value = structure.ArrayItemName + ArrayMarker;
                worksheet.Range(currentRow, 2, currentRow, structure.MaxColumns).Merge();
                worksheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.Green;
                currentRow++;
                
                // 배열 항목 객체 마커
                worksheet.Cell(currentRow, 1).Value = IgnoreMarker;
                worksheet.Cell(currentRow, 2).Value = IgnoreMarker;
                worksheet.Cell(currentRow, 3).Value = ObjectMarker;
                worksheet.Range(currentRow, 3, currentRow, structure.MaxColumns).Merge();
                worksheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.LightGreen;
                currentRow++;
                
                // 컬럼 헤더 작성
                currentRow = WriteColumnHeaders(worksheet, structure.ItemStructure, currentRow, 3);
            }
            else
            {
                // 단일 객체 스키마 작성
                int col = 1;
                foreach (var column in structure.Columns)
                {
                    worksheet.Cell(currentRow, col).Value = column.Name;
                    col++;
                }
                currentRow++;
            }
            
            return currentRow;
        }

        /// <summary>
        /// 컬럼 헤더를 작성합니다.
        /// </summary>
        private int WriteColumnHeaders(IXLWorksheet worksheet, ElementStructure structure, int row, int startCol)
        {
            int col = startCol;
            
            // 무시 컬럼들 (A, B열)
            worksheet.Cell(row, 1).Value = IgnoreMarker;
            worksheet.Cell(row, 2).Value = IgnoreMarker;
            
            // 속성 컬럼
            foreach (var attr in structure.Attributes)
            {
                worksheet.Cell(row, col).Value = AttributePrefix + attr;
                col++;
            }
            
            // 단순 요소들 먼저 처리
            foreach (var child in structure.ChildElements.Where(c => !c.Value.IsArray && !c.Value.ChildElements.Any()))
            {
                worksheet.Cell(row, col).Value = child.Key;
                col++;
            }
            
            // 텍스트 내용
            if (structure.HasTextContent)
            {
                worksheet.Cell(row, col).Value = TextContentKey;
                col++;
            }
            
            // 복잡한 자식 요소 (중첩 객체)
            foreach (var child in structure.ChildElements.Where(c => !c.Value.IsArray && c.Value.ChildElements.Any()))
            {
                // 중첩 객체 헤더
                int nestedStartCol = col;
                var nestedColumns = GetNestedColumnCount(child.Value);
                
                worksheet.Cell(row, col).Value = child.Key + ObjectMarker;
                if (nestedColumns > 1)
                {
                    worksheet.Range(row, col, row, col + nestedColumns - 1).Merge();
                }
                
                // 중첩 객체의 하위 구조
                int subRow = row + 1;
                // A, B열까지 무시 컬럼으로 채우기
                for (int i = 1; i < nestedStartCol; i++)
                {
                    worksheet.Cell(subRow, i).Value = IgnoreMarker;
                }
                
                int subCol = nestedStartCol;
                foreach (var nestedChild in child.Value.ChildElements)
                {
                    worksheet.Cell(subRow, subCol).Value = nestedChild.Key;
                    subCol++;
                }
                
                col += nestedColumns;
            }
            
            // 배열 요소
            foreach (var child in structure.ChildElements.Where(c => c.Value.IsArray))
            {
                // 배열 헤더
                int arrayStartCol = col;
                int arrayColumns = CalculateElementColumns(child.Value);
                
                worksheet.Cell(row, col).Value = child.Key + ArrayMarker;
                if (arrayColumns > 1)
                {
                    worksheet.Range(row, col, row, col + arrayColumns - 1).Merge();
                }
                
                // 배열 항목 구조
                WriteArrayItemSchema(worksheet, child.Value, row + 1, col, arrayColumns);
                
                col += arrayColumns;
            }
            
            return row + 2; // 중첩 구조를 위해 추가 행 필요
        }
        
        /// <summary>
        /// 중첩된 객체의 컬럼 수를 계산합니다.
        /// </summary>
        private int GetNestedColumnCount(ElementStructure structure)
        {
            return structure.ChildElements.Count; // Requirements의 경우 Level, Class = 2개
        }

        /// <summary>
        /// 배열 항목의 스키마를 작성합니다.
        /// </summary>
        private void WriteArrayItemSchema(IXLWorksheet worksheet, ElementStructure structure, int row, int startCol, int totalColumns)
        {
            // 예시로 3개 항목을 위한 스키마 작성
            int itemColumns = totalColumns / 3;
            for (int i = 0; i < 3; i++)
            {
                int col = startCol + (i * itemColumns);
                worksheet.Cell(row, col).Value = ObjectMarker;
                if (itemColumns > 1)
                {
                    worksheet.Range(row, col, row, col + itemColumns - 1).Merge();
                }
            }
        }

        /// <summary>
        /// XML 데이터를 Excel에 작성합니다.
        /// </summary>
        private void WriteData(IXLWorksheet worksheet, XElement root, XmlStructureInfo structure, int startRow)
        {
            if (structure.IsArray)
            {
                // 배열 데이터 작성
                int row = startRow;
                foreach (var element in root.Elements())
                {
                    WriteElementData(worksheet, element, structure.ItemStructure, row, 3); // 3번째 컬럼에서 시작
                    row++;
                }
            }
            else
            {
                // 단일 객체 데이터 작성
                WriteObjectData(worksheet, root, structure.Columns, startRow, 1);
            }
        }

        /// <summary>
        /// 요소 데이터를 작성합니다.
        /// </summary>
        private void WriteElementData(IXLWorksheet worksheet, XElement element, ElementStructure structure, int row, int startCol)
        {
            int col = startCol;
            
            // 앞의 두 열을 빈 셀로 처리 (A, B열)
            worksheet.Cell(row, 1).Value = "";
            worksheet.Cell(row, 2).Value = "";
            
            // 속성 값
            foreach (var attrName in structure.Attributes)
            {
                var attr = element.Attribute(attrName);
                worksheet.Cell(row, col).Value = attr?.Value ?? "";
                col++;
            }
            
            // 단순 요소들 먼저 처리
            foreach (var child in structure.ChildElements.Where(c => !c.Value.IsArray && !c.Value.ChildElements.Any()))
            {
                var childElement = element.Element(child.Key);
                if (childElement != null)
                {
                    worksheet.Cell(row, col).Value = childElement.Value;
                }
                col++;
            }
            
            // 텍스트 내용
            if (structure.HasTextContent)
            {
                worksheet.Cell(row, col).Value = element.Value;
                col++;
            }
            
            // 복잡한 자식 요소 (중첩 객체)
            foreach (var child in structure.ChildElements.Where(c => !c.Value.IsArray && c.Value.ChildElements.Any()))
            {
                var childElement = element.Element(child.Key);
                if (childElement != null)
                {
                    // 중첩 객체의 자식 요소들
                    foreach (var nestedChild in child.Value.ChildElements)
                    {
                        var nestedElement = childElement.Element(nestedChild.Key);
                        if (nestedElement != null)
                        {
                            worksheet.Cell(row, col).Value = nestedElement.Value;
                        }
                        col++;
                    }
                }
                else
                {
                    col += GetNestedColumnCount(child.Value) - 1; // -1 because we don't skip ignore column here
                }
            }
            
            // 배열 요소
            foreach (var child in structure.ChildElements.Where(c => c.Value.IsArray))
            {
                var childElements = element.Elements(child.Key).ToList();
                
                // 배열 항목들을 가로로 나열
                foreach (var childElement in childElements)
                {
                    WriteElementData(worksheet, childElement, child.Value, row, col);
                    col += CalculateElementColumns(child.Value);
                }
            }
        }

        /// <summary>
        /// 객체 데이터를 작성합니다.
        /// </summary>
        private void WriteObjectData(IXLWorksheet worksheet, XElement element, List<ColumnInfo> columns, int row, int startCol)
        {
            int col = startCol;
            
            foreach (var column in columns)
            {
                if (column.Name == IgnoreMarker)
                {
                    worksheet.Cell(row, col).Value = "";
                }
                else if (column.Name.StartsWith(AttributePrefix))
                {
                    var attrName = column.Name.Substring(AttributePrefix.Length);
                    var attr = element.Attribute(attrName);
                    worksheet.Cell(row, col).Value = attr?.Value ?? "";
                }
                else if (column.Name == TextContentKey)
                {
                    worksheet.Cell(row, col).Value = element.Value;
                }
                else
                {
                    var childElement = element.Element(column.Name.Replace(ArrayMarker, ""));
                    if (childElement != null)
                    {
                        worksheet.Cell(row, col).Value = childElement.Value;
                    }
                }
                col++;
            }
        }

        // 내부 클래스들
        private class XmlStructureInfo
        {
            public string ElementName { get; set; }
            public bool IsArray { get; set; }
            public string ArrayItemName { get; set; }
            public List<ColumnInfo> Columns { get; set; }
            public ElementStructure ItemStructure { get; set; }
            public int MaxColumns { get; set; }
        }

        private class ColumnInfo
        {
            public string Name { get; set; }
            public int Level { get; set; }
            public bool IsArray { get; set; }
            public bool IsObject { get; set; }
        }

        private class ElementStructure
        {
            public List<string> Attributes { get; set; } = new List<string>();
            public Dictionary<string, ElementStructure> ChildElements { get; set; } = new Dictionary<string, ElementStructure>();
            public bool HasTextContent { get; set; }
            public bool IsArray { get; set; }
        }
    }
}