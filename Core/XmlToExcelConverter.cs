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
                
                // 1행은 주석 행으로 비워둠 (SchemeParser 규칙)
                worksheet.Cell(1, 1).Value = ""; // COMMENT_ROW_NUM = 0 (1행)
                
                // XML 구조 분석
                var structure = AnalyzeXmlStructure(root);
                
                // 스키마 생성 (2행부터 시작 - SchemeParser 규칙 준수)
                int currentRow = 2;
                currentRow = WriteSchema(worksheet, structure, currentRow);
                
                // 실제 사용된 컬럼 수 확인
                int actualMaxColumns = GetActualUsedColumns(worksheet, currentRow - 1);
                Logger.Information($"실제 사용된 컬럼 수: {actualMaxColumns}, 계산된 MaxColumns: {structure.MaxColumns}");
                
                // 스키마 종료 마커 추가 (실제 사용된 컬럼까지 병합)
                worksheet.Cell(currentRow, 1).Value = SchemeEndMarker;
                if (actualMaxColumns > 1)
                {
                    worksheet.Range(currentRow, 1, currentRow, actualMaxColumns).Merge();
                }
                Logger.Information($"스키마 종료 마커 작성: 행={currentRow}, 실제MaxColumns={actualMaxColumns}");
                
                // 이전 행들의 병합도 실제 컬럼 수로 수정
                UpdateSchemaMerging(worksheet, structure, actualMaxColumns, currentRow - 1);
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
            bool hasTextContent = !string.IsNullOrWhiteSpace(element.Value) && !element.HasElements;
            Logger.Debug($"요소 '{element.Name.LocalName}' 텍스트 내용 분석: Value='{element.Value}', HasElements={element.HasElements}, HasTextContent={hasTextContent}");
            
            var structure = new ElementStructure
            {
                Attributes = element.Attributes()
                    .Where(a => !a.IsNamespaceDeclaration)
                    .Select(a => a.Name.LocalName).ToList(),
                ChildElements = new Dictionary<string, ElementStructure>(),
                HasTextContent = hasTextContent
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
            var allChildNames = new List<string>();
            var complexChildNames = new HashSet<string>();
            
            Logger.Information($"전체 {elements.Count}개 요소 구조 분석 시작");
            
            // 모든 요소에서 속성과 자식 요소 수집
            foreach (var element in elements)
            {
                Logger.Debug($"요소 분석: {element.Name.LocalName}");
                
                // 속성 수집 (xmlns 같은 네임스페이스 속성 제외)
                foreach (var attr in element.Attributes())
                {
                    var attrName = attr.Name.LocalName;
                    if (!attr.IsNamespaceDeclaration && !structure.Attributes.Contains(attrName))
                    {
                        structure.Attributes.Add(attrName);
                        Logger.Debug($"속성 추가: {attrName}");
                    }
                }
                
                // 자식 요소 수집 (순서 보장)
                foreach (var child in element.Elements())
                {
                    string childName = child.Name.LocalName;
                    if (!allChildNames.Contains(childName))
                    {
                        allChildNames.Add(childName);
                        Logger.Debug($"자식 요소 추가: {childName}");
                    }
                    
                    if (child.HasElements || child.HasAttributes)
                    {
                        complexChildNames.Add(childName);
                        Logger.Debug($"복잡한 요소로 분류: {childName} (속성: {child.Attributes().Count()}, 자식: {child.Elements().Count()})");
                    }
                    
                    if (!structure.ChildElements.ContainsKey(childName))
                    {
                        structure.ChildElements[childName] = AnalyzeElementStructure(child);
                        Logger.Debug($"새 구조 생성: {childName}");
                    }
                    else
                    {
                        // 기존 구조와 병합 (같은 이름이지만 다른 구조일 수 있음)
                        Logger.Debug($"구조 병합: {childName}");
                        MergeElementStructures(structure.ChildElements[childName], child);
                    }
                }
                
                // 텍스트 내용 확인
                if (!string.IsNullOrWhiteSpace(element.Value) && !element.HasElements)
                {
                    structure.HasTextContent = true;
                }
            }
            
            Logger.Information($"구조 분석 완료 - 속성: {structure.Attributes.Count}, 자식: {structure.ChildElements.Count}, 복잡한 요소: {complexChildNames.Count}");
            
            return structure;
        }

        /// <summary>
        /// 두 ElementStructure를 병합합니다.
        /// </summary>
        private void MergeElementStructures(ElementStructure existing, XElement newElement)
        {
            Logger.Debug($"구조 병합 시작: {newElement.Name.LocalName}");
            
            // 새로운 요소의 구조 분석
            var newStructure = AnalyzeElementStructure(newElement);
            
            // 속성 병합
            int initialAttrCount = existing.Attributes.Count;
            foreach (var attr in newStructure.Attributes)
            {
                if (!existing.Attributes.Contains(attr))
                {
                    existing.Attributes.Add(attr);
                    Logger.Debug($"병합: 새 속성 추가 {attr}");
                }
            }
            Logger.Debug($"속성 병합 완료: {initialAttrCount} → {existing.Attributes.Count}");
            
            // 자식 요소 병합
            int initialChildCount = existing.ChildElements.Count;
            foreach (var kvp in newStructure.ChildElements)
            {
                if (!existing.ChildElements.ContainsKey(kvp.Key))
                {
                    existing.ChildElements[kvp.Key] = kvp.Value;
                    Logger.Debug($"병합: 새 자식 요소 추가 {kvp.Key}");
                }
                else
                {
                    // 재귀적으로 병합
                    Logger.Debug($"병합: 기존 자식 요소 병합 {kvp.Key}");
                    MergeElementStructuresRecursive(existing.ChildElements[kvp.Key], kvp.Value);
                }
            }
            Logger.Debug($"자식 요소 병합 완료: {initialChildCount} → {existing.ChildElements.Count}");
            
            // 텍스트 내용 병합
            if (newStructure.HasTextContent)
            {
                existing.HasTextContent = true;
                Logger.Debug($"병합: 텍스트 내용 추가");
            }
            
            // 배열 여부 병합
            if (newStructure.IsArray)
            {
                existing.IsArray = true;
            }
        }
        
        /// <summary>
        /// 두 ElementStructure를 재귀적으로 병합합니다.
        /// </summary>
        private void MergeElementStructuresRecursive(ElementStructure existing, ElementStructure newStructure)
        {
            // 속성 병합
            foreach (var attr in newStructure.Attributes)
            {
                if (!existing.Attributes.Contains(attr))
                {
                    existing.Attributes.Add(attr);
                }
            }
            
            // 자식 요소 병합
            foreach (var kvp in newStructure.ChildElements)
            {
                if (!existing.ChildElements.ContainsKey(kvp.Key))
                {
                    existing.ChildElements[kvp.Key] = kvp.Value;
                }
                else
                {
                    // 재귀적으로 병합
                    MergeElementStructuresRecursive(existing.ChildElements[kvp.Key], kvp.Value);
                }
            }
            
            // 텍스트 내용 병합
            if (newStructure.HasTextContent)
            {
                existing.HasTextContent = true;
                Logger.Debug($"병합: 텍스트 내용 추가");
            }
            
            // 배열 여부 병합
            if (newStructure.IsArray)
            {
                existing.IsArray = true;
            }
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
                int totalColumns = CalculateElementColumns(info.ItemStructure) + 2; // +2 for A, B ignore columns
                Logger.Information($"최대 컬럼 수 계산: 배열 구조, 총 {totalColumns}개 컬럼");
                return totalColumns;
            }
            Logger.Information($"최대 컬럼 수 계산: 단일 객체, 총 {info.Columns.Count}개 컬럼");
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
                else
                {
                    // 모든 중첩 요소는 GetNestedColumnCount와 동일하게 계산
                    count += GetNestedColumnCount(child);
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
                if (structure.MaxColumns > 1)
                {
                    worksheet.Range(currentRow, 1, currentRow, structure.MaxColumns).Merge();
                }
                worksheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.LightGreen;
                currentRow++;
                
                // 배열 항목 마커 (예: Item$[])
                worksheet.Cell(currentRow, 1).Value = IgnoreMarker;
                worksheet.Cell(currentRow, 2).Value = structure.ArrayItemName + ArrayMarker;
                if (structure.MaxColumns > 2)
                {
                    worksheet.Range(currentRow, 2, currentRow, structure.MaxColumns).Merge();
                }
                worksheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.Green;
                currentRow++;
                
                // 배열 항목 객체 마커
                worksheet.Cell(currentRow, 1).Value = IgnoreMarker;
                worksheet.Cell(currentRow, 2).Value = IgnoreMarker;
                worksheet.Cell(currentRow, 3).Value = ObjectMarker;
                if (structure.MaxColumns > 3)
                {
                    worksheet.Range(currentRow, 3, currentRow, structure.MaxColumns).Merge();
                }
                worksheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.LightGreen;
                currentRow++;
                
                // 컬럼 헤더 작성
                Logger.Information($"컬럼 헤더 작성 전: currentRow={currentRow}");
                currentRow = WriteColumnHeaders(worksheet, structure.ItemStructure, currentRow, 3);
                Logger.Information($"컬럼 헤더 작성 후: currentRow={currentRow}");
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
        /// 컬럼 헤더를 작성합니다. (무한 행 확장 지원)
        /// </summary>
        private int WriteColumnHeaders(IXLWorksheet worksheet, ElementStructure structure, int row, int startCol)
        {
            int currentRow = row;
            int col = startCol;
            
            // 무시 컬럼들 (A, B열)
            worksheet.Cell(currentRow, 1).Value = IgnoreMarker;
            worksheet.Cell(currentRow, 2).Value = IgnoreMarker;
            
            // 속성 컬럼
            foreach (var attr in structure.Attributes)
            {
                worksheet.Cell(currentRow, col).Value = AttributePrefix + attr;
                col++;
            }
            
            // XML 태그 순서를 보장하기 위해 자식 요소들을 원래 순서대로 처리
            int maxNestedRows = currentRow;
            bool hasAttributeOnlyObjects = structure.ChildElements.Any(c => !c.Value.IsArray && c.Value.Attributes.Any() && !c.Value.ChildElements.Any());
            bool attributeRowInitialized = false;
            
            // 복잡한 객체가 있으면 6행을 ^ 마커로 미리 초기화
            bool hasComplexObjects = structure.ChildElements.Any(c => !c.Value.IsArray && (c.Value.Attributes.Any() || c.Value.HasTextContent));
            if (hasComplexObjects)
            {
                int attrRow = currentRow + 1;
                worksheet.Cell(attrRow, 1).Value = IgnoreMarker;
                worksheet.Cell(attrRow, 2).Value = IgnoreMarker;
                
                // 모든 컬럼을 ^ 마커로 초기화 (나중에 복잡한 객체들이 덮어씀)
                int totalCols = col + structure.ChildElements.Sum(c => GetNestedColumnCount(c.Value));
                for (int i = startCol; i <= totalCols; i++)
                {
                    worksheet.Cell(attrRow, i).Value = IgnoreMarker;
                }
                
                attributeRowInitialized = true;
                maxNestedRows = Math.Max(maxNestedRows, attrRow);
                Logger.Debug($"6행 ^ 마커로 미리 초기화: 행={attrRow}, 컬럼={startCol}~{totalCols}");
            }
            
            foreach (var child in structure.ChildElements)
            {
                if (child.Value.IsArray) continue; // 배열은 나중에 처리
                
                // 단순 요소 (속성도 자식 요소도 없음)
                if (!child.Value.ChildElements.Any() && !child.Value.Attributes.Any())
                {
                    worksheet.Cell(currentRow, col).Value = child.Key;
                    col++;
                }
                // 복잡한 요소 (속성이나 자식 요소가 있음)
                else
                {
                    // 중첩 객체 헤더
                    int nestedStartCol = col;
                        var nestedColumns = GetNestedColumnCount(child.Value);
                    
                    worksheet.Cell(currentRow, col).Value = child.Key + ObjectMarker;
                    if (nestedColumns > 1)
                    {
                        worksheet.Range(currentRow, col, currentRow, col + nestedColumns - 1).Merge();
                    }
                    worksheet.Row(currentRow).Style.Fill.BackgroundColor = XLColor.LightGreen;
                    
                    // 속성만 있는 경우와 자식 요소가 있는 경우를 구분해서 처리
                    if (child.Value.Attributes.Any() && !child.Value.ChildElements.Any() && !child.Value.HasTextContent)
                    {
                        // 속성만 있는 경우(텍스트 내용 없음): 바로 다음 행에 속성 헤더 작성
                        int attrRow = currentRow + 1;
                        Logger.Debug($"속성만 있는 객체 처리: {child.Key}, 속성 수={child.Value.Attributes.Count}");
                        
                        // 속성 행 초기화 (한 번만)
                        if (!attributeRowInitialized)
                        {
                            // A, B열에 무시 마커 작성
                            worksheet.Cell(attrRow, 1).Value = IgnoreMarker;
                            worksheet.Cell(attrRow, 2).Value = IgnoreMarker;
                            attributeRowInitialized = true;
                            Logger.Debug($"속성 행 초기화: 행={attrRow}");
                        }
                        
                        // 속성들 작성 (무시 마커 덮어쓰지 않음)
                        int attrCol = nestedStartCol;
                        foreach (var attr in child.Value.Attributes)
                        {
                            worksheet.Cell(attrRow, attrCol).Value = AttributePrefix + attr;
                            Logger.Debug($"속성 헤더 작성: 행={attrRow}, 열={attrCol}, 속성={AttributePrefix + attr}");
                            attrCol++;
                        }
                        
                        maxNestedRows = Math.Max(maxNestedRows, attrRow);
                    }
                    else
                    {
                        // 자식 요소가 있는 경우나 속성+텍스트가 모두 있는 경우: 속성 행에 속성과 텍스트 내용 쓰기
                        int attrRow = currentRow + 1;
                        
                        // 속성 행 초기화 (한 번만)
                        if (!attributeRowInitialized)
                        {
                            worksheet.Cell(attrRow, 1).Value = IgnoreMarker;
                            worksheet.Cell(attrRow, 2).Value = IgnoreMarker;
                            attributeRowInitialized = true;
                            Logger.Debug($"속성 행 초기화: 행={attrRow}");
                        }
                        
                        // 속성들과 텍스트 내용을 속성 행에 작성
                        int attrCol = nestedStartCol;
                        foreach (var attr in child.Value.Attributes)
                        {
                            worksheet.Cell(attrRow, attrCol).Value = AttributePrefix + attr;
                            Logger.Debug($"복잡 객체 속성 헤더 작성: 행={attrRow}, 열={attrCol}, 속성={AttributePrefix + attr}");
                            attrCol++;
                        }
                        
                        // 텍스트 내용이 있는 경우
                        if (child.Value.HasTextContent)
                        {
                            worksheet.Cell(attrRow, attrCol).Value = TextContentKey;
                            Logger.Debug($"복잡 객체 텍스트 헤더 작성: 행={attrRow}, 열={attrCol}");
                            attrCol++;
                        }
                        
                        maxNestedRows = Math.Max(maxNestedRows, attrRow);
                        
                        // 더 깊은 중첩이 있는 경우 재귀 처리
                        if (child.Value.ChildElements.Any())
                        {
                            int nestedRowEnd = WriteNestedObjectHeaders(worksheet, child.Value, attrRow + 1, nestedStartCol, maxNestedRows);
                            maxNestedRows = Math.Max(maxNestedRows, nestedRowEnd);
                        }
                    }
                    
                    col += nestedColumns;
                }
            }
            
            // 텍스트 내용
            if (structure.HasTextContent)
            {
                worksheet.Cell(currentRow, col).Value = TextContentKey;
                col++;
            }
            
            // 배열 요소
            foreach (var child in structure.ChildElements.Where(c => c.Value.IsArray))
            {
                // 배열 헤더
                int arrayStartCol = col;
                int arrayColumns = CalculateElementColumns(child.Value);
                
                worksheet.Cell(currentRow, col).Value = child.Key + ArrayMarker;
                if (arrayColumns > 1)
                {
                    worksheet.Range(currentRow, col, currentRow, col + arrayColumns - 1).Merge();
                }
                
                // 배열 항목 구조
                WriteArrayItemSchema(worksheet, child.Value, currentRow + 1, col, arrayColumns);
                
                col += arrayColumns;
            }
            
            // 텍스트 내용 처리 후 최종 행 계산
            int finalRow = Math.Max(maxNestedRows, currentRow);
            
            Logger.Information($"스키마 헤더 작성 완료: 시작행={row}, 현재행={currentRow}, 최종행={finalRow}, 중첩최대행={maxNestedRows}, 반환값={finalRow + 1}");
            
            return finalRow + 1; // 다음 스키마 요소를 위한 행
        }
        
        /// <summary>
        /// 중첩 객체의 헤더를 재귀적으로 작성합니다. (무한 확장 지원)
        /// </summary>
        private int WriteNestedObjectHeaders(IXLWorksheet worksheet, ElementStructure structure, int row, int startCol, int currentMaxRow)
        {
            Logger.Debug($"중첩 헤더 작성: 행={row}, 시작열={startCol}, 속성수={structure.Attributes.Count}, 자식수={structure.ChildElements.Count}");
            
            int maxRow = Math.Max(row, currentMaxRow);
            
            // A, B열부터 시작 컬럼 전까지 무시 마커(^)로 채우기
            for (int i = 1; i < startCol; i++)
            {
                worksheet.Cell(row, i).Value = IgnoreMarker;
            }
            
            int col = startCol;
            
            // 속성들 먼저 처리 (_속성명 형태)
            foreach (var attr in structure.Attributes)
            {
                worksheet.Cell(row, col).Value = AttributePrefix + attr;
                Logger.Debug($"속성 헤더 작성: 행={row}, 열={col}, 속성={AttributePrefix + attr}");
                col++;
            }
            
            // 단순 자식 요소들 처리
            foreach (var child in structure.ChildElements.Where(c => !c.Value.IsArray && !c.Value.ChildElements.Any() && !c.Value.Attributes.Any()))
            {
                worksheet.Cell(row, col).Value = child.Key;
                Logger.Debug($"단순 요소 헤더 작성: 행={row}, 열={col}, 요소={child.Key}");
                col++;
            }
            
            // 텍스트 내용이 있는 경우
            if (structure.HasTextContent)
            {
                worksheet.Cell(row, col).Value = TextContentKey;
                Logger.Debug($"텍스트 헤더 작성: 행={row}, 열={col}");
                col++;
            }
            
            // 더 깊은 중첩 객체들 재귀 처리
            foreach (var child in structure.ChildElements.Where(c => !c.Value.IsArray && (c.Value.ChildElements.Any() || c.Value.Attributes.Any())))
            {
                int nestedColumns = GetNestedColumnCount(child.Value);
                
                Logger.Debug($"더 깊은 중첩 처리: {child.Key}, 컬럼수={nestedColumns}");
                
                // 재귀적으로 더 깊은 레벨 처리 (현재 행의 속성들 건너뛰고)
                int nestedRowEnd = WriteNestedObjectHeaders(worksheet, child.Value, row, col, maxRow);
                maxRow = Math.Max(maxRow, nestedRowEnd);
                
                col += nestedColumns;
            }
            
            // 현재 행에 내용이 있으면 사용된 행으로 카운트
            bool hasContent = structure.Attributes.Any() || 
                             structure.ChildElements.Any(c => !c.Value.IsArray && !c.Value.ChildElements.Any() && !c.Value.Attributes.Any()) ||
                             structure.HasTextContent;
            
            if (hasContent)
            {
                maxRow = Math.Max(maxRow, row);
                Logger.Debug($"내용 있는 행 확인: 행={row}, 최대행={maxRow}");
            }
            
            return maxRow;
        }

        /// <summary>
        /// 중첩된 객체의 컬럼 수를 계산합니다.
        /// </summary>
        private int GetNestedColumnCount(ElementStructure structure)
        {
            int count = 0;
            
            // 속성 수
            count += structure.Attributes.Count;
            
            // 자식 요소 수
            count += structure.ChildElements.Count;
            
            // 텍스트 내용이 있으면 +1
            if (structure.HasTextContent)
            {
                count++;
            }
            
            return Math.Max(count, 1); // 최소 1개 컬럼
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
            
            // XML 태그 순서를 보장하기 위해 자식 요소들을 원래 순서대로 처리
            foreach (var child in structure.ChildElements)
            {
                if (child.Value.IsArray) continue; // 배열은 나중에 처리
                
                var childElement = element.Element(child.Key);
                
                // 단순 요소 (속성도 자식 요소도 없음)
                if (!child.Value.ChildElements.Any() && !child.Value.Attributes.Any())
                {
                    if (childElement != null)
                    {
                        worksheet.Cell(row, col).Value = childElement.Value;
                    }
                    col++;
                }
                // 복잡한 요소 (속성이나 자식 요소가 있음)
                else
                {
                    col = WriteNestedElementData(worksheet, childElement, child.Value, row, col);
                }
            }
            
            // 텍스트 내용
            if (structure.HasTextContent)
            {
                worksheet.Cell(row, col).Value = element.Value;
                col++;
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
        /// 중첩 요소 데이터를 재귀적으로 작성합니다.
        /// </summary>
        private int WriteNestedElementData(IXLWorksheet worksheet, XElement element, ElementStructure structure, int row, int startCol)
        {
            int col = startCol;
            
            if (element != null)
            {
                // 속성들 먼저 처리
                foreach (var attr in structure.Attributes)
                {
                    var attribute = element.Attribute(attr);
                    if (attribute != null)
                    {
                        worksheet.Cell(row, col).Value = attribute.Value;
                    }
                    col++;
                }
                
                // 단순 자식 요소들
                foreach (var nestedChild in structure.ChildElements.Where(c => !c.Value.IsArray && !c.Value.ChildElements.Any() && !c.Value.Attributes.Any()))
                {
                    var nestedElement = element.Element(nestedChild.Key);
                    if (nestedElement != null)
                    {
                        worksheet.Cell(row, col).Value = nestedElement.Value;
                    }
                    col++;
                }
                
                // 텍스트 내용이 있는 경우
                if (structure.HasTextContent)
                {
                    if (!string.IsNullOrWhiteSpace(element.Value) && !element.HasElements)
                    {
                        worksheet.Cell(row, col).Value = element.Value;
                    }
                    col++;
                }
                
                // 더 깊은 중첩 객체들 재귀 처리
                foreach (var nestedChild in structure.ChildElements.Where(c => !c.Value.IsArray && (c.Value.ChildElements.Any() || c.Value.Attributes.Any())))
                {
                    var nestedElement = element.Element(nestedChild.Key);
                    col = WriteNestedElementData(worksheet, nestedElement, nestedChild.Value, row, col);
                }
            }
            else
            {
                // 요소가 없는 경우 빈 컬럼들 건너뛰기
                col += GetNestedColumnCount(structure);
            }
            
            return col;
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

        /// <summary>
        /// 워크시트에서 실제 사용된 컬럼 수를 확인합니다.
        /// </summary>
        private int GetActualUsedColumns(IXLWorksheet worksheet, int lastSchemaRow)
        {
            int maxCol = 1;
            
            // 스키마 행들(2행부터 lastSchemaRow까지)을 검사
            for (int row = 2; row <= lastSchemaRow; row++)
            {
                var usedRange = worksheet.Row(row).LastCellUsed();
                if (usedRange != null)
                {
                    maxCol = Math.Max(maxCol, usedRange.Address.ColumnNumber);
                }
            }
            
            Logger.Debug($"실제 사용된 최대 컬럼: {maxCol}");
            return maxCol;
        }

        /// <summary>
        /// 스키마 행들의 병합을 실제 컬럼 수로 업데이트합니다.
        /// </summary>
        private void UpdateSchemaMerging(IXLWorksheet worksheet, XmlStructureInfo structure, int actualMaxColumns, int lastSchemaRow)
        {
            if (!structure.IsArray) return;

            try
            {
                // 기존 병합 제거 및 새로운 병합 적용
                // 2행: ${} 마커
                ClearAndRemerge(worksheet, 2, 1, actualMaxColumns);
                
                // 3행: Mission$[] 마커
                ClearAndRemerge(worksheet, 3, 2, actualMaxColumns);
                
                // 4행: ${} 마커
                ClearAndRemerge(worksheet, 4, 3, actualMaxColumns);
                
                Logger.Information($"스키마 병합 업데이트 완료: 실제 컬럼 수={actualMaxColumns}");
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "스키마 병합 업데이트 중 오류 발생");
            }
        }

        /// <summary>
        /// 기존 병합을 제거하고 새로운 범위로 병합합니다.
        /// </summary>
        private void ClearAndRemerge(IXLWorksheet worksheet, int row, int startCol, int endCol)
        {
            try
            {
                // 해당 행의 기존 병합 찾아서 제거
                var rangeToUnmerge = worksheet.MergedRanges
                    .FirstOrDefault(r => r.FirstRow().RowNumber() == row && r.FirstColumn().ColumnNumber() == startCol);
                
                if (rangeToUnmerge != null)
                {
                    rangeToUnmerge.Unmerge();
                    Logger.Debug($"기존 병합 제거: 행={row}, 시작열={startCol}");
                }
                
                // 새로운 병합 적용
                if (endCol > startCol)
                {
                    worksheet.Range(row, startCol, row, endCol).Merge();
                    Logger.Debug($"새 병합 적용: 행={row}, 시작열={startCol}, 끝열={endCol}");
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, $"병합 업데이트 실패: 행={row}, 시작열={startCol}, 끝열={endCol}");
            }
        }
    }
}