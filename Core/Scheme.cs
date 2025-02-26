using Microsoft.Extensions.Logging;
using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;

namespace ExcelToJsonAddin.Core
{
    public class Scheme : IEnumerable<SchemeNode>
    {
        private static readonly ILogger<Scheme> Logger = CreateLogger();

        private static ILogger<Scheme> CreateLogger()
        {
            // 간단한 로거 팩토리 생성
            var loggerFactory = LoggerFactory.Create(builder => 
            {
                // AddConsole 대신 디버그 로깅만 사용
                builder.SetMinimumLevel(LogLevel.Debug);
            });
            
            return loggerFactory.CreateLogger<Scheme>();
        }

        private readonly SchemeNode root;
        private readonly IXLWorksheet sheet;
        private readonly int contentStartRowNum;
        private readonly int endRowNum;
        private readonly LinkedList<SchemeNode> linearNodes;

        public Scheme(IXLWorksheet sheet, SchemeNode root, int contentStartRowNum, int endRowNum)
        {
            this.sheet = sheet ?? throw new ArgumentNullException(nameof(sheet));
            this.root = root ?? throw new ArgumentNullException(nameof(root));
            this.contentStartRowNum = contentStartRowNum;
            this.endRowNum = endRowNum;
            
            this.linearNodes = root.Linear() ?? new LinkedList<SchemeNode>();
            
            Logger.LogInformation("Scheme 생성: 루트={RootName}, 시작행={StartRow}, 끝행={EndRow}, 노드 수={NodeCount}", 
                root.Key, contentStartRowNum, endRowNum, linearNodes.Count);
        }

        public Scheme(IXLWorksheet sheet)
        {
            if (sheet == null) throw new ArgumentNullException(nameof(sheet));
            
            this.sheet = sheet;
            
            // SchemeParser 생성 및 노드 파싱
            var parser = new SchemeParser(sheet);
            var parsed = parser.Parse();
            
            // 파싱된 스키마에서 값 가져오기
            this.root = parsed.Root;
            this.contentStartRowNum = parsed.ContentStartRowNum;
            this.endRowNum = parsed.EndRowNum;
            this.linearNodes = new LinkedList<SchemeNode>(parsed.GetLinearNodes());
            
            Logger.LogInformation("Scheme 생성(자동 파싱): 루트={RootName}, 시작행={StartRow}, 끝행={EndRow}, 노드 수={NodeCount}", 
                root.Key, contentStartRowNum, endRowNum, linearNodes.Count);
        }

        public SchemeNode Root => root;
        public IXLWorksheet Sheet => sheet;
        public int ContentStartRowNum => contentStartRowNum;
        public int EndRowNum => endRowNum;
        
        public IEnumerator<SchemeNode> GetEnumerator()
        {
            return linearNodes.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        /// <summary>
        /// 모든 스키마 노드를 선형 순서로 반환합니다.
        /// </summary>
        /// <returns>선형 순서로 정렬된 스키마 노드 목록</returns>
        public LinkedList<SchemeNode> GetLinearNodes()
        {
            return root.Linear();
        }

        public LinkedList<SchemeNode> ToList() => GetLinearNodes();
    }
}
