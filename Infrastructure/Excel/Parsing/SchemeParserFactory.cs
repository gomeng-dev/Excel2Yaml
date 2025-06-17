using ClosedXML.Excel;
using ExcelToYamlAddin.Infrastructure.Logging;
using System;

namespace ExcelToYamlAddin.Infrastructure.Excel.Parsing
{
    /// <summary>
    /// SchemeParser 인스턴스를 생성하는 팩토리 클래스
    /// </summary>
    public static class SchemeParserFactory
    {
        /// <summary>
        /// 기본 설정으로 SchemeParser를 생성합니다.
        /// </summary>
        public static SchemeParser Create(IXLWorksheet worksheet)
        {
            return new SchemeParser(worksheet);
        }

        /// <summary>
        /// 의존성을 주입하여 SchemeParser를 생성합니다.
        /// </summary>
        public static SchemeParser CreateWithDependencies(
            IXLWorksheet worksheet,
            ISimpleLogger logger = null,
            ISchemeEndMarkerFinder endMarkerFinder = null,
            IMergedCellHandler mergedCellHandler = null,
            ISchemeNodeBuilder nodeBuilder = null)
        {
            // 기본 구현체 제공
            logger = logger ?? SimpleLoggerFactory.CreateLogger<SchemeParser>();
            endMarkerFinder = endMarkerFinder ?? new SchemeEndMarkerFinder(logger);
            mergedCellHandler = mergedCellHandler ?? new MergedCellHandler(logger);
            nodeBuilder = nodeBuilder ?? new SchemeNodeBuilder(logger);

            return new SchemeParser(worksheet, logger, endMarkerFinder, mergedCellHandler, nodeBuilder);
        }

        /// <summary>
        /// 테스트용 SchemeParser를 생성합니다.
        /// </summary>
        public static SchemeParser CreateForTesting(
            IXLWorksheet worksheet,
            ISimpleLogger logger,
            ISchemeEndMarkerFinder endMarkerFinder,
            IMergedCellHandler mergedCellHandler,
            ISchemeNodeBuilder nodeBuilder)
        {
            if (worksheet == null)
                throw new ArgumentNullException(nameof(worksheet));
            if (logger == null)
                throw new ArgumentNullException(nameof(logger));
            if (endMarkerFinder == null)
                throw new ArgumentNullException(nameof(endMarkerFinder));
            if (mergedCellHandler == null)
                throw new ArgumentNullException(nameof(mergedCellHandler));
            if (nodeBuilder == null)
                throw new ArgumentNullException(nameof(nodeBuilder));

            return new SchemeParser(worksheet, logger, endMarkerFinder, mergedCellHandler, nodeBuilder);
        }
    }
}