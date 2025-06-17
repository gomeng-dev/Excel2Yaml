using System;

namespace ExcelToYamlAddin.Domain.Constants
{
    /// <summary>
    /// 에러 메시지 상수를 정의하는 클래스
    /// </summary>
    public static class ErrorMessages
    {
        /// <summary>
        /// 스키마 관련 에러 메시지
        /// </summary>
        public static class Schema
        {
            /// <summary>
            /// 스키마 종료 마커를 찾을 수 없을 때
            /// </summary>
            public const string SchemeEndNotFound = "Scheme end marker not found.";

            /// <summary>
            /// 스키마 시작 행을 찾을 수 없을 때
            /// </summary>
            public const string SchemeStartRowNotFound = "Scheme start row (2) not found.";

            /// <summary>
            /// 알 수 없는 노드 유형 에러
            /// </summary>
            public const string UnknownNodeType = "알 수 없는 노드 유형: ";

            /// <summary>
            /// 빈 키 값 에러
            /// </summary>
            public const string EmptyKeyError = "오류: JSON/YAML 표준에서는 객체 내에 이름 없는 속성을 가질 수 없습니다. 키 값이 비어있습니다.";

            /// <summary>
            /// JSON/YAML 표준 에러 제목
            /// </summary>
            public const string JsonYamlStandardError = "JSON/YAML 표준 오류";

            /// <summary>
            /// 이름 없는 값 에러
            /// </summary>
            public const string UnnamedValueError = "오류: JSON/YAML 표준에서는 객체 내에 이름 없는 값을 가질 수 없습니다. 노드 타입: ";

            /// <summary>
            /// 중첩 배열 경고
            /// </summary>
            public const string NestedArrayWarning = "배열 안에 직접 배열을 추가하는 것은 일부 파서에서 문제가 될 수 있습니다. 가능하면 이름 있는 객체로 감싸는 것이 좋습니다.";

            /// <summary>
            /// 스키마 이름이 비어있음
            /// </summary>
            public const string SchemeNameIsEmpty = "스키마 이름이 비어있습니다.";

            /// <summary>
            /// 루트 노드가 null임
            /// </summary>
            public const string RootNodeIsNull = "루트 노드가 null입니다.";

            /// <summary>
            /// 컨텐츠 시작 행이 잘못됨
            /// </summary>
            public const string InvalidContentStartRow = "컨텐츠 시작 행은 1 이상이어야 합니다.";

            /// <summary>
            /// 종료 행이 시작 행보다 작음
            /// </summary>
            public const string EndRowLessThanStartRow = "종료 행은 시작 행보다 커야 합니다.";

            /// <summary>
            /// 스키마가 유효하지 않음
            /// </summary>
            public const string SchemaIsInvalid = "스키마가 유효하지 않습니다.";

            /// <summary>
            /// 스키마에 노드가 없음
            /// </summary>
            public const string NoNodesInSchema = "스키마에 노드가 없습니다.";

            /// <summary>
            /// 순환 참조 발견
            /// </summary>
            public const string CircularReferenceFound = "스키마에 순환 참조가 있습니다.";

            /// <summary>
            /// 중복된 키 발견
            /// </summary>
            public const string DuplicateKeysFound = "중복된 키가 있습니다: {0}";
        }

        /// <summary>
        /// 변환 관련 에러 메시지
        /// </summary>
        public static class Conversion
        {
            /// <summary>
            /// 사용자 취소 메시지
            /// </summary>
            public const string UserCancelled = "사용자 요청에 의해 변환 작업이 중단되었습니다.";

            /// <summary>
            /// Excel 변환 중 오류
            /// </summary>
            public const string ExcelConversionError = "Excel 변환 중 오류: ";

            /// <summary>
            /// 시트 분석 중 오류
            /// </summary>
            public const string SheetAnalysisError = "시트 분석 중 오류 발생";

            /// <summary>
            /// 시트 분석 중 오류 (시트명 포함)
            /// </summary>
            public const string SheetAnalysisErrorWithName = "시트 분석 중 오류 발생: {0}";

            /// <summary>
            /// XML to YAML 변환 오류
            /// </summary>
            public const string XmlToYamlConversionError = "XML to YAML 변환 중 오류 발생: ";

            /// <summary>
            /// XML to Dictionary 변환 오류
            /// </summary>
            public const string XmlToDictionaryConversionError = "XML to Dictionary 변환 중 오류 발생: ";

            /// <summary>
            /// XML to Excel 변환 오류
            /// </summary>
            public const string XmlToExcelConversionError = "XML to Excel 변환 중 오류 발생: ";

            /// <summary>
            /// YAML to XML 변환 오류
            /// </summary>
            public const string YamlToXmlConversionError = "YAML to XML 변환 중 오류 발생";

            /// <summary>
            /// 병합 처리 중 오류
            /// </summary>
            public const string MergeProcessingError = "병합 처리 중 오류 발생";

            /// <summary>
            /// 플로우 스타일 처리 중 오류
            /// </summary>
            public const string FlowStyleProcessingError = "플로우 스타일 처리 중 오류 발생";
        }

        /// <summary>
        /// 파일 관련 에러 메시지
        /// </summary>
        public static class File
        {
            /// <summary>
            /// 엑셀 파일을 찾을 수 없음
            /// </summary>
            public const string ExcelFileNotFound = "엑셀 파일을 찾을 수 없습니다.";

            /// <summary>
            /// 시트를 찾을 수 없음
            /// </summary>
            public const string SheetNotFound = "'{0}' 시트를 찾을 수 없습니다.";

            /// <summary>
            /// 시트가 중복됨
            /// </summary>
            public const string DuplicateSheet = "'{0}' 시트가 중복되었습니다!";

            /// <summary>
            /// 파일 변환 중 오류
            /// </summary>
            public const string FileConversionError = "파일 변환 중 오류 발생: ";

            /// <summary>
            /// 임시 파일 저장 중 오류
            /// </summary>
            public const string TempFileSaveError = "임시 파일 저장 중 오류 발생: ";

            /// <summary>
            /// HTML 내보내기 오류
            /// </summary>
            public const string HtmlExportError = "Excel을 HTML로 내보내는 중 오류 발생";
        }

        /// <summary>
        /// 설정 관련 에러 메시지
        /// </summary>
        public static class Configuration
        {
            /// <summary>
            /// 설정 시트를 찾을 수 없음
            /// </summary>
            public const string ConfigSheetNotFound = "설정 시트를 찾을 수 없습니다: ";

            /// <summary>
            /// 설정 로드 실패
            /// </summary>
            public const string ConfigLoadFailed = "설정을 로드하는 데 실패했습니다.";

            /// <summary>
            /// 설정 저장 실패
            /// </summary>
            public const string ConfigSaveFailed = "설정을 저장하는 데 실패했습니다.";

            /// <summary>
            /// 잘못된 설정 값
            /// </summary>
            public const string InvalidConfigValue = "잘못된 설정 값입니다: ";

            /// <summary>
            /// 설정 업데이트 중 오류
            /// </summary>
            public const string ConfigUpdateError = "설정 업데이트 중 오류 발생";
        }

        /// <summary>
        /// 검증 관련 에러 메시지
        /// </summary>
        public static class Validation
        {
            /// <summary>
            /// 워크북이 null임
            /// </summary>
            public const string WorkbookIsNull = "워크북이 null입니다.";

            /// <summary>
            /// 시트가 null임
            /// </summary>
            public const string SheetIsNull = "시트가 null입니다.";

            /// <summary>
            /// 시트에 데이터가 없음
            /// </summary>
            public const string NoDataInSheet = "시트에 데이터가 없습니다.";

            /// <summary>
            /// XML 루트 요소를 찾을 수 없음
            /// </summary>
            public const string XmlRootNotFound = "XML 루트 요소를 찾을 수 없습니다.";

            /// <summary>
            /// YAML 변환 결과가 비어있음
            /// </summary>
            public const string EmptyYamlResult = "XML을 YAML로 변환한 결과가 비어있습니다.";

            /// <summary>
            /// 잘못된 시트 이름
            /// </summary>
            public const string InvalidSheetName = "잘못된 시트 이름입니다: ";

            /// <summary>
            /// 잘못된 파일 경로
            /// </summary>
            public const string InvalidFilePath = "잘못된 파일 경로입니다: ";

            /// <summary>
            /// 행 번호가 1보다 작음
            /// </summary>
            public const string RowLessThanOne = "행 번호는 1 이상이어야 합니다.";

            /// <summary>
            /// 열 번호가 1보다 작음
            /// </summary>
            public const string ColumnLessThanOne = "열 번호는 1 이상이어야 합니다.";

            /// <summary>
            /// 셀 주소가 비어있음
            /// </summary>
            public const string CellAddressIsEmpty = "셀 주소가 비어있습니다.";

            /// <summary>
            /// 잘못된 셀 주소 형식
            /// </summary>
            public const string InvalidCellAddressFormat = "잘못된 셀 주소 형식입니다: {0}";

            /// <summary>
            /// 노드 타입 값이 비어있음
            /// </summary>
            public const string NodeTypeValueIsEmpty = "노드 타입 값이 비어있습니다.";

            /// <summary>
            /// 노드 타입이 없음
            /// </summary>
            public const string NodeTypeIsNull = "노드 타입이 없습니다.";

            /// <summary>
            /// 노드 위치가 없음
            /// </summary>
            public const string NodePositionIsNull = "노드 위치가 없습니다.";

            /// <summary>
            /// 컨테이너 노드에 자식이 없음
            /// </summary>
            public const string ContainerNodeHasNoChildren = "컨테이너 노드에 자식이 없습니다: {0}";

            /// <summary>
            /// VALUE 노드에 자식 추가 불가
            /// </summary>
            public const string CannotAddChildToValueNode = "VALUE 노드에는 자식을 추가할 수 없습니다: {0}";

            /// <summary>
            /// IGNORE 노드에 자식 추가 불가
            /// </summary>
            public const string CannotAddChildToIgnoreNode = "IGNORE 노드에는 자식을 추가할 수 없습니다: {0}";

            /// <summary>
            /// 잘못된 자식 노드 타입
            /// </summary>
            public const string InvalidChildNodeType = "{0} 노드에는 {1} 노드를 추가할 수 없습니다: {2} -> {3}";
        }

        /// <summary>
        /// 일반 에러 메시지
        /// </summary>
        public static class General
        {
            /// <summary>
            /// 예기치 않은 오류
            /// </summary>
            public const string UnexpectedError = "예기치 않은 오류가 발생했습니다.";

            /// <summary>
            /// 작업 실패
            /// </summary>
            public const string OperationFailed = "작업이 실패했습니다.";

            /// <summary>
            /// 초기화 실패
            /// </summary>
            public const string InitializationFailed = "초기화에 실패했습니다.";

            /// <summary>
            /// 정리 중 오류
            /// </summary>
            public const string CleanupError = "정리 중 오류가 발생했습니다.";
        }

        /// <summary>
        /// 로깅 관련 메시지
        /// </summary>
        public static class Logging
        {
            /// <summary>
            /// 변환 시작
            /// </summary>
            public const string ConversionStarted = "변환 작업을 시작합니다.";

            /// <summary>
            /// 변환 완료
            /// </summary>
            public const string ConversionCompleted = "변환 작업이 완료되었습니다.";

            /// <summary>
            /// 시트 처리 중
            /// </summary>
            public const string ProcessingSheet = "시트를 처리 중입니다: ";

            /// <summary>
            /// 파일 생성 완료
            /// </summary>
            public const string FileCreated = "파일이 생성되었습니다: ";

            /// <summary>
            /// 설정 로드 중
            /// </summary>
            public const string LoadingConfiguration = "설정을 로드하는 중입니다.";

            /// <summary>
            /// 설정 저장 중
            /// </summary>
            public const string SavingConfiguration = "설정을 저장하는 중입니다.";
        }
    }
}