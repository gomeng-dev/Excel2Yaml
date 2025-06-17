namespace ExcelToYamlAddin.Domain.Interfaces
{
    /// <summary>
    /// 도메인 서비스의 기본 인터페이스
    /// </summary>
    public interface IDomainService
    {
        /// <summary>
        /// 서비스 이름
        /// </summary>
        string ServiceName { get; }
    }
}