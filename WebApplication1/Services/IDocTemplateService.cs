using System.Threading.Tasks;

namespace WebApplication1.Services
{
    public interface IDocTemplateService
    {
        /// <summary>
        /// 템플릿코드 기준으로 최신버전 메타를 읽어
        /// - descriptorJson (Inputs/Approvals)
        /// - previewJson (엑셀 미리보기; 없으면 ClosedXML로 생성)
        /// - templateTitle (문서 제목)
        /// - versionId (저장 시 참조)
        /// - excelFilePath (실제 엑셀 경로; 저장 시 씀)
        /// 을 반환합니다.
        /// </summary>
        Task<(string descriptorJson, string previewJson, string templateTitle, long versionId, string? excelFilePath)>
            LoadMetaAsync(string templateCode);
    }
}
