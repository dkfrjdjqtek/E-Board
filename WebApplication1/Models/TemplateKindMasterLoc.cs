using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace WebApplication1.Models
{
    // PK = (Id, LangCode) / FK(Id) -> TemplateKindMaster(Id)
    public class TemplateKindMasterLoc
    {
        public int Id { get; set; }                             // FK → Masters.Id (ON DELETE CASCADE)
        [MaxLength(10)] public string CompCd { get; set; } = null!;
        public int DepartmentId { get; set; }                   // Masters.DepartmentId 그대로
        [MaxLength(10)] public string LangCode { get; set; } = null!;
        [MaxLength(64)] public string Name { get; set; } = null!;
        // 여기엔 절대 TemplateKindMasterId, MasterId 등의 속성/어트리뷰트를 두지 마세요.
        // RowVersion은 섀도우 속성(모델에 정의하지 않음)
    }
}