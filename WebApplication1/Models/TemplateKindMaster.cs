// File: Models/TemplateKindMaster.cs
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace WebApplication1.Models
{
    public class TemplateKindMaster
    {
        public int Id { get; set; }                             // IDENTITY (DB)
        [MaxLength(10)] public string CompCd { get; set; } = null!;
        public int DepartmentId { get; set; }                   // 기본값 0
        [MaxLength(32)] public string Code { get; set; } = null!;
        [MaxLength(64)] public string Name { get; set; } = null!;
        public bool IsActive { get; set; } = true;
        public int SortOrder { get; set; } = 0;
        // RowVersion은 섀도우 속성(모델에 정의하지 않음)
    }
}