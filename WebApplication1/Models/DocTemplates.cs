using System;

namespace WebApplication1.Models
{
    // dbo.DocTemplateMaster
    public class DocTemplateMaster
    {
        public int Id { get; set; }
        public string CompCd { get; set; } = default!;
        public int DepartmentId { get; set; }   // 공용이면 0
        public string? KindCode { get; set; }
        public string DocCode { get; set; } = default!;
        public string DocName { get; set; } = default!;
        public string? Title { get; set; }
        public int ApprovalCount { get; set; }
        public bool IsActive { get; set; }
        public DateTime CreatedAt { get; set; }
        public string? CreatedBy { get; set; }
        public DateTime? UpdatedAt { get; set; }
        public string? UpdatedBy { get; set; }
    }

    // dbo.DocTemplateVersion
    public class DocTemplateVersion
    {
        public long Id { get; set; }
        public int TemplateId { get; set; }
        public int VersionNo { get; set; }
        public string? DescriptorJson { get; set; }
        public string? PreviewJson { get; set; }
        public bool Templated { get; set; }     // 스키마에 있으면 사용, 없으면 false 디폴트
        public DateTime CreatedAt { get; set; }
        public string? CreatedBy { get; set; }
    }

    // dbo.DocTemplateFile
    public class DocTemplateFile
    {
        public long Id { get; set; }
        public long VersionId { get; set; }
        public string FileRole { get; set; } = default!;     // DescriptorJson / PreviewJson / ExcelFile
        public string Storage { get; set; } = default!;      // Db / Disk
        public string? FileName { get; set; }
        public string? FilePath { get; set; }
        public int?  FileSize { get; set; }                  // (있으면) KB 등
        public long? FileSizeBytes { get; set; }             // (있으면) 바이트
        public string? ContentType { get; set; }
        public string? Contents { get; set; }                // Storage = Db 일 때
        public DateTime CreatedAt { get; set; }
        public string? CreatedBy { get; set; }
        // Blob 등 바이너리 컬럼이 있어도 현재 컨트롤러에서는 string Contents 만 읽으므로 생략
    }

    // dbo.DocTemplateApproval (이번 화면에선 직접 조회 안 하지만 참조 대비)
    public class DocTemplateApproval
    {
        public long Id { get; set; }
        public long VersionId { get; set; }
        public int Slot { get; set; }
        public string Part { get; set; } = default!;
        public string? A1 { get; set; }
        public int? Row { get; set; }
        public int? Column { get; set; }
        public string? CellA1 { get; set; }
        public int? CellRow { get; set; }
        public int? CellColumn { get; set; }
    }
}
