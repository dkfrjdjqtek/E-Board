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
        public int IsActive { get; set; }
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

        // 2026.06.11 Added: 템플릿 보호 재적용 및 실제 xlsx 기준 표시 메트릭
        public DateTime? PreparedAt { get; set; }
        public string? TemplateFileHash { get; set; }
        public string? ProtectionRuleCode { get; set; }
        public string? VisualMetricRuleCode { get; set; }
        public string? VisualSource { get; set; }
        public string? VisualRangeA1 { get; set; }
        public int? VisualWidthPx { get; set; }
        public int? VisualHeightPx { get; set; }
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
        public int? FileSize { get; set; }                  // (있으면) KB 등
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

    // 2026.06.16 Added: 전결 규칙 엔티티 추가 Contents 템플릿 버전별 전결권자 차수와 생략 대상 차수 및 조건 유형을 저장
    public class DocTemplateDelegationRule
    {
        public long Id { get; set; }
        public int TemplateId { get; set; }
        public long TemplateVersionId { get; set; }
        public string? RuleName { get; set; }
        public string ConditionType { get; set; } = default!;
        public int DelegationStepOrder { get; set; }
        public int SkipFromStepOrder { get; set; }
        public int SkipToStepOrder { get; set; }
        public int Priority { get; set; }
        public bool IsActive { get; set; }
        public string? Note { get; set; }
        public string CreatedBy { get; set; } = default!;
        public DateTime CreatedAt { get; set; }
        public string? UpdatedBy { get; set; }
        public DateTime? UpdatedAt { get; set; }
    }

    // 2026.06.16 Added: 전결 금액 조건 엔티티 추가 Contents 금액 조건 전결의 금액 필드와 통화 필드 및 통화별 기준 금액을 저장
    public class DocTemplateDelegationAmountRule
    {
        public long Id { get; set; }
        public long RuleId { get; set; }
        public string AmountFieldKey { get; set; } = default!;
        public string CurrencyFieldKey { get; set; } = default!;

        // 2026.06.18 Added: 전결 금액 조건 셀 직접 참조 추가 Contents 입력 필드가 아닌 수식 셀과 통화 셀을 참조하기 위한 셀 주소를 저장
        public string? AmountCellA1 { get; set; }
        public string? CurrencyCellA1 { get; set; }

        public string CurrencyCode { get; set; } = default!;
        public decimal LimitAmount { get; set; }
        public bool IsActive { get; set; }
        public string CreatedBy { get; set; } = default!;
        public DateTime CreatedAt { get; set; }
        public string? UpdatedBy { get; set; }
        public DateTime? UpdatedAt { get; set; }
    }

    // 2026.06.16 Added: 문서별 전결 적용 결과 엔티티 추가 Contents 문서 작성 및 승인 시점의 전결 후보와 적용 결과를 스냅샷으로 저장
    public class DocumentDelegationResult
    {
        public long Id { get; set; }
        public string DocId { get; set; } = default!;
        public long? RuleId { get; set; }
        public long? TemplateVersionId { get; set; }
        public string ConditionType { get; set; } = default!;
        public int DelegationStepOrder { get; set; }
        public int SkipFromStepOrder { get; set; }
        public int SkipToStepOrder { get; set; }
        public string? AmountFieldKey { get; set; }
        public decimal? AmountValue { get; set; }
        public string? CurrencyFieldKey { get; set; }

        // 2026.06.18 Added: 문서별 전결 금액 조건 셀 직접 참조 추가 Contents 실제 문서 파일에서 비교한 금액 셀과 통화 셀 주소를 저장
        public string? AmountCellA1 { get; set; }
        public string? CurrencyCellA1 { get; set; }

        public string? CurrencyCode { get; set; }
        public decimal? LimitAmount { get; set; }
        public string AppliedStatus { get; set; } = default!;
        public string? AppliedBy { get; set; }
        public DateTime? AppliedAt { get; set; }
        public string? CancelledBy { get; set; }
        public DateTime? CancelledAt { get; set; }
        public string? ResultMessageKey { get; set; }
        public string? DetailJson { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
    }
}