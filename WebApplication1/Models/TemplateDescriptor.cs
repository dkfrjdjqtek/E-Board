using System.Collections.Generic;

namespace WebApplication1.Models
{
    // 결재권자 유형
    public enum ApproverType
    {
        Person = 0,   // 특정 사용자
        Role = 1,   // 역할 코드
        Rule = 2    // 규칙(JSON)
    }

    // A1 셀 위치 모델 (프로젝트에 이미 있다면 재사용하세요)
    public class CellMap
    {
        public string? A1 { get; set; }
    }

    // 입력 필드 매핑 항목 (프로젝트에 이미 있다면 재사용하세요)
    public class FieldMapItem
    {
        public string Key { get; set; } = "";
        public string Type { get; set; } = "Text"; // Text/Date/Num
        public CellMap? Cell { get; set; }
    }

    // ★ 결재란 매핑 항목 (이번 단계의 핵심)
    public class ApprovalMapItem
    {
        public int Slot { get; set; }                    // 결재 순번(열)
        public string Part { get; set; } = "";           // 항목명(결재1, 결재2…)
        public CellMap? Cell { get; set; }               // 사인/도장 위치(A1)

        // 결재권자 지정
        public ApproverType ApproverType { get; set; } = ApproverType.Person;

        // ApproverType == Person
        public string? PersonUserId { get; set; }

        // ApproverType == Role
        public string? RoleCode { get; set; }

        // ApproverType == Rule (When/Then JSON)
        public string? RuleJson { get; set; }
    }

    // 템플릿 디스크립터(전체)
    public class TemplateDescriptor
    {
        public List<FieldMapItem> Fields { get; set; } = new();
        public List<ApprovalMapItem> Approvals { get; set; } = new();
    }
}
