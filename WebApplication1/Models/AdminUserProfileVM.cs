using Microsoft.AspNetCore.Mvc.Rendering;
using System.ComponentModel.DataAnnotations;

namespace WebApplication1.Models.ViewModels
{
    public class AdminUserProfileVM
    {
        // 계정 선택
        public string? SelectedUserId { get; set; }
        public List<SelectListItem> Accounts { get; set; } = new();

        // 읽기 전용
        public string? Email { get; set; }
        public string? UserName { get; set; }

        // 편집 필드
        [Required]
        public string DisplayName { get; set; } = string.Empty;

        [Required]
        public string CompCd { get; set; } = string.Empty;

        // 프로젝트 실제 스키마에 맞춤: int? (스크린샷 오류 원인)
        public int? DepartmentId { get; set; }
        public int? PositionId { get; set; }

        public string? PhoneNumber { get; set; }

        public bool IsCreate { get; set; }          // ← 추가: 생성 모드 여부
        // 권한 체크박스 (IsAdmin 0/1 매핑)
        //public bool IsAdminChecked { get; set; }
        public int AdminLevel { get; set; }  // ⬅️ 추가

        // 드롭다운
        public List<SelectListItem> CompList { get; set; } = new();
        public List<SelectListItem> DeptList { get; set; } = new();
        public List<SelectListItem> PosList { get; set; } = new();

        // 검색어(옵션)
        public string? Q { get; set; }
    }
}
