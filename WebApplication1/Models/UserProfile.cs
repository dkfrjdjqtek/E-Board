// 2025.09.23 Changed: RowVersion에 [Timestamp] 추가, UserId 고유 인덱스 부여(조회 일관성 보장)
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using Microsoft.EntityFrameworkCore;

namespace WebApplication1.Models   // ← ApplicationUser 와 동일 네임스페이스 유지
{
    // UserId로 조회/업데이트하는 패턴을 안전화하기 위해 유니크 인덱스 부여
    [Index(nameof(UserId), IsUnique = true)]
    public class UserProfile
    {
        // PK: Guid (기존 테이블이 Guid PK 구조라고 가정; 그렇지 않다면 마이그레이션 필요)
        [Key]
        public Guid Id { get; set; } = Guid.NewGuid();

        // AspNetUsers(Id)와의 FK(문자열 450 길이)
        [Required, MaxLength(450)]
        public string UserId { get; set; } = default!;

        [Required, MaxLength(10)]
        public string CompCd { get; set; } = default!;

        [MaxLength(64)]
        public string? DisplayName { get; set; }

        public int? DepartmentId { get; set; }
        public int? PositionId { get; set; }

        [ForeignKey(nameof(DepartmentId))] public DepartmentMaster? Department { get; set; }
        [ForeignKey(nameof(PositionId))] public PositionMaster? Position { get; set; }

        [ForeignKey(nameof(UserId))] public ApplicationUser User { get; set; } = default!;

        // 0: 일반, 1: 관리자, 2: 슈퍼관리자
        public int IsAdmin { get; set; } = 0;

        // 2025.09.23 Added: 결재권자 여부(기본값: false = 아니오)
        public bool IsApprover { get; set; } = false;

        // 동시성 제어: 실제 rowversion 컬럼과 매핑
        [Timestamp]
        public byte[]? RowVersion { get; set; }
    }
}
