using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace WebApplication1.Models; // 프로젝트 네임스페이스에 맞추세요

public class PositionMaster
{
    public int Id { get; set; } // IDENTITY

    [Required, MaxLength(10)]
    public string CompCd { get; set; } = default!;        // 회사코드

    [Required, MaxLength(32)]
    public string Code { get; set; } = default!;          // 직책코드(회사 내 유니크)

    [Required, MaxLength(64)]
    public string Name { get; set; } = default!;          // 기본명(기본 언어)

    public bool IsActive { get; set; } = true;

    /// <summary>
    /// 직급 레벨(우선순위) — 값이 클수록 상위 직급으로 취급합니다.
    /// 정렬 시 RankLevel DESC, SortOrder ASC 권장
    /// </summary>
    [Range(0, short.MaxValue)]
    public short RankLevel { get; set; } = 0;

    /// <summary>리뷰/결재 권한 보유 여부</summary>
    public bool IsApprover { get; set; } = false;

    /// <summary>동일 레벨 내 표시 순서(오름차순)</summary>
    public int SortOrder { get; set; } = 0;

    // i18n
    public ICollection<PositionMasterLoc> Locs { get; set; } = new List<PositionMasterLoc>();
}
