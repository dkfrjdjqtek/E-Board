using System.ComponentModel.DataAnnotations;

public class DepartmentMaster
{
    public int Id { get; set; } // IDENTITY

    [Required, MaxLength(10)]
    public string CompCd { get; set; } = default!;        // 회사코드

    [Required, MaxLength(32)]
    public string Code { get; set; } = default!;          // 부서코드(회사내 유니크)

    [Required, MaxLength(64)]
    public string Name { get; set; } = default!;          // 기본명(기본 언어)

    public bool IsActive { get; set; } = true;
    public int SortOrder { get; set; } = 0;

    // 다국어
    public ICollection<DepartmentMasterLoc> Locs { get; set; } = new List<DepartmentMasterLoc>();
}