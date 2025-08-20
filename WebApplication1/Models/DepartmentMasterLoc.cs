using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

public class DepartmentMasterLoc
{
    public int Id { get; set; } // IDENTITY

    [Required]
    public int DepartmentId { get; set; }                 // FK -> DepartmentMaster

    [Required, MaxLength(10)]
    public string LangCode { get; set; } = default!;      // 예: "ko", "en", "ko-KR", "en-US"

    [Required, MaxLength(64)]
    public string Name { get; set; } = default!;          // 번역명

    [MaxLength(32)]
    public string? ShortName { get; set; }                // (선택) 약어

    [ForeignKey(nameof(DepartmentId))]
    public DepartmentMaster Department { get; set; } = default!;
}
