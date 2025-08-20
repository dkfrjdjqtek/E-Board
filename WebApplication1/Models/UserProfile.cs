using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace WebApplication1.Models;  // ← ApplicationUser 와 동일해야 함

public class UserProfile
{
    [Key] public Guid Id { get; set; } = Guid.NewGuid();

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
}