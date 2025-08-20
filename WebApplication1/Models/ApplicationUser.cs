using Microsoft.AspNetCore.Identity;

namespace WebApplication1.Models;

public class ApplicationUser : IdentityUser
{
    public string? DisplayName { get; set; }
    public int? DepartmentId { get; set; }
    public DepartmentMaster? Department { get; set; }
    public int? PositionId { get; set; }
    public PositionMaster? Position { get; set; }
    public int? IsAdmin { get; set; } = 0; // 0: 일반, 1: 관리자, 2: 슈퍼관리자
}