using Microsoft.AspNetCore.Identity;

namespace WebApplication1.Models;  // ← 프로젝트 기본 네임스페이스에 맞추세요

public class ApplicationUser : IdentityUser
{
    public UserProfile? Profile { get; set; }   // 1:1 내비게이션
}