using Microsoft.AspNetCore.Identity;
using Microsoft.Extensions.Options;
using System.Security.Claims;
using WebApplication1.Data;
using WebApplication1.Models;
using Microsoft.EntityFrameworkCore;

public class CustomUserClaimsPrincipalFactory
    : UserClaimsPrincipalFactory<ApplicationUser, IdentityRole>
{
    private readonly ApplicationDbContext _db;
    public CustomUserClaimsPrincipalFactory(
        UserManager<ApplicationUser> userManager,
        RoleManager<IdentityRole> roleManager,
        IOptions<IdentityOptions> optionsAccessor,
        ApplicationDbContext db) : base(userManager, roleManager, optionsAccessor)
    {
        _db = db;
    }

    protected override async Task<ClaimsIdentity> GenerateClaimsAsync(ApplicationUser user)
    {
        var id = await base.GenerateClaimsAsync(user);

        var prof = await _db.UserProfiles.AsNoTracking()
                     .Where(p => p.UserId == user.Id)
                     .Select(p => new { p.IsAdmin })
                     .FirstOrDefaultAsync();

        // 2025.09.11 CS1501/CS8604 대응: nullable → int 안전값으로 보정
        int safeAdmin = (prof?.IsAdmin ?? user.IsAdmin ?? 0);
        id.AddClaim(new Claim("is_admin", safeAdmin.ToString()));

        return id;
    }
}
