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
        // 우선순위: UserProfiles.IsAdmin → 없으면 ApplicationUser.IsAdmin
        var prof = await _db.UserProfiles.AsNoTracking()
                     .Where(p => p.UserId == user.Id)
                     .Select(p => new { p.IsAdmin })
                     .FirstOrDefaultAsync();

        var isAdmin = prof?.IsAdmin ?? user.IsAdmin; // int (0/1/2)
        id.AddClaim(new Claim("is_admin", isAdmin.ToString()));  // "0"|"1"|"2"
        return id;
    }
}
