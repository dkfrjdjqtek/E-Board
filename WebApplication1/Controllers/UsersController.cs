// Controllers/UsersController.cs
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WebApplication1.Data;
using WebApplication1.Models;

[Route("admin/users")]
public class UsersController : Controller
{
    private readonly ApplicationDbContext _db;
    public UsersController(ApplicationDbContext db) => _db = db;

    // 목록
    [HttpGet("")]
    public async Task<IActionResult> Index(string comp = "0001")
    {
        var list = await _db.UserProfiles
            .Include(p => p.User)
            .Include(p => p.Department).ThenInclude(d => d.Locs)
            .Include(p => p.Position).ThenInclude(p => p.Locs)
            .Where(p => p.CompCd == comp)
            .Select(p => new UserRowVM
            {
                UserId = p.UserId,
                Email = p.User.Email!,
                DisplayName = p.DisplayName,
                Department = p.Department!.Locs.Where(l => l.LangCode == "ko").Select(l => l.Name).FirstOrDefault() ?? p.Department!.Name,
                Position = p.Position!.Locs.Where(l => l.LangCode == "ko").Select(l => l.Name).FirstOrDefault() ?? p.Position!.Name
            })
            .ToListAsync();

        return View(list);
    }

    // 편집 폼
    [HttpGet("edit/{userId}")]
    public async Task<IActionResult> Edit(string userId, string comp = "0001")
    {
        var p = await _db.UserProfiles.Include(x => x.User).FirstAsync(x => x.UserId == userId);
        var vm = new UserEditVM
        {
            UserId = p.UserId,
            Email = p.User!.Email!,
            DisplayName = p.DisplayName,
            CompCd = p.CompCd,
            DepartmentId = p.DepartmentId,
            PositionId = p.PositionId,
            Departments = await _db.DepartmentMasters.Where(d => d.CompCd == comp && d.IsActive)
                          .OrderBy(d => d.SortOrder).Select(d => new KeyValuePair<int, string>(d.Id, d.Name)).ToListAsync(),
            Positions = await _db.PositionMasters.Where(t => t.CompCd == comp && t.IsActive)
                          .OrderByDescending(t => t.RankLevel).ThenBy(t => t.SortOrder)
                          .Select(t => new KeyValuePair<int, string>(t.Id, t.Name)).ToListAsync()
        };
        return View(vm);
    }

    [HttpPost("edit/{userId}")]
    public async Task<IActionResult> Edit(string userId, UserEditVM vm)
    {
        if (!ModelState.IsValid) return View(vm);
        var p = await _db.UserProfiles.FirstAsync(x => x.UserId == userId);
        p.DisplayName = vm.DisplayName;
        p.CompCd = vm.CompCd;
        p.DepartmentId = vm.DepartmentId;
        p.PositionId = vm.PositionId;
        await _db.SaveChangesAsync();
        return RedirectToAction(nameof(Index));
    }
}

public record UserRowVM
{
    public string UserId { get; init; } = default!;
    public string Email { get; init; } = default!;
    public string? DisplayName { get; init; }
    public string Department { get; init; } = "";
    public string Position { get; init; } = "";
}

public class UserEditVM
{
    public string UserId { get; set; } = default!;
    public string Email { get; set; } = default!;
    public string? DisplayName { get; set; }
    public string CompCd { get; set; } = "0001";
    public int? DepartmentId { get; set; }
    public int? PositionId { get; set; }

    public List<KeyValuePair<int, string>> Departments { get; set; } = new();
    public List<KeyValuePair<int, string>> Positions { get; set; } = new();
}
