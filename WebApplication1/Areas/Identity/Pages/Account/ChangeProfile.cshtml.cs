// Areas/Identity/Pages/Account/ChangeProfile.cshtml.cs
using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using WebApplication1.Data;
using WebApplication1.Models;

namespace WebApplication1.Areas.Identity.Pages.Account;

[Authorize]
public class ChangeProfileModel : PageModel
{
    private readonly UserManager<ApplicationUser> _userManager;
    private readonly SignInManager<ApplicationUser> _signInManager;
    private readonly ApplicationDbContext _db;

    public ChangeProfileModel(
        UserManager<ApplicationUser> userManager,
        SignInManager<ApplicationUser> signInManager,
        ApplicationDbContext db)
    {
        _userManager = userManager;
        _signInManager = signInManager;
        _db = db;
    }

    [BindProperty] public InputModel Input { get; set; } = new();
    public SelectList CompOptions { get; set; } = default!;
    public SelectList DepartmentOptions { get; set; } = default!;
    public SelectList PositionOptions { get; set; } = default!;

    public class InputModel
    {
        [EmailAddress, Display(Name = "이메일")]
        public string? Email { get; set; }

        [Display(Name = "사용자명")]
        public string? UserName { get; set; }   // 표시용(읽기전용)

        [Display(Name = "부서")] public int? DepartmentId { get; set; }
        [Display(Name = "직급")] public int? PositionId { get; set; }
        [Display(Name = "표시 이름")] public string? DisplayName { get; set; }
        [Phone, Display(Name = "전화번호")] public string? PhoneNumber { get; set; }
        [Display(Name = "지사")]
        public string? CompCd { get; set; }

        // 비밀번호 변경(선택)
        [DataType(DataType.Password), Display(Name = "현재 비밀번호")]
        public string? CurrentPassword { get; set; }

        [DataType(DataType.Password), Display(Name = "새 비밀번호")]
        public string? NewPassword { get; set; }

        [DataType(DataType.Password), Display(Name = "새 비밀번호 확인")]
        [Compare("NewPassword", ErrorMessage = "새 비밀번호가 일치하지 않습니다.")]
        public string? ConfirmNewPassword { get; set; }
    }

    public async Task<IActionResult> OnGetAsync()
    {
        var user = await _userManager.GetUserAsync(User);
        if (user is null) return RedirectToPage("/Account/Login");

        await LoadOptionsAsync();

        // 조회 전용(추적 X) - 프로필 없으면 빈 값으로 표시
        var profile = await _db.UserProfiles
            .AsNoTracking()
            .SingleOrDefaultAsync(p => p.UserId == user.Id);

        Input = new InputModel
        {
            Email = user.Email,
            UserName = user.UserName,
            PhoneNumber = user.PhoneNumber,
            DisplayName = profile?.DisplayName,
            DepartmentId = profile?.DepartmentId,
            PositionId = profile?.PositionId,
            CompCd = profile?.CompCd 
        };
        return Page();
    }

    public async Task<IActionResult> OnPostAsync()
    {
        var user = await _userManager.GetUserAsync(User);
        if (user is null) return RedirectToPage("/Account/Login");

        await LoadOptionsAsync();
        if (!ModelState.IsValid) return Page();

        // 1) 서버측 필수 + 유효성 검사
        if (string.IsNullOrWhiteSpace(Input.CompCd))
            ModelState.AddModelError("Input.CompCd", "지사를 선택하세요.");
        else if (!await _db.CompMasters.AnyAsync(c => c.CompCd == Input.CompCd))
            ModelState.AddModelError("Input.CompCd", "유효하지 않은 지사입니다.");

        if (!ModelState.IsValid) return Page();

        // ❶ 프로필 확보(없으면 생성) — POST에서만
        var profile = await _db.UserProfiles.SingleOrDefaultAsync(p => p.UserId == user.Id);
        if (profile is null)
        {
            profile = new UserProfile
            {
                UserId = user.Id,
                CompCd = Input.CompCd!.Trim().ToUpperInvariant()
            };
            _db.UserProfiles.Add(profile);
        }

        // ❷ 기본 정보 저장
        profile.DisplayName = Input.DisplayName?.Trim();
        profile.DepartmentId = Input.DepartmentId;
        profile.PositionId = Input.PositionId;

        // ❸ 사용자(User) 기본 필드(이메일/전화)
        if (!string.Equals(user.Email, Input.Email, StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(Input.Email))
        {
            var er = await _userManager.SetEmailAsync(user, Input.Email);
            if (!er.Succeeded) { foreach (var e in er.Errors) ModelState.AddModelError("", e.Description); return Page(); }
        }
        if (Input.PhoneNumber != user.PhoneNumber)
        {
            var pr = await _userManager.SetPhoneNumberAsync(user, Input.PhoneNumber);
            if (!pr.Succeeded) { foreach (var e in pr.Errors) ModelState.AddModelError("", e.Description); return Page(); }
        }

        var ur = await _userManager.UpdateAsync(user);
        if (!ur.Succeeded) { foreach (var e in ur.Errors) ModelState.AddModelError("", e.Description); return Page(); }

        // ❹ 비밀번호 변경(입력 시에만)
        if (!string.IsNullOrWhiteSpace(Input.NewPassword))
        {
            if (string.IsNullOrWhiteSpace(Input.CurrentPassword))
            {
                ModelState.AddModelError("Input.CurrentPassword", "현재 비밀번호를 입력하세요.");
                return Page();
            }
            var cr = await _userManager.ChangePasswordAsync(user, Input.CurrentPassword, Input.NewPassword);
            if (!cr.Succeeded) { foreach (var e in cr.Errors) ModelState.AddModelError("", e.Description); return Page(); }
        }

        await _db.SaveChangesAsync();
        await _signInManager.RefreshSignInAsync(user);

        TempData["StatusMessage"] = "프로필이 저장되었습니다.";
        return RedirectToPage();
    }

    private async Task LoadOptionsAsync()
    {
        var comps = await _db.CompMasters.Where(c => c.IsActive).OrderBy(c => c.CompCd).ToListAsync();
        var deps = await _db.DepartmentMasters.Where(d => d.IsActive).OrderBy(d => d.Name).ToListAsync();
        var poss = await _db.PositionMasters.Where(p => p.IsActive).OrderBy(p => p.RankLevel).ToListAsync();
        CompOptions = new SelectList(comps, "CompCd", "Name");      // ⬅️ 추가
        DepartmentOptions = new SelectList(deps, "Id", "Name");
        PositionOptions = new SelectList(poss, "Id", "Name");

    }
}
