// Areas/Identity/Pages/Account/ChangeProfile.cshtml.cs
using System.ComponentModel.DataAnnotations;
using System.Text;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.WebUtilities;
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
    [BindProperty] public EmailChangeInput EmailChange { get; set; } = new();

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
        [Display(Name = "지사")] public string? CompCd { get; set; }

        // 비밀번호 변경(선택)
        [DataType(DataType.Password), Display(Name = "현재 비밀번호")]
        public string? CurrentPassword { get; set; }

        [DataType(DataType.Password), Display(Name = "새 비밀번호")]
        public string? NewPassword { get; set; }

        [DataType(DataType.Password), Display(Name = "새 비밀번호 확인")]
        [Compare("NewPassword", ErrorMessage = "새 비밀번호가 일치하지 않습니다.")]
        public string? ConfirmNewPassword { get; set; }
    }

    public class EmailChangeInput
    {
        [Required, EmailAddress, Display(Name = "새 이메일")]
        public string? NewEmail { get; set; }

        [Required, DataType(DataType.Password), Display(Name = "현재 비밀번호")]
        public string? CurrentPassword { get; set; }
    }

    public async Task<IActionResult> OnGetAsync()
    {
        var user = await _userManager.GetUserAsync(User);
        if (user is null) return RedirectToPage("/Account/Login");

        await LoadOptionsAsync();

        var profile = await _db.UserProfiles.AsNoTracking()
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
            ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.CompCd)}", "지사를 선택하세요.");
        else if (!await _db.CompMasters.AnyAsync(c => c.CompCd == Input.CompCd))
            ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.CompCd)}", "유효하지 않은 지사입니다.");

        if (Input.DepartmentId is null)
            ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.DepartmentId)}", "부서를 선택하세요.");
        else if (!await _db.DepartmentMasters.AnyAsync(d => d.Id == Input.DepartmentId))
            ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.DepartmentId)}", "유효하지 않은 부서입니다.");

        if (Input.PositionId is null)
            ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.PositionId)}", "직급을 선택하세요.");
        else if (!await _db.PositionMasters.AnyAsync(p => p.Id == Input.PositionId))
            ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.PositionId)}", "유효하지 않은 직급입니다.");

        if (string.IsNullOrWhiteSpace(Input.DisplayName))
            ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.DisplayName)}", "이름을 입력하세요.");

        // 2) 비밀번호 변경을 시도하는지 판단 + 검증
        var wantsPwChange =
            !string.IsNullOrWhiteSpace(Input.NewPassword) ||
            !string.IsNullOrWhiteSpace(Input.ConfirmNewPassword);

        if (wantsPwChange)
        {
            if (string.IsNullOrWhiteSpace(Input.NewPassword))
                ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.NewPassword)}", "새 비밀번호를 입력하세요.");

            if (string.IsNullOrWhiteSpace(Input.ConfirmNewPassword))
                ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.ConfirmNewPassword)}", "새 비밀번호 확인을 입력하세요.");

            if (!string.IsNullOrWhiteSpace(Input.NewPassword) &&
                !string.IsNullOrWhiteSpace(Input.ConfirmNewPassword) &&
                Input.NewPassword != Input.ConfirmNewPassword)
            {
                ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.ConfirmNewPassword)}", "새 비밀번호가 일치하지 않습니다.");
            }

            if (string.IsNullOrWhiteSpace(Input.CurrentPassword))
                ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.CurrentPassword)}", "현재 비밀번호를 입력하세요.");
        }

        if (!ModelState.IsValid) return Page();

        // 실제 비밀번호 변경 호출(요청 시에만)
        if (wantsPwChange)
        {
            var cr = await _userManager.ChangePasswordAsync(user, Input.CurrentPassword!, Input.NewPassword!);
            if (!cr.Succeeded)
            {
                foreach (var e in cr.Errors) ModelState.AddModelError("", e.Description);
                return Page();
            }
        }

        // 3) 프로필 확보(없으면 생성)
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

        // 4) 기본 정보 저장
        profile.DisplayName = Input.DisplayName?.Trim();
        profile.DepartmentId = Input.DepartmentId;
        profile.PositionId = Input.PositionId;

        // 5) 사용자 기본 필드(이메일/전화)
        if (!string.IsNullOrWhiteSpace(Input.Email) &&
            !string.Equals(user.Email, Input.Email, StringComparison.OrdinalIgnoreCase))
        {
            var er = await _userManager.SetEmailAsync(user, Input.Email);
            if (!er.Succeeded)
            {
                foreach (var e in er.Errors) ModelState.AddModelError("", e.Description);
                return Page();
            }
        }

        if (Input.PhoneNumber != user.PhoneNumber)
        {
            var pr = await _userManager.SetPhoneNumberAsync(user, Input.PhoneNumber);
            if (!pr.Succeeded)
            {
                foreach (var e in pr.Errors) ModelState.AddModelError("", e.Description);
                return Page();
            }
        }

        var ur = await _userManager.UpdateAsync(user);
        if (!ur.Succeeded)
        {
            foreach (var e in ur.Errors) ModelState.AddModelError("", e.Description);
            return Page();
        }

        await _db.SaveChangesAsync();
        await _signInManager.RefreshSignInAsync(user);

        TempData["StatusMessage"] = "프로필이 저장되었습니다.";
        return RedirectToPage();
    }

    // 이메일 변경 확인 메일 발송 (모달)
    public async Task<IActionResult> OnPostSendEmailChangeAsync()
    {
        var user = await _userManager.GetUserAsync(User);
        if (user is null) return RedirectToPage("/Account/Login");

        await LoadOptionsAsync();

        if (!ModelState.IsValid) return Page();

        // 비밀번호 확인
        var pwOk = await _userManager.CheckPasswordAsync(user, EmailChange.CurrentPassword!);
        if (!pwOk)
        {
            ModelState.AddModelError($"{nameof(EmailChange)}.{nameof(EmailChange.CurrentPassword)}", "현재 비밀번호가 올바르지 않습니다.");
            return Page();
        }

        // 기존 이메일과 동일한지
        if (string.Equals(user.Email, EmailChange.NewEmail, StringComparison.OrdinalIgnoreCase))
        {
            ModelState.AddModelError($"{nameof(EmailChange)}.{nameof(EmailChange.NewEmail)}", "기존 이메일과 동일합니다.");
            return Page();
        }

        // 토큰 생성 + 확인 링크
        var token = await _userManager.GenerateChangeEmailTokenAsync(user, EmailChange.NewEmail!);
        var code = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(token));

        var callbackUrl = Url.Page(
            "/Account/ConfirmEmailChange",
            pageHandler: null,
            values: new { userId = user.Id, email = EmailChange.NewEmail, code },
            protocol: Request.Scheme);

        // TODO: IEmailSender로 전송
        // await _emailSender.SendEmailAsync(EmailChange.NewEmail!, "이메일 변경 확인",
        //     $"아래 링크를 클릭하여 이메일 변경을 완료하세요: <a href='{HtmlEncoder.Default.Encode(callbackUrl!)}'>확인</a>");

        TempData["StatusMessage"] = "확인 메일을 보냈습니다. (개발환경: 링크는 서버 콘솔에서 확인)";
        Console.WriteLine($"[EMAIL-LINK] {callbackUrl}");

        return RedirectToPage();
    }

    private async Task LoadOptionsAsync()
    {
        var comps = await _db.CompMasters.Where(c => c.IsActive).OrderBy(c => c.CompCd).ToListAsync();
        var deps = await _db.DepartmentMasters.Where(d => d.IsActive).OrderBy(d => d.Name).ToListAsync();
        var poss = await _db.PositionMasters.Where(p => p.IsActive).OrderBy(p => p.RankLevel).ToListAsync();

        CompOptions = new SelectList(comps, "CompCd", "Name");
        DepartmentOptions = new SelectList(deps, "Id", "Name");
        PositionOptions = new SelectList(poss, "Id", "Name");
    }
}
