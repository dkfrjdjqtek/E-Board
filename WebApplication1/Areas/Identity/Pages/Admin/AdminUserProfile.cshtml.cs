// 2025.09.19 Changed: TempData와 AddModelError에 LocalizedString을 직접 넣던 부분을 .Value로 치환 직렬화 예외 방지
// 2025.09.19 Changed: 메일 제목과 본문 템플릿도 .Value로 명시 적용
// 2025.09.19 Changed: NormalizeModelStateMessages의 리소스 맵 값도 .Value로 통일

using System.Security.Claims;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Options;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Models.ViewModels;

namespace WebApplication1.Areas.Identity.Pages.Admin
{
    [Authorize(Policy = "AdminOnly")]
    public class AdminUserProfileModel : PageModel
    {
        private readonly ApplicationDbContext _db;
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly SignInManager<ApplicationUser> _signInManager;
        private readonly IEmailSender _email;
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IOptionsMonitor<DataProtectionTokenProviderOptions> _tokenOpts;

        public AdminUserProfileModel(
            ApplicationDbContext db,
            UserManager<ApplicationUser> userManager,
            SignInManager<ApplicationUser> signInManager,
            IEmailSender email,
            IStringLocalizer<SharedResource> S,
            IOptionsMonitor<DataProtectionTokenProviderOptions> tokenOpts)
        {
            _db = db;
            _userManager = userManager;
            _signInManager = signInManager;
            _email = email;
            _S = S;
            _tokenOpts = tokenOpts;
        }

        [BindProperty] public AdminUserProfileVM VM { get; set; } = new();

        // 2025.09.19 Added: ModelState 기본 영문 Required Range 등 메시지를 리소스 기반으로 교체
        private void NormalizeModelStateMessages()
        {
            static IEnumerable<string> Keys(string prop) => new[] { $"VM.{prop}", prop };

            var map = new Dictionary<string, string>
            {
                { "CompCd", _S["ACP_CompCd_Req"].Value },
                { "DisplayName", _S["A_Alert_Require_Name"].Value },
                { "newEmail", _S["ACP_NewEmail_Req"].Value }
            };

            foreach (var kv in map)
            {
                foreach (var key in Keys(kv.Key))
                {
                    if (!ModelState.TryGetValue(key, out var entry)) continue;

                    var isEmpty =
                        (kv.Key == "newEmail" && string.IsNullOrWhiteSpace(Request.Form["newEmail"])) ||
                        (kv.Key != "newEmail" && (VM?.GetType().GetProperty(kv.Key)?.GetValue(VM) is string s ? string.IsNullOrWhiteSpace(s) : VM?.GetType().GetProperty(kv.Key)?.GetValue(VM) is null));

                    if (!isEmpty || entry.Errors.Count == 0) continue;

                    var msg = kv.Value;
                    entry.Errors.Clear();
                    ModelState.Remove(key);
                    ModelState.AddModelError(key, msg);
                    break;
                }
            }
        }

        private async Task BindMastersAsync()
        {
            VM.CompList = await _db.CompMasters
                .Where(x => x.IsActive)
                .OrderBy(x => x.CompCd)
                .Select(x => new SelectListItem { Value = x.CompCd, Text = x.Name })
                .ToListAsync();

            VM.DeptList = await _db.DepartmentMasters
                .Where(x => x.IsActive)
                .OrderBy(x => x.Name)
                .Select(x => new SelectListItem { Value = x.Id.ToString(), Text = x.Name })
                .ToListAsync();

            VM.PosList = await _db.PositionMasters
                .Where(x => x.IsActive)
                .OrderBy(x => x.RankLevel)
                .Select(x => new SelectListItem { Value = x.Id.ToString(), Text = x.Name })
                .ToListAsync();

            var usersQ = _db.Users.AsNoTracking();
            if (!string.IsNullOrWhiteSpace(VM.Q))
            {
                var q = VM.Q.Trim();
                usersQ = usersQ.Where(u => u.UserName!.Contains(q) || u.Email!.Contains(q));
            }
            VM.Accounts = await usersQ
                .OrderBy(u => u.UserName)
                .Select(u => new SelectListItem { Value = u.Id, Text = $"{u.UserName} ({u.Email})", Selected = u.Id == VM.SelectedUserId })
                .ToListAsync();
        }

        private async Task LoadUserAsync(string id)
        {
            var user = await _userManager.FindByIdAsync(id);
            if (user is null) return;

            var profile = await _db.UserProfiles.AsNoTracking().FirstOrDefaultAsync(p => p.UserId == id);

            VM.Email = user.Email;
            VM.UserName = user.UserName;
            VM.DisplayName = profile?.DisplayName ?? string.Empty;
            VM.CompCd = profile?.CompCd ?? string.Empty;
            VM.DepartmentId = profile?.DepartmentId;
            VM.PositionId = profile?.PositionId;
            VM.PhoneNumber = user.PhoneNumber;

            var claims = await _userManager.GetClaimsAsync(user);
            var ac = claims.FirstOrDefault(c => c.Type == "is_admin")?.Value;
            VM.AdminLevel = ac == "2" ? 2 : ac == "1" ? 1 : Math.Clamp(profile?.IsAdmin ?? 0, 0, 2);
        }

        public async Task<IActionResult> OnGetAsync(string? id, string? q)
        {
            VM.SelectedUserId = id;
            VM.Q = q;

            await BindMastersAsync();

            if (!string.IsNullOrEmpty(id))
                await LoadUserAsync(id);

            return Page();
        }

        public async Task<IActionResult> OnPostSaveAsync()
        {
            await BindMastersAsync();

            if (string.IsNullOrEmpty(VM.SelectedUserId))
            {
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_SelectUser_Req"].Value);
            }

            NormalizeModelStateMessages();

            if (!ModelState.IsValid)
                return Page();

            var user = await _userManager.FindByIdAsync(VM.SelectedUserId!);
            var profile = await _db.UserProfiles.FirstOrDefaultAsync(p => p.UserId == VM.SelectedUserId);
            if (user is null || profile is null)
            {
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_UserOrProfile_NotFound"].Value);
                return Page();
            }

            if (!await _db.CompMasters.AnyAsync(x => x.IsActive && x.CompCd == VM.CompCd))
                ModelState.AddModelError(nameof(VM.CompCd), _S["ACP_Invalid_Comp"].Value);

            if (VM.DepartmentId.HasValue &&
                !await _db.DepartmentMasters.AnyAsync(x => x.IsActive && x.Id == VM.DepartmentId.Value))
                ModelState.AddModelError(nameof(VM.DepartmentId), _S["ACP_Invalid_Dept"].Value);

            if (VM.PositionId.HasValue &&
                !await _db.PositionMasters.AnyAsync(x => x.IsActive && x.Id == VM.PositionId.Value))
                ModelState.AddModelError(nameof(VM.PositionId), _S["ACP_Invalid_Pos"].Value);

            if (!ModelState.IsValid)
                return Page();

            profile.DisplayName = VM.DisplayName;
            profile.CompCd = VM.CompCd;
            profile.DepartmentId = VM.DepartmentId;
            profile.PositionId = VM.PositionId;
            profile.IsAdmin = Math.Clamp(VM.AdminLevel, 0, 2);

            user.PhoneNumber = VM.PhoneNumber;
            await _db.SaveChangesAsync();

            await UpdateIsAdminClaimAsync(user, VM.AdminLevel);

            await _userManager.UpdateSecurityStampAsync(user);
            if (User.FindFirstValue(ClaimTypes.NameIdentifier) == user.Id)
                await _signInManager.RefreshSignInAsync(user);

            TempData["StatusMessage"] = _S["_CM_Saved"].Value; // TempData 문자열만
            return RedirectToPage(new { id = VM.SelectedUserId, q = VM.Q });
        }

        // 2025.09.19 Changed: 임시 비밀번호 발행 메일로 전환 리소스 키 교체 및 본문 파라미터 유지
        public async Task<IActionResult> OnPostResetPasswordAsync()
        {
            if (string.IsNullOrEmpty(VM.SelectedUserId))
            {
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_SelectUser_Req"].Value);
                await BindMastersAsync();
                return Page();
            }

            var user = await _userManager.FindByIdAsync(VM.SelectedUserId);
            if (user is null)
            {
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_User_NotFound"].Value);
                await BindMastersAsync();
                return Page();
            }

            // 임시 비밀번호 생성 및 즉시 설정
            var tokenForTemp = await _userManager.GeneratePasswordResetTokenAsync(user);
            var tempPw = "Temp!" + Guid.NewGuid().ToString("N")[..8];

            var result = await _userManager.ResetPasswordAsync(user, tokenForTemp, tempPw);
            if (!result.Succeeded)
            {
                var msg = result.Errors.FirstOrDefault()?.Description ?? "Reset failed";
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_TempPw_Failed", msg].Value);
                await BindMastersAsync();
                await LoadUserAsync(VM.SelectedUserId);
                return Page();
            }

            // 보안 스탬프 갱신 및 즉시 로그인 차단
            await _userManager.UpdateSecurityStampAsync(user);
            await _userManager.SetLockoutEndDateAsync(user, DateTimeOffset.UtcNow);

            // 사용자에게 임시 비밀번호 통지 후 비밀번호 변경 링크 제공
            var resetToken = await _userManager.GeneratePasswordResetTokenAsync(user);
            var encoded = Microsoft.AspNetCore.WebUtilities.WebEncoders.Base64UrlEncode(System.Text.Encoding.UTF8.GetBytes(resetToken));
            var resetUrl = $"{Request.Scheme}://{Request.Host}/Identity/Account/ResetPassword?code={encoded}&email={Uri.EscapeDataString(user.Email ?? string.Empty)}";

            var minutes = (int)_tokenOpts.CurrentValue.TokenLifespan.TotalMinutes;

            var displayName = await _db.UserProfiles
                .Where(p => p.UserId == user.Id)
                .Select(p => p.DisplayName)
                .FirstOrDefaultAsync() ?? (user.UserName ?? user.Email ?? "User");

            // 2025.09.19 Changed: 임시 비밀번호 발행 전용 리소스 키 사용
            // {0}=표시명 {1}=임시비밀번호 {2}=비밀번호변경링크 {3}=유효시간분
            var subject = _S["ACP_TempPW_Subject"].Value;
            var template = _S["ACP_TempPW_BodyHtml"].Value;
            var body = string.Format(
                template,
                HtmlEncoder.Default.Encode(displayName),
                HtmlEncoder.Default.Encode(tempPw),
                HtmlEncoder.Default.Encode(resetUrl),
                minutes
            );

            if (!string.IsNullOrWhiteSpace(user.Email))
                await _email.SendEmailAsync(user.Email!, subject, body);

            TempData["StatusMessage"] = _S["ACP_TempPw_Sent"].Value;
            return RedirectToPage(new { id = VM.SelectedUserId, q = VM.Q });
        }


        public async Task<IActionResult> OnPostSendChangeEmailAsync(string newEmail)
        {
            if (string.IsNullOrEmpty(VM.SelectedUserId))
            {
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_SelectUser_Req"].Value);
                await BindMastersAsync();
                return Page();
            }

            var user = await _userManager.FindByIdAsync(VM.SelectedUserId);
            if (user is null)
            {
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_User_NotFound"].Value);
                await BindMastersAsync();
                return Page();
            }

            if (string.IsNullOrWhiteSpace(newEmail))
            {
                ModelState.AddModelError("newEmail", _S["ACP_NewEmail_Req"].Value);
                await BindMastersAsync();
                await LoadUserAsync(VM.SelectedUserId);
                return Page();
            }

            var exists = await _userManager.FindByEmailAsync(newEmail.Trim());
            if (exists is not null && exists.Id != user.Id)
            {
                ModelState.AddModelError("newEmail", _S["ACP_EmailChange_InUse"].Value);
                await BindMastersAsync();
                await LoadUserAsync(VM.SelectedUserId);
                return Page();
            }

            var code = await _userManager.GenerateChangeEmailTokenAsync(user, newEmail);
            var encoded = Microsoft.AspNetCore.WebUtilities.WebEncoders.Base64UrlEncode(System.Text.Encoding.UTF8.GetBytes(code));
            var url = Url.Page("/Account/ConfirmEmailChange", null,
                new { area = "Identity", userId = user.Id, email = newEmail, code = encoded }, Request.Scheme)!;

            var minutes = (int)_tokenOpts.CurrentValue.TokenLifespan.TotalMinutes;

            // 2025.09.19 Added: 본문 {0} 표시용 사용자 표시명 로드
            var displayName = await _db.UserProfiles
                .Where(p => p.UserId == user.Id)
                .Select(p => p.DisplayName)
                .FirstOrDefaultAsync() ?? (user.UserName ?? user.Email ?? "User");

            var subject = _S["ACP_EmailChange_Subject"].Value;
            var template = _S["ACP_EmailChange_BodyHtml"].Value;

            // 2025.09.19 Changed: 템플릿 파라미터 순서 반영 {0}=DisplayName {1}=NewEmail {2}=Url {3}=Minutes
            var body = string.Format(
                template,
                HtmlEncoder.Default.Encode(displayName),
                HtmlEncoder.Default.Encode(newEmail),
                HtmlEncoder.Default.Encode(url),
                minutes
            );

            await _email.SendEmailAsync(newEmail, subject, body);

            TempData["StatusMessage"] = _S["ACP_ChangeEmail_Sent"].Value;
            return RedirectToPage(new { id = VM.SelectedUserId, q = VM.Q });
        }

        private async Task UpdateIsAdminClaimAsync(ApplicationUser user, int level)
        {
            var claims = await _userManager.GetClaimsAsync(user);
            foreach (var c in claims.Where(c => c.Type == "is_admin"))
                await _userManager.RemoveClaimAsync(user, c);

            level = Math.Clamp(level, 0, 2);
            if (level > 0)
                await _userManager.AddClaimAsync(user, new Claim("is_admin", level.ToString()));
        }
    }
}
