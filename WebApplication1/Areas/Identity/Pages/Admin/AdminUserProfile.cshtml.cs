// 2025.09.23 Changed: IsApprover 필드 추가 로드 저장 매핑 적용 기본값 아니오 유지
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

namespace WebApplication1.Areas.Identity.Pages.Admin
{
    // 이 페이지 전용 내장 ViewModel
    public class AdminUserProfileVM
    {
        public string? SelectedUserId { get; set; }
        public string? Q { get; set; }

        public string? CompCd { get; set; }
        public string? DisplayName { get; set; }
        public int? DepartmentId { get; set; }
        public int? PositionId { get; set; }
        public string? PhoneNumber { get; set; }
        public int AdminLevel { get; set; } = 0;

        public string? UserName { get; set; }
        public string? Email { get; set; }
        public bool IsActive { get; set; } = true;

        // 2025.09.23 Added: 결재권자 여부 토글 기본값 아니오
        public bool IsApprover { get; set; } = false;

        public List<SelectListItem> CompList { get; set; } = new();
        public List<SelectListItem> DeptList { get; set; } = new();
        public List<SelectListItem> PosList { get; set; } = new();
        public List<SelectListItem> Accounts { get; set; } = new();
    }

    [Authorize(Policy = "AdminOnly")]
    public class AdminUserProfileModel : PageModel
    {
        public const string NewUserId = "__new";

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

        private void NormalizeModelStateMessages()
        {
            static IEnumerable<string> Keys(string prop) => new[] { $"VM.{prop}", prop };

            var map = new Dictionary<string, string>
            {
                { "CompCd", _S["ACP_CompCd_Req"].Value },
                { "DisplayName", _S["_Alert_Require_Name"].Value },
                { "newEmail", _S["_Alert_Require_NewMail"].Value }
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

            // 활성 비활성 상태 계산
            VM.IsActive = IsActiveFromLockout(user);

            // 2025.09.23 Added: 결재권자 여부 로드
            VM.IsApprover = profile?.IsApprover ?? false;

            var claims = await _userManager.GetClaimsAsync(user);
            var ac = claims.FirstOrDefault(c => c.Type == "is_admin")?.Value;
            VM.AdminLevel = ac == "2" ? 2 : ac == "1" ? 1 : Math.Clamp(profile?.IsAdmin ?? 0, 0, 2);
        }

        public async Task<IActionResult> OnGetAsync(string? id, string? q)
        {
            VM.SelectedUserId = id;
            VM.Q = q;

            await BindMastersAsync();

            if (string.IsNullOrEmpty(id))
                return Page();

            if (string.Equals(id, NewUserId, StringComparison.Ordinal))
            {
                VM.IsActive = true;
                // 2025.09.23 Added: 신규 기본값 아니오
                VM.IsApprover = false;
                return Page();
            }

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

            if (string.IsNullOrWhiteSpace(VM.CompCd))
                ModelState.AddModelError(nameof(VM.CompCd), _S["ACP_CompCd_Req"].Value);

            if (string.IsNullOrWhiteSpace(VM.DisplayName))
                ModelState.AddModelError(nameof(VM.DisplayName), _S["_Alert_Require_Name"].Value);

            var isCreate = string.Equals(VM.SelectedUserId, NewUserId, StringComparison.Ordinal);
            if (isCreate)
            {
                if (string.IsNullOrWhiteSpace(VM.UserName))
                    ModelState.AddModelError(nameof(VM.UserName), _S["_Alert_Require_ID"].Value);

                if (string.IsNullOrWhiteSpace(VM.Email))
                    ModelState.AddModelError(nameof(VM.Email), _S["_Alert_Require_NewMail"].Value);
            }

            if (!ModelState.IsValid)
                return Page();

            if (isCreate)
            {
                // 신규 사용자 생성
                if (await _userManager.FindByNameAsync(VM.UserName!) != null)
                {
                    ModelState.AddModelError(nameof(VM.UserName), _S["ACP_UserName_Dup"].Value);
                    return Page();
                }
                if (!string.IsNullOrWhiteSpace(VM.Email) && await _userManager.FindByEmailAsync(VM.Email!) != null)
                {
                    ModelState.AddModelError(nameof(VM.Email), _S["ACP_Email_Dup"].Value);
                    return Page();
                }

                var user = new ApplicationUser
                {
                    UserName = VM.UserName!,
                    Email = VM.Email!,
                    PhoneNumber = VM.PhoneNumber ?? string.Empty
                };

                var tempPw = "Temp!" + Guid.NewGuid().ToString("N")[..8];
                var createResult = await _userManager.CreateAsync(user, tempPw);
                if (!createResult.Succeeded)
                {
                    var msg = createResult.Errors.FirstOrDefault()?.Description ?? "Create failed";
                    ModelState.AddModelError(nameof(VM.UserName), msg);
                    return Page();
                }

                var profile = new UserProfile
                {
                    UserId = user.Id,
                    DisplayName = VM.DisplayName ?? string.Empty,
                    CompCd = VM.CompCd ?? string.Empty,
                    DepartmentId = VM.DepartmentId,
                    PositionId = VM.PositionId,
                    IsAdmin = Math.Clamp(VM.AdminLevel, 0, 2),
                    // 2025.09.23 Added: 결재권자 여부 저장
                    IsApprover = VM.IsApprover
                };
                _db.UserProfiles.Add(profile);
                await _db.SaveChangesAsync();

                await UpdateIsAdminClaimAsync(user, VM.AdminLevel);

                // 활성 비활성 반영
                if (VM.IsActive)
                    await _userManager.SetLockoutEndDateAsync(user, null);
                else
                    await _userManager.SetLockoutEndDateAsync(user, DateTimeOffset.UtcNow.AddYears(100));

                await _userManager.UpdateSecurityStampAsync(user);

                TempData["StatusMessage"] = _S["_CM_Saved"].Value;

                return RedirectToPage(new { id = user.Id, q = VM.Q });
            }
            else
            {
                // 기존 사용자 수정
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

                profile.DisplayName = VM.DisplayName ?? string.Empty;
                profile.CompCd = VM.CompCd ?? string.Empty;
                profile.DepartmentId = VM.DepartmentId;
                profile.PositionId = VM.PositionId;
                profile.IsAdmin = Math.Clamp(VM.AdminLevel, 0, 2);
                // 2025.09.23 Changed: 결재권자 여부 갱신
                profile.IsApprover = VM.IsApprover;

                user.PhoneNumber = VM.PhoneNumber ?? string.Empty;

                await _db.SaveChangesAsync();

                await UpdateIsAdminClaimAsync(user, VM.AdminLevel);

                // 활성 비활성 반영
                if (VM.IsActive)
                    await _userManager.SetLockoutEndDateAsync(user, null);
                else
                    await _userManager.SetLockoutEndDateAsync(user, DateTimeOffset.UtcNow.AddYears(100));

                await _userManager.UpdateSecurityStampAsync(user);

                if (User.FindFirstValue(ClaimTypes.NameIdentifier) == user.Id)
                    await _signInManager.RefreshSignInAsync(user);

                TempData["StatusMessage"] = _S["_CM_Saved"].Value;
                return RedirectToPage(new { id = VM.SelectedUserId, q = VM.Q });
            }
        }

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

            await _userManager.UpdateSecurityStampAsync(user);
            await _userManager.SetLockoutEndDateAsync(user, DateTimeOffset.UtcNow);

            var resetToken = await _userManager.GeneratePasswordResetTokenAsync(user);
            var encoded = Microsoft.AspNetCore.WebUtilities.WebEncoders.Base64UrlEncode(System.Text.Encoding.UTF8.GetBytes(resetToken));
            var resetUrl = Url.Page("/Account/ResetPassword", null,
                new { area = "Identity", userId = user.Id, email = user.Email, code = encoded }, Request.Scheme)!;

            var minutes = (int)_tokenOpts.CurrentValue.TokenLifespan.TotalMinutes;

            var displayName = await _db.UserProfiles
                .Where(p => p.UserId == user.Id)
                .Select(p => p.DisplayName)
                .FirstOrDefaultAsync() ?? (user.UserName ?? user.Email ?? "User");

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
                ModelState.AddModelError("newEmail", _S["_Alert_Require_NewMail"].Value);
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

            var displayName = await _db.UserProfiles
                .Where(p => p.UserId == user.Id)
                .Select(p => p.DisplayName)
                .FirstOrDefaultAsync() ?? (user.UserName ?? user.Email ?? "User");

            var subject = _S["ACP_EmailChange_Subject"].Value;
            var template = _S["ACP_EmailChange_BodyHtml"].Value;

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

        private static bool IsActiveFromLockout(ApplicationUser user)
        {
            // LockoutEnd가 없거나 과거면 활성 true 미래면 비활성 false
            return !user.LockoutEnd.HasValue || user.LockoutEnd <= DateTimeOffset.UtcNow;
        }
    }
}
