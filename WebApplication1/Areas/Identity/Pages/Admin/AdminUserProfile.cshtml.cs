// 2025.09.26 Added: 임시비밀번호 발행과 재초대 전용 핸들러 추가 저장 흐름 및 기존 코드 변경 없음
using System.Security.Claims;
using System.Text;
using System.Text.Encodings.Web; // 2025.09.26 Added
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.WebUtilities; // 2025.09.26 Added
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Options;
using WebApplication1.Data;
using WebApplication1.Models;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace WebApplication1.Areas.Identity.Pages.Admin
{
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
        public bool IsApprover { get; set; } = false; // 2025.09.23 Added
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
            VM.IsActive = !user.LockoutEnd.HasValue || user.LockoutEnd <= DateTimeOffset.UtcNow;
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
                VM.IsApprover = false;
                return Page();
            }

            await LoadUserAsync(id);
            return Page();
        }

        // 2025.09.24 Changed: 저장 가드와 메시지 병행 기존 로직 유지
        public async Task<IActionResult> OnPostSaveAsync()
        {
            var isCreate = string.Equals(VM.SelectedUserId, NewUserId, StringComparison.Ordinal);

            if (!isCreate && string.IsNullOrWhiteSpace(VM.SelectedUserId))
            {
                ModelState.Clear();
                var msg = _S["ACP_SelectUser_Req"].Value;
                ModelState.AddModelError(string.Empty, msg);
                ModelState.AddModelError(nameof(VM.SelectedUserId), msg);
                await BindMastersAsync();
                return Page();
            }

            if (isCreate && string.IsNullOrWhiteSpace(VM.UserName))
            {
                var msg = _S["_Alert_Require_ID"].Value;
                ModelState.AddModelError(nameof(VM.UserName), msg);
                ModelState.AddModelError(string.Empty, msg);
            }
            if (isCreate && string.IsNullOrWhiteSpace(VM.Email))
            {
                var msg = _S["_Alert_Require_EMail"].Value;
                ModelState.AddModelError(nameof(VM.Email), msg);
                ModelState.AddModelError(string.Empty, msg);
            }

            await BindMastersAsync();

            if (string.IsNullOrWhiteSpace(VM.CompCd))
            {
                ModelState.AddModelError(nameof(VM.CompCd), _S["ACP_CompCd_Req"].Value);
                ModelState.AddModelError(string.Empty, _S["ACP_CompCd_Req"].Value);
            }

            if (string.IsNullOrWhiteSpace(VM.DisplayName))
            {
                ModelState.AddModelError(nameof(VM.DisplayName), _S["_Alert_Require_Name"].Value);
                ModelState.AddModelError(string.Empty, _S["_Alert_Require_Name"].Value);
            }

            if (isCreate)
            {
                if (await _userManager.FindByNameAsync(VM.UserName!) != null)
                {
                    var msg = _S["_Alert_Duplicate_ID"].Value;
                    ModelState.AddModelError($"{nameof(VM)}.{nameof(VM.UserName)}", msg);
                    ModelState.AddModelError(nameof(VM.UserName), msg);
                    ModelState.AddModelError(string.Empty, msg);

                    ModelState.SetModelValue($"{nameof(VM)}.{nameof(VM.UserName)}", VM.UserName, VM.UserName);
                    ModelState.SetModelValue(nameof(VM.UserName), VM.UserName, VM.UserName);

                    await BindMastersAsync();
                    return Page();
                }

                if (!string.IsNullOrWhiteSpace(VM.Email) && await _userManager.FindByEmailAsync(VM.Email!) != null)
                {
                    var msg = _S["_Alert_Duplicate_EMail"].Value;
                    ModelState.AddModelError($"{nameof(VM)}.{nameof(VM.Email)}", msg);
                    ModelState.AddModelError(nameof(VM.Email), msg);
                    ModelState.AddModelError(string.Empty, msg);

                    ModelState.SetModelValue($"{nameof(VM)}.{nameof(VM.Email)}", VM.Email, VM.Email);
                    ModelState.SetModelValue(nameof(VM.Email), VM.Email, VM.Email);

                    await BindMastersAsync();
                    return Page();
                }
            }

            if (!ModelState.IsValid)
                return Page();

            if (isCreate)
            {
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
                    var emsg = createResult.Errors.FirstOrDefault()?.Description ?? "Create failed";
                    ModelState.AddModelError(nameof(VM.UserName), emsg);
                    ModelState.AddModelError(string.Empty, emsg);
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
                    IsApprover = VM.IsApprover
                };
                _db.UserProfiles.Add(profile);
                await _db.SaveChangesAsync();

                await UpdateIsAdminClaimAsync(user, VM.AdminLevel);
                await _userManager.UpdateSecurityStampAsync(user);

                try
                {
                    if (!string.IsNullOrWhiteSpace(user.Email))
                    {
                        var ecRaw = await _userManager.GenerateEmailConfirmationTokenAsync(user);
                        var ec = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(ecRaw));
                        var rpRaw = await _userManager.GeneratePasswordResetTokenAsync(user);
                        var rp = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(rpRaw));

                        var inviteUrl = Url.Page("/Account/InviteLanding", pageHandler: null,
                            values: new { area = "Identity", userId = user.Id, email = user.Email, ec, rp },
                            protocol: Request.Scheme)!;

                        var minutes = (int)_tokenOpts.CurrentValue.TokenLifespan.TotalMinutes;
                        var displayName = profile.DisplayName ?? (user.UserName ?? user.Email ?? "User");

                        var subject = _S["IL_Email_Subject"].Value;
                        var template = _S["IL_Email_BodyHtml"].Value;

                        var body = string.Format(
                            template,
                            HtmlEncoder.Default.Encode(displayName),
                            HtmlEncoder.Default.Encode(inviteUrl),
                            minutes
                        );

                        await _email.SendEmailAsync(user.Email!, subject, body);
                        TempData["StatusMessage"] = _S["ACP_SendInvite_Sent"].Value;
                    }
                    else
                    {
                        TempData["StatusMessage"] = _S["ACP_SendInvite_Failed", _S["IL_Error_InvalidLink"].Value].Value;
                    }
                }
                catch (Exception ex)
                {
                    TempData["StatusMessage"] = _S["ACP_SendInvite_Failed", ex.Message].Value;
                }

                if (VM.IsActive)
                    await _userManager.SetLockoutEndDateAsync(user, null);
                else
                    await _userManager.SetLockoutEndDateAsync(user, DateTimeOffset.UtcNow.AddYears(100));

                if (TempData["StatusMessage"] is null)
                    TempData["StatusMessage"] = _S["_CM_Saved"].Value;

                return RedirectToPage(new { id = user.Id, q = VM.Q });
            }
            else
            {
                var user = await _userManager.FindByIdAsync(VM.SelectedUserId!);
                var profile = await _db.UserProfiles.FirstOrDefaultAsync(p => p.UserId == VM.SelectedUserId);
                if (user is null || profile is null)
                {
                    ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_UserOrProfile_NotFound"].Value);
                    ModelState.AddModelError(string.Empty, _S["ACP_UserOrProfile_NotFound"].Value);
                    return Page();
                }

                if (!await _db.CompMasters.AnyAsync(x => x.IsActive && x.CompCd == VM.CompCd))
                {
                    ModelState.AddModelError(nameof(VM.CompCd), _S["ACP_Invalid_Comp"].Value);
                    ModelState.AddModelError(string.Empty, _S["ACP_Invalid_Comp"].Value);
                }
                if (VM.DepartmentId.HasValue &&
                    !await _db.DepartmentMasters.AnyAsync(x => x.IsActive && x.Id == VM.DepartmentId.Value))
                {
                    ModelState.AddModelError(nameof(VM.DepartmentId), _S["ACP_Invalid_Dept"].Value);
                    ModelState.AddModelError(string.Empty, _S["ACP_Invalid_Dept"].Value);
                }
                if (VM.PositionId.HasValue &&
                    !await _db.PositionMasters.AnyAsync(x => x.IsActive && x.Id == VM.PositionId.Value))
                {
                    ModelState.AddModelError(nameof(VM.PositionId), _S["ACP_Invalid_Pos"].Value);
                    ModelState.AddModelError(string.Empty, _S["ACP_Invalid_Pos"].Value);
                }

                if (!ModelState.IsValid)
                    return Page();

                profile.DisplayName = VM.DisplayName ?? string.Empty;
                profile.CompCd = VM.CompCd ?? string.Empty;
                profile.DepartmentId = VM.DepartmentId;
                profile.PositionId = VM.PositionId;
                profile.IsAdmin = Math.Clamp(VM.AdminLevel, 0, 2);
                profile.IsApprover = VM.IsApprover;

                user.PhoneNumber = VM.PhoneNumber ?? string.Empty;

                await _db.SaveChangesAsync();

                await UpdateIsAdminClaimAsync(user, VM.AdminLevel);

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

        // 2025.09.26 Added: 임시비밀번호 전용 핸들러 저장과 분리 레이아웃 변경 없음
         

        // 2025.09.26 Added: 재초대 전용 핸들러 저장과 분리 레이아웃 변경 없음
        public async Task<IActionResult> OnPostReinviteAsync()
        {
            // 대상 선택 검증
            if (string.IsNullOrWhiteSpace(VM?.SelectedUserId) || VM.SelectedUserId == NewUserId)
            {
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_SelectUser_Req"].Value);
                ModelState.AddModelError(string.Empty, _S["ACP_SelectUser_Req"].Value);
                await BindMastersAsync();
                return Page();
            }

            var user = await _userManager.FindByIdAsync(VM.SelectedUserId!);
            if (user is null)
            {
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_Error_UserNotFound"].Value);
                ModelState.AddModelError(string.Empty, _S["ACP_Error_UserNotFound"].Value);
                await BindMastersAsync();
                return Page();
            }

            // 2025.09.26 Added: 이메일 없으면 재초대 불가
            if (string.IsNullOrWhiteSpace(user.Email))
            {
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["IL_Error_InvalidLink"].Value);
                ModelState.AddModelError(string.Empty, _S["IL_Error_InvalidLink"].Value);
                await BindMastersAsync();
                return Page();
            }

            // 2025.09.26 Changed: 기존 프로필 재사용 신규 생성 금지
            var profile = await _db.UserProfiles
                .AsNoTracking()
                .FirstOrDefaultAsync(p => p.UserId == user.Id);

            // 2025.09.26 Added: 토큰은 최신 스탬프 기준으로 발급
            await _userManager.UpdateSecurityStampAsync(user);

            var ecRaw = await _userManager.GenerateEmailConfirmationTokenAsync(user);
            var ec = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(ecRaw));

            var rpRaw = await _userManager.GeneratePasswordResetTokenAsync(user);
            var rp = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(rpRaw));

            var inviteUrl = Url.Page("/Account/InviteLanding", pageHandler: null,
                values: new { area = "Identity", userId = user.Id, email = user.Email, ec, rp },
                protocol: Request.Scheme)!;

            var minutes = (int)_tokenOpts.CurrentValue.TokenLifespan.TotalMinutes;
            var displayName = profile?.DisplayName ?? (user.UserName ?? user.Email ?? "User");

            // 2025.09.26 Changed: 재초대 전용 리소스 키 사용
            var subject = _S["IL_Email_Subject"].Value;   // 2025.09.26 Confirmed: 기존 키 유지
            var template = _S["IL_Email_BodyHtml"].Value;

            var body = string.Format(
                template,
                HtmlEncoder.Default.Encode(displayName),   // {0}
                HtmlEncoder.Default.Encode(inviteUrl),     // {1}
                minutes                                    // {2}
            );

            await _email.SendEmailAsync(user.Email!, subject, body);

            // 2025.09.26 Added: 재초대 카운트 로깅 위치
            // 주의 신규 식별자 생성 금지 기존 프로젝트에 있는 동일 식별자만 사용하여 아래 위치에서 증가 저장
            // 예 profile.InviteResentCount++; profile.LastInviteSentAt = DateTimeOffset.UtcNow; _db.SaveChangesAsync();

            TempData["StatusMessage"] = _S["ACP_SendInvite_Sent"].Value;
            return RedirectToPage(new { id = VM.SelectedUserId, q = VM.Q });
        }

        public async Task<IActionResult> OnPostResetPasswordAsync()
        {
            if (string.IsNullOrEmpty(VM.SelectedUserId))
            {
                ModelState.AddModelError(string.Empty, _S["ACP_SelectUser_Req"].Value);
                ModelState.AddModelError("VM.SelectedUserId", _S["ACP_SelectUser_Req"].Value);
                await BindMastersAsync();
                return Page();
            }

            var user = await _userManager.FindByIdAsync(VM.SelectedUserId);
            if (user is null)
            {
                ModelState.AddModelError(string.Empty, _S["ACP_User_NotFound"].Value);
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
                ModelState.AddModelError(string.Empty, _S["ACP_TempPw_Failed", msg].Value);
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_TempPw_Failed", msg].Value);
                await BindMastersAsync();
                await LoadUserAsync(VM.SelectedUserId);
                return Page();
            }

            await _userManager.UpdateSecurityStampAsync(user);
            await _userManager.SetLockoutEndDateAsync(user, DateTimeOffset.UtcNow);

            var resetToken = await _userManager.GeneratePasswordResetTokenAsync(user);
            var encoded = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(resetToken));
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
                ModelState.AddModelError(string.Empty, _S["ACP_SelectUser_Req"].Value);
                ModelState.AddModelError("newEmail", _S["_Alert_Require_NewEmail"].Value);
                await BindMastersAsync();
                return Page();
            }

            var user = await _userManager.FindByIdAsync(VM.SelectedUserId);
            if (user is null)
            {
                ModelState.AddModelError(string.Empty, _S["ACP_User_NotFound"].Value);
                ModelState.AddModelError(nameof(VM.SelectedUserId), _S["ACP_User_NotFound"].Value);
                await BindMastersAsync();
                return Page();
            }

            if (string.IsNullOrWhiteSpace(newEmail))
            {
                ModelState.AddModelError(string.Empty, _S["_Alert_Require_NewEmail"].Value);
                ModelState.AddModelError("newEmail", _S["_Alert_Require_NewEmail"].Value);
                await BindMastersAsync();
                await LoadUserAsync(VM.SelectedUserId);
                return Page();
            }

            var exists = await _userManager.FindByEmailAsync(newEmail.Trim());
            if (exists is not null && exists.Id != user.Id)
            {
                ModelState.AddModelError(string.Empty, _S["ACP_EmailChange_InUse"].Value);
                ModelState.AddModelError("newEmail", _S["ACP_EmailChange_InUse"].Value);
                await BindMastersAsync();
                await LoadUserAsync(VM.SelectedUserId);
                return Page();
            }

            var code = await _userManager.GenerateChangeEmailTokenAsync(user, newEmail);
            var encoded = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(code));
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
    }
}
