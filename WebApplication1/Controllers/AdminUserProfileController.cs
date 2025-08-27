using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Models.ViewModels;
using Microsoft.AspNetCore.Identity.UI.Services;

namespace WebApplication1.Controllers
{
    [Authorize(Policy = "AdminOnly")]
    public class AdminUserProfileController : Controller
    {
        private readonly ApplicationDbContext _db;
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly SignInManager<ApplicationUser> _signInManager;
        private readonly IEmailSender _emailSender;

        public AdminUserProfileController(
            ApplicationDbContext db,
            UserManager<ApplicationUser> userManager,
            SignInManager<ApplicationUser> signInManager,
            IEmailSender emailSender)
        {
            _db = db;
            _userManager = userManager;
            _signInManager = signInManager;
            _emailSender = emailSender;
        }
        //
        [HttpGet]
        public async Task<IActionResult> Create()
        {
            var vm = new AdminUserProfileVM { IsCreate = true, AdminLevel = 0 };

            // 상단 콤보는 유지(선택하면 기존 사용자 편집으로 전환)
            vm.Accounts = await _db.Users.AsNoTracking()
                .OrderBy(u => u.UserName)
                .Select(u => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem
                {
                    Value = u.Id,
                    Text = $"{u.UserName} ({u.Email})"
                })
                .ToListAsync();

            await BindMasters(vm);
            return View("Index", vm);                  // 기존 Index.cshtml 재사용
        }

        // POST: /AdminUserProfile/Create  (Insert)
        [HttpPost, ValidateAntiForgeryToken]
        public async Task<IActionResult> Create(AdminUserProfileVM vm)
        {
            vm.IsCreate = true;                        // 뷰로 되돌릴 때 모드 유지
            await BindMasters(vm);

            // --- 서버 검증 ---
            if (string.IsNullOrWhiteSpace(vm.UserName))
                ModelState.AddModelError(nameof(vm.UserName), "ID(사용자명)을 입력하세요.");
            if (string.IsNullOrWhiteSpace(vm.Email))
                ModelState.AddModelError(nameof(vm.Email), "E-Mail을 입력하세요.");
            if (string.IsNullOrWhiteSpace(vm.DisplayName))
                ModelState.AddModelError(nameof(vm.DisplayName), "이름을 입력하세요.");

            if (!string.IsNullOrWhiteSpace(vm.UserName) &&
                await _userManager.FindByNameAsync(vm.UserName) != null)
                ModelState.AddModelError(nameof(vm.UserName), "이미 존재하는 ID입니다.");

            if (!string.IsNullOrWhiteSpace(vm.Email) &&
                await _userManager.FindByEmailAsync(vm.Email) != null)
                ModelState.AddModelError(nameof(vm.Email), "이미 사용 중인 E-Mail입니다.");

            if (string.IsNullOrWhiteSpace(vm.CompCd) ||
                !await _db.CompMasters.AnyAsync(c => c.IsActive && c.CompCd == vm.CompCd))
                ModelState.AddModelError(nameof(vm.CompCd), "유효한 사업장을 선택하세요.");

            if (vm.DepartmentId.HasValue &&
                !await _db.DepartmentMasters.AnyAsync(d => d.IsActive && d.Id == vm.DepartmentId.Value))
                ModelState.AddModelError(nameof(vm.DepartmentId), "유효한 부서입니다.");

            if (vm.PositionId.HasValue &&
                !await _db.PositionMasters.AnyAsync(p => p.IsActive && p.Id == vm.PositionId.Value))
                ModelState.AddModelError(nameof(vm.PositionId), "유효한 직급입니다.");

            if (!ModelState.IsValid)
            {
                // 상단 계정 콤보 채워서 같은 화면으로
                vm.Accounts = await _db.Users.AsNoTracking()
                    .OrderBy(u => u.UserName)
                    .Select(u => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem
                    {
                        Value = u.Id,
                        Text = $"{u.UserName} ({u.Email})"
                    }).ToListAsync();
                return View("Index", vm);
            }

            // --- Insert Users ---
            var user = new ApplicationUser
            {
                UserName = vm.UserName!.Trim(),
                Email = vm.Email!.Trim(),
                PhoneNumber = vm.PhoneNumber,
                EmailConfirmed = false
            };

            var tempPw = "Temp!" + Guid.NewGuid().ToString("N")[..8];
            var cr = await _userManager.CreateAsync(user, tempPw);
            if (!cr.Succeeded)
            {
                foreach (var e in cr.Errors) ModelState.AddModelError("", e.Description);
                vm.Accounts = await _db.Users.AsNoTracking()
                    .OrderBy(u => u.UserName)
                    .Select(u => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem
                    {
                        Value = u.Id,
                        Text = $"{u.UserName} ({u.Email})"
                    }).ToListAsync();
                return View("Index", vm);
            }

            // --- Insert UserProfiles ---
            var profile = new UserProfile
            {
                UserId = user.Id,
                DisplayName = vm.DisplayName,
                CompCd = vm.CompCd!,
                DepartmentId = vm.DepartmentId,
                PositionId = vm.PositionId,
                IsAdmin = Math.Clamp(vm.AdminLevel, 0, 2)
            };
            _db.UserProfiles.Add(profile);
            await _db.SaveChangesAsync();

            // --- is_admin 클레임 부여 ---
            await UpdateIsAdminClaimAsync(user, vm.AdminLevel);

            // --- 비밀번호 설정 메일 발송(초대) ---
            var resetToken = await _userManager.GeneratePasswordResetTokenAsync(user);
            var resetUrl = $"{Request.Scheme}://{Request.Host}/Identity/Account/ResetPassword" +
                           $"?code={Uri.EscapeDataString(resetToken)}&email={Uri.EscapeDataString(user.Email!)}";
            await _emailSender.SendEmailAsync(
                user.Email!,
                "[E-Board] 초기 비밀번호 설정",
                $@"아래 링크에서 초기 비밀번호를 설정하세요: <a href=""{resetUrl}"">비밀번호 설정</a>");

            TempData["StatusMessage"] = "사용자를 추가했습니다. 초기 비밀번호 설정 메일을 발송했습니다.";
            return RedirectToAction(nameof(Index), new { id = user.Id });   // 곧바로 편집화면으로 이동
        }
        //
        // GET: /AdminUserProfile?id={userId}&q={검색}
        public async Task<IActionResult> Index(string? id, string? q)
        {
            var vm = new AdminUserProfileVM { SelectedUserId = id, Q = q };

            // 계정 콤보
            var usersQuery = _db.Users.AsNoTracking();
            if (!string.IsNullOrWhiteSpace(q))
            {
                q = q.Trim();
                usersQuery = usersQuery.Where(u => u.UserName!.Contains(q) || u.Email!.Contains(q));
            }
            vm.Accounts = await usersQuery
                .OrderBy(u => u.UserName)
                .Select(u => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem
                {
                    Value = u.Id,
                    Text = $"{u.UserName} ({u.Email})"
                })
                .ToListAsync();

            if (string.IsNullOrEmpty(id))
                return View(vm);

            var user = await _userManager.FindByIdAsync(id);
            if (user == null)
            {
                TempData["StatusMessage"] = "사용자를 찾을 수 없습니다.";
                return View(vm);
            }

            var profile = await _db.UserProfiles.AsNoTracking().FirstOrDefaultAsync(p => p.UserId == id);
            if (profile == null)
            {
                TempData["StatusMessage"] = "사용자 프로필이 없습니다.";
                return View(vm);
            }

            // 화면 바인딩
            vm.Email = user.Email;
            vm.UserName = user.UserName;
            vm.DisplayName = profile.DisplayName ?? string.Empty;
            vm.CompCd = profile.CompCd ?? string.Empty;
            vm.DepartmentId = profile.DepartmentId;   // int?
            vm.PositionId = profile.PositionId;     // int?
            vm.PhoneNumber = user.PhoneNumber;

            // 권한 레벨 바인딩: is_admin 클레임(우선) -> 프로필.IsAdmin(백업)
            var claims = await _userManager.GetClaimsAsync(user);
            var adminClaim = claims.FirstOrDefault(c => c.Type == "is_admin");
            int adminLevel = 0;
            if (adminClaim?.Value == "1") adminLevel = 1;
            else if (adminClaim?.Value == "2") adminLevel = 2;
            else
            {
                // 프로필 보정: 0/1/2로 수렴
                if (profile.IsAdmin >= 2) adminLevel = 2;
                else if (profile.IsAdmin == 1) adminLevel = 1;
            }
            vm.AdminLevel = adminLevel;

            await BindMasters(vm);
            return View(vm);
        }

        // POST 저장
        [HttpPost, ValidateAntiForgeryToken]
        public async Task<IActionResult> Save(AdminUserProfileVM vm)
        {
            if (string.IsNullOrEmpty(vm.SelectedUserId))
            {
                TempData["StatusMessage"] = "수정할 계정을 선택하세요.";
                return RedirectToAction(nameof(Index));
            }

            var user = await _userManager.FindByIdAsync(vm.SelectedUserId);
            if (user == null)
            {
                TempData["StatusMessage"] = "사용자를 찾을 수 없습니다.";
                return RedirectToAction(nameof(Index));
            }

            var profile = await _db.UserProfiles.FirstOrDefaultAsync(p => p.UserId == vm.SelectedUserId);
            if (profile == null)
            {
                TempData["StatusMessage"] = "사용자 프로필이 없습니다.";
                return RedirectToAction(nameof(Index), new { id = vm.SelectedUserId });
            }

            // 마스터 유효성(활성만)
            if (!await _db.CompMasters.AnyAsync(x => x.IsActive && x.CompCd == vm.CompCd))
                ModelState.AddModelError(nameof(vm.CompCd), "유효하지 않은 지사 코드입니다.");

            if (vm.DepartmentId.HasValue &&
                !await _db.DepartmentMasters.AnyAsync(x => x.IsActive && x.Id == vm.DepartmentId.Value))
                ModelState.AddModelError(nameof(vm.DepartmentId), "유효하지 않은 부서입니다.");

            if (vm.PositionId.HasValue &&
                !await _db.PositionMasters.AnyAsync(x => x.IsActive && x.Id == vm.PositionId.Value))
                ModelState.AddModelError(nameof(vm.PositionId), "유효하지 않은 직급입니다.");

            if (!ModelState.IsValid)
            {
                await BindMasters(vm);
                vm.Accounts = await _db.Users.AsNoTracking()
                    .OrderBy(u => u.UserName)
                    .Select(u => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem
                    {
                        Value = u.Id,
                        Text = $"{u.UserName} ({u.Email})",
                        Selected = u.Id == vm.SelectedUserId
                    })
                    .ToListAsync();

                return View("Index", vm);
            }

            // 저장
            profile.DisplayName = vm.DisplayName;
            profile.CompCd = vm.CompCd;
            profile.DepartmentId = vm.DepartmentId;     // int?
            profile.PositionId = vm.PositionId;       // int?
            profile.IsAdmin = Math.Clamp(vm.AdminLevel, 0, 2); // DB에도 0/1/2 반영

            user.PhoneNumber = vm.PhoneNumber;

            await _db.SaveChangesAsync();

            // is_admin 클레임 갱신 (0=제거, 1/2=셋)
            await UpdateIsAdminClaimAsync(user, vm.AdminLevel);

            // 권한/클레임 반영
            await _userManager.UpdateSecurityStampAsync(user);
            if (User.FindFirstValue(ClaimTypes.NameIdentifier) == user.Id)
                await _signInManager.RefreshSignInAsync(user);

            TempData["StatusMessage"] = "저장되었습니다.";
            return RedirectToAction(nameof(Index), new { id = vm.SelectedUserId });
        }

        // POST 비밀번호 초기화
        [HttpPost, ValidateAntiForgeryToken]
        public async Task<IActionResult> ResetPassword(string id)
        {
            var user = await _userManager.FindByIdAsync(id);
            if (user == null)
            {
                TempData["StatusMessage"] = "사용자를 찾을 수 없습니다.";
                return RedirectToAction(nameof(Index));
            }

            // 1) 임시 비밀번호로 즉시 초기화 (현재 동작 유지)
            var tokenForTemp = await _userManager.GeneratePasswordResetTokenAsync(user);
            var tempPw = "Temp!" + Guid.NewGuid().ToString("N")[..8];

            var result = await _userManager.ResetPasswordAsync(user, tokenForTemp, tempPw);
            if (!result.Succeeded)
            {
                TempData["StatusMessage"] = "비밀번호 초기화 실패: " + string.Join(", ", result.Errors.Select(e => e.Description));
                return RedirectToAction(nameof(Index), new { id });
            }

            // 보안 스탬프 갱신 + 잠금 즉시 해제(원하시면 제거)
            await _userManager.UpdateSecurityStampAsync(user);
            await _userManager.SetLockoutEndDateAsync(user, DateTimeOffset.UtcNow);

            // 2) 사용자에게 비밀번호 변경 링크(30분 유효) 메일 전송
            //    (DataProtectionTokenProviderOptions.TokenLifespan = 30분 으로 Program.cs에 이미 설정되어 있어야 합니다)
            var resetToken = await _userManager.GeneratePasswordResetTokenAsync(user);
            var resetUrl =
                $"{Request.Scheme}://{Request.Host}/Identity/Account/ResetPassword" +
                $"?code={Uri.EscapeDataString(resetToken)}&email={Uri.EscapeDataString(user.Email ?? string.Empty)}";

            // 표시용 이름
            var displayName = await _db.UserProfiles
                .Where(p => p.UserId == user.Id)
                .Select(p => p.DisplayName)
                .FirstOrDefaultAsync() ?? (user.UserName ?? user.Email ?? "사용자");

            var subject = "[Han Young E-Board] 비밀번호 초기화 안내";
            var body = $@"
                        <!doctype html>
                        <html lang=""ko"">
                          <body style=""font-family:Segoe UI,Arial,sans-serif;font-size:14px;margin:0;padding:16px;"">
                            <h2 style=""margin:0 0 12px"">비밀번호 초기화 안내</h2>

                            <p>안녕하세요, <strong>{System.Text.Encodings.Web.HtmlEncoder.Default.Encode(displayName)}</strong> 님.</p>

                            <p>관리자가 회원님의 계정 비밀번호를 임시 비밀번호로 초기화했습니다.</p>     
                            <p style=""margin:12px 0;padding:10px;background:#f8f9fa;border:1px solid #eee;border-radius:6px;"">
                              <strong>임시 비밀번호:</strong> {System.Text.Encodings.Web.HtmlEncoder.Default.Encode(tempPw)}
                            </p>

                            <p>아래 버튼을 클릭하여 <strong>새 비밀번호</strong>로 즉시 변경해 주세요.</p>

                            <p style=""margin:20px 0"">
                              <a href=""{System.Text.Encodings.Web.HtmlEncoder.Default.Encode(resetUrl)}""
                                 style=""background:#0d6efd;color:#fff;padding:10px 16px;border-radius:6px;
                                        text-decoration:none;display:inline-block"">비밀번호 변경하기</a>
                            </p>

                            <p style=""color:#6c757d;margin:10px 0;"">
                              ※ 이 메일의 비밀번호 변경 링크는 <strong>수신 시각부터 30분 동안</strong>만 유효합니다.
                              만료되면 관리자에게 재요청하시거나 로그인 후 설정에서 변경해 주세요.
                            </p>

                            <hr style=""margin:24px 0;border:none;border-top:1px solid #eee""/>

                            <p style=""color:#6c757d"">본인이 요청하지 않았다면 이 메일은 무시하셔도 됩니다.</p>
                          </body>
                        </html>";

            if (!string.IsNullOrWhiteSpace(user.Email))
                await _emailSender.SendEmailAsync(user.Email, subject, body);

            // 3) 관리자 화면 알림
            TempData["StatusMessage"] = "임시 비밀번호를 발급하고 안내 메일(30분 유효)을 전송했습니다.";
            return RedirectToAction(nameof(Index), new { id });
        }

        // POST 이메일 변경 링크 발송(강제 변경 아님)
        [HttpPost, ValidateAntiForgeryToken]
        public async Task<IActionResult> SendChangeEmail(string id, string newEmail)
        {
            if (string.IsNullOrWhiteSpace(newEmail))
            {
                TempData["StatusMessage"] = "새 이메일을 입력하세요.";
                return RedirectToAction(nameof(Index), new { id });
            }

            var user = await _userManager.FindByIdAsync(id);
            if (user == null)
            {
                TempData["StatusMessage"] = "사용자를 찾을 수 없습니다.";
                return RedirectToAction(nameof(Index));
            }

            var code = await _userManager.GenerateChangeEmailTokenAsync(user, newEmail);

            var path = $"/Identity/Account/ConfirmEmailChange" +
                       $"?userId={Uri.EscapeDataString(user.Id)}" +
                       $"&email={Uri.EscapeDataString(newEmail)}" +
                       $"&code={Uri.EscapeDataString(code)}";
            var callbackUrl = $"{Request.Scheme}://{Request.Host}{path}";

            await _emailSender.SendEmailAsync(newEmail,
                "[E-Board] 이메일 변경 확인",
                $@"이메일 변경을 확인하려면 <a href=""{callbackUrl}"">여기</a>를 클릭하세요.");

            TempData["StatusMessage"] = "변경 확인 메일을 발송했습니다.";
            return RedirectToAction(nameof(Index), new { id });
        }

        // ---- Helpers ----
        private async Task UpdateIsAdminClaimAsync(ApplicationUser user, int level)
        {
            // 기존 is_admin 모두 제거
            var claims = await _userManager.GetClaimsAsync(user);
            var toRemove = claims.Where(c => c.Type == "is_admin").ToList();
            foreach (var c in toRemove)
                await _userManager.RemoveClaimAsync(user, c);

            // 1/2만 추가 (0은 일반 → 미보유)
            level = Math.Clamp(level, 0, 2);
            if (level > 0)
                await _userManager.AddClaimAsync(user, new Claim("is_admin", level.ToString()));
        }

        private async Task BindMasters(AdminUserProfileVM vm)
        {
            vm.CompList = await _db.CompMasters
                .Where(x => x.IsActive)
                .OrderBy(x => x.CompCd)
                .Select(x => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem
                {
                    Value = x.CompCd,
                    Text = x.Name
                })
                .ToListAsync();

            vm.DeptList = await _db.DepartmentMasters
                .Where(x => x.IsActive)
                .OrderBy(x => x.Name)
                .Select(x => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem
                {
                    Value = x.Id.ToString(),
                    Text = x.Name
                })
                .ToListAsync();

            vm.PosList = await _db.PositionMasters
                .Where(x => x.IsActive)
                .OrderBy(x => x.RankLevel)
                .Select(x => new Microsoft.AspNetCore.Mvc.Rendering.SelectListItem
                {
                    Value = x.Id.ToString(),
                    Text = x.Name
                })
                .ToListAsync();
        }
    }
}
