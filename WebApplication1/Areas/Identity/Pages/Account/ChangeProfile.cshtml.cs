// Areas/Identity/Pages/Account/ChangeProfile.cshtml.cs
using System.ComponentModel.DataAnnotations;
using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.EntityFrameworkCore;
using WebApplication1.Data;
using WebApplication1.Models;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Identity.UI.Services;
using System.Text.Encodings.Web;

namespace WebApplication1.Areas.Identity.Pages.Account
{
    [Authorize]
    public class ChangeProfileModel : PageModel
    {
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly SignInManager<ApplicationUser> _signInManager;
        private readonly ApplicationDbContext _db;
        private readonly ILogger<ChangeProfileModel> _logger;
        private readonly IEmailSender _emailSender;

        [TempData] public string? DevEmailLink { get; set; }

        public ChangeProfileModel(
            UserManager<ApplicationUser> userManager,
            SignInManager<ApplicationUser> signInManager,
            ApplicationDbContext db,
            ILogger<ChangeProfileModel> logger,
            IEmailSender emailSender)
        {
            _userManager = userManager;
            _signInManager = signInManager;
            _db = db;
            _logger = logger;
            _emailSender = emailSender;
        }

        // ===== ViewModels =====
        public class InputModel
        {
            [EmailAddress, Display(Name = "E-Mail")]
            public string? Email { get; set; }

            [Display(Name = "ID")]
            public string? UserName { get; set; }   // 읽기전용 표시

            [Display(Name = "부서")] public int? DepartmentId { get; set; }
            [Display(Name = "직급")] public int? PositionId { get; set; }
            [Display(Name = "이름")] public string? DisplayName { get; set; }
            [Phone, Display(Name = "전화번호")] public string? PhoneNumber { get; set; }
            [Display(Name = "사업장")] public string? CompCd { get; set; }

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

        // ===== Bindings =====
        [BindProperty] public InputModel Input { get; set; } = new();

        // EmailChange는 메인 저장(Post)에서 자동 바인딩/검증되지 않도록 둡니다.
        // (모달 핸들러에서 [FromForm] 파라미터로만 검증/사용)
        public EmailChangeInput? EmailChange { get; set; }

        // ===== Dropdowns =====
        public SelectList CompOptions { get; set; } = default!;
        public SelectList DepartmentOptions { get; set; } = default!;
        public SelectList PositionOptions { get; set; } = default!;

        // ===== GET =====
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

            EmailChange = new EmailChangeInput(); // 뷰 렌더용
            return Page();
        }

        // ===== POST: 프로필 저장 =====
        public async Task<IActionResult> OnPostAsync()
        {
            var user = await _userManager.GetUserAsync(User);
            if (user is null) return RedirectToPage("/Account/Login");

            // 드롭다운(검증 실패 시에도 리스트가 비지 않도록 먼저 바인딩)
            await LoadOptionsAsync();

            // 비번 입력값 공백 정규화
            Input.NewPassword = (Input.NewPassword ?? string.Empty).Trim();
            Input.ConfirmNewPassword = (Input.ConfirmNewPassword ?? string.Empty).Trim();
            Input.CurrentPassword = (Input.CurrentPassword ?? string.Empty).Trim();

            // 1) 서버측 유효성 (활성만 허용)
            if (string.IsNullOrWhiteSpace(Input.CompCd) ||
                !await _db.CompMasters.AnyAsync(c => c.IsActive && c.CompCd == Input.CompCd))
                ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.CompCd)}", "유효한 사업장을 선택하세요.");

            if (Input.DepartmentId is null ||
                !await _db.DepartmentMasters.AnyAsync(d => d.IsActive && d.Id == Input.DepartmentId))
                ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.DepartmentId)}", "유효한 부서를 선택하세요.");

            if (Input.PositionId is null ||
                !await _db.PositionMasters.AnyAsync(p => p.IsActive && p.Id == Input.PositionId))
                ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.PositionId)}", "유효한 직급을 선택하세요.");

            if (string.IsNullOrWhiteSpace(Input.DisplayName))
                ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.DisplayName)}", "이름을 입력하세요.");

            // 2) 비밀번호 변경 의사가 있는 경우에만 검증/변경 수행
            var wantsPwChange = Input.NewPassword.Length > 0 || Input.ConfirmNewPassword.Length > 0;
            if (wantsPwChange)
            {
                if (Input.NewPassword.Length == 0)
                    ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.NewPassword)}", "새 비밀번호를 입력하세요.");

                if (Input.ConfirmNewPassword.Length == 0)
                    ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.ConfirmNewPassword)}", "새 비밀번호 확인을 입력하세요.");

                if (Input.NewPassword.Length > 0 &&
                    Input.ConfirmNewPassword.Length > 0 &&
                    Input.NewPassword != Input.ConfirmNewPassword)
                {
                    ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.ConfirmNewPassword)}", "새 비밀번호가 일치하지 않습니다.");
                }

                if (Input.CurrentPassword.Length == 0)
                    ModelState.AddModelError($"{nameof(Input)}.{nameof(Input.CurrentPassword)}", "현재 비밀번호를 입력하세요.");
            }

            if (!ModelState.IsValid)
            {
                EmailChange = new EmailChangeInput(); // 뷰 다시 렌더용
                return Page();
            }

            // 3) 비밀번호 변경
            if (wantsPwChange)
            {
                var cr = await _userManager.ChangePasswordAsync(user, Input.CurrentPassword!, Input.NewPassword!);
                if (!cr.Succeeded)
                {
                    foreach (var e in cr.Errors) ModelState.AddModelError("", e.Description);
                    EmailChange = new EmailChangeInput();
                    return Page();
                }
            }

            // 4) 프로필 로드/저장
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
            else
            {
                // 기존 프로필에도 CompCd 반영
                profile.CompCd = Input.CompCd!.Trim().ToUpperInvariant();
            }

            profile.DisplayName = Input.DisplayName?.Trim();
            profile.DepartmentId = Input.DepartmentId;
            profile.PositionId = Input.PositionId;

            // 5) 사용자 기본 필드
            if (!string.IsNullOrWhiteSpace(Input.Email) &&
                !string.Equals(user.Email, Input.Email, StringComparison.OrdinalIgnoreCase))
            {
                var er = await _userManager.SetEmailAsync(user, Input.Email);
                if (!er.Succeeded)
                {
                    foreach (var e in er.Errors) ModelState.AddModelError("", e.Description);
                    EmailChange = new EmailChangeInput();
                    return Page();
                }
            }

            if (Input.PhoneNumber != user.PhoneNumber)
            {
                var pr = await _userManager.SetPhoneNumberAsync(user, Input.PhoneNumber);
                if (!pr.Succeeded)
                {
                    foreach (var e in pr.Errors) ModelState.AddModelError("", e.Description);
                    EmailChange = new EmailChangeInput();
                    return Page();
                }
            }

            var ur = await _userManager.UpdateAsync(user);
            if (!ur.Succeeded)
            {
                foreach (var e in ur.Errors) ModelState.AddModelError("", e.Description);
                EmailChange = new EmailChangeInput();
                return Page();
            }

            await _db.SaveChangesAsync();
            await _signInManager.RefreshSignInAsync(user);

            TempData["StatusMessage"] = "프로필이 저장되었습니다.";
            return RedirectToPage(); // GET으로 리다이렉트하여 최신 값 표시
        }

        // ===== POST: 이메일 변경 확인 메일 발송(모달) =====
        public async Task<IActionResult> OnPostSendEmailChangeAsync([FromForm] EmailChangeInput emailChange)
        {
            var user = await _userManager.GetUserAsync(User);
            if (user is null) return RedirectToPage("/Account/Login");

            await LoadOptionsAsync();

            if (!TryValidateModel(emailChange)) { EmailChange = emailChange; return Page(); }

            var pwOk = await _userManager.CheckPasswordAsync(user, emailChange.CurrentPassword!);
            if (!pwOk)
            {
                ModelState.AddModelError("EmailChange.CurrentPassword", "현재 비밀번호가 올바르지 않습니다.");
                EmailChange = emailChange;
                return Page();
            }

            if (string.Equals(user.Email, emailChange.NewEmail, StringComparison.OrdinalIgnoreCase))
            {
                ModelState.AddModelError("EmailChange.NewEmail", "기존 이메일과 동일합니다.");
                EmailChange = emailChange;
                return Page();
            }

            var token = await _userManager.GenerateChangeEmailTokenAsync(user, emailChange.NewEmail!);
            var code = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(token));
            var callbackUrl = Url.Page("/Account/ConfirmEmailChange", null,
                new { userId = user.Id, email = emailChange.NewEmail, code }, Request.Scheme);

            // === 여기부터 커스텀 제목/본문 ===
            var displayName = (await _db.UserProfiles
                .AsNoTracking()
                .Where(p => p.UserId == user.Id)
                .Select(p => p.DisplayName)
                .FirstOrDefaultAsync()) ?? user.UserName ?? "사용자";
            string subject = "[Han Young E-Board] 이메일 변경 확인 안내";

            string body = $@"
                        <!doctype html>
                        <html lang=""ko"">
                          <body style=""font-family:Segoe UI,Arial,sans-serif;font-size:14px;margin:0;padding:16px;"">
                            <h2 style=""margin:0 0 12px"">이메일 변경 확인</h2>

                            <p>안녕하세요, <strong>{displayName}</strong> 님.</p>
                            <p>아래 버튼을 클릭하시면 해당 계정의 이메일을
                               <strong>{HtmlEncoder.Default.Encode(emailChange.NewEmail!)}</strong> 로 변경합니다.</p>

                            <p style=""margin:20px 0"">
                              <a href=""{HtmlEncoder.Default.Encode(callbackUrl!)}""
                                 style=""background:#0d6efd;color:#fff;padding:10px 16px;border-radius:6px;
                                        text-decoration:none;display:inline-block"">이메일 변경 확인</a>
                            </p>

                            <!-- ▼ 유효시간 안내 추가 -->
                            <p style=""color:#6c757d;margin:8px 0 0"">
                              이 메일의 확인 링크는 <strong>수신된 시간으로부터 30분간</strong> 유효합니다.
                            </p>
                            <!-- ▲ 유효시간 안내 추가 -->

                            <p>버튼이 동작하지 않으면 다음 주소를 복사해 브라우저 주소창에 붙여 넣으세요.</p>
                            <p style=""word-break:break-all;color:#555"">{HtmlEncoder.Default.Encode(callbackUrl!)}</p>

                            <hr style=""margin:24px 0;border:none;border-top:1px solid #eee""/>

                            <p style=""color:#6c757d"">
                              만약 본인이 요청하지 않았다면 이 메일은 무시하셔도 됩니다.
                            </p>
                          </body>
                        </html>";


            await _emailSender.SendEmailAsync(emailChange.NewEmail!, subject, body);

            TempData["StatusMessage"] = "확인 메일을 보냈습니다. 메일함을 확인하세요.";
            return RedirectToPage();
        }

        // ===== 공통 =====
        private async Task LoadOptionsAsync()
        {
            var comps = await _db.CompMasters.Where(c => c.IsActive).OrderBy(c => c.CompCd).ToListAsync();
            var deps = await _db.DepartmentMasters.Where(d => d.IsActive).OrderBy(d => d.Name).ToListAsync();
            var poss = await _db.PositionMasters.Where(p => p.IsActive).OrderBy(p => p.RankLevel).ToListAsync();

            // 지사명 컬럼명이 Name으로 정의되어 있다고 가정 (다르면 여기만 바꾸세요)
            CompOptions = new SelectList(comps, "CompCd", "Name");
            DepartmentOptions = new SelectList(deps, "Id", "Name");
            PositionOptions = new SelectList(poss, "Id", "Name");
        }
    }
}
