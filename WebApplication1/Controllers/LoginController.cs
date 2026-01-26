// 2025.09.17 Changed: Login(POST) 초입에 서버측 Required 가드 추가(Username/Password 비었을 때 즉시 에러 스탬핑); 기존 흐름/리소스/헬퍼 유지
// 2025.09.16 Changed: EBValidate가 소비할 필드별 에러와 요약을 ViewData로 전달하는 로직 추가
// 2025.09.16 Changed: 로그인 실패 시 View 반환 전에 EBValidate 데이터 Stamp 호출
// 2025.09.16 Added: StampEbValidateData 헬퍼 추가
// 2025.09.16 Checked: 기존 AddErrorOnce MapAndAddRegisterError 유지
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using System.Text.Encodings.Web;
using System.Threading.Tasks;
using WebApplication1.Models;
using Microsoft.Extensions.Localization;
using System.Linq;
using System.Collections.Generic;
using System.Text.Json;

namespace WebApplication1.Controllers
{
    [AllowAnonymous]
    public class AccountController : Controller
    {
        private readonly SignInManager<ApplicationUser> _signInManager;
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly IStringLocalizer<SharedResource> _S;

        public AccountController(SignInManager<ApplicationUser> signInManager,
                                 UserManager<ApplicationUser> userManager,
                                 IStringLocalizer<SharedResource> S)
        {
            _signInManager = signInManager;
            _userManager = userManager;
            _S = S;
        }

        [Authorize]
        [HttpGet("/Account/LoginNotifications")]
        public async Task<IActionResult> LoginNotifications()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (string.IsNullOrEmpty(userId))
                return Json(new { approvalPending = 0, sharedUnread = 0 });

            var approvalPending = await GetApprovalPendingCountAsync(userId);
            var sharedUnread = await GetSharedUnreadCountAsync(userId);

            return Json(new
            {
                approvalPending,
                sharedUnread
            });
        }

        // 2025.12.15 Added: returnUrl 정규화 Login 자기자신 및 초대리셋 경유 URL 차단
        private string NormalizeReturnUrl(string? returnUrl)
        {
            var fallback = "/Doc/Board";

            if (string.IsNullOrWhiteSpace(returnUrl))
                return fallback;

            // 오픈 리다이렉트 방지
            if (!Url.IsLocalUrl(returnUrl))
                return fallback;

            // 로그인 페이지로 되돌아오는 케이스 차단
            if (returnUrl.Contains("/Identity/Account/Login", System.StringComparison.OrdinalIgnoreCase) ||
                returnUrl.Contains("/Account/Login", System.StringComparison.OrdinalIgnoreCase))
                return fallback;

            // 초대 및 비밀번호 설정 관련 페이지로 되돌아오는 케이스 차단
            if (returnUrl.Contains("/Identity/Account/InviteLanding", System.StringComparison.OrdinalIgnoreCase) ||
                returnUrl.Contains("/Identity/Account/ResetPassword", System.StringComparison.OrdinalIgnoreCase) ||
                returnUrl.Contains("/Identity/Account/SetPassword", System.StringComparison.OrdinalIgnoreCase))
                return fallback;

            return returnUrl;
        }

        [HttpGet]
        public IActionResult Login(string? returnUrl = null)
        {
            ViewData["ReturnUrl"] = NormalizeReturnUrl(returnUrl);

            StampEbValidateData();
            return View(new LoginViewModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(LoginViewModel model, string? returnUrl = null)
        {
            var normalizedReturnUrl = NormalizeReturnUrl(returnUrl);
            ViewData["ReturnUrl"] = normalizedReturnUrl;

            if (string.IsNullOrWhiteSpace(model.UserName))
                AddErrorOnce(nameof(LoginViewModel.UserName), _S["Req"]);
            if (string.IsNullOrWhiteSpace(model.Password))
                AddErrorOnce(nameof(LoginViewModel.Password), _S["Req"]);

            if (!ModelState.IsValid)
            {
                StampEbValidateData();
                return View(model);
            }

            var loginText = (model.UserName ?? string.Empty).Trim();

            ApplicationUser? user = loginText.Contains('@')
                ? await _userManager.FindByEmailAsync(loginText)
                : await _userManager.FindByNameAsync(loginText);

            if (user is null)
            {
                AddErrorOnce(nameof(LoginViewModel.UserName), _S["Login_UserNotExist"]);
                StampEbValidateData();
                return View(model);
            }

            if (string.IsNullOrWhiteSpace(model.Password))
            {
                AddErrorOnce(nameof(LoginViewModel.Password), _S["Req"]);
                StampEbValidateData();
                return View(model);
            }

            var result = await _signInManager.PasswordSignInAsync(
                user, model.Password, model.RememberMe, lockoutOnFailure: true);

            if (result.Succeeded)
                return LocalRedirect(normalizedReturnUrl);

            if (result.RequiresTwoFactor)
            {
                var encodedReturnUrl = UrlEncoder.Default.Encode(normalizedReturnUrl);
                return Redirect($"/Identity/Account/LoginWith2fa?ReturnUrl={encodedReturnUrl}&RememberMe={model.RememberMe}");
            }

            if (result.IsLockedOut)
            {
                AddErrorOnce(nameof(LoginViewModel.UserName), _S["Login_LockedOut"]);
                StampEbValidateData();
                return View(model);
            }

            if (result.IsNotAllowed)
            {
                AddErrorOnce(nameof(LoginViewModel.UserName), _S["Login_NotAllowed"]);
                StampEbValidateData();
                return View(model);
            }

            AddErrorOnce(nameof(LoginViewModel.Password), _S["Login_WrongPassword"]);
            StampEbValidateData();
            return View(model);
        }

        [HttpGet]
        public IActionResult Register() => View(new RegisterViewModel());

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Register(RegisterViewModel model)
        {
            if (!ModelState.IsValid)
            {
                StampEbValidateData();
                return View(model);
            }

            var user = new ApplicationUser { UserName = model.UserName, Email = model.Email };
            var result = await _userManager.CreateAsync(user, model.Password);
            if (result.Succeeded) return RedirectToAction("Login");

            foreach (var e in result.Errors)
                MapAndAddRegisterError(e);

            StampEbValidateData();
            return View(model);
        }

        // ✅ 추가: GET /Account/Logout 이 들어와도 400 대신 로그인으로 보냄
        [HttpGet]
        public async Task<IActionResult> Logout(string? returnUrl = null)
        {
            await _signInManager.SignOutAsync();

            // 세션 만료/자동 리다이렉트 상황에서는 "로그인 화면"이 목적
            // returnUrl이 있으면 Login으로 넘겨서 로그인 후 원위치 가능(정규화)
            var normalizedReturnUrl = NormalizeReturnUrl(returnUrl);
            return RedirectToAction("Login", new { returnUrl = normalizedReturnUrl });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Logout()
        {
            await _signInManager.SignOutAsync();
            // 기존 Home 이동 대신 Login으로 보내는 것이 UX/일관성에 유리
            return RedirectToAction("Login");
        }

        private void AddErrorOnce(string fieldKey, string message)
        {
            if (string.IsNullOrWhiteSpace(fieldKey) || string.IsNullOrWhiteSpace(message)) return;
            if (!ModelState.TryGetValue(fieldKey, out var entry) || !entry.Errors.Any(e => e.ErrorMessage == message))
                ModelState.AddModelError(fieldKey, message);
        }

        private void MapAndAddRegisterError(IdentityError e)
        {
            var code = e.Code ?? string.Empty;
            var desc = e.Description ?? "Invalid";

            string key;
            if (code.Contains("UserName", System.StringComparison.OrdinalIgnoreCase))
                key = nameof(RegisterViewModel.UserName);
            else if (code.Contains("Email", System.StringComparison.OrdinalIgnoreCase))
                key = nameof(RegisterViewModel.Email);
            else if (code.Equals("PasswordMismatch", System.StringComparison.OrdinalIgnoreCase))
                key = nameof(RegisterViewModel.ConfirmPassword);
            else if (code.StartsWith("Password", System.StringComparison.OrdinalIgnoreCase))
                key = nameof(RegisterViewModel.Password);
            else
                key = nameof(RegisterViewModel.UserName);

            AddErrorOnce(key, desc);
        }

        private void StampEbValidateData()
        {
            var fieldErrors = ModelState
                .Where(kv => !string.IsNullOrEmpty(kv.Key) && kv.Value is { Errors.Count: > 0 })
                .ToDictionary(
                    kv => GetFieldId(kv.Key),
                    kv => kv.Value!.Errors.Select(e => e.ErrorMessage).ToArray()
                );

            var summaryErrors = ModelState.Values
                .SelectMany(v => v.Errors)
                .Select(e => e.ErrorMessage)
                .Where(m => !string.IsNullOrWhiteSpace(m))
                .Distinct()
                .ToList();

            ViewData["EBV_FieldErrors"] = fieldErrors;
            ViewData["EBV_FieldErrorsJson"] = JsonSerializer.Serialize(fieldErrors);
            ViewData["EBV_SummaryErrors"] = summaryErrors;
            ViewData["EBV_SummaryErrorsJson"] = JsonSerializer.Serialize(summaryErrors);
        }

        private static string GetFieldId(string key)
        {
            var idx = key.LastIndexOf('.');
            return idx >= 0 ? key.Substring(idx + 1) : key;
        }
    }
}
