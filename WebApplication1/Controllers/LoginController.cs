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
            // BEFORE:
            // ViewData["ReturnUrl"] = returnUrl ?? Url.Content("~/");

            // AFTER:
            ViewData["ReturnUrl"] = NormalizeReturnUrl(returnUrl);

            StampEbValidateData();
            return View(new LoginViewModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(LoginViewModel model, string? returnUrl = null)
        {
            // BEFORE:
            // returnUrl ??= Url.Content("~/");
            // ViewData["ReturnUrl"] = returnUrl;

            // AFTER:
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
            {
                // BEFORE:
                // var target = (returnUrl != null && Url.IsLocalUrl(returnUrl)) ? returnUrl : "/Doc/Board";
                // return Redirect(target);

                // AFTER:
                return LocalRedirect(normalizedReturnUrl);
            }

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
                // 2025.09.16 Added: 등록 폼도 동일하게 EBValidate 데이터 Stamp
                StampEbValidateData();
                return View(model);
            }

            var user = new ApplicationUser { UserName = model.UserName, Email = model.Email };
            var result = await _userManager.CreateAsync(user, model.Password);
            if (result.Succeeded) return RedirectToAction("Login");

            // 2025.09.16 Changed: Identity 오류를 필드 키로 매핑
            foreach (var e in result.Errors)
                MapAndAddRegisterError(e);

            StampEbValidateData(); // 2025.09.16 Added
            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Logout()
        {
            await _signInManager.SignOutAsync();
            return RedirectToAction("Index", "Home");
        }

        // 2025.09.16 Added: 중복 방지용 헬퍼 동일 메시지 1회만
        private void AddErrorOnce(string fieldKey, string message)
        {
            if (string.IsNullOrWhiteSpace(fieldKey) || string.IsNullOrWhiteSpace(message)) return;
            if (!ModelState.TryGetValue(fieldKey, out var entry) || !entry.Errors.Any(e => e.ErrorMessage == message))
                ModelState.AddModelError(fieldKey, message);
        }

        // 2025.09.16 Added: Identity 오류코드 매핑
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

        // 2025.09.16 Added: EBValidate가 바로 사용할 수 있는 데이터 ViewData에 Stamp
        // 요약 에러는 EBValidate.showAlert 로 사용
        // 필드별 에러는 EBValidate.setInvalid 대상 입력 id와 매칭
        private void StampEbValidateData()
        {
            // 2025.09.16 Added: 필드별 사전 생성 키는 마지막 토큰 사용
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

            // 2025.09.16 Added: 뷰에서 바로 쓰거나 스크립트에서 읽을 수 있도록 이중 형태로 제공
            ViewData["EBV_FieldErrors"] = fieldErrors;
            ViewData["EBV_FieldErrorsJson"] = JsonSerializer.Serialize(fieldErrors);
            ViewData["EBV_SummaryErrors"] = summaryErrors;
            ViewData["EBV_SummaryErrorsJson"] = JsonSerializer.Serialize(summaryErrors);
        }

        // 2025.09.16 Added: a.b.c 키가 올 경우 마지막 토큰을 id로 사용
        private static string GetFieldId(string key)
        {
            var idx = key.LastIndexOf('.');
            return idx >= 0 ? key.Substring(idx + 1) : key;
        }
    }
}
