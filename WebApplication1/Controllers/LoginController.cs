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
using WebApplication1.Models; // ApplicationUser, LoginViewModel, RegisterViewModel
using Microsoft.Extensions.Localization;
using System.Linq; // 2025.09.16 Added: Linq 사용
using System.Collections.Generic; // 2025.09.16 Added: 사전과 목록 사용
using System.Text.Json; // 2025.09.16 Added: Json 직렬화

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

        [HttpGet]
        public IActionResult AccessDenied(string? returnUrl = null)
        {
            Response.StatusCode = 403;                 // 상태코드 403
            ViewData["ReturnUrl"] = returnUrl;
            return View();
        }

        [HttpGet]
        public IActionResult Login(string? returnUrl = null)
        {
            ViewData["ReturnUrl"] = returnUrl ?? Url.Content("~/");
            // 2025.09.16 Added: 초기 진입 시 에러 정보 초기화
            StampEbValidateData(); // 비어있는 상태로 전달
            return View(new LoginViewModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(LoginViewModel model, string? returnUrl = null)
        {
            returnUrl ??= Url.Content("~/");
            ViewData["ReturnUrl"] = returnUrl;

            // 2025.09.17 Added: 서버측 Required 가드(클라이언트 검증 누락/비활성 대비)
            if (string.IsNullOrWhiteSpace(model.UserName))
                AddErrorOnce(nameof(LoginViewModel.UserName), _S["Req"]);
            if (string.IsNullOrWhiteSpace(model.Password))
                AddErrorOnce(nameof(LoginViewModel.Password), _S["Req"]);

            if (!ModelState.IsValid)
            {
                // 2025.09.16 Changed: 클라이언트 EBValidate와 요약에 동일 데이터 전달
                StampEbValidateData();
                return View(model);
            }

            var loginText = (model.UserName ?? string.Empty).Trim();

            // 아이디 이메일 모두 허용
            ApplicationUser? user = loginText.Contains('@')
                ? await _userManager.FindByEmailAsync(loginText)
                : await _userManager.FindByNameAsync(loginText);

            if (user is null)
            {
                // 2025.09.16 Changed: 빈 키 대신 UserName 키로 에러 매핑
                AddErrorOnce(nameof(LoginViewModel.UserName), _S["Login_UserNotExist"]);
                StampEbValidateData(); // 2025.09.16 Added: ViewData에 필드별 에러와 요약 Stamp
                return View(model);
            }
            // 2025.09.17 Added: 호출 직전 재가드 비밀번호 공란이면 WrongPassword 대신 Required를 우선 노출
            if (string.IsNullOrWhiteSpace(model.Password))
            {
                AddErrorOnce(nameof(LoginViewModel.Password), _S["Req"]);
                StampEbValidateData();
                return View(model);
            }

            var result = await _signInManager.PasswordSignInAsync(
                user, model.Password, model.RememberMe, lockoutOnFailure: true);

            if (result.Succeeded)
                return LocalRedirect(returnUrl ?? Url.Content("~/"));

            if (result.RequiresTwoFactor)
            {
                var encodedReturnUrl = UrlEncoder.Default.Encode(returnUrl ?? "/");
                return Redirect($"/Identity/Account/LoginWith2fa?ReturnUrl={encodedReturnUrl}&RememberMe={model.RememberMe}");
            }

            if (result.IsLockedOut)
            {
                // 2025.09.16 Changed: 잠금은 UserName 키로 매핑
                AddErrorOnce(nameof(LoginViewModel.UserName), _S["Login_LockedOut"]);
                StampEbValidateData(); // 2025.09.16 Added
                return View(model);
            }

            if (result.IsNotAllowed)
            {
                // 2025.09.16 Changed: 허용 안됨도 UserName 키로 매핑
                AddErrorOnce(nameof(LoginViewModel.UserName), _S["Login_NotAllowed"]);
                StampEbValidateData(); // 2025.09.16 Added
                return View(model);
            }

            // 비밀번호 틀림
            // 2025.09.16 Changed: Password 키로 명확히 매핑
            AddErrorOnce(nameof(LoginViewModel.Password), _S["Login_WrongPassword"]);
            StampEbValidateData(); // 2025.09.16 Added
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
