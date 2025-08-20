using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using System.Text.Encodings.Web;
using System.Threading.Tasks;
using WebApplication1.Models; // ApplicationUser, LoginViewModel

namespace WebApplication1.Controllers
{
    [AllowAnonymous]
    public class AccountController : Controller
    {
        private readonly SignInManager<ApplicationUser> _signInManager;
        private readonly UserManager<ApplicationUser> _userManager;

        public AccountController(SignInManager<ApplicationUser> signInManager,
                                 UserManager<ApplicationUser> userManager)
        {
            _signInManager = signInManager;
            _userManager = userManager;
        }
        [HttpGet]
        public IActionResult AccessDenied(string? returnUrl = null)
        {
            Response.StatusCode = 403;                 // 상태코드 403
            ViewData["ReturnUrl"] = returnUrl;
            return View();
        }

        [HttpGet]
        public IActionResult Login([FromQuery] string? returnUrl = null)
        {
            ViewData["ReturnUrl"] = returnUrl ?? Url.Content("~/");
            return View(new LoginViewModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Login(LoginViewModel model, string? returnUrl = null)
        {
            returnUrl ??= Url.Content("~/");
            ViewData["ReturnUrl"] = returnUrl;

            if (!ModelState.IsValid) return View(model);

            var loginText = (model.UserName ?? string.Empty).Trim();

            // 아이디/이메일 모두 허용
            ApplicationUser? user = loginText.Contains('@')
                ? await _userManager.FindByEmailAsync(loginText)
                : await _userManager.FindByNameAsync(loginText);

            if (user is null)
            {
                ModelState.AddModelError(nameof(LoginViewModel.UserName), "존재하지 않는 사용자입니다.");
                return View(model); // 같은 URL에서 표시
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
                ModelState.AddModelError(string.Empty, "비밀번호 오류 누적으로 잠금 상태입니다. 잠시 후 다시 시도하세요.");
                return View(model);
            }

            if (result.IsNotAllowed)
            {
                ModelState.AddModelError(string.Empty, "로그인이 허용되지 않았습니다. (이메일 확인 또는 2단계 인증 필요)");
                return View(model);
            }

            // 비밀번호 틀림
            ModelState.AddModelError(nameof(LoginViewModel.Password), "비밀번호가 올바르지 않습니다.");
            return View(model);
        }

        [HttpGet]
        public IActionResult Register() => View(new RegisterViewModel());

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Register(RegisterViewModel model)
        {
            if (!ModelState.IsValid) return View(model);

            var user = new ApplicationUser { UserName = model.UserName, Email = model.Email };
            var result = await _userManager.CreateAsync(user, model.Password);
            if (result.Succeeded) return RedirectToAction("Login");

            foreach (var e in result.Errors)
                ModelState.AddModelError(string.Empty, e.Description);
            return View(model);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Logout()
        {
            await _signInManager.SignOutAsync();
            return RedirectToAction("Index", "Home");
        }
    }
}
