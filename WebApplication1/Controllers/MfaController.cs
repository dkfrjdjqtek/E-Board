using System;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using QRCoder;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    [Authorize]
    public class MfaController : Controller
    {
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly SignInManager<ApplicationUser> _signInManager;

        public MfaController(UserManager<ApplicationUser> userManager,
                             SignInManager<ApplicationUser> signInManager)
        {
            _userManager = userManager;
            _signInManager = signInManager;
        }

        // 등록(Setup) 화면: 앱에 계정 추가용 QR + 수동 키 표기
        [HttpGet]
        public async Task<IActionResult> Setup()
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null) return Challenge();

            // 사용자의 Authenticator Key 준비
            var key = await _userManager.GetAuthenticatorKeyAsync(user);
            if (string.IsNullOrWhiteSpace(key))
            {
                await _userManager.ResetAuthenticatorKeyAsync(user);
                key = await _userManager.GetAuthenticatorKeyAsync(user);
            }

            ViewBag.SharedKey = InsertSpaces(key);
            return View();
        }

        // QR PNG 생성 (현재 로그인 사용자 기준으로 otpauth URL을 내부에서 재계산)
        [HttpGet]
        public async Task<IActionResult> Qr()
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null) return Unauthorized();

            var key = await _userManager.GetAuthenticatorKeyAsync(user);
            if (string.IsNullOrWhiteSpace(key))
            {
                await _userManager.ResetAuthenticatorKeyAsync(user);
                key = await _userManager.GetAuthenticatorKeyAsync(user);
            }

            var issuer = "Han Young E-Board"; // 앱에 표시될 발급자명
            var account = await _userManager.GetEmailAsync(user) ?? user.UserName ?? user.Id;

            var otpauth =
                $"otpauth://totp/{Uri.EscapeDataString(issuer)}:{Uri.EscapeDataString(account)}" +
                $"?secret={key}&issuer={Uri.EscapeDataString(issuer)}&digits=6";

            using var gen = new QRCodeGenerator();
            using var data = gen.CreateQrCode(otpauth, QRCodeGenerator.ECCLevel.Q);
            using var png = new PngByteQRCode(data);
            var bytes = png.GetGraphic(5);
            return File(bytes, "image/png");
        }

        // 앱에서 생성된 6자리 코드 검증
        [HttpPost]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Verify(string code)
        {
            var user = await _userManager.GetUserAsync(User);
            if (user == null) return Unauthorized();

            code = (code ?? "").Replace(" ", "").Replace("-", "");

            var ok = await _userManager.VerifyTwoFactorTokenAsync(
                user, TokenOptions.DefaultAuthenticatorProvider, code);

            if (!ok)
            {
                TempData["Error"] = "코드가 올바르지 않습니다.";
                return RedirectToAction(nameof(Setup));
            }

            await _userManager.SetTwoFactorEnabledAsync(user, true);
            await _signInManager.RefreshSignInAsync(user);

            TempData["Msg"] = "2단계 인증이 활성화되었습니다.";
            return RedirectToAction(nameof(Setup));
        }

        //private static string InsertSpaces(string s) =>
        //    Regex.Replace(s ?? "", ".{4}", "$0 ").Trim();
        private static string InsertSpaces(string? input, int group = 4)
        {
            if (string.IsNullOrWhiteSpace(input))
                return string.Empty;

            var s = input.Replace(" ", "").Trim();      // 기존 공백 제거
            var sb = new StringBuilder(s.Length + s.Length / group);

            for (int i = 0; i < s.Length; i++)
            {
                if (i > 0 && i % group == 0)
                    sb.Append(' ');
                sb.Append(s[i]);
            }
            return sb.ToString().ToUpperInvariant();    // 보통 대문자로 표기
        }
    }
}
