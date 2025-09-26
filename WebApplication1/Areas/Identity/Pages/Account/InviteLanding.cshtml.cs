using System;
using System.Linq; // 2025.09.26 Added
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Localization;
using WebApplication1.Models;

namespace WebApplication1.Areas.Identity.Pages.Account
{
    [AllowAnonymous]
    public class InviteLandingModel : PageModel
    {
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly IStringLocalizer<SharedResource> _S;

        public InviteLandingModel(
            UserManager<ApplicationUser> userManager,
            IStringLocalizer<SharedResource> S)
        {
            _userManager = userManager;
            _S = S;
        }

        // 2025.09.25 Added: 바인딩 모델 단일 클래스
        public class InviteVM
        {
            public string? UserId { get; set; }
            public string? Email { get; set; }
            public string? Ec { get; set; } // 이메일 확인 토큰 base64url
            public string? Rp { get; set; } // 비밀번호 재설정 토큰 base64url

            public string? Password { get; set; }
            public string? ConfirmPassword { get; set; }
        }

        [BindProperty]
        public InviteVM VM { get; set; } = new();

        // 2025.09.25 Added: 요약 메시지 일관화를 위한 헬퍼 필드와 요약에 동시에 등록
        // 2025.09.26 Changed: 기존 AddErr → 단일 구체 키로만 등록(빈 키/이중 키 등록 제거)
        private void AddErr(string prop, string message)
        {
            // 2025.09.26 Changed: 항상 "VM.<prop>" 형태의 구체 키만 사용하여 등록(요약은 프레임워크가 자동 집계)
            if (!string.IsNullOrWhiteSpace(prop))
                AddErrorOnce($"VM.{prop}", message);
        }

        public IActionResult OnGet(string? userId, string? email, string? ec, string? rp)
        {
            VM.UserId = userId;
            VM.Email = email;
            VM.Ec = ec;
            VM.Rp = rp;

            // 2025.09.25 Added: 필수 파라미터 검증
            if (string.IsNullOrWhiteSpace(VM.UserId) ||
                string.IsNullOrWhiteSpace(VM.Email) ||
                string.IsNullOrWhiteSpace(VM.Ec) ||
                string.IsNullOrWhiteSpace(VM.Rp))
            {
                AddErr("", _S["IL_Error_InvalidLink"].Value);
            }

            return Page();
        }

        // 2025.09.26 Added: 중복 방지 헬퍼(동일 키에 동일 메시지 1회만 등록)
        private void AddErrorOnce(string fieldKey, string message)
        {
            if (string.IsNullOrWhiteSpace(fieldKey) || string.IsNullOrWhiteSpace(message)) return; // 2025.09.26 Added: 빈 키 금지
            if (!ModelState.TryGetValue(fieldKey, out var entry) || !entry.Errors.Any(e => e.ErrorMessage == message))
                ModelState.AddModelError(fieldKey, message);
        }

        public async Task<IActionResult> OnPostAsync()
        {
            // 2025.09.26 Changed: 필수값 검증도 구체 키만 사용(빈 키 제거)
            if (string.IsNullOrWhiteSpace(VM.UserId) ||
                string.IsNullOrWhiteSpace(VM.Email) ||
                string.IsNullOrWhiteSpace(VM.Ec) ||
                string.IsNullOrWhiteSpace(VM.Rp))
            {
                AddErr(nameof(VM.Email), _S["IL_Error_InvalidLink"].Value);
                return Page();
            }

            if (string.IsNullOrWhiteSpace(VM.Password))
                AddErr(nameof(VM.Password), _S["_Alert_Require_NewPassword"].Value);
            if (string.IsNullOrWhiteSpace(VM.ConfirmPassword))
                AddErr(nameof(VM.ConfirmPassword), _S["_Alert_Require_ConfirmPassword"].Value);
            if (!string.IsNullOrWhiteSpace(VM.Password) &&
                !string.IsNullOrWhiteSpace(VM.ConfirmPassword) &&
                VM.Password != VM.ConfirmPassword)
            {
                AddErr(nameof(VM.ConfirmPassword), _S["IL_Confirm_Mismatch"].Value);
            }

            if (!ModelState.IsValid)
                return Page();

            var user = await _userManager.FindByIdAsync(VM.UserId!);
            if (user is null)
            {
                AddErr(nameof(VM.Email), _S["IL_Error_UserNotFound"].Value);
                return Page();
            }

            if (!string.Equals(user.Email, VM.Email, StringComparison.OrdinalIgnoreCase))
            {
                AddErr(nameof(VM.Email), _S["IL_Error_InvalidLink"].Value);
                return Page();
            }

            string? ecRaw = null;
            string? rpRaw = null;
            try
            {
                ecRaw = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(VM.Ec!));
                rpRaw = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(VM.Rp!));
            }
            catch
            {
                // 2025.09.26 Changed: 일반 토큰 디코드 실패도 구체 키에만 매핑(요약은 자동)
                AddErr(nameof(VM.Email), _S["IL_Error_InvalidLink"].Value);
                return Page();
            }

            if (!user.EmailConfirmed)
            {
                var ecResult = await _userManager.ConfirmEmailAsync(user, ecRaw!);
                if (!ecResult.Succeeded)
                {
                    var msg = ecResult.Errors.FirstOrDefault()?.Description ?? _S["IL_Error_Token"].Value;
                    AddErr(nameof(VM.Email), _S["IL_Error_ConfirmEmail", msg].Value);
                    return Page();
                }
            }

            var rpResult = await _userManager.ResetPasswordAsync(user, rpRaw!, VM.Password!);
            if (!rpResult.Succeeded)
            {
                var msgs = rpResult.Errors.Select(e => e.Description).ToArray();
                if (msgs.Length == 0) msgs = new[] { _S["IL_Error_Token"].Value };
                foreach (var m in msgs)
                    AddErr(nameof(VM.Password), m);
                return Page();
            }

            await _userManager.SetLockoutEndDateAsync(user, null);
            await _userManager.UpdateSecurityStampAsync(user);

            TempData["StatusMessage"] = _S["IL_Completed"].Value;
            return RedirectToPage("/Account/Login", new { area = "Identity" });
        }
    }
}