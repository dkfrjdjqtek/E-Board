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

        // 2025.09.25 Added: ���ε� �� ���� Ŭ����
        public class InviteVM
        {
            public string? UserId { get; set; }
            public string? Email { get; set; }
            public string? Ec { get; set; } // �̸��� Ȯ�� ��ū base64url
            public string? Rp { get; set; } // ��й�ȣ �缳�� ��ū base64url

            public string? Password { get; set; }
            public string? ConfirmPassword { get; set; }
        }

        [BindProperty]
        public InviteVM VM { get; set; } = new();

        // 2025.09.25 Added: ��� �޽��� �ϰ�ȭ�� ���� ���� �ʵ�� ��࿡ ���ÿ� ���
        // 2025.09.26 Changed: ���� AddErr �� ���� ��ü Ű�θ� ���(�� Ű/���� Ű ��� ����)
        private void AddErr(string prop, string message)
        {
            // 2025.09.26 Changed: �׻� "VM.<prop>" ������ ��ü Ű�� ����Ͽ� ���(����� �����ӿ�ũ�� �ڵ� ����)
            if (!string.IsNullOrWhiteSpace(prop))
                AddErrorOnce($"VM.{prop}", message);
        }

        public IActionResult OnGet(string? userId, string? email, string? ec, string? rp)
        {
            VM.UserId = userId;
            VM.Email = email;
            VM.Ec = ec;
            VM.Rp = rp;

            // 2025.09.25 Added: �ʼ� �Ķ���� ����
            if (string.IsNullOrWhiteSpace(VM.UserId) ||
                string.IsNullOrWhiteSpace(VM.Email) ||
                string.IsNullOrWhiteSpace(VM.Ec) ||
                string.IsNullOrWhiteSpace(VM.Rp))
            {
                AddErr("", _S["IL_Error_InvalidLink"].Value);
            }

            return Page();
        }

        // 2025.09.26 Added: �ߺ� ���� ����(���� Ű�� ���� �޽��� 1ȸ�� ���)
        private void AddErrorOnce(string fieldKey, string message)
        {
            if (string.IsNullOrWhiteSpace(fieldKey) || string.IsNullOrWhiteSpace(message)) return; // 2025.09.26 Added: �� Ű ����
            if (!ModelState.TryGetValue(fieldKey, out var entry) || !entry.Errors.Any(e => e.ErrorMessage == message))
                ModelState.AddModelError(fieldKey, message);
        }

        public async Task<IActionResult> OnPostAsync()
        {
            // 2025.09.26 Changed: �ʼ��� ������ ��ü Ű�� ���(�� Ű ����)
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
                // 2025.09.26 Changed: �Ϲ� ��ū ���ڵ� ���е� ��ü Ű���� ����(����� �ڵ�)
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