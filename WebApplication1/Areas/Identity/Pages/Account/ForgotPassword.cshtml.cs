using System.Text;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Localization;
using WebApplication1.Models; // ApplicationUser

namespace WebApplication1.Areas.Identity.Pages.Account
{
    [AllowAnonymous]
    public class ForgotPasswordModel : PageModel
    {
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly IEmailSender _email;
        private readonly IStringLocalizer<SharedResource> _S;

        public ForgotPasswordModel(
            UserManager<ApplicationUser> userManager,
            IEmailSender email,
            IStringLocalizer<SharedResource> S)
        {
            _userManager = userManager;
            _email = email;
            _S = S;
        }

        [BindProperty]
        public InputModel Input { get; set; } = new();

        public class InputModel
        {
            // ▼ DataAnnotations의 Name/ErrorMessage는 “값=리소스키”로 씁니다.
            [System.ComponentModel.DataAnnotations.Display(Name = "FP_UserName_Label")]
            [System.ComponentModel.DataAnnotations.Required(ErrorMessage = "FP_UserName_Required")]
            public string UserName { get; set; } = string.Empty;

            [System.ComponentModel.DataAnnotations.Display(Name = "FP_Email_Label")]
            [System.ComponentModel.DataAnnotations.Required(ErrorMessage = "FP_Email_Required")]
            [System.ComponentModel.DataAnnotations.EmailAddress(ErrorMessage = "FP_Email_Invalid")]
            public string Email { get; set; } = string.Empty;
        }

        public void OnGet() { }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid) return Page();

            var user = await _userManager.FindByNameAsync(Input.UserName.Trim());

            // 계정 존재/불일치 여부를 노출하지 않기 위해 항상 동일한 결과로
            if (user is null ||
                !string.Equals(user.Email, Input.Email.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return RedirectToPage("./ForgotPasswordConfirmation");
            }

            if (!user.EmailConfirmed)
                return RedirectToPage("./ForgotPasswordConfirmation");

            var subject = _S["FP_Email_Subject"];

            // 표시 이름(없으면 Email 사용)
            var displayName = string.IsNullOrWhiteSpace(user.UserName) ? user.Email! : user.UserName;

            // 링크 토큰
            var token = await _userManager.GeneratePasswordResetTokenAsync(user);
            var encoded = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(token));
            var callbackUrl = Url.Page(
                "/Account/ResetPassword",
                null,
                new { area = "Identity", code = encoded, userId = user.Id },
                Request.Scheme)!;

            // 자리표시자 {0}={displayName}, {1}={callbackUrl}
            var body = string.Format(
                _S["FP_Email_BodyHtml"],
                HtmlEncoder.Default.Encode(displayName),
                HtmlEncoder.Default.Encode(callbackUrl)
            );

            await _email.SendEmailAsync(user.Email!, subject, body);

            return RedirectToPage("./ForgotPasswordConfirmation");
        }
    }
}
