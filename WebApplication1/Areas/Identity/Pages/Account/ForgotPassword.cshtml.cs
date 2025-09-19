// File: Areas/Identity/Pages/Account/ForgotPassword.cshtml.cs
// 2025.09.19 Changed: UserProfiles.DisplayName 조회로 실제 이름 사용(없으면 UserName→Email); 최소 패치로 DI/using/쿼리 추가
using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Options;
using Microsoft.EntityFrameworkCore;                 // 2025.09.19 Added
using WebApplication1.Models;
using WebApplication1.Data;                          // 2025.09.19 Added

namespace WebApplication1.Areas.Identity.Pages.Account
{
    [AllowAnonymous]
    public class ForgotPasswordModel : PageModel
    {
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly IEmailSender _email;
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly TimeSpan _tokenLifespan;
        private readonly ApplicationDbContext _db;    // 2025.09.19 Added

        // 2025.09.19 Changed: ApplicationDbContext 주입
        public ForgotPasswordModel(
            UserManager<ApplicationUser> userManager,
            IEmailSender email,
            IStringLocalizer<SharedResource> S,
            IOptions<DataProtectionTokenProviderOptions> tokenOptions,
            ApplicationDbContext db)
        {
            _userManager = userManager;
            _email = email;
            _S = S;
            _tokenLifespan = tokenOptions.Value.TokenLifespan;
            _db = db;                                  // 2025.09.19 Added
        }

        [BindProperty]
        public InputModel Input { get; set; } = new();

        public class InputModel
        {
            [System.ComponentModel.DataAnnotations.Display(Name = "_CM_Label_ID")]
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

            if (user is null ||
                !string.Equals(user.Email, Input.Email.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                return RedirectToPage("./ForgotPasswordConfirmation");
            }

            if (!user.EmailConfirmed)
                return RedirectToPage("./ForgotPasswordConfirmation");

            var subject = _S["FP_Email_Subject"];

            // 2025.09.19 Changed: UserProfiles.DisplayName(실제 이름) 우선 사용
            string displayName = user.Email!;
            var profName = await _db.UserProfiles
                                    .AsNoTracking()
                                    .Where(p => p.UserId == user.Id)
                                    .Select(p => p.DisplayName)
                                    .FirstOrDefaultAsync();
            if (!string.IsNullOrWhiteSpace(profName)) displayName = profName;
            else if (!string.IsNullOrWhiteSpace(user.UserName)) displayName = user.UserName!;

            // 링크 토큰 생성 및 인코딩
            var token = await _userManager.GeneratePasswordResetTokenAsync(user);
            var encoded = WebEncoders.Base64UrlEncode(Encoding.UTF8.GetBytes(token));
            var callbackUrl = Url.Page(
                "/Account/ResetPassword",
                null,
                new { area = "Identity", code = encoded, userId = user.Id },
                Request.Scheme)!;

            var minutes = ((int)Math.Round(_tokenLifespan.TotalMinutes)).ToString(CultureInfo.InvariantCulture);

            var body = string.Format(
                _S["FP_Email_BodyHtml"],
                HtmlEncoder.Default.Encode(displayName),
                callbackUrl,
                minutes
            );

            await _email.SendEmailAsync(user.Email!, subject, body);
            return RedirectToPage("./ForgotPasswordConfirmation");
        }
    }
}
