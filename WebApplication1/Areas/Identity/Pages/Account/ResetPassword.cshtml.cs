// File: Areas/Identity/Pages/Account/ResetPassword.cshtml.cs
// 2025.09.18 Changed: ������Ʈ�� code�� Base64Url ���ڵ��Ͽ� Input.Code�� ����(Invalid token ����); OnGet���� UserId ����
using System.ComponentModel.DataAnnotations;
using System.Text;                         // �� �߰�
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;   // �� �߰�
using WebApplication1.Models;
using Microsoft.Extensions.Localization;

namespace WebApplication1.Areas.Identity.Pages.Account
{
    public class ResetPasswordModel : PageModel
    {
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly IStringLocalizer<SharedResource> _S;

        public ResetPasswordModel(UserManager<ApplicationUser> userManager, IStringLocalizer<SharedResource> s)
        {
            _userManager = userManager;
            _S = s;
        }

        public class InputModel
        {
            public string? UserId { get; set; }
            public string? Code { get; set; }

            [EmailAddress(ErrorMessage = "RP_Email_Invalid")]
            [Display(Name = "RP_Email_Label")]
            [Required(ErrorMessage = "RP_Email_Req")]
            public string Email { get; set; } = "";

            [Display(Name = "RP_Pwd_Label")]
            [Required(ErrorMessage = "RP_Pwd_Req")]
            [StringLength(100, ErrorMessage = "RP_Pwd_Len", MinimumLength = 4)]
            [DataType(DataType.Password)]
            public string Password { get; set; } = "";

            [Display(Name = "RP_PwdConfirm_Label")]
            [Required(ErrorMessage = "RP_PwdConfirm_Req")]
            [DataType(DataType.Password)]
            [Compare(nameof(Password), ErrorMessage = "RP_PwdConfirm_NotMatch")]
            public string ConfirmPassword { get; set; } = "";

        }

        [BindProperty] public InputModel Input { get; set; } = new();

        public IActionResult OnGet(string userId, string code)
        {
            if (string.IsNullOrEmpty(userId) || string.IsNullOrEmpty(code))
            //return BadRequest("Invalid password reset token.");
            {
                ModelState.AddModelError(string.Empty, _S["_CM_InvalidToken"]); // �� �ٱ��� �޽���
                return Page();
            }

            // 2025.09.18 ����: ���� ��ũ�� code�� Base64Url ���ڵ��� ���ڿ� �� ���ڵ��ؼ� �����ؾ� ResetPasswordAsync ���� ���
            string decodedCode;
            try
            {
                decodedCode = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(code));
                Input.UserId = userId;
                Input.Code = decodedCode;
            }
            catch
            {
                //return BadRequest("Invalid password reset token.");
                ModelState.AddModelError(string.Empty, _S["_CM_InvalidToken"]);
            }            
            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid) return Page();
            if (string.IsNullOrEmpty(Input.Code))
            {
                ModelState.AddModelError(string.Empty, _S["_CM_InvalidToken"]); 
                return Page();
            }

            var user = await _userManager.FindByEmailAsync(Input.Email);
            user ??= await _userManager.FindByIdAsync(Input.UserId!);
            if (user is null)
                return RedirectToPage("./ResetPasswordConfirmation");

            var result = await _userManager.ResetPasswordAsync(user, Input.Code!, Input.Password);
            if (result.Succeeded)
                return RedirectToPage("./ResetPasswordConfirmation");

            foreach (var e in result.Errors)
            {
                // ��ū ������ ������ ���ҽ� �޽����� ġȯ
                if (string.Equals(e.Code, nameof(IdentityErrorDescriber.InvalidToken), StringComparison.OrdinalIgnoreCase))
                    ModelState.AddModelError(string.Empty, _S["_CM_InvalidToken"]);
                else
                    ModelState.AddModelError(string.Empty, e.Description);
            }
            return Page();
        }
    }
}
