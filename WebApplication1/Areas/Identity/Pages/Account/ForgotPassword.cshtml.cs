using System.ComponentModel.DataAnnotations;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using WebApplication1.Models;

public class ForgotPasswordModel : PageModel
{
    private readonly UserManager<ApplicationUser> _userManager;
    private readonly IEmailSender _emailSender;

    public ForgotPasswordModel(UserManager<ApplicationUser> userManager,
                               IEmailSender emailSender)
    {
        _userManager = userManager;
        _emailSender = emailSender;
    }

    public class InputModel
    {
        [Display(Name = "FP_User_Label")]
        [Required(ErrorMessage = "FP_User_Req")]
        public string UserNameOrEmail { get; set; } = "";
    }

    [BindProperty] public InputModel Input { get; set; } = new();

    public void OnGet() { }

    public async Task<IActionResult> OnPostAsync()
    {
        if (!ModelState.IsValid) return Page();

        // 아이디 또는 이메일로 사용자 찾기 (계정 존재 노출 방지)
        var user = await _userManager.FindByEmailAsync(Input.UserNameOrEmail)
                   ?? await _userManager.FindByNameAsync(Input.UserNameOrEmail);

        if (user != null && !string.IsNullOrEmpty(user.Email))
        {
            var token = await _userManager.GeneratePasswordResetTokenAsync(user);
            var callbackUrl = Url.Page(
                "/Account/ResetPassword",
                pageHandler: null,
                values: new { area = "Identity", userId = user.Id, code = token },
                protocol: Request.Scheme);

            var subject = "Reset your password";
            var body = $"Click <a href='{HtmlEncoder.Default.Encode(callbackUrl)}'>here</a> to reset your password.";
            await _emailSender.SendEmailAsync(user.Email, subject, body);
        }

        // 존재 여부와 관계없이 같은 확인 페이지로
        return RedirectToPage("./ForgotPasswordConfirmation");
    }
}
