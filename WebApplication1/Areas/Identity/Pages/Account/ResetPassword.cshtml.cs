using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using WebApplication1.Models;

namespace WebApplication1.Areas.Identity.Pages.Account
{
    public class ResetPasswordModel : PageModel
    {
        private readonly UserManager<ApplicationUser> _userManager;

        public ResetPasswordModel(UserManager<ApplicationUser> userManager)
        {
            _userManager = userManager;
        }

        public class InputModel
        {
            public string? UserId { get; set; }
            public string? Code { get; set; }

            [EmailAddress(ErrorMessage = "RP_Email_Invalid")]   // ¡ç Ãß°¡
            [Display(Name = "RP_Email_Label")]
            [Required(ErrorMessage = "RP_Email_Req")]
            public string Email { get; set; } = "";

            [Display(Name = "RP_Pwd_Label")]
            [Required(ErrorMessage = "RP_Pwd_Req")]
            [StringLength(100, ErrorMessage = "RP_Pwd_Len", MinimumLength = 4)]
            [DataType(DataType.Password)]
            public string Password { get; set; } = "";

            [Display(Name = "RP_PwdConfirm_Label")]
            [DataType(DataType.Password)]
            [Compare(nameof(Password), ErrorMessage = "RP_PwdConfirm_NotMatch")]
            public string ConfirmPassword { get; set; } = "";
        }

        [BindProperty] public InputModel Input { get; set; } = new();

        public IActionResult OnGet(string userId, string code)
        {
            if (string.IsNullOrEmpty(userId) || string.IsNullOrEmpty(code))
                return BadRequest("Invalid password reset token.");

            Input.UserId = userId;
            Input.Code = code;
            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid) return Page();
            if (string.IsNullOrEmpty(Input.Code)) return BadRequest("Invalid token.");

            var user = await _userManager.FindByEmailAsync(Input.Email);
            user ??= await _userManager.FindByIdAsync(Input.UserId!);

            if (user is null)
                return RedirectToPage("./ResetPasswordConfirmation");

            var result = await _userManager.ResetPasswordAsync(user, Input.Code!, Input.Password);
            if (result.Succeeded)
                return RedirectToPage("./ResetPasswordConfirmation");

            foreach (var e in result.Errors)
                ModelState.AddModelError(string.Empty, e.Description);

            return Page();
        }
    }
}
