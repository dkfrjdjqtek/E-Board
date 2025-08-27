using System.ComponentModel.DataAnnotations;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using WebApplication1.Models;

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

        [EmailAddress]
        [Display(Name = "RP_Email_Label")]
        [Required(ErrorMessage = "RP_Email_Req")]
        public string Email { get; set; } = "";

        [Display(Name = "RP_Pwd_Label")]
        [Required(ErrorMessage = "RP_Pwd_Req")]
        [StringLength(100, ErrorMessage = "RP_Pwd_Len", MinimumLength = 6)]
        [DataType(DataType.Password)]
        public string Password { get; set; } = "";

        [Display(Name = "RP_PwdConfirm_Label")]
        [DataType(DataType.Password)]
        [Compare(nameof(Password), ErrorMessage = "RP_PwdConfirm_NotMatch")]
        public string ConfirmPassword { get; set; } = "";
    }


    [BindProperty] public InputModel Input { get; set; } = new();

    // 링크로부터 userId, code를 받아 숨김필드에 보관
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

        // 일반적으로 이메일로 사용자 확인
        var user = await _userManager.FindByEmailAsync(Input.Email);
        // (예외적으로) 이메일 변경 등으로 못 찾으면 링크의 userId로 시도
        user ??= await _userManager.FindByIdAsync(Input.UserId!);

        // 계정 존재 여부 노출 방지: 없으면 그냥 성공 페이지로
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
