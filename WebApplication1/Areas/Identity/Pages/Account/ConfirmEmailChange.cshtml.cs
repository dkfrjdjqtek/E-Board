using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using WebApplication1.Models;

namespace WebApplication1.Areas.Identity.Pages.Account;

[AllowAnonymous]
public class ConfirmEmailChangeModel : PageModel
{
    private readonly UserManager<ApplicationUser> _userManager;
    private readonly SignInManager<ApplicationUser> _signInManager;

    public ConfirmEmailChangeModel(
        UserManager<ApplicationUser> userManager,
        SignInManager<ApplicationUser> signInManager)
    {
        _userManager = userManager;
        _signInManager = signInManager;
    }

    public async Task<IActionResult> OnGetAsync(string? userId, string? email, string? code)
    {
        // 파라미터 검증
        if (string.IsNullOrWhiteSpace(userId) ||
            string.IsNullOrWhiteSpace(email) ||
            string.IsNullOrWhiteSpace(code))
        {
            TempData["StatusMessage"] = "요청이 올바르지 않습니다.";
            return RedirectToPage("/Account/ChangeProfile");
        }

        var user = await _userManager.FindByIdAsync(userId);
        if (user is null)
        {
            TempData["StatusMessage"] = "사용자를 찾을 수 없습니다.";
            return RedirectToPage("/Account/ChangeProfile");
        }

        // 토큰 디코딩
        var decoded = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(code));

        // 이메일 변경 + 확인 처리
        var result = await _userManager.ChangeEmailAsync(user, email, decoded);
        if (!result.Succeeded)
        {
            TempData["StatusMessage"] = string.Join(" ",
                result.Errors.Select(e => e.Description));
            return RedirectToPage("/Account/ChangeProfile");
        }

        // (선택) 이메일을 사용자명과 동일하게 쓰는 정책이면 아래도 수행
        // await _userManager.SetUserNameAsync(user, email);

        await _signInManager.RefreshSignInAsync(user);
        TempData["StatusMessage"] = "이메일이 변경/확인되었습니다.";
        return RedirectToPage("/Account/ChangeProfile");
    }
}
