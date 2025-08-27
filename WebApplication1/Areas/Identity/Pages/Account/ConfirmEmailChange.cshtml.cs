using System.Text;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.WebUtilities;
using WebApplication1.Models;

namespace WebApplication1.Areas.Identity.Pages.Account;

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

    public async Task<IActionResult> OnGetAsync(string userId, string email, string code)
    {
        if (userId is null || email is null || code is null)
        {
            TempData["StatusMessage"] = "잘못된 요청입니다.";
            return RedirectToPage("/Account/ChangeProfile");
        }

        var user = await _userManager.FindByIdAsync(userId);
        if (user is null)
        {
            TempData["StatusMessage"] = "사용자를 찾을 수 없습니다.";
            return RedirectToPage("/Account/ChangeProfile");
        }

        // 토큰 디코드
        var decoded = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(code));

        // 이메일 변경 확정
        var result = await _userManager.ChangeEmailAsync(user, email, decoded);
        if (!result.Succeeded)
        {
            TempData["StatusMessage"] = "이메일 변경 링크가 유효하지 않거나 만료되었습니다.";
            return RedirectToPage("/Account/ChangeProfile");
        }

        // (선택) UserName도 이메일로 맞출 경우
        // await _userManager.SetUserNameAsync(user, email);

        // 로그인 갱신
        await _signInManager.RefreshSignInAsync(user);

        TempData["StatusMessage"] = "이메일이 변경되었습니다.";
        return RedirectToPage("/Account/ChangeProfile");
    }
}
