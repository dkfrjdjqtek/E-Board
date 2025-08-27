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
            TempData["StatusMessage"] = "�߸��� ��û�Դϴ�.";
            return RedirectToPage("/Account/ChangeProfile");
        }

        var user = await _userManager.FindByIdAsync(userId);
        if (user is null)
        {
            TempData["StatusMessage"] = "����ڸ� ã�� �� �����ϴ�.";
            return RedirectToPage("/Account/ChangeProfile");
        }

        // ��ū ���ڵ�
        var decoded = Encoding.UTF8.GetString(WebEncoders.Base64UrlDecode(code));

        // �̸��� ���� Ȯ��
        var result = await _userManager.ChangeEmailAsync(user, email, decoded);
        if (!result.Succeeded)
        {
            TempData["StatusMessage"] = "�̸��� ���� ��ũ�� ��ȿ���� �ʰų� ����Ǿ����ϴ�.";
            return RedirectToPage("/Account/ChangeProfile");
        }

        // (����) UserName�� �̸��Ϸ� ���� ���
        // await _userManager.SetUserNameAsync(user, email);

        // �α��� ����
        await _signInManager.RefreshSignInAsync(user);

        TempData["StatusMessage"] = "�̸����� ����Ǿ����ϴ�.";
        return RedirectToPage("/Account/ChangeProfile");
    }
}
