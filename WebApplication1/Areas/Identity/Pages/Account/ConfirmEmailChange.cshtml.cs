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
        // �Ķ���� ����
        if (string.IsNullOrWhiteSpace(userId) ||
            string.IsNullOrWhiteSpace(email) ||
            string.IsNullOrWhiteSpace(code))
        {
            TempData["StatusMessage"] = "��û�� �ùٸ��� �ʽ��ϴ�.";
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

        // �̸��� ���� + Ȯ�� ó��
        var result = await _userManager.ChangeEmailAsync(user, email, decoded);
        if (!result.Succeeded)
        {
            TempData["StatusMessage"] = string.Join(" ",
                result.Errors.Select(e => e.Description));
            return RedirectToPage("/Account/ChangeProfile");
        }

        // (����) �̸����� ����ڸ�� �����ϰ� ���� ��å�̸� �Ʒ��� ����
        // await _userManager.SetUserNameAsync(user, email);

        await _signInManager.RefreshSignInAsync(user);
        TempData["StatusMessage"] = "�̸����� ����/Ȯ�εǾ����ϴ�.";
        return RedirectToPage("/Account/ChangeProfile");
    }
}
