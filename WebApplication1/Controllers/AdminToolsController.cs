using System.ComponentModel.DataAnnotations;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Localization;
using WebApplication1;
using WebApplication1.Models;

[Authorize(Policy = "AdminOnly")]
public class AdminToolsController : Controller
{
    private readonly UserManager<ApplicationUser> _userManager;
    private readonly IEmailSender _emailSender;
    private readonly IStringLocalizer<SharedResource> _S;

    public AdminToolsController(
        UserManager<ApplicationUser> userManager,
        IEmailSender emailSender,
        IStringLocalizer<SharedResource> S)
    {
        _userManager = userManager;
        _emailSender = emailSender;
        _S = S;
    }

    [HttpGet]
    public IActionResult SendResetLink() => View(new AdminSendResetLinkVM());

    [HttpPost]
    [ValidateAntiForgeryToken]
    public async Task<IActionResult> SendResetLink(AdminSendResetLinkVM vm)
    {
        if (!ModelState.IsValid) return View(vm);

        var user = await _userManager.FindByEmailAsync(vm.UserNameOrEmail)
                   ?? await _userManager.FindByNameAsync(vm.UserNameOrEmail);

        if (user is null || string.IsNullOrEmpty(user.Email))
        {
            // LocalizedString를 TempData/ModelState에 직접 넣지 말고 .Value 사용
            ModelState.AddModelError(string.Empty, _S["Admin_ResetLink_UserNotFound"].Value);
            return View(vm);
        }

        var token = await _userManager.GeneratePasswordResetTokenAsync(user);

        // Url.Page가 null 반환 가능 → 널 가드
        var resetUrl = Url.Page("/Account/ResetPassword", pageHandler: null,
            values: new { area = "Identity", userId = user.Id, code = token },
            protocol: Request.Scheme);

        if (string.IsNullOrEmpty(resetUrl))
        {
            ModelState.AddModelError(string.Empty, _S["Admin_ResetLink_GenerateUrl_Failed"].Value);
            return View(vm);
        }

        if (vm.SendEmail)
        {
            var subject = _S["Admin_ResetLink_Email_Subject"].Value;

            // 본문 리소스는 {0} 자리표시자 유무와 무관하게 안전 (없으면 그대로 출력됨)
            var bodyTemplate = _S["Admin_ResetLink_Email_BodyHtml"].Value;
            var safeUrl = HtmlEncoder.Default.Encode(resetUrl);
            var body = string.Format(bodyTemplate, safeUrl);

            await _emailSender.SendEmailAsync(user.Email, subject, body);
            TempData["Msg"] = _S["Admin_ResetLink_Email_Sent"].Value;
        }

        TempData["Link"] = resetUrl;
        return RedirectToAction(nameof(SendResetLink));
    }

    public class AdminSendResetLinkVM
    {
        [Required(ErrorMessage = "Admin_ResetLink_User_Req")]
        [Display(Name = "Admin_ResetLink_User_Label")]
        public string UserNameOrEmail { get; set; } = "";

        [Display(Name = "Admin_ResetLink_SendEmail_Label")]
        public bool SendEmail { get; set; } = true;
    }
}
