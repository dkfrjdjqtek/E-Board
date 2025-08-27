using System.ComponentModel.DataAnnotations;
using System.Text.Encodings.Web;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using WebApplication1.Models;
using Microsoft.Extensions.Localization;
using WebApplication1;

[Authorize(Policy = "AdminOnly")]
public class AdminToolsController : Controller
{
    private readonly UserManager<ApplicationUser> _userManager;
    private readonly IEmailSender _emailSender;
    private readonly IStringLocalizer<SharedResource> _S;

    public AdminToolsController(UserManager<ApplicationUser> userManager,
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
            ModelState.AddModelError(string.Empty, _S["Admin_ResetLink_UserNotFound"]);
            return View(vm);
        }

        var token = await _userManager.GeneratePasswordResetTokenAsync(user);
        var url = Url.Page("/Account/ResetPassword", null,
            new { area = "Identity", userId = user.Id, code = token },
            Request.Scheme);

        if (vm.SendEmail)
        {
            var subject = _S["Admin_ResetLink_Email_Subject"];
            var body = string.Format(_S["Admin_ResetLink_Email_BodyHtml"],
                                     System.Text.Encodings.Web.HtmlEncoder.Default.Encode(url));
            await _emailSender.SendEmailAsync(user.Email, subject, body);
            TempData["Msg"] = _S["Admin_ResetLink_Email_Sent"];
        }

        TempData["Link"] = url;
        return RedirectToAction(nameof(SendResetLink));
    }

    public class AdminSendResetLinkVM
    {
        [Required(ErrorMessage = "Admin_ResetLink_User_Req")]                 // ← 키 사용
        [Display(Name = "Admin_ResetLink_User_Label")]                        // ← 키 사용
        public string UserNameOrEmail { get; set; } = "";

        [Display(Name = "Admin_ResetLink_SendEmail_Label")]                   // ← 키 사용
        public bool SendEmail { get; set; } = true;
    }
}