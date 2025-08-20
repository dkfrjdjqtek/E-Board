using System.ComponentModel.DataAnnotations;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.CodeAnalysis.Elfie.Diagnostics;
using WebApplication1.Models;

namespace YourApp.Areas.Identity.Pages.Account
{
    [Authorize]
    public class ChangeProfileModel : PageModel
    {
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly SignInManager<ApplicationUser> _signInManager;

        public ChangeProfileModel(
            UserManager<ApplicationUser> userManager,
            SignInManager<ApplicationUser> signInManager)
        {
            _userManager = userManager;
            _signInManager = signInManager;
        }

        [BindProperty]
        //public InputModel Input { get; set; }
        public InputModel Input { get; set; } = new(); // 생성자 종료 시 초기화

        public class InputModel
        {
            [Display(Name = "이메일")]
            [EmailAddress]
            public string Email { get; set; } = string.Empty;

            [Display(Name = "사용자명")]
            public string UserName { get; set; } = string.Empty;

            [Display(Name = "부서")]
            public string Department { get; set; } = string.Empty;

            [Display(Name = "직급")]
            public string Position { get; set; } = string.Empty;

            [Display(Name = "현재 비밀번호")]
            [DataType(DataType.Password)]
            public string CurrentPassword { get; set; } = string.Empty;

            [Display(Name = "새 비밀번호")]
            [DataType(DataType.Password)]
            public string NewPassword { get; set; } = string.Empty;   // 기본값

            [Display(Name = "새 비밀번호 확인")]
            [Compare("NewPassword", ErrorMessage = "새 비밀번호가 일치하지 않습니다.")]
            [DataType(DataType.Password)]
            public string ConfirmNewPassword { get; set; } = string.Empty;   // 기본값
        }

        public async Task<IActionResult> OnGetAsync()
        {
            var user = await _userManager.GetUserAsync(User);
            if (user is null) return Challenge();              // CS8602 방지

            //Input = new InputModel
            //{
            //    Email = user.Email ?? string.Empty,             // CS8601 방지
            //    UserName = user.UserName,
            //    Department = "",// ""user.Claims.FirstOrDefault(c => c.Type == "Department")?.Value,
            //    Position = ""//user.Claims.FirstOrDefault(c => c.Type == "Position")?.Value
            //}; 
            //return Page();
            Input = new InputModel
            {
                Email = user.Email ?? string.Empty, // UserName이 null이면 Email, 그것도 없으면 Id로 대체(절대 null 아님)
                UserName = user.UserName ?? user.Email ?? user.Id,

                // 선택: 클레임에서 읽고 없으면 빈 문자열
                Department = User.FindFirst("Department")?.Value ?? string.Empty,
                Position = User.FindFirst("Position")?.Value ?? string.Empty,
            };
            return Page();
        }

        public async Task<IActionResult> OnPostAsync()
        {
            if (!ModelState.IsValid)
                return Page();

            var user = await _userManager.GetUserAsync(User);
            if (user is null) return Challenge();              // CS8602 방지

            // 사용자명 변경
            if (user.UserName != Input.UserName)
            {
                user.UserName = Input.UserName;
            }

               // 부서/직급 변경 (클레임 방식 예시)
            await _userManager.RemoveClaimAsync(user, new Claim("Department", ""));
            await _userManager.AddClaimAsync(user, new Claim("Department", Input.Department ?? ""));
            await _userManager.RemoveClaimAsync(user, new Claim("Position", ""));
            await _userManager.AddClaimAsync(user, new Claim("Position", Input.Position ?? ""));

            // 1) 이메일 변경
            if (user.Email != Input.Email)
            {
                var setEmail = await _userManager.SetEmailAsync(user, Input.Email);
                if (!setEmail.Succeeded)
                {
                    foreach (var e in setEmail.Errors)
                        ModelState.AddModelError(string.Empty, e.Description);
                    return Page();
                }
            }
            
            // 전체 프로필 업데이트
            var updateResult = await _userManager.UpdateAsync(user);
               if (!updateResult.Succeeded)
                   {
                       foreach (var e in updateResult.Errors)
                    ModelState.AddModelError(string.Empty, e.Description);
                       return Page();
                   }

            // 2) 비밀번호 변경
            if (!string.IsNullOrEmpty(Input.NewPassword))
            {
                var changePwd = await _userManager.ChangePasswordAsync(
                    user, Input.CurrentPassword, Input.NewPassword);
                if (!changePwd.Succeeded)
                {
                    foreach (var e in changePwd.Errors)
                        ModelState.AddModelError(string.Empty, e.Description);
                    return Page();
                }
                // 비밀번호 변경 후 자동 재로그인
                await _signInManager.RefreshSignInAsync(user);
            }

            TempData["StatusMessage"] = "프로필이 성공적으로 업데이트되었습니다.";
            return RedirectToPage();
        }
    }
}
