// Models/LoginViewModel.cs
using System.ComponentModel.DataAnnotations;

namespace YourApp.Models;

public class LoginViewModel
{
    [Required(ErrorMessage = "아이디/이메일을 입력하세요.")]
    [Display(Name = "아이디 또는 이메일")]
    public string Login { get; set; } = "";

    [Required(ErrorMessage = "비밀번호를 입력하세요.")]
    [DataType(DataType.Password)]
    [Display(Name = "비밀번호")]
    public string Password { get; set; } = "";

    [Display(Name = "로그인 유지")]
    public bool RememberMe { get; set; }

    public string? ReturnUrl { get; set; }
}
