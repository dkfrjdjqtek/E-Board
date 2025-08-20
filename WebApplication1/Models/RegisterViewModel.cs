using System.ComponentModel.DataAnnotations;

namespace WebApplication1.Models
{
    public class RegisterViewModel
    {

        [Required]
        [Display(Name = "사용자명")]
        public string UserName { get; set; } = string.Empty;

        [Required]
        [EmailAddress]
        [Display(Name = "이메일")]
        public string Email { get; set; } = string.Empty;

        [Required]
        [DataType(DataType.Password)]
        [Display(Name = "비밀번호")]
        public string Password { get; set; } = string.Empty;

        [DataType(DataType.Password)]
        [Display(Name = "비밀번호 확인")]
        [Compare("Password", ErrorMessage = "비밀번호가 일치하지 않습니다.")]
        public string ConfirmPassword { get; set; } = string.Empty;

    }
}
