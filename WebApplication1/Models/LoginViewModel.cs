// Models/LoginViewModel.cs
using System.ComponentModel.DataAnnotations;

namespace WebApplication1.Models
{
    public class LoginViewModel
    {
        [Required(ErrorMessage = "Login_UserName_Required")]
        [Display(Name = "Login_UserName_Label")]
        public string UserName { get; set; } = string.Empty;   // ← 컨트롤러가 model.UserName 사용

        [Required(ErrorMessage = "Login_Password_Required")]
        [DataType(DataType.Password)]
        [Display(Name = "Login_Password_Label")]
        public string Password { get; set; } = string.Empty;

        [Display(Name = "Login_RememberMe_Label")]
        public bool RememberMe { get; set; }
    }
}
