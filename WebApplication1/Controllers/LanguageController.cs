using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Localization;
using Microsoft.AspNetCore.Mvc;

namespace WebApplication1.Controllers
{
    [AllowAnonymous] // 컨트롤러 전체 익명 허용
    public class LanguageController : Controller
    {
        [HttpGet]                    // GET만으로 충분
        [IgnoreAntiforgeryToken]     // (GET이면 사실 필요 없음)
        public IActionResult Set(string culture, string? returnUrl = "/")
        {
            var targetCulture = string.IsNullOrWhiteSpace(culture) ? "ko-KR" : culture;
            var targetReturnUrl = string.IsNullOrWhiteSpace(returnUrl) ? "/" : returnUrl;

            var cookie = CookieRequestCultureProvider
                         .MakeCookieValue(new RequestCulture(targetCulture));

            Response.Cookies.Append(
                CookieRequestCultureProvider.DefaultCookieName,
                cookie,
                new CookieOptions
                {
                    Path = "/",
                    IsEssential = true,
                    Expires = DateTimeOffset.UtcNow.AddYears(1),
                    SameSite = SameSiteMode.Lax,
                    Secure = HttpContext.Request.IsHttps
                });

            if (!Url.IsLocalUrl(targetReturnUrl)) targetReturnUrl = "/";
            return LocalRedirect(targetReturnUrl);
        }
    }
}
