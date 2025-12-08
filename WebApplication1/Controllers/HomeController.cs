// 2025.12.03 Changed: 로그인 사용자 Doc Board 리다이렉트 유지 및 홈 타이틀 리소스 다시 적용 기타 로직 변경 없음
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Localization;
using System.Diagnostics;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IStringLocalizer<SharedResource> _S;

        public HomeController(
            ILogger<HomeController> logger,
            IStringLocalizer<SharedResource> s)
        {
            _logger = logger;
            _S = s;
        }

        public IActionResult Index()
        {
            // 로그인 상태면 바로 전자결재 보드로 이동
            if (User?.Identity?.IsAuthenticated == true)
            {
                return RedirectToAction("Board", "Doc");
            }

            // 비로그인 홈 화면 타이틀 (리소스 사용)
            ViewBag.PageTitle = _S["Home_Title"];
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel
            {
                RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier
            });
        }
    }
}
