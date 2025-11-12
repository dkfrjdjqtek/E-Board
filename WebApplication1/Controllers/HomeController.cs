using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;
using WebApplication1.Models;
using Microsoft.Extensions.Localization;

namespace WebApplication1.Controllers
{
    public class HomeController : Controller
    {
        private readonly IStringLocalizer<SharedResource> _S;

        public HomeController(
       ILogger<HomeController> logger,
       IStringLocalizer<SharedResource> s)
        {
            _logger = logger;
            _S = s;
        }


        private readonly ILogger<HomeController> _logger;

        
        public IActionResult Index()
        {
            //ViewBag.PageTitle = _S["Home_Title"];
            //return View();
            if (User?.Identity?.IsAuthenticated == true)
                return Redirect("/Doc/Board");
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
