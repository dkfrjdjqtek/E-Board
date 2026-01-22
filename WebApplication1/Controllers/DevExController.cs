using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("DevX")]
    public class DevExController : Controller
    {
        [HttpGet("Spreadsheet")]
        public IActionResult Spreadsheet()
        {
            return View("Spreadsheet");
        }
    }
}
