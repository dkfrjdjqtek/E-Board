using DevExpress.AspNetCore.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace WebApplication1.Controllers
{
    // DevExpress 내부 콜백은 인증/인가 미들웨어에서 분기 처리하지만,
    // 컨트롤러 레벨에서도 안전하게 AllowAnonymous로 둡니다(데모/검증용).
    [AllowAnonymous]
    public class SpreadsheetController : Controller
    {
        [HttpGet]
        [HttpPost]
        public IActionResult DxSpreadsheetRequest()
        {
            return SpreadsheetRequestProcessor.GetResponse(HttpContext);
        }
    }
}