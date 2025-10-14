// 2025.10.14 Added: 템플릿 선택 화면(Select.cshtml)에서 사용하는 목록 API 신설
// - 경로: GET /DocumentTemplates/list?active=1
// - 응답 스키마: [{ code, title, description }]

using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Linq;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("[controller]")]
    public class DocumentTemplatesController : Controller
    {
        // 예: /DocumentTemplates/list?active=1&site=...&dept=...&category=...
        [HttpGet("list")]
        public IActionResult List([FromQuery] int? active = null,
                                  [FromQuery] string? site = null,
                                  [FromQuery] string? dept = null,
                                  [FromQuery] string? category = null)
        {
            // TODO: 실제 저장소/서비스 연동
            // var templates = _templateService.Search(new TemplateQuery { Active = active, Site = site, Dept = dept, Category = category });
            // var items = templates.Select(t => new { code = t.Code, title = t.Title, description = t.Description });

            var items = Enumerable.Empty<object>(); // 임시 더미(빈 목록)

            return Json(items);
        }
    }
}
