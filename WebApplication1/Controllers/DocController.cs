// 2025.10.14 Added: 템플릿 기반 문서 작성 폼용 컨트롤러 스켈레톤 추가 Compose 뷰 연결 및 Create 저장 엔드포인트 기본 구조
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Localization;
using System.Text.Json;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    [Authorize]
    public class DocController : Controller
    {
        private readonly IStringLocalizer<SharedResource> _S;

        public DocController(IStringLocalizer<SharedResource> S)
        {
            _S = S;
        }

        // GET: /Doc/Create?templateCode=TPL_xxx
        // 뷰 파일: Views/DocTL/Compose.cshtml  ← cshtml은 View, 컨트롤러는 cs 파일이 맞습니다
        [HttpGet]
        public IActionResult Create(string templateCode)
        {
            // TODO: templateCode로 템플릿 메타 로드 및 디스크립터/프리뷰 JSON 채우기
            // 기존에 사용 중인 DocTLViewModel을 그대로 사용합니다. 새 키나 변수를 임의 생성하지 않습니다.
            var vm = new DocTLViewModel
            {
                TemplateCode = templateCode ?? string.Empty,
                TemplateTitle = null,                 // TODO: 리소스 키 기반 제목 매핑
                DescriptorJson = "{}",               // TODO: 저장된 디스크립터 JSON
                PreviewJson = "{}"                   // TODO: 저장된 프리뷰 JSON
            };
            return View("~/Views/DocTL/DocCompose.cshtml", vm);
        }

        // POST: /Doc/Create
        // Compose.cshtml에서 JSON으로 호출됨 fetch('/Doc/Create', POST)
        [HttpPost]
        public IActionResult Create([FromBody] ComposePostDto dto)
        {
            // 서버 유효성 검증: EB-VALIDATE 표준에 맞춰 ModelState 필드 키로 추가
            if (dto is null)
            {
                ModelState.AddModelError("Payload", "DOC_Err_InvalidPayload");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(dto.templateCode))
                    ModelState.AddModelError("TemplateCode", "DOC_Val_Required");

                // 디스크립터에 지정된 필수 입력값이 있다면 동일 키로 검증 추가
                // TODO: 디스크립터 로드 후 필수 필드 반복 검증
            }

            if (!ModelState.IsValid)
            {
                // EB-VALIDATE 규격: 요약 메시지와 필드별 오류 모두 반환
                var fieldErrors = ModelState
                    .Where(kv => kv.Value?.Errors?.Count > 0)
                    .ToDictionary(
                        kv => kv.Key,
                        kv => kv.Value!.Errors.Select(e => e.ErrorMessage).Distinct().ToArray()
                    );

                var summaryMsgs = fieldErrors.Values.SelectMany(v => v).Distinct().ToArray();
                return BadRequest(new
                {
                    messages = summaryMsgs,   // Validation Summary 용
                    fieldErrors               // 각 필드 Valid State 용
                });
            }

            // TODO: 저장 로직
            // - dto.templateCode로 템플릿 리비전 확인
            // - dto.inputs, dto.approvals 저장
            // - 문서 코드 생성 및 리비전 기록
            // - 저장 후 이동할 상세 보기 주소 생성

            var redirectUrl = Url.Action("Details", "Doc", new { id = "DOC_xxx" }); // TODO: 실제 문서 코드로 교체
            return Json(new { redirectUrl });
        }

        // 요청 페이로드 DTO: Compose.cshtml에서 전송하는 구조와 동일 키 사용
        public class ComposePostDto
        {
            public string? templateCode { get; set; }
            public Dictionary<string, string>? inputs { get; set; }
            public Dictionary<string, string>? approvals { get; set; }
            public string? descriptorVersion { get; set; }
        }
    }
}
