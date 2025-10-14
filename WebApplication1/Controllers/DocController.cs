// 2025.10.14 Added: 서버측 필수값 검증 구현(디스크립터 기반)
// 2025.10.14 Added: 디스크립터 로더 헬퍼(실서비스 연동 전 임시 스텁) 및 DTO 추가
// 2025.10.14 Added: System.Text.Json 사용으로 디스크립터 파싱
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Localization;
using System.Linq;
using System.Text.Json;               // 2025.10.14 Added
using System.Text.Json.Serialization; // 2025.10.14 Added
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

        [HttpGet]
        public IActionResult New()
        {
            return View("Select");
        }

        [HttpGet]
        public IActionResult Create(string templateCode)
        {
            // 템플릿 코드 누락 방어(기존)
            if (string.IsNullOrWhiteSpace(templateCode))
            {
                TempData["NewDocAlert"] = "DOC_Val_TemplateRequired";
                return RedirectToAction(nameof(New));
            }

            var tplCode = templateCode;
            ViewBag.TemplateCode = tplCode;

            // TODO: 실제 저장소에서 tplCode 기준으로 메타/디스크립터/프리뷰 로드
            // --------------------------------------------------------------------
            // 아래 더미를 실제 구현으로 교체하십시오.
            // notFound 플래그는 예시입니다. 저장소 조회 결과에 따라 설정하세요.
            bool notFound = false; // ← TODO: 저장소 조회 후 true/false 결정
            string descriptorJson = "{}"; // ← TODO
            string previewJson = "{}"; // ← TODO
            string templateTitle = "";   // ← TODO
                                         // --------------------------------------------------------------------

            // 2025.10.14 Added: 템플릿 미존재 시 선택 화면으로 안전 리다이렉트
            if (notFound)
            {
                TempData["NewDocAlert"] = "DOC_Err_TemplateNotFound";
                return RedirectToAction(nameof(New));
            }

            ViewBag.DescriptorJson = descriptorJson;
            ViewBag.PreviewJson = previewJson;
            ViewBag.TemplateTitle = templateTitle;

            return View("Compose", new WebApplication1.Models.DocTLViewModel());
        }

        // 2025.10.14 Changed: 서버 검증 로직 추가(디스크립터 기반)
        [HttpPost]
        [ValidateAntiForgeryToken] // 2025.10.14 Added
        public IActionResult Create([FromBody] ComposePostDto dto)
        {
            if (dto is null)
            {
                ModelState.AddModelError("Payload", "DOC_Err_InvalidPayload");
            }
            else
            {
                if (string.IsNullOrWhiteSpace(dto.templateCode))
                    ModelState.AddModelError("TemplateCode", "DOC_Val_Required");

                var descriptor = LoadDescriptor(dto.templateCode ?? string.Empty);

                foreach (var f in descriptor.Inputs.Where(x => x.Required))
                {
                    var ok = dto.inputs != null
                             && dto.inputs.TryGetValue(f.Key ?? string.Empty, out var v)
                             && !string.IsNullOrWhiteSpace(v);
                    if (!ok)
                        ModelState.AddModelError($"Inputs[{f.Key}]", "DOC_Val_Required");
                }
                foreach (var ap in descriptor.Approvals.Where(x => x.Required))
                {
                    var ok = dto.approvals != null
                             && dto.approvals.TryGetValue(ap.RoleKey ?? string.Empty, out var v)
                             && !string.IsNullOrWhiteSpace(v);
                    if (!ok)
                        ModelState.AddModelError($"Approvals[{ap.RoleKey}]", "DOC_Val_ApproverRequired");
                }
            }

            if (!ModelState.IsValid)
            {
                var fieldErrors = ModelState
                    .Where(kv => kv.Value?.Errors?.Count > 0)
                    .ToDictionary(
                        kv => kv.Key,
                        kv => kv.Value!.Errors.Select(e => e.ErrorMessage).Distinct().ToArray()
                    );
                var summaryMsgs = fieldErrors.Values.SelectMany(v => v).Distinct().ToArray();

                return BadRequest(new { messages = summaryMsgs, fieldErrors });
            }

            var docId = "DOC_xxx"; // TODO: 실제 생성된 문서 코드
            var redirectUrl = Url.Action(nameof(Details), "Doc", new { id = docId });
            return Json(new { redirectUrl });
        }

        [HttpGet]
        public IActionResult Details(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return RedirectToAction(nameof(New));

            ViewBag.DocId = id;
            return View("Detail");
        }

        // ----------------- 내부 헬퍼/DTO -----------------

        // 2025.10.14 Added: 클라이언트 페이로드 DTO(키명 변경 금지)
        public class ComposePostDto
        {
            public string? templateCode { get; set; }
            public System.Collections.Generic.Dictionary<string, string>? inputs { get; set; }
            public System.Collections.Generic.Dictionary<string, string>? approvals { get; set; }
            public string? descriptorVersion { get; set; }
        }

        // 2025.10.14 Added: 디스크립터 구조(서버 검증용 최소 필드만)
        private record Descriptor(
            System.Collections.Generic.List<InputField> Inputs,
            System.Collections.Generic.List<ApprovalField> Approvals,
            string? Version
        )
        {
            public static Descriptor Empty => new(new(), new(), null);
        }

        private record InputField(
            [property: JsonPropertyName("key")] string? Key,
            [property: JsonPropertyName("type")] string? Type,
            [property: JsonPropertyName("required")] bool Required,
            [property: JsonPropertyName("a1")] string? A1
        );

        private record ApprovalField(
            [property: JsonPropertyName("roleKey")] string? RoleKey,
            [property: JsonPropertyName("approverType")] string? ApproverType,
            [property: JsonPropertyName("required")] bool Required,
            [property: JsonPropertyName("value")] string? Value
        );

        // 2025.10.14 Added: 디스크립터 로딩 스텁
        // 실제 구현 시 템플릿 저장소(예: DB/파일/캐시)에서 templateCode 기준으로 JSON을 가져오십시오.
        private Descriptor LoadDescriptor(string templateCode)
        {
            // 1) 실제 저장소 호출로 교체
            // var json = _templateStore.GetDescriptorJson(templateCode);

            // 2) 현재는 안전 폴백("{}")
            var json = "{}";

            if (string.IsNullOrWhiteSpace(json))
                return Descriptor.Empty;

            try
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                var inputs = new System.Collections.Generic.List<InputField>();
                if (root.TryGetProperty("inputs", out var inArr) && inArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var el in inArr.EnumerateArray())
                    {
                        inputs.Add(new InputField(
                            el.TryGetProperty("key", out var k) ? k.GetString() : null,
                            el.TryGetProperty("type", out var t) ? t.GetString() : null,
                            el.TryGetProperty("required", out var rq) && rq.GetBoolean(),
                            el.TryGetProperty("a1", out var a1) ? a1.GetString() : null
                        ));
                    }
                }

                var approvals = new System.Collections.Generic.List<ApprovalField>();
                if (root.TryGetProperty("approvals", out var apArr) && apArr.ValueKind == JsonValueKind.Array)
                {
                    foreach (var el in apArr.EnumerateArray())
                    {
                        approvals.Add(new ApprovalField(
                            el.TryGetProperty("roleKey", out var rk) ? rk.GetString() : null,
                            el.TryGetProperty("approverType", out var at) ? at.GetString() : null,
                            el.TryGetProperty("required", out var rq) && rq.GetBoolean(),
                            el.TryGetProperty("value", out var vv) ? vv.GetString() : null
                        ));
                    }
                }

                var version = root.TryGetProperty("version", out var ver) ? ver.GetString() : null;
                return new Descriptor(inputs, approvals, version);
            }
            catch
            {
                // 파싱 실패 시 빈 디스크립터로 검증을 건너뜀(클라이언트와 동일 결과)
                return Descriptor.Empty;
            }
        }
    }
}
