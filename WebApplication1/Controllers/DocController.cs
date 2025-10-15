// 2025.10.15 Changed: New 액션에서 Model.CompOptions 및 DepartmentOptions를 ViewBag에서 안전 매핑하여 View로 전달 DocTL과 동일 서버사이드 렌더로 값 표시 유지 나머지 로직 변경 없음

using System;
using System.Linq;
using System.Data;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Security.Claims;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Localization;
using WebApplication1.Models;
using WebApplication1.Services;
using Microsoft.AspNetCore.Antiforgery;
using Microsoft.AspNetCore.Hosting;
using System.IO;
using Microsoft.AspNetCore.Mvc.Rendering; // 2025.10.15 Added

namespace WebApplication1.Controllers
{
    [Authorize]
    public class DocController : Controller
    {
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IAuditLogger _audit;
        private readonly IConfiguration _cfg;
        private readonly IAntiforgery _antiforgery;
        private readonly IWebHostEnvironment _env;

        public DocController(IStringLocalizer<SharedResource> S, IAuditLogger audit, IConfiguration cfg, IAntiforgery antiforgery, IWebHostEnvironment env)
        {
            _S = S;
            _audit = audit;
            _cfg = cfg;
            _antiforgery = antiforgery;
            _env = env; // 2025.10.15 Added
        }

        // 1) 템플릿 선택
        [HttpGet]
        public IActionResult New()
        {
            // 2025.10.15 Changed: DocTL과 동일하게 서버사이드로 콤보 옵션을 즉시 렌더할 수 있도록 뷰모델에 채워 전달
            var vm = new DocTLViewModel();

            // 사업장: ViewBag.Sites가 SelectListItem 또는 (value,text) 튜플일 수 있어 둘 다 대응
            if (ViewBag.Sites is IEnumerable<SelectListItem> siteItems && siteItems.Any())
            {
                vm.CompOptions = siteItems.ToList(); // ← 변경
            }
            else if (ViewBag.Sites is IEnumerable<(string value, string text)> siteTuples && siteTuples.Any())
            {
                vm.CompOptions = siteTuples
                    .Select(t => new SelectListItem { Value = t.value, Text = t.text })
                    .ToList(); // ← 변경
            }


            // 부서: ViewBag.Departments 동일 패턴 매핑
            if (ViewBag.Departments is IEnumerable<SelectListItem> deptItems && deptItems.Any())
            {
                vm.DepartmentOptions = deptItems.ToList(); // ← 변경
            }
            else if (ViewBag.Departments is IEnumerable<(string value, string text)> deptTuples && deptTuples.Any())
            {
                vm.DepartmentOptions = deptTuples
                    .Select(t => new SelectListItem { Value = t.value, Text = t.text })
                    .ToList(); // ← 변경
            }

            // 카테고리(ViewBag.Kinds), 템플릿(ViewBag.Templates)은 기존 공급 로직을 그대로 사용(뷰에서 직접 렌더링)
            return View("Select", vm);
        }

        // 2) 선택된 템플릿으로 작성 화면
        [HttpGet]
        public IActionResult Create(string templateCode)
        {
            if (string.IsNullOrWhiteSpace(templateCode))
            {
                TempData["NewDocAlert"] = "DOC_Val_TemplateRequired";
                return RedirectToAction(nameof(New));
            }

            ViewBag.TemplateCode = templateCode;

            // 2025.10.15 Changed: 실제 파일 저장소에서 템플릿 메타 로드
            var (descriptorJson, previewJson, templateTitle) = LoadTemplateMeta(templateCode);

            ViewBag.DescriptorJson = descriptorJson;
            ViewBag.PreviewJson = previewJson;
            ViewBag.TemplateTitle = templateTitle;

            return View("Compose", new DocTLViewModel());
        }

        // 2025.10.15 Added: CSRF 토큰 발급
        [HttpGet]
        [Produces("application/json")]
        public IActionResult Csrf()
        {
            var tokens = _antiforgery.GetAndStoreTokens(HttpContext);
            return Json(new { headerName = "RequestVerificationToken", token = tokens.RequestToken });
        }

        // 3) 작성본 저장(JSON POST)
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> Create([FromBody] ComposePostDto? dto)
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

                // 입력 필수 검증
                foreach (var f in descriptor.Inputs.Where(x => x.Required))
                {
                    var ok = dto.inputs != null
                             && dto.inputs.TryGetValue(f.Key ?? string.Empty, out var v)
                             && !string.IsNullOrWhiteSpace(v);
                    if (!ok)
                        ModelState.AddModelError($"Inputs[{f.Key}]", "DOC_Val_Required");
                }

                // 결재 필수 검증
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

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;

            // 문서 키 예시
            var docId = $"DOC_{DateTime.UtcNow:yyyyMMddHHmmssfff}";

            // (옵션) 컨텍스트 값
            string? departmentId = null;
            string? compCd = null;

            // 2025.10.15 Changed: 템플릿 메타 실제 값 로드(파일 없으면 폴백)
            var meta = LoadTemplateMeta(dto?.templateCode ?? string.Empty);
            string templateTitle = meta.templateTitle;
            string descriptorJson = meta.descriptorJson;
            string previewJson = meta.previewJson;

            // ✅ DB 저장부에서 사용할 널-안전 지역 변수
            var d = dto ?? new ComposePostDto();

            var connStr = _cfg.GetConnectionString("DefaultConnection");
            using (var conn = new SqlConnection(connStr))
            {
                await conn.OpenAsync();
                using (var tx = conn.BeginTransaction())
                {
                    // Documents
                    const string SQL_INS_DOC = @"
INSERT INTO dbo.Documents
    (DocId, TemplateCode, TemplateTitle, Status, DepartmentId, CompCd,
     CreatedBy, DescriptorJson, PreviewJson)
VALUES
    (@DocId, @TemplateCode, @TemplateTitle, @Status, @DepartmentId, @CompCd,
     @CreatedBy, @DescriptorJson, @PreviewJson);";

                    using (var cmd = new SqlCommand(SQL_INS_DOC, conn, tx))
                    {
                        cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                        cmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 50) { Value = d.templateCode ?? string.Empty });
                        cmd.Parameters.Add(new SqlParameter("@TemplateTitle", SqlDbType.NVarChar, 200) { Value = (object?)templateTitle ?? DBNull.Value });
                        cmd.Parameters.Add(new SqlParameter("@Status", SqlDbType.VarChar, 20) { Value = "Submitted" });
                        cmd.Parameters.Add(new SqlParameter("@DepartmentId", SqlDbType.VarChar, 12) { Value = (object?)departmentId ?? DBNull.Value });
                        cmd.Parameters.Add(new SqlParameter("@CompCd", SqlDbType.VarChar, 12) { Value = (object?)compCd ?? DBNull.Value });
                        cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 64) { Value = userId });
                        cmd.Parameters.Add(new SqlParameter("@DescriptorJson", SqlDbType.NVarChar, -1) { Value = descriptorJson });
                        cmd.Parameters.Add(new SqlParameter("@PreviewJson", SqlDbType.NVarChar, -1) { Value = previewJson });
                        cmd.ExecuteNonQuery();
                    }

                    // DocumentInputs — 널 안정성(패턴 매칭)
                    if (d.inputs is { Count: > 0 } inputs)
                    {
                        const string SQL_INS_INPUT = @"
INSERT INTO dbo.DocumentInputs
    (DocId, FieldKey, FieldValue)
VALUES
    (@DocId, @FieldKey, @FieldValue);";

                        using (var cmd = new SqlCommand(SQL_INS_INPUT, conn, tx))
                        {
                            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                            var pKey = cmd.Parameters.Add("@FieldKey", SqlDbType.NVarChar, 100);
                            var pVal = cmd.Parameters.Add("@FieldValue", SqlDbType.NVarChar, -1);

                            foreach (var kv in inputs)
                            {
                                pKey.Value = kv.Key ?? string.Empty;
                                pVal.Value = (object?)(kv.Value ?? string.Empty) ?? DBNull.Value;
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }

                    // 2025.10.15 Added: Seq 컬럼 존재 여부 확인 후 INSERT 분기
                    bool hasSeq;
                    using (var chkSeq = new SqlCommand(
                        "SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('dbo.DocumentApprovals') AND name = 'Seq'", conn, tx))
                    {
                        hasSeq = (await chkSeq.ExecuteScalarAsync()) is not null;
                    }

                    string sqlInsAppr = hasSeq
                        ? @"INSERT INTO dbo.DocumentApprovals (DocId, Seq, RoleKey, ApproverType, Value, Status)
                           VALUES (@DocId, @Seq, @RoleKey, @ApproverType, @Value, @Status);"
                        : @"INSERT INTO dbo.DocumentApprovals (DocId, RoleKey, ApproverType, Value, Status)
                           VALUES (@DocId, @RoleKey, @ApproverType, @Value, @Status);";

                    var approvalsMap = d.approvals;
                    using (var cmd = new SqlCommand(sqlInsAppr, conn, tx))
                    {
                        cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                        SqlParameter? pSeq = null;
                        if (hasSeq) pSeq = cmd.Parameters.Add("@Seq", SqlDbType.Int);
                        var pRole = cmd.Parameters.Add("@RoleKey", SqlDbType.NVarChar, 50);
                        var pType = cmd.Parameters.Add("@ApproverType", SqlDbType.VarChar, 20);
                        var pValue = cmd.Parameters.Add("@Value", SqlDbType.NVarChar, 128);
                        var pStat = cmd.Parameters.Add("@Status", SqlDbType.VarChar, 20);

                        var descriptor2 = LoadDescriptor(d.templateCode ?? string.Empty);
                        for (int i = 0; i < descriptor2.Approvals.Count; i++)
                        {
                            var ap = descriptor2.Approvals[i];
                            var roleKey = ap.RoleKey ?? string.Empty;

                            string? chosen = null;
                            approvalsMap?.TryGetValue(roleKey, out chosen);

                            if (hasSeq && pSeq != null) pSeq.Value = i + 1;
                            pRole.Value = roleKey;
                            pType.Value = ap.ApproverType ?? "Person";
                            pValue.Value = string.IsNullOrWhiteSpace(chosen) ? (object)DBNull.Value : chosen;
                            pStat.Value = "Pending";
                            cmd.ExecuteNonQuery();
                        }
                    }

                    tx.Commit();
                }
            }

            // 감사로그
            try { await _audit.LogAsync(docId, userId, "Created", null); } catch { /* ignore */ }

            var redirectUrl = Url.Action(nameof(Details), "Doc", new { id = docId });
            return Json(new { redirectUrl });
        }

        // 4) 문서 상세(게시)
        [HttpGet]
        public async Task<IActionResult> Details(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                TempData["NewDocAlert"] = "DOC_Err_DocumentNotFound";
                return RedirectToAction(nameof(New));
            }

            var connStr = _cfg.GetConnectionString("DefaultConnection");
            using (var conn = new SqlConnection(connStr))
            {
                await conn.OpenAsync();

                // 문서 본문
                const string SQL_DOC = @"
SELECT DocId, TemplateCode, TemplateTitle, Status, DepartmentId, CompCd,
       DescriptorJson, PreviewJson
FROM dbo.Documents WITH (NOLOCK)
WHERE DocId = @DocId";

                using (var cmd = new SqlCommand(SQL_DOC, conn))
                {
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = id });
                    using var rdr = await cmd.ExecuteReaderAsync();
                    if (!await rdr.ReadAsync())
                    {
                        TempData["NewDocAlert"] = "DOC_Err_DocumentNotFound";
                        return RedirectToAction(nameof(New));
                    }

                    ViewBag.DocumentId = rdr["DocId"]?.ToString() ?? id;
                    ViewBag.TemplateCode = rdr["TemplateCode"]?.ToString() ?? "";
                    ViewBag.TemplateTitle = rdr["TemplateTitle"]?.ToString() ?? "";
                    ViewBag.Status = rdr["Status"]?.ToString() ?? "Draft";
                    ViewBag.DescriptorJson = rdr["DescriptorJson"]?.ToString() ?? "{}";
                    ViewBag.PreviewJson = rdr["PreviewJson"]?.ToString() ?? "{}";
                    ViewBag.DepartmentId = rdr["DepartmentId"]?.ToString();
                    ViewBag.CompCd = rdr["CompCd"]?.ToString();
                }

                // 입력값
                const string SQL_INPUTS = @"
SELECT FieldKey, FieldValue
FROM dbo.DocumentInputs WITH (NOLOCK)
WHERE DocId = @DocId";

                var inputs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                using (var cmd = new SqlCommand(SQL_INPUTS, conn))
                {
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = id });
                    using var rdr = await cmd.ExecuteReaderAsync();
                    while (await rdr.ReadAsync())
                    {
                        var k = rdr["FieldKey"]?.ToString() ?? "";
                        var v = rdr["FieldValue"]?.ToString() ?? "";
                        if (!string.IsNullOrEmpty(k)) inputs[k] = v ?? "";
                    }
                }
                ViewBag.InputsJson = JsonSerializer.Serialize(inputs);

                // 2025.10.15 Added: Seq 유무에 따라 ORDER BY 분기
                string sqlAppr = @"
IF COL_LENGTH('dbo.DocumentApprovals','Seq') IS NOT NULL
    SELECT RoleKey, ApproverType, Value, Status
    FROM dbo.DocumentApprovals WITH (NOLOCK)
    WHERE DocId = @DocId
    ORDER BY Seq ASC;
ELSE
    SELECT RoleKey, ApproverType, Value, Status
    FROM dbo.DocumentApprovals WITH (NOLOCK)
    WHERE DocId = @DocId
    ORDER BY RoleKey ASC;";

                var appr = new List<object>();
                using (var cmd = new SqlCommand(sqlAppr, conn))
                {
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = id });
                    using var rdr = await cmd.ExecuteReaderAsync();
                    while (await rdr.ReadAsync())
                    {
                        appr.Add(new
                        {
                            roleKey = rdr["RoleKey"]?.ToString() ?? "",
                            approverType = rdr["ApproverType"]?.ToString() ?? "",
                            value = rdr["Value"]?.ToString() ?? "",
                            status = rdr["Status"]?.ToString() ?? "Pending"
                        });
                    }
                }
                ViewBag.ApprovalsJson = JsonSerializer.Serialize(appr);
            }

            // 감사 로그
            try
            {
                var actorId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
                await _audit.LogAsync(id, actorId, "Viewed", null);
            }
            catch { /* ignore */ }

            return View("Detail", new DocTLViewModel());
        }

        // ----------------- 내부 헬퍼/DTO -----------------

        // 클라이언트 페이로드 DTO(키명 변경 금지)
        public class ComposePostDto
        {
            public string? templateCode { get; set; }
            public Dictionary<string, string>? inputs { get; set; }
            public Dictionary<string, string>? approvals { get; set; }
            public string? descriptorVersion { get; set; }
        }

        // 디스크립터 구조(서버 검증용 최소 필드)
        private record Descriptor(
            List<InputField> Inputs,
            List<ApprovalField> Approvals,
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

        // 2025.10.15 Changed: 파일 기반 디스크립터 사용
        private Descriptor LoadDescriptor(string templateCode)
        {
            var meta = LoadTemplateMeta(templateCode);
            var json = string.IsNullOrWhiteSpace(meta.descriptorJson) ? "{}" : meta.descriptorJson;

            try
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                var inputs = new List<InputField>();
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

                var approvals = new List<ApprovalField>();
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
                return Descriptor.Empty;
            }
        }

        // 2025.10.15 Added: 템플릿 메타 파일 로더
        private (string descriptorJson, string previewJson, string templateTitle) LoadTemplateMeta(string templateCode)
        {
            try
            {
                var root = Path.Combine(_env.WebRootPath ?? string.Empty, "templates", templateCode ?? string.Empty);
                var descPath = Path.Combine(root, "descriptor.json");
                var prevPath = Path.Combine(root, "preview.json");
                var titlePath = Path.Combine(root, "title.txt");

                string descriptorJson = System.IO.File.Exists(descPath) ? System.IO.File.ReadAllText(descPath) : "{}";
                string previewJson = System.IO.File.Exists(prevPath) ? System.IO.File.ReadAllText(prevPath) : "{}";
                string templateTitle = System.IO.File.Exists(titlePath) ? (System.IO.File.ReadAllText(titlePath)?.Trim() ?? string.Empty) : string.Empty;

                if (string.IsNullOrWhiteSpace(descriptorJson)) descriptorJson = "{}";
                if (string.IsNullOrWhiteSpace(previewJson)) previewJson = "{}";
                return (descriptorJson, previewJson, templateTitle);
            }
            catch
            {
                return ("{}", "{}", string.Empty);
            }
        }

        // ====================== Comments API ======================
        #region Comments

        // 목록: GET /Doc/Comments?docId=...
        [HttpGet]
        [Produces("application/json")]
        public async Task<IActionResult> Comments([FromQuery] string docId)
        {
            if (string.IsNullOrWhiteSpace(docId))
                return BadRequest(new { messages = new[] { "DOC_Err_DocumentNotFound" } });

            var connStr = _cfg.GetConnectionString("DefaultConnection");
            var items = new List<object>();

            using (var conn = new SqlConnection(connStr))
            {
                await conn.OpenAsync();

                // 문서 존재 확인
                const string SQL_DOC_EXISTS = @"SELECT 1 FROM dbo.Documents WITH (NOLOCK) WHERE DocId=@DocId";
                using (var cmd = new SqlCommand(SQL_DOC_EXISTS, conn))
                {
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                    var ok = await cmd.ExecuteScalarAsync();
                    if (ok == null) return NotFound(new { messages = new[] { "DOC_Err_DocumentNotFound" } });
                }

                // 댓글 + 파일
                const string SQL_COMMENTS = @"
;WITH C AS
(
    SELECT c.CommentId, c.DocId, c.ParentCommentId, c.Body, c.CreatedAt, c.CreatedBy, 0 AS Depth
    FROM dbo.DocumentComments AS c WITH (NOLOCK)
    WHERE c.DocId = @DocId AND c.ParentCommentId IS NULL
    UNION ALL
    SELECT ch.CommentId, ch.DocId, ch.ParentCommentId, ch.Body, ch.CreatedAt, ch.CreatedBy, p.Depth + 1
    FROM dbo.DocumentComments AS ch WITH (NOLOCK)
    JOIN C AS p ON ch.ParentCommentId = p.CommentId
)
SELECT c.CommentId, c.ParentCommentId, c.Body, c.CreatedAt, c.CreatedBy, c.Depth,
       f.FileId, f.FileKey, f.OriginalName, f.ContentType, f.ByteSize
FROM C AS c
LEFT JOIN dbo.DocumentCommentFiles AS f WITH (NOLOCK) ON f.CommentId = c.CommentId
ORDER BY c.CreatedAt ASC, c.CommentId ASC, f.FileId ASC;";

                var map = new Dictionary<long, dynamic>();
                using (var cmd = new SqlCommand(SQL_COMMENTS, conn))
                {
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                    using var rdr = await cmd.ExecuteReaderAsync();
                    while (await rdr.ReadAsync())
                    {
                        var cmtId = Convert.ToInt64(rdr["CommentId"]);
                        if (!map.TryGetValue(cmtId, out var node))
                        {
                            node = new
                            {
                                commentId = cmtId,
                                parentCommentId = rdr["ParentCommentId"] == DBNull.Value ? (long?)null : Convert.ToInt64(rdr["ParentCommentId"]),
                                body = rdr["Body"]?.ToString() ?? "",
                                createdAt = ((DateTime)rdr["CreatedAt"]).ToString("yyyy-MM-dd HH:mm:ss"),
                                createdBy = rdr["CreatedBy"]?.ToString() ?? "",
                                depth = Convert.ToInt32(rdr["Depth"]),
                                files = new List<object>()
                            };
                            map[cmtId] = node;
                        }

                        if (rdr["FileId"] != DBNull.Value)
                        {
                            ((List<object>)node.files).Add(new
                            {
                                fileId = Convert.ToInt64(rdr["FileId"]),
                                fileKey = rdr["FileKey"]?.ToString() ?? "",
                                originalName = rdr["OriginalName"]?.ToString() ?? "",
                                contentType = rdr["ContentType"]?.ToString() ?? "",
                                byteSize = rdr["ByteSize"] == DBNull.Value ? 0L : Convert.ToInt64(rdr["ByteSize"])
                            });
                        }
                    }
                }

                items = map.Values.Cast<object>().ToList();
            }

            return Json(new { items });
        }

        // 작성: POST /Doc/Comments
        public sealed class CommentPostDto
        {
            public string? docId { get; set; }
            public long? parentCommentId { get; set; }
            public string? body { get; set; }
            public List<CommentFileDto>? files { get; set; }
        }
        public sealed class CommentFileDto
        {
            public string? fileKey { get; set; }
            public string? originalName { get; set; }
            public string? contentType { get; set; }
            public long? byteSize { get; set; }
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> Comments([FromBody] CommentPostDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.docId))
                return BadRequest(new { messages = new[] { "DOC_Err_InvalidPayload" } });
            if (string.IsNullOrWhiteSpace(dto.body))
                return BadRequest(new { messages = new[] { "DOC_Val_Required" }, fieldErrors = new { body = new[] { "DOC_Val_Required" } } });

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";

            var connStr = _cfg.GetConnectionString("DefaultConnection");
            using (var conn = new SqlConnection(connStr))
            {
                await conn.OpenAsync();
                using (var tx = conn.BeginTransaction())
                {
                    // 문서 존재 확인
                    const string SQL_DOC_EXISTS = @"SELECT 1 FROM dbo.Documents WITH (NOLOCK) WHERE DocId=@DocId";
                    using (var chk = new SqlCommand(SQL_DOC_EXISTS, conn, tx))
                    {
                        chk.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = dto.docId! });
                        var ok = await chk.ExecuteScalarAsync();
                        if (ok == null)
                        {
                            tx.Rollback();
                            return NotFound(new { messages = new[] { "DOC_Err_DocumentNotFound" } });
                        }
                    }

                    // 댓글 저장
                    const string SQL_INS_C = @"
INSERT INTO dbo.DocumentComments (DocId, ParentCommentId, Body, CreatedAt, CreatedBy)
VALUES (@DocId, @ParentCommentId, @Body, SYSUTCDATETIME(), @CreatedBy);
SELECT CAST(SCOPE_IDENTITY() AS BIGINT);";

                    long newId;
                    using (var cmd = new SqlCommand(SQL_INS_C, conn, tx))
                    {
                        cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = dto.docId! });
                        cmd.Parameters.Add(new SqlParameter("@ParentCommentId", SqlDbType.BigInt) { Value = (object?)dto.parentCommentId ?? DBNull.Value });
                        cmd.Parameters.Add(new SqlParameter("@Body", SqlDbType.NVarChar, -1) { Value = dto.body! });
                        cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 64) { Value = userId });

                        object? scalar = await cmd.ExecuteScalarAsync();
                        if (scalar is null || scalar == DBNull.Value)
                            throw new InvalidOperationException("Identity value was not returned.");

                        newId = scalar switch
                        {
                            decimal d => (long)d,
                            long l => l,
                            int i => i,
                            _ => Convert.ToInt64(scalar)
                        };
                    }

                    // 파일 메타 저장
                    if (dto.files is { Count: > 0 })
                    {
                        const string SQL_INS_F = @"
INSERT INTO dbo.DocumentCommentFiles (CommentId, FileKey, OriginalName, ContentType, ByteSize)
VALUES (@CommentId, @FileKey, @OriginalName, @ContentType, @ByteSize);";

                        using (var cmd = new SqlCommand(SQL_INS_F, conn, tx))
                        {
                            var pCid = cmd.Parameters.Add("@CommentId", SqlDbType.BigInt);
                            var pKey = cmd.Parameters.Add("@FileKey", SqlDbType.NVarChar, 200);
                            var pNm = cmd.Parameters.Add("@OriginalName", SqlDbType.NVarChar, 260);
                            var pCt = cmd.Parameters.Add("@ContentType", SqlDbType.NVarChar, 100);
                            var pSz = cmd.Parameters.Add("@ByteSize", SqlDbType.BigInt);
                            pSz.IsNullable = true;

                            foreach (var f in dto.files.Where(x => x != null))
                            {
                                pCid.Value = newId;
                                pKey.Value = (object?)f!.fileKey ?? DBNull.Value;
                                pNm.Value = (object?)f.originalName ?? DBNull.Value;
                                pCt.Value = (object?)f.contentType ?? DBNull.Value;
                                pSz.Value = f.byteSize is long bs ? bs : DBNull.Value;
                                await cmd.ExecuteNonQueryAsync();
                            }
                        }
                    }

                    tx.Commit();

                    try { await _audit.LogAsync(dto.docId!, userId, "CommentCreated", newId.ToString()); } catch { /* ignore */ }
                }
            }

            return Json(new { ok = true });
        }

        // 삭제: DELETE /Doc/Comments/{id}?docId=...
        [HttpDelete("/Doc/Comments/{id:long}")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteComment([FromRoute] long id, [FromQuery] string docId)
        {
            if (string.IsNullOrWhiteSpace(docId))
                return BadRequest(new { messages = new[] { "DOC_Err_DocumentNotFound" } });

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            var connStr = _cfg.GetConnectionString("DefaultConnection");

            using (var conn = new SqlConnection(connStr))
            {
                await conn.OpenAsync();
                using (var tx = conn.BeginTransaction())
                {
                    // 간단 권한 체크
                    const string SQL_GET_OWNER = @"SELECT CreatedBy FROM dbo.DocumentComments WITH (NOLOCK) WHERE CommentId=@Id AND DocId=@DocId";
                    string? owner = null;
                    using (var cmd = new SqlCommand(SQL_GET_OWNER, conn, tx))
                    {
                        cmd.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = id });
                        cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                        owner = (string?)await cmd.ExecuteScalarAsync();
                    }
                    if (owner == null)
                    {
                        tx.Rollback();
                        return NotFound(new { messages = new[] { "DOC_Err_NotFound" } });
                    }
                    if (!string.Equals(owner, userId, StringComparison.OrdinalIgnoreCase))
                    {
                        tx.Rollback();
                        return Forbid();
                    }

                    // 자식 댓글/파일부터 삭제
                    const string SQL_DEL_FILES = @"
DELETE f FROM dbo.DocumentCommentFiles f
WHERE f.CommentId IN (
    WITH RECURSIVE_CTE AS
    (
        SELECT CommentId FROM dbo.DocumentComments WHERE CommentId=@Id AND DocId=@DocId
        UNION ALL
        SELECT c.CommentId FROM dbo.DocumentComments c
        JOIN RECURSIVE_CTE p ON c.ParentCommentId = p.CommentId
    )
    SELECT CommentId FROM RECURSIVE_CTE
);";
                    using (var cmd = new SqlCommand(SQL_DEL_FILES.Replace("WITH RECURSIVE_CTE", "C"), conn, tx))
                    {
                        cmd.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = id });
                        cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                        await cmd.ExecuteNonQueryAsync();
                    }

                    const string SQL_DEL_CMTS = @"
;WITH C AS
(
    SELECT CommentId FROM dbo.DocumentComments WHERE CommentId=@Id AND DocId=@DocId
    UNION ALL
    SELECT c.CommentId FROM dbo.DocumentComments c JOIN C p ON c.ParentCommentId = p.CommentId
)
DELETE FROM dbo.DocumentComments WHERE CommentId IN (SELECT CommentId FROM C);";
                    using (var cmd = new SqlCommand(SQL_DEL_CMTS, conn, tx))
                    {
                        cmd.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = id });
                        cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                        await cmd.ExecuteNonQueryAsync();
                    }

                    tx.Commit();
                    try { await _audit.LogAsync(docId, userId, "CommentDeleted", id.ToString()); } catch { /* ignore */ }
                }
            }

            return Json(new { ok = true });
        }

        #endregion
        // ==================== /Comments API =======================
    }
}
