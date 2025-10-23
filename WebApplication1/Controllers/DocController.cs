// 2025.10.16 Final: ToHex() 제거 → ToHexIfRgb(IXLCell) 헬퍼 도입
// 2025.10.16 Final: BuildPreviewJsonFromExcel()로 스타일/병합/열폭 포함 미리보기 생성
// 2025.10.16 Final: DocTL 디스크립터 → Compose 스키마 자동 변환 및 PreviewJson 안전 보장
// 2025.10.16 Final: flowGroups 자동 생성(ORDER BY [Key], A1) 주입 → 템플릿별 3B 흘려쓰기 공통화

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Antiforgery;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using ClosedXML.Excel;
using Microsoft.EntityFrameworkCore;
using WebApplication1.Models;
using WebApplication1.Data;
using WebApplication1.Services;
using System.Text.Encodings.Web;

namespace WebApplication1.Controllers
{
    [Authorize]
    public class DocController : Controller
    {
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IConfiguration _cfg;
        private readonly IAntiforgery _antiforgery;
        private readonly IWebHostEnvironment _env;
        private readonly ApplicationDbContext _db;
        private readonly IDocTemplateService _tpl;

        public DocController(
            IStringLocalizer<SharedResource> S,
            IConfiguration cfg,
            IAntiforgery antiforgery,
            IWebHostEnvironment env,
            ApplicationDbContext db,
            IDocTemplateService tpl)
        {
            _S = S;
            _cfg = cfg;
            _antiforgery = antiforgery;
            _env = env;
            _db = db;
            _tpl = tpl;
        }

        // ---------- New ----------

        [HttpGet]
        public IActionResult New()
        {
            var vm = new DocTLViewModel();

            // 1) 클레임 우선
            var userComp = User.FindFirstValue("compCd") ?? "";
            var userDept = User.FindFirstValue("departmentId") ?? "";

            // 2) DB 보정
            if (string.IsNullOrWhiteSpace(userComp) || string.IsNullOrWhiteSpace(userDept))
            {
                try
                {
                    var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
                    if (!string.IsNullOrEmpty(userId))
                    {
                        var cs = _cfg.GetConnectionString("DefaultConnection");
                        using var conn = new SqlConnection(cs);
                        using var cmd = new SqlCommand(
                            @"SELECT TOP 1 CompCd, DepartmentId
                              FROM dbo.UserProfiles
                              WHERE UserId = @userId", conn);
                        cmd.Parameters.Add(new SqlParameter("@userId", SqlDbType.NVarChar, 450) { Value = userId });
                        conn.Open();
                        using var rd = cmd.ExecuteReader();
                        if (rd.Read())
                        {
                            if (string.IsNullOrWhiteSpace(userComp) && !rd.IsDBNull(0))
                                userComp = rd.GetString(0) ?? "";
                            if (string.IsNullOrWhiteSpace(userDept) && !rd.IsDBNull(1))
                                userDept = rd.GetInt32(1).ToString();
                        }
                    }
                }
                catch { /* ignore */ }
            }

            // 기본 placeholder
            vm.CompOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "", Text = _S["_CM_Select"].Value, Selected = true }
            };

            // 3) 사업장 목록
            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                using var cmd = new SqlCommand("SELECT CompCd, Name FROM dbo.CompMasters ORDER BY CompCd", conn);
                conn.Open();
                using var rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    var compCd = rd.GetString(0);
                    var name = rd.IsDBNull(1) ? compCd : rd.GetString(1);
                    vm.CompOptions.Add(new SelectListItem
                    {
                        Value = compCd,
                        Text = name,
                        Selected = !string.IsNullOrWhiteSpace(userComp)
                                   && string.Equals(userComp, compCd, StringComparison.OrdinalIgnoreCase)
                    });
                }
                if (vm.CompOptions.Any(o => o.Selected))
                    vm.CompOptions[0].Selected = false;
            }
            catch { /* ignore */ }

            // 부서 옵션
            vm.DepartmentOptions = new List<SelectListItem>
            {
                new SelectListItem { Value = "__SELECT__", Text = $"-- {_S["_CM_Select"]} --", Selected = true },
                new SelectListItem { Value = "", Text = _S["_CM_Common"].Value, Selected = string.IsNullOrEmpty(userDept) }
            };

            ViewBag.Templates = Array.Empty<(string code, string title)>();
            ViewBag.Kinds = Array.Empty<(string value, string text)>();
            ViewBag.Departments = Array.Empty<(string value, string text)>();
            ViewBag.Sites = Array.Empty<(string value, string text)>();
            ViewBag.UserComp = userComp;

            return View("Select", vm);
        }

        // ---------- Create (GET) ----------

        public async Task<IActionResult> Create(string templateCode)
        {
            var meta = await _tpl.LoadMetaAsync(templateCode);

            // tuple은 null 아님. 각 필드는 비어 있으면 안전 폴백 적용
            var descriptorJson = string.IsNullOrWhiteSpace(meta.descriptorJson) ? "{}" : meta.descriptorJson;
            var previewJson = string.IsNullOrWhiteSpace(meta.previewJson) ? "{}" : meta.previewJson;
            var excelAbsPath = meta.excelFilePath ?? string.Empty;

            // previewJson이 비거나 cells가 없으면 Controller의 로컬 생성기 사용
            if (string.IsNullOrWhiteSpace(previewJson) || !HasCells(previewJson))
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(excelAbsPath) && System.IO.File.Exists(excelAbsPath))
                    {
                        var rebuilt = BuildPreviewJsonFromExcel(excelAbsPath); // ← 서비스가 아니라 컨트롤러의 헬퍼 사용
                        if (!string.IsNullOrWhiteSpace(rebuilt) && HasCells(rebuilt))
                        {
                            previewJson = rebuilt;
                        }
                    }
                }
                catch { /* 폴백 유지 */ }
            }

            // ★ DocTL 스키마 → Compose 변환 + flowGroups 자동 주입 (ORDER BY [Key], A1)
            descriptorJson = BuildDescriptorJsonWithFlowGroups(templateCode, descriptorJson);

            ViewBag.DescriptorJson = descriptorJson;
            ViewBag.PreviewJson = previewJson;
            ViewBag.TemplateTitle = meta.templateTitle ?? string.Empty;
            ViewBag.TemplateCode = templateCode;

            // 엔드유저 화면에서는 입력 테이블 숨김 플래그 유지
            ViewBag.HideCellPicker = true;

            return View("Compose");

            // ----- 로컬 헬퍼 유지 -----
            static bool HasCells(string json)
            {
                if (string.IsNullOrWhiteSpace(json)) return false;
                if (!json.Contains("\"cells\"", StringComparison.Ordinal)) return false;
                try
                {
                    using var doc = System.Text.Json.JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("cells", out var cells)
                        && cells.ValueKind == System.Text.Json.JsonValueKind.Array
                        && cells.GetArrayLength() > 0)
                    {
                        var firstRow = cells[0];
                        return firstRow.ValueKind == System.Text.Json.JsonValueKind.Array;
                    }
                }
                catch { return false; }
                return false;
            }
        }

        // ---------- CSRF ----------

        [HttpGet]
        [Produces("application/json")]
        public IActionResult Csrf()
        {
            var tokens = _antiforgery.GetAndStoreTokens(HttpContext);
            return Json(new { headerName = "RequestVerificationToken", token = tokens.RequestToken });
        }

        // ---------- Create (POST) ----------
        [HttpPost]
        [ValidateAntiForgeryToken]
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> Create([FromBody] ComposePostDto? dto)
        {
            if (dto is null || string.IsNullOrWhiteSpace(dto.templateCode))
                return BadRequest(new { messages = new[] { "DOC_Val_TemplateRequired" } });

            // 1) 템플릿 메타 재조회 (versionId, excelPath 확보)
            var (descriptorJson, previewJson, title, versionId, excelPath)
                = await _tpl.LoadMetaAsync(dto.templateCode);

            if (versionId <= 0 || string.IsNullOrWhiteSpace(excelPath) || !System.IO.File.Exists(excelPath))
                return BadRequest(new { messages = new[] { "DOC_Err_TemplateNotReady" } });

            // 2) 엑셀 생성 (입력값 채워넣기)
            string? outputPath;
            try
            {
                outputPath = await GenerateExcelFromInputsAsync(versionId, excelPath, dto.inputs ?? new(), dto.descriptorVersion);
            }
            catch
            {
                return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" } });
            }

            // 3) 저장 후 이동
            var redirectUrl = Url.Action(nameof(Details), "Doc", new { id = System.IO.Path.GetFileNameWithoutExtension(outputPath) })
                              ?? "/Doc/Details";
            return Json(new { redirectUrl });
        }

        // ---------- Details ----------
        [HttpGet]
        public IActionResult Details(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                TempData["NewDocAlert"] = "DOC_Err_DocumentNotFound";
                return RedirectToAction(nameof(New));
            }

            ViewBag.DocumentId = id;
            ViewBag.TemplateCode = "";
            ViewBag.TemplateTitle = "";
            ViewBag.Status = "Draft";
            ViewBag.DescriptorJson = "{}";
            ViewBag.PreviewJson = "{}";
            ViewBag.InputsJson = "{}";
            ViewBag.ApprovalsJson = "[]";

            return View("Detail", new DocTLViewModel());
        }

        // ---------- DTOs ----------

        public class ComposePostDto
        {
            public string? templateCode { get; set; }
            public Dictionary<string, string>? inputs { get; set; }
            public Dictionary<string, string>? approvals { get; set; }
            public string? descriptorVersion { get; set; }
        }

        private record Descriptor(List<InputField> Inputs, List<ApprovalField> Approvals, string? Version)
        {
            public static Descriptor Empty => new(new(), new(), null);
        }

        private record InputField(string? Key, string? Type, bool Required, string? A1);
        private record ApprovalField(string? RoleKey, string? ApproverType, bool Required, string? Value);

        // Compose 디스크립터 DTO (flowGroups 주입용)
        private sealed class DescriptorDto
        {
            public string? version { get; set; }
            public List<InputFieldDto>? inputs { get; set; }
            public Dictionary<string, object>? styles { get; set; }
            public List<object>? approvals { get; set; }
            public List<FlowGroupDto>? flowGroups { get; set; }
        }
        private sealed class InputFieldDto
        {
            public string key { get; set; } = "";
            public string? a1 { get; set; }
            public string? type { get; set; }
        }
        private sealed class FlowGroupDto
        {
            public string id { get; set; } = "";
            public List<string> keys { get; set; } = new();
        }

        private Descriptor LoadDescriptor(string templateCode)
        {
            var meta = LoadTemplateMeta(templateCode);
            var raw = string.IsNullOrWhiteSpace(meta.descriptorJson) ? "{}" : meta.descriptorJson;
            var json = ConvertDescriptorIfNeeded(raw);

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

        // ---------- Template Meta Loader ----------
        // 최신 버전 메타를 DB에서 직접 읽고, PreviewJson이 없으면 ExcelFilePath로 미리보기 생성
        private (string descriptorJson, string previewJson, string templateTitle) LoadTemplateMeta(string templateCode)
        {
            string descriptorJson = "{}";
            string previewJson = "{}";
            string templateTitle = string.Empty;

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            if (string.IsNullOrWhiteSpace(cs))
                return (descriptorJson, previewJson, templateTitle);

            using var conn = new SqlConnection(cs);
            conn.Open();

            // 1) 최신 버전 한 건 조회 (DocCode -> 최신 Version)
            using (var cmd = new SqlCommand(@"
SELECT TOP 1
    m.DocName,
    v.DescriptorJson,
    v.PreviewJson,
    v.ExcelFilePath
FROM DocTemplateMaster m
JOIN DocTemplateVersion v ON v.TemplateId = m.Id
WHERE m.DocCode = @code
ORDER BY v.VersionNo DESC;", conn))
            {
                cmd.Parameters.Add(new SqlParameter("@code", SqlDbType.NVarChar, 100) { Value = templateCode ?? string.Empty });

                using var rd = cmd.ExecuteReader();
                if (rd.Read())
                {
                    templateTitle = rd["DocName"] as string ?? string.Empty;
                    descriptorJson = rd["DescriptorJson"] as string ?? "{}";
                    previewJson = rd["PreviewJson"] as string ?? "{}";

                    // 2) PreviewJson이 비어 있으면 ExcelFilePath로 미리보기 생성
                    if (string.IsNullOrWhiteSpace(previewJson) || previewJson.Trim() == "{}")
                    {
                        var excelPath = rd["ExcelFilePath"] as string ?? string.Empty;

                        // ContentRoot 기준 상대경로라면 절대경로로 보정
                        if (!string.IsNullOrWhiteSpace(excelPath) && !System.IO.Path.IsPathRooted(excelPath))
                        {
                            var baseDir = _env.ContentRootPath ?? AppContext.BaseDirectory;
                            excelPath = System.IO.Path.Combine(baseDir, excelPath);
                        }

                        if (!string.IsNullOrWhiteSpace(excelPath) && System.IO.File.Exists(excelPath))
                        {
                            try
                            {
                                previewJson = BuildPreviewJsonFromExcel(excelPath); // ← 이미 넣어둔 ClosedXML 기반 생성기 사용
                            }
                            catch
                            {
                                previewJson = "{}";
                            }
                        }
                    }
                }
            }

            // 3) 웹루트 템플릿 폴더에 파일이 있는 경우는 그대로 우선 사용(선택)
            try
            {
                var root = System.IO.Path.Combine(_env.WebRootPath ?? string.Empty, "templates", templateCode ?? string.Empty);
                var descPath = System.IO.File.Exists(System.IO.Path.Combine(root, "descriptor.json"))
                               ? System.IO.Path.Combine(root, "descriptor.json") : null;
                var prevPath = System.IO.File.Exists(System.IO.Path.Combine(root, "preview.json"))
                               ? System.IO.Path.Combine(root, "preview.json") : null;
                var titlePath = System.IO.File.Exists(System.IO.Path.Combine(root, "title.txt"))
                                ? System.IO.Path.Combine(root, "title.txt") : null;

                if (descPath != null) descriptorJson = System.IO.File.ReadAllText(descPath);
                if (prevPath != null) previewJson = System.IO.File.ReadAllText(prevPath);
                if (titlePath != null) templateTitle = (System.IO.File.ReadAllText(titlePath) ?? "").Trim();
            }
            catch { /* 무시 */ }

            // 4) 널/빈값 방어
            if (string.IsNullOrWhiteSpace(descriptorJson)) descriptorJson = "{}";
            if (string.IsNullOrWhiteSpace(previewJson)) previewJson = "{}";

            return (descriptorJson, previewJson, templateTitle);
        }

        // ---------- DocTL → Compose 변환기 ----------

        private static string ConvertDescriptorIfNeeded(string? json)
        {
            // 최소 구조
            const string EMPTY = "{\"inputs\":[],\"approvals\":[],\"version\":\"converted\"}";
            if (!TryParseJsonFlexible(json, out var doc)) return EMPTY;

            try
            {
                var root = doc.RootElement;

                // 이미 Compose 스키마
                if (root.TryGetProperty("inputs", out _)) return json!;

                // DocTL 스키마 → Compose 스키마로 변환
                List<(string key, string type, string a1)> inputs = new();
                if (root.TryGetProperty("Fields", out var fieldsEl) && fieldsEl.ValueKind == JsonValueKind.Array)
                {
                    foreach (var f in fieldsEl.EnumerateArray())
                    {
                        var key = f.TryGetProperty("Key", out var k) ? (k.GetString() ?? "").Trim() : "";
                        if (string.IsNullOrEmpty(key)) continue;
                        var type = f.TryGetProperty("Type", out var t) ? (t.GetString() ?? "Text") : "Text";
                        string a1 = "";
                        if (f.TryGetProperty("Cell", out var cell) && cell.ValueKind == JsonValueKind.Object)
                            a1 = cell.TryGetProperty("A1", out var a) ? (a.GetString() ?? "") : "";
                        inputs.Add((key, MapType(type), a1));
                    }
                }

                var approvalsDict = new Dictionary<int, (string approverType, string? value)>();
                if (root.TryGetProperty("Approvals", out var apprEl) && apprEl.ValueKind == JsonValueKind.Array)
                {
                    foreach (var a in apprEl.EnumerateArray())
                    {
                        int slot = a.TryGetProperty("Slot", out var s) && s.TryGetInt32(out var v) ? v : 0;
                        if (slot <= 0) continue;
                        var typ = a.TryGetProperty("ApproverType", out var at) ? (at.GetString() ?? "Person") : "Person";
                        var val = a.TryGetProperty("ApproverValue", out var av) ? av.GetString() : null;
                        approvalsDict[slot] = (MapApproverType(typ), val);
                    }
                }

                var approvals = approvalsDict
                    .OrderBy(kv => kv.Key)
                    .Select(kv => new { roleKey = $"A{kv.Key}", approverType = kv.Value.approverType, required = false, value = kv.Value.value ?? "" })
                    .ToList();

                var obj = new
                {
                    inputs = inputs.Select(x => new { key = x.key, type = x.type, required = false, a1 = x.a1 }).ToList(),
                    approvals,
                    version = "converted"
                };
                return JsonSerializer.Serialize(obj);
            }
            catch { return EMPTY; }
            finally { doc.Dispose(); }

            static string MapType(string t)
            {
                t = (t ?? "").Trim().ToLowerInvariant();
                if (t.StartsWith("date")) return "Date";
                if (t.StartsWith("num") || t.Contains("number") || t.Contains("decimal") || t.Contains("integer")) return "Num";
                return "Text";
            }
            static string MapApproverType(string t)
            {
                t = (t ?? "").Trim();
                return (t == "Person" || t == "Role" || t == "Rule") ? t : "Person";
            }
        }

        private static bool TryParseJsonFlexible(string? json, out JsonDocument doc)
        {
            doc = null!;
            if (string.IsNullOrWhiteSpace(json)) return false;

            try
            {
                var first = JsonDocument.Parse(json);
                if (first.RootElement.ValueKind == JsonValueKind.String)
                {
                    var inner = first.RootElement.GetString();
                    first.Dispose();
                    if (string.IsNullOrWhiteSpace(inner)) return false;
                    doc = JsonDocument.Parse(inner!);
                    return true;
                }
                doc = first;
                return true;
            }
            catch
            {
                return false;
            }
        }

        // ---------- Preview JSON 안전 보장 ----------

        private static string EnsurePreviewJsonSafe(string? previewJson)
        {
            if (TryParseJsonFlexible(previewJson, out var doc))
            {
                doc.Dispose();            // 파싱만 확인했으면 그대로 사용
                return previewJson!;
            }

            // 빈 그리드 폴백
            int rows = 50, cols = 26;
            var cells = new List<List<string>>(rows);
            for (int r = 0; r < rows; r++)
            {
                var row = new List<string>(cols);
                for (int c = 0; c < cols; c++) row.Add(string.Empty);
                cells.Add(row);
            }
            var colW = Enumerable.Repeat(12.0, cols).ToList();
            return JsonSerializer.Serialize(new { sheet = "Sheet1", rows, cols, cells, merges = Array.Empty<int[]>(), colW });
        }

        // ---------- 엑셀 → Preview JSON 생성기 ----------

        private static string BuildPreviewJsonFromExcel(string excelPath, int maxRows = 50, int maxCols = 26)
        {
            using var wb = new XLWorkbook(excelPath);
            var ws0 = wb.Worksheets.First();

            // 워크시트 기본값
            double defaultRowPt = ws0.RowHeight; if (defaultRowPt <= 0) defaultRowPt = 15.0;
            double defaultColChar = ws0.ColumnWidth; if (defaultColChar <= 0) defaultColChar = 8.43;

            // 1) 셀 값
            var cells = new List<List<string>>(maxRows);
            for (int r = 1; r <= maxRows; r++)
            {
                var row = new List<string>(maxCols);
                for (int c = 1; c <= maxCols; c++)
                    row.Add(ws0.Cell(r, c).GetString());
                cells.Add(row);
            }

            // 2) 병합
            var merges = new List<int[]>();
            foreach (var mr in ws0.MergedRanges)
            {
                var a = mr.RangeAddress;
                int r1 = a.FirstAddress.RowNumber, c1 = a.FirstAddress.ColumnNumber;
                int r2 = a.LastAddress.RowNumber, c2 = a.LastAddress.ColumnNumber;
                if (r1 > maxRows || c1 > maxCols) continue;
                r2 = Math.Min(r2, maxRows);
                c2 = Math.Min(c2, maxCols);
                if (r1 <= r2 && c1 <= c2) merges.Add(new[] { r1, c1, r2, c2 });
            }

            // 3) 열 너비
            var colW = new List<double>(maxCols);
            for (int c = 1; c <= maxCols; c++)
            {
                var w = ws0.Column(c).Width;
                if (w <= 0) w = defaultColChar;
                colW.Add(w);
            }

            // 4) 행 높이(포인트)
            var rowH = new List<double>(maxRows);
            for (int r = 1; r <= maxRows; r++)
            {
                var h = ws0.Row(r).Height;
                if (h <= 0) h = defaultRowPt;
                rowH.Add(h);
            }

            // 5) 스타일
            var styles = new Dictionary<string, object>();
            for (int r = 1; r <= maxRows; r++)
                for (int c = 1; c <= maxCols; c++)
                {
                    var cell = ws0.Cell(r, c);
                    var st = cell.Style;
                    string? bgHex = ToHexIfRgb(cell);

                    styles[$"{r},{c}"] = new
                    {
                        font = new
                        {
                            name = st.Font.FontName,
                            size = st.Font.FontSize,
                            bold = st.Font.Bold,
                            italic = st.Font.Italic,
                            underline = st.Font.Underline != XLFontUnderlineValues.None
                        },
                        align = new
                        {
                            h = st.Alignment.Horizontal.ToString(),
                            v = st.Alignment.Vertical.ToString(),
                            wrap = st.Alignment.WrapText
                        },
                        border = new
                        {
                            l = st.Border.LeftBorder.ToString(),
                            r = st.Border.RightBorder.ToString(),
                            t = st.Border.TopBorder.ToString(),
                            b = st.Border.BottomBorder.ToString()
                        },
                        fill = new { bg = bgHex }
                    };
                }

            // 6) 직렬화 — rowH 포함
            return JsonSerializer.Serialize(new
            {
                sheet = ws0.Name,
                rows = maxRows,
                cols = maxCols,
                cells,
                merges,
                colW,
                rowH,
                styles
            });
        }

        // ---------- 색상 HEX 헬퍼 ----------

        private static string? ToHexIfRgb(IXLCell cell)
        {
            try
            {
                var bg = cell?.Style?.Fill?.BackgroundColor;
                if (bg != null && bg.ColorType == XLColorType.Color)
                {
                    var c = bg.Color; // System.Drawing.Color
                    return $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                }
            }
            catch { /* ignore */ }
            return null;
        }

        /// <summary>
        /// DocTemplateField(VersionId) 매핑을 기준으로 inputs를 엑셀에 채워 산출물을 만든다.
        /// </summary>
        private async Task<string> GenerateExcelFromInputsAsync(
            long versionId,
            string templateExcelPath,
            Dictionary<string, string> inputs,
            string? descriptorVersion)
        {
            // 1) 필드 매핑 불러오기 (A1 또는 Row/Col → A1)
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var maps = new List<(string key, string a1, string type)>();

            await using (var conn = new SqlConnection(cs))
            {
                await conn.OpenAsync();
                await using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
SELECT [Key], [Type], A1, CellRow, CellColumn
FROM dbo.DocTemplateField
WHERE VersionId = @vid
ORDER BY Id;";
                    cmd.Parameters.Add(new SqlParameter("@vid", SqlDbType.BigInt) { Value = versionId });

                    await using var rd = await cmd.ExecuteReaderAsync();
                    while (await rd.ReadAsync())
                    {
                        var key = rd["Key"] as string ?? "";
                        var typ = rd["Type"] as string ?? "Text";

                        string a1 = rd["A1"] as string ?? "";
                        if (string.IsNullOrWhiteSpace(a1))
                        {
                            int r = rd["CellRow"] is DBNull ? 0 : Convert.ToInt32(rd["CellRow"]);
                            int c = rd["CellColumn"] is DBNull ? 0 : Convert.ToInt32(rd["CellColumn"]);
                            if (r > 0 && c > 0) a1 = A1FromRowCol(r, c);
                        }

                        if (!string.IsNullOrWhiteSpace(key) && !string.IsNullOrWhiteSpace(a1))
                            maps.Add((key, a1, MapFieldType(typ)));
                    }
                }
            }

            // 2) 엑셀 열기
            using var wb = new XLWorkbook(templateExcelPath);
            var ws = wb.Worksheets.First();

            // 3) 값 채우기
            foreach (var m in maps)
            {
                if (!inputs.TryGetValue(m.key, out var raw)) continue;
                var cell = ws.Cell(m.a1);

                if (string.IsNullOrEmpty(raw))
                {
                    cell.Clear(XLClearOptions.Contents); // 비우기 선택
                    continue;
                }

                // 줄바꿈 감지 시 wrap
                if (raw.Contains('\n') || raw.Contains('\r')) cell.Style.Alignment.WrapText = true;

                switch (m.type.ToLowerInvariant())
                {
                    case "date":
                        if (DateTime.TryParse(raw, out var dt))
                        {
                            cell.SetValue(dt);
                        }
                        else
                        {
                            cell.Value = raw;
                        }
                        break;

                    case "num":
                        if (decimal.TryParse(raw, out var dec))
                        {
                            cell.Value = dec;
                        }
                        else
                        {
                            cell.Value = raw;
                        }
                        break;

                    default:
                        cell.Value = raw;
                        break;
                }
            }

            // 4) 저장 경로
            var outDir = System.IO.Path.Combine(_env.ContentRootPath ?? AppContext.BaseDirectory, "App_Data", "Docs");
            Directory.CreateDirectory(outDir);
            var outName = $"DOC_{DateTime.Now:yyyyMMddHHmmssfff}.xlsx";
            var outPath = System.IO.Path.Combine(outDir, outName);

            wb.SaveAs(outPath);
            return outPath;

            // ---- local helpers ----
            static string A1FromRowCol(int row, int col)
            {
                if (row < 1 || col < 1) return "";
                string letters = "";
                int n = col;
                while (n > 0)
                {
                    int r = (n - 1) % 26;
                    letters = (char)('A' + r) + letters;
                    n = (n - 1) / 26;
                }
                return $"{letters}{row}";
            }
            static string MapFieldType(string? t)
            {
                var s = (t ?? "").Trim().ToLowerInvariant();
                if (s.StartsWith("date")) return "Date";
                if (s.StartsWith("num") || s.Contains("number") || s.Contains("decimal") || s.Contains("integer")) return "Num";
                return "Text";
            }
        }

        // ================== 여기부터 flowGroups 자동 주입 로직 ==================

        /// <summary>
        /// DocTL/Compose 여부와 상관없이 descriptor JSON을 Compose 스키마로 정규화한 뒤
        /// DB의 DocTemplateField를 사용해 flowGroups를 자동 생성하여 주입합니다.
        /// </summary>
        private string BuildDescriptorJsonWithFlowGroups(string templateCode, string rawDescriptorJson)
        {
            // 1) 스키마 정규화
            var normalized = ConvertDescriptorIfNeeded(rawDescriptorJson);

            // 2) 최신 VersionId
            var versionId = GetLatestVersionId(templateCode);
            if (versionId <= 0) return normalized;

            // 3) 그룹 계산
            var groups = BuildFlowGroupsForTemplate(versionId);

            // 4) JSON에 주입
            DescriptorDto desc;
            try
            {
                desc = JsonSerializer.Deserialize<DescriptorDto>(normalized) ?? new DescriptorDto();
            }
            catch
            {
                desc = new DescriptorDto();
            }
            desc.flowGroups = groups;

            var json = JsonSerializer.Serialize(
                desc,
                new JsonSerializerOptions
                {
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                    WriteIndented = false
                });

            return json;
        }

        /// <summary>
        /// 템플릿 코드로 최신 VersionId 조회
        /// </summary>
        private long GetLatestVersionId(string templateCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            if (string.IsNullOrWhiteSpace(cs)) return 0;

            try
            {
                using var conn = new SqlConnection(cs);
                conn.Open();
                // 스키마에 따라 v.Id 또는 v.VersionId 사용
                using var cmd = new SqlCommand(@"
SELECT TOP 1
    COALESCE(CAST(v.Id AS BIGINT), CAST(v.VersionId AS BIGINT)) AS VersionId
FROM DocTemplateMaster m
JOIN DocTemplateVersion v ON v.TemplateId = m.Id
WHERE m.DocCode = @code
ORDER BY v.VersionNo DESC;", conn);
                cmd.Parameters.Add(new SqlParameter("@code", SqlDbType.NVarChar, 100) { Value = templateCode ?? string.Empty });
                var obj = cmd.ExecuteScalar();
                if (obj is long l) return l;
                if (obj is int i) return i;
                if (obj != null && long.TryParse(obj.ToString(), out var p)) return p;
            }
            catch { /* ignore */ }
            return 0;
        }

        /// <summary>
        /// DocTemplateField(VersionId)에서 [Key], A1, Row/Col을 읽어
        /// 1) GROUP: 접두어(끝 "_숫자" 제거) 기준
        /// 2) ORDER: (a) 숫자 인덱스 있으면 숫자순 → (b) 없으면 A1 → Row → Col
        /// 로 Keys 배열을 만들고, 한 개짜리 그룹은 제외합니다.
        /// </summary>
        private List<FlowGroupDto> BuildFlowGroupsForTemplate(long versionId)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var rows = new List<(string Key, string? A1, int Row, int Col, string BaseKey, int? Index)>();

            using (var conn = new SqlConnection(cs))
            {
                conn.Open();
                using var cmd = new SqlCommand(@"
SELECT [Key], A1, CellRow, CellColumn
FROM dbo.DocTemplateField
WHERE VersionId = @vid AND [Key] IS NOT NULL
ORDER BY [Key], A1;", conn);
                cmd.Parameters.Add(new SqlParameter("@vid", SqlDbType.BigInt) { Value = versionId });
                using var rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    var key = rd["Key"] as string ?? "";
                    var a1 = rd["A1"] as string;
                    var r = rd["CellRow"] is DBNull ? 0 : Convert.ToInt32(rd["CellRow"]);
                    var c = rd["CellColumn"] is DBNull ? 0 : Convert.ToInt32(rd["CellColumn"]);
                    var (baseKey, idx) = ParseKey(key);
                    rows.Add((key, a1, r, c, baseKey, idx));
                }
            }

            var groups = rows
                .GroupBy(x => x.BaseKey)
                .Select(g =>
                {
                    var ordered = g
                        .OrderBy(x => x.Index.HasValue ? 0 : 1)
                        .ThenBy(x => x.Index)
                        .ThenBy(x => x.A1)         // A1이 null이면 아래로
                        .ThenBy(x => x.Row)
                        .ThenBy(x => x.Col)
                        .Select(x => x.Key)
                        .ToList();

                    return new FlowGroupDto { id = g.Key, keys = ordered };
                })
                .Where(g => g.keys.Count > 1) // 한 개짜리는 제외(의미 없음)
                .ToList();

            return groups;
        }

        /// <summary>
        /// Key 끝에 "_숫자"가 있으면 (접두어, 숫자) 반환. 없으면 (원키, null)
        /// 예) "진행업무_26" → ("진행업무", 26)
        /// </summary>
        private static (string BaseKey, int? Index) ParseKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return ("", null);
            var i = key.LastIndexOf('_');
            if (i > 0 && int.TryParse(key[(i + 1)..], out var n))
                return (key[..i], n);
            return (key, null);
        }
    }
}
