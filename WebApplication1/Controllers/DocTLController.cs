using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Localization;
using System.Security.Claims;
using System.Text.Json;
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using WebApplication1.Data;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{

    [Authorize]
    [Route("DocumentTemplates")]
    public class DocTLController : Controller
    {
        private readonly ApplicationDbContext _db;
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IWebHostEnvironment _env;

        public DocTLController(ApplicationDbContext db, IStringLocalizer<SharedResource> S, IWebHostEnvironment env)
        {
            _db = db;
            _S = S;
            _env = env;
        }

        private static bool IsExcelOpenXml(IFormFile f)
        {
            if (f is null || f.Length == 0) return false;
            var ext = Path.GetExtension(f.FileName).ToLowerInvariant();
            return ext == ".xlsx" || ext == ".xlsm";
        }

        private string? CurrentUserId() => User?.FindFirst(ClaimTypes.NameIdentifier)?.Value;

        private async Task<(bool found, string compCd, string compName, int? deptId, string? deptName, int adminLevel, string userName)> GetUserContextAsync()
        {
            var uid = CurrentUserId();
            if (string.IsNullOrEmpty(uid))
                return (false, "", "", null, null, 0, "");

            var profile = await _db.UserProfiles.AsNoTracking().FirstOrDefaultAsync(p => p.UserId == uid);

            var userName = User?.Identity?.Name ?? profile?.DisplayName ?? "";

            string compCd = profile?.CompCd ?? "";
            string compName = "";
            if (!string.IsNullOrWhiteSpace(compCd))
            {
                compName = await _db.CompMasters.Where(c => c.CompCd == compCd).Select(c => c.Name).FirstOrDefaultAsync() ?? "";
            }

            int? deptId = profile?.DepartmentId;
            string? deptName = null;
            if (deptId.HasValue)
            {
                deptName = await _db.DepartmentMasters.Where(d => d.Id == deptId.Value).Select(d => d.Name).FirstOrDefaultAsync();
            }

            var adminLevel = profile?.IsAdmin ?? 0;

            return (true, compCd, compName, deptId, deptName, adminLevel, userName);
        }

        [HttpGet("", Name = "DocumentTemplates.Index")]
        public async Task<IActionResult> Index()
        {
            var vm = new DocTLViewModel();

            var ctx = await GetUserContextAsync();

            ViewBag.IsAdmin = ctx.adminLevel > 0;
            ViewBag.AdminLevel = ctx.adminLevel;
            ViewBag.UserName = ctx.userName;
            ViewBag.UserCompCd = ctx.compCd;
            ViewBag.CompName = ctx.compName;
            ViewBag.UserDepartmentId = ctx.deptId;
            ViewBag.DeptName = ctx.deptName ?? _S["_CM_Common"];

            // 2025.09.09 사업장 콤보(DB 직접)
            vm.CompOptions = await _db.CompMasters
                .OrderBy(c => c.CompCd)
                .Select(c => new SelectListItem
                {
                    Value = c.CompCd,
                    Text = c.Name,
                    Selected = c.CompCd == ctx.compCd
                })
                .ToListAsync();

            // 2025.09.09 부서 콤보(DB 직접, 초기값 포함)
            vm.DepartmentOptions.Add(new SelectListItem
            {
                Value = "__SELECT__",
                Text = $"-- {_S["_CM_Select"]} --",
                Selected = true
            });
            vm.DepartmentOptions.Add(new SelectListItem
            {
                Value = "",
                Text = $"{_S["_CM_Common"]}",
                Selected = !ctx.deptId.HasValue
            });

            if (!string.IsNullOrWhiteSpace(ctx.compCd))
            {
                var depts = await _db.DepartmentMasters
                    .Where(d => d.CompCd == ctx.compCd)
                    .OrderBy(d => d.Name)
                    .Select(d => new { d.Id, d.Name })
                    .ToListAsync();

                foreach (var d in depts)
                {
                    vm.DepartmentOptions.Add(new SelectListItem
                    {
                        Value = d.Id.ToString(),
                        Text = d.Name,
                        Selected = ctx.deptId.HasValue && ctx.deptId.Value == d.Id
                    });
                }

                // 2025.09.09 초기에는 "--선택--" 대신 실제 값이 하나라도 선택되도록 조정
                if (vm.DepartmentOptions.Any(o => o.Selected))
                {
                    foreach (var o in vm.DepartmentOptions) if (o.Value == "__SELECT__") o.Selected = false;
                }
            }

            return View("DocTL", vm);
        }

        [HttpGet("get-departments")]
        public async Task<IActionResult> GetDepartments([FromQuery] string compCd)
        {
            var ctx = await GetUserContextAsync();
            var isAdmin = ctx.adminLevel > 0;

            // 비관리자는 자신의 사업장으로 강제
            if (!isAdmin) compCd = ctx.compCd;
            compCd = (compCd ?? "").Trim();

            // 1) 선택 사업장의 부서 목록 로드 (없을 수도 있음)
            var list = new List<object>();
            if (!string.IsNullOrEmpty(compCd))
            {
                var raw = await _db.DepartmentMasters
                    .Where(d => (d.CompCd ?? "").Trim() == compCd)
                    .OrderBy(d => d.Name)
                    .Select(d => new { d.Id, d.Name })
                    .ToListAsync();

                list.AddRange(raw.Select(d => new { id = d.Id, text = d.Name ?? "" }));
            }

            // 2) 안전장치: 현재 사용자 부서를 목록에 **항상 포함**
            if (ctx.deptId.HasValue && list.All(o => (o as dynamic).id != ctx.deptId.Value))
            {
                var mine = await _db.DepartmentMasters
                    .Where(d => d.Id == ctx.deptId.Value)
                    .Select(d => new { d.Id, d.Name })
                    .FirstOrDefaultAsync();

                if (mine != null)
                    list.Insert(0, new { id = mine.Id, text = mine.Name ?? "" });
            }

            // 3) 기본 옵션 + 목록
            var items = new List<object>
            {
                new { id = "__SELECT__", text = $"-- {_S["_CM_Select"]} --" },
                new { id = (int?)null,   text = $"{_S["_CM_Common"]}" }
            };
            items.AddRange(list);

            // 4) 초기 선택 값: 같은 사업장일 때만 사용자 부서를 바로 선택
            var selectedValue =
                (!isAdmin || string.Equals(compCd, ctx.compCd, StringComparison.OrdinalIgnoreCase))
                    ? (ctx.deptId?.ToString() ?? "")
                    : "__SELECT__";

            return Ok(new { items, selectedValue });
        }

        [HttpGet("get-documents")]
        public async Task<IActionResult> GetDocuments([FromQuery] string compCd,
                                             [FromQuery] int? departmentId,
                                             [FromQuery] string? kind)
        {
            var ctx = await GetUserContextAsync();
            if (ctx.adminLevel == 0)
                compCd = ctx.compCd;

            // 2025.09.09 실제 데이터 연결 전까지 빈 목록
            await Task.CompletedTask;
            return Ok(new { items = Array.Empty<object>() });
        }

        [HttpGet("new-template")]
        public IActionResult NewTemplate([FromQuery] string? compCd, [FromQuery] int? departmentId)
        {
            var rv = new RouteValueDictionary();
            if (!string.IsNullOrWhiteSpace(compCd)) rv["compCd"] = compCd;
            if (departmentId.HasValue) rv["departmentId"] = departmentId.Value;
            rv["openNewTemplate"] = "1";
            return RedirectToRoute("DocumentTemplates.Index", rv);
        }

        [HttpGet("open")]
        public IActionResult Open([FromQuery] string compCd, [FromQuery] int? departmentId, [FromQuery] string docCode)
        {
            var dept = departmentId.HasValue ? departmentId.Value.ToString() : $"{_S["_CM_Common"]}";
            return Content($"[OPEN] compCd={compCd}, dept={dept}, doc={docCode}");
        }

        [HttpPost("new-template")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> NewTemplatePost([FromForm] string? compCd,
                                                         [FromForm] int? departmentId,
                                                         [FromForm] string? kind,
                                                         [FromForm] string docName,
                                                         [FromForm] IFormFile? excelFile)
        {
            var ctx = await GetUserContextAsync();
            if (ctx.adminLevel == 0)
                compCd = ctx.compCd;

            if (string.IsNullOrWhiteSpace(compCd))
            {
                TempData["Alert"] = _S["DTL_Alert_SiteRequired"];
                return RedirectToRoute("DocumentTemplates.Index");
            }
            if (string.IsNullOrWhiteSpace(docName))
            {
                TempData["Alert"] = _S["DTL_Alert_EnterDocName"];
                return RedirectToRoute("DocumentTemplates.Index");
            }
            if (excelFile is null || excelFile.Length == 0)
            {
                TempData["Alert"] = _S["DTL_Alert_ExcelRequired"];
                return RedirectToRoute("DocumentTemplates.Index");
            }
            if (!IsExcelOpenXml(excelFile))
            {
                TempData["Alert"] = _S["DTL_Alert_ExcelOpenXmlOnly"];
                return RedirectToRoute("DocumentTemplates.Index");
            }

            var baseDir = Path.Combine(_env.ContentRootPath, "App_Data", "DocTemplates", "files");
            Directory.CreateDirectory(baseDir);
            static string Safe(string s) => string.Concat((s ?? string.Empty).Where(ch => char.IsLetterOrDigit(ch) || ch is '-' or '_'));
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var ext = Path.GetExtension(excelFile.FileName).ToLowerInvariant();
            var fileName = $"{Safe(compCd!)}_{Safe(docName)}_{stamp}_{Guid.NewGuid():N}{ext}";
            var excelPath = Path.Combine(baseDir, fileName);
            using (var fs = System.IO.File.Create(excelPath))
                await excelFile.CopyToAsync(fs);
            ViewBag.ExcelPath = excelPath;

            using var wb = new XLWorkbook(excelPath);

            var meta = ReadMetaCX(wb);
            var parsed = ParseByCommentsCX(wb);

            if (meta.ApprovalCount == null)
                meta.ApprovalCount = parsed.MaxApprovalSlot;

            var title = parsed.Title ?? ResolveTitleByNameOrMetaCX(wb, meta);

            // 2025.09.09 GRID PREVIEW JSON 생성
            var ws0 = wb.Worksheets.FirstOrDefault();
            if (ws0 != null)
            {
                const int PREV_MAX_ROWS = 50;
                const int PREV_MAX_COLS = 26;

                var cells = new List<List<string>>(PREV_MAX_ROWS);
                for (int r = 1; r <= PREV_MAX_ROWS; r++)
                {
                    var row = new List<string>(PREV_MAX_COLS);
                    for (int c = 1; c <= PREV_MAX_COLS; c++)
                        row.Add(ws0.Cell(r, c).GetString());
                    cells.Add(row);
                }

                var merges = new List<int[]>();
                foreach (var mr in ws0.MergedRanges)
                {
                    var a = mr.RangeAddress;
                    int r1 = a.FirstAddress.RowNumber, c1 = a.FirstAddress.ColumnNumber;
                    int r2 = a.LastAddress.RowNumber, c2 = a.LastAddress.ColumnNumber;
                    if (r1 > PREV_MAX_ROWS || c1 > PREV_MAX_COLS) continue;
                    r2 = Math.Min(r2, PREV_MAX_ROWS);
                    c2 = Math.Min(c2, PREV_MAX_COLS);
                    if (r1 <= r2 && c1 <= c2) merges.Add(new[] { r1, c1, r2, c2 });
                }

                var colW = new List<double>(PREV_MAX_COLS);
                for (int c = 1; c <= PREV_MAX_COLS; c++)
                    colW.Add(ws0.Column(c).Width);

                var preview = new { sheet = ws0.Name, rows = PREV_MAX_ROWS, cols = PREV_MAX_COLS, cells, merges, colW };
                ViewBag.PreviewJson = JsonSerializer.Serialize(preview);
            }

            // (2) 템플릿 디스크립터 만들기
            var descriptor = new TemplateDescriptor
            {
                CompCd = compCd!,
                DepartmentId = departmentId,
                Kind = kind,
                DocName = docName,
                Title = title,
                ApprovalCount = meta.ApprovalCount ?? 0,
                Fields = parsed.Fields,
                Approvals = parsed.Approvals
            };

            // (3) 뷰로 넘길 JSON
            ViewBag.DescriptorJson = JsonSerializer.Serialize(descriptor, new JsonSerializerOptions { WriteIndented = true });

            // (4) 매핑 화면으로 이동
            return View("DocTLMap");
        }

        private sealed class MetaInfo
        {
            public int? ApprovalCount { get; set; }
            public string? TitleCell { get; set; }
        }

        private MetaInfo ReadMetaCX(XLWorkbook wb)
        {
            var meta = new MetaInfo();
            var ws = wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, "EB_META", StringComparison.OrdinalIgnoreCase));
            if (ws == null) return meta;

            for (int r = 1; r <= 1000; r++)
            {
                var key = ws.Cell(r, 1).GetString().Trim();
                if (string.IsNullOrEmpty(key)) continue;
                var val = ws.Cell(r, 2).GetString().Trim();

                if (string.Equals(key, "ApprovalCount", StringComparison.OrdinalIgnoreCase) && int.TryParse(val, out var n))
                    meta.ApprovalCount = n;
                if (string.Equals(key, "TitleCell", StringComparison.OrdinalIgnoreCase))
                    meta.TitleCell = val;
            }
            return meta;
        }

        private string? ResolveTitleByNameOrMetaCX(XLWorkbook wb, MetaInfo meta)
        {
            var nr = wb.NamedRanges.FirstOrDefault(n => string.Equals(n.Name, "F_Title", StringComparison.OrdinalIgnoreCase));
            if (nr != null && nr.Ranges.Any())
                return nr.Ranges.First().FirstCell().GetString().Trim();

            if (!string.IsNullOrWhiteSpace(meta.TitleCell))
            {
                try
                {
                    var parts = meta.TitleCell.Split('!');
                    var addr = parts.Length == 2 ? parts[1] : parts[0];
                    var ws = parts.Length == 2 ? wb.Worksheets.Worksheet(parts[0]) : wb.Worksheets.First();
                    return ws.Cell(addr).GetString().Trim();
                }
                catch { }
            }
            return null;
        }

        private sealed class CommentParseResultCX
        {
            public string? Title { get; set; }
            public int MaxApprovalSlot { get; set; }
            public List<FieldDef> Fields { get; set; } = new();
            public List<ApprovalDef> Approvals { get; set; } = new();
        }

        private static Dictionary<string, string> ParseCommentTags(string text)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(text)) return dict;

            foreach (var raw in text.Replace("\r", "").Split('\n'))
            {
                var line = raw.Trim();
                if (line.Length == 0) continue;
                var eq = line.IndexOf('=');
                var col = line.IndexOf(':');
                var pos = (eq >= 0 && col >= 0) ? Math.Min(eq, col) : (eq >= 0 ? eq : col);
                if (pos <= 0) continue;
                var k = line[..pos].Trim();
                var v = line[(pos + 1)..].Trim();
                if (k.Length == 0) continue;
                dict[k] = v;
            }
            return dict;
        }

        private CommentParseResultCX ParseByCommentsCX(XLWorkbook wb)
        {
            var result = new CommentParseResultCX();

            foreach (var ws in wb.Worksheets)
            {
                if (string.Equals(ws.Name, "EB_META", StringComparison.OrdinalIgnoreCase))
                    continue;

                foreach (var cell in ws.CellsUsed(c => c.HasComment))
                {
                    var text = cell.GetComment().Text?.Trim() ?? "";
                    var tags = ParseCommentTags(text);
                    if (tags.Count == 0) continue;

                    if (tags.TryGetValue("Title", out var tf) && tf.Equals("true", StringComparison.OrdinalIgnoreCase))
                        result.Title ??= cell.GetString().Trim();

                    if (tags.TryGetValue("Field", out var key) && !string.IsNullOrWhiteSpace(key))
                    {
                        var type = tags.TryGetValue("Type", out var t) && !string.IsNullOrWhiteSpace(t)
                                   ? NormalizeType(t)
                                   : TryInferTypeFromValidationCX(ws, cell) ?? "Text";

                        result.Fields.Add(new FieldDef { Key = key.Trim(), Type = type, Cell = ToCellRef(cell) });
                        continue;
                    }

                    if (tags.TryGetValue("Approval", out var slotStr) && int.TryParse(slotStr, out int slot)
                        && tags.TryGetValue("Part", out var part) && !string.IsNullOrWhiteSpace(part))
                    {
                        result.Approvals.Add(new ApprovalDef { Slot = slot, Part = part.Trim(), Cell = ToCellRef(cell) });
                        if (slot > result.MaxApprovalSlot) result.MaxApprovalSlot = slot;
                        continue;
                    }

                    if (tags.TryGetValue("ApprovalKey", out var ak) && TryParseApprovalKey(ak, out int s, out string p))
                    {
                        result.Approvals.Add(new ApprovalDef { Slot = s, Part = p, Cell = ToCellRef(cell) });
                        if (s > result.MaxApprovalSlot) result.MaxApprovalSlot = s;
                        continue;
                    }
                }
            }

            result.Fields = result.Fields
                .GroupBy(f => f.Key, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.Last())
                .OrderBy(f => f.Key, StringComparer.OrdinalIgnoreCase)
                .ToList();

            result.Approvals = result.Approvals
                 .GroupBy(a => new { a.Slot, Part = a.Part.ToLowerInvariant() })
                 .Select(g => g.Last())
                 .OrderBy(a => a.Slot)
                 .ThenBy(a => a.Part, StringComparer.OrdinalIgnoreCase)
                 .ToList();

            return result;
        }

        private static string NormalizeType(string t)
        {
            t = t.Trim().ToLowerInvariant();
            if (t.StartsWith("date")) return "Date";
            if (t.StartsWith("num") || t.Contains("number") || t.Contains("decimal") || t.Contains("integer")) return "Num";
            return "Text";
        }

        private static string? TryInferTypeFromValidationCX(IXLWorksheet ws, IXLCell cell)
        {
            foreach (var dv in ws.DataValidations)
            {
                if (dv.Ranges.Any(r => r.Cells().Contains(cell)))
                {
                    return dv.AllowedValues switch
                    {
                        XLAllowedValues.Date => "Date",
                        XLAllowedValues.Decimal or XLAllowedValues.WholeNumber => "Num",
                        _ => null
                    };
                }
            }
            return null;
        }

        private static bool TryParseApprovalKey(string input, out int slot, out string part)
        {
            slot = 0; part = "";
            if (string.IsNullOrWhiteSpace(input)) { part = ""; return false; }
            var m = Regex.Match(input, @"^A(\d+)_(\w+)$", RegexOptions.IgnoreCase);
            if (!m.Success) return false;
            slot = int.Parse(m.Groups[1].Value);
            part = m.Groups[2].Value;
            return true;
        }

        private static CellRef ToCellRef(IXLCell cell)
        {
            var ws = cell.Worksheet;
            var range = cell.MergedRange() ?? cell.AsRange();
            var first = range.RangeAddress.FirstAddress;
            var last = range.RangeAddress.LastAddress;

            var a1 = first.Equals(last) ? first.ToStringRelative() : $"{first.ToStringRelative()}:{last.ToStringRelative()}";

            return new CellRef
            {
                Sheet = ws.Name,
                Row = first.RowNumber - 1,
                Column = first.ColumnNumber - 1,
                RowSpan = range.RowCount(),
                ColSpan = range.ColumnCount(),
                A1 = a1
            };
        }

        public sealed class TemplateDescriptor
        {
            public string CompCd { get; set; } = default!;
            public int? DepartmentId { get; set; }
            public string? Kind { get; set; }
            public string DocName { get; set; } = default!;
            public string? Title { get; set; }
            public int ApprovalCount { get; set; }
            public List<FieldDef> Fields { get; set; } = new();
            public List<ApprovalDef> Approvals { get; set; } = new();
        }

        public sealed class FieldDef
        {
            public string Key { get; set; } = default!;
            public string Type { get; set; } = "Text";
            public CellRef Cell { get; set; } = default!;
        }

        public sealed class ApprovalDef
        {
            public int Slot { get; set; }
            public string Part { get; set; } = "";
            public CellRef Cell { get; set; } = default!;
        }

        public sealed class CellRef
        {
            public string Sheet { get; set; } = default!;
            public int Row { get; set; }
            public int Column { get; set; }
            public int RowSpan { get; set; } = 1;
            public int ColSpan { get; set; } = 1;
            public string A1 { get; set; } = "";
        }

        [HttpPost("map-save")]
        [ValidateAntiForgeryToken]
        public IActionResult MapSave([FromForm] string descriptor, [FromForm] string? excelPath)
        {
            if (string.IsNullOrWhiteSpace(descriptor))
                return BadRequest("No descriptor");

            TemplateDescriptor? model;
            try { model = JsonSerializer.Deserialize<TemplateDescriptor>(descriptor); }
            catch { return BadRequest("Invalid descriptor"); }
            if (model == null) return BadRequest("Empty descriptor");

            var baseDir = Path.Combine(_env.ContentRootPath, "App_Data", "DocTemplates");
            Directory.CreateDirectory(baseDir);

            static string Safe(string s) => string.Concat(s.Where(ch => char.IsLetterOrDigit(ch) || ch is '-' or '_'));
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var name = $"{Safe(model.CompCd)}_{Safe(model.DocName)}_{stamp}_{Guid.NewGuid():N}.json";
            var path = Path.Combine(baseDir, name);

            System.IO.File.WriteAllText(path, JsonSerializer.Serialize(model, new JsonSerializerOptions { WriteIndented = true }));
            return Content($"[MAP-SAVE]\nfile={path}\nexcel={excelPath}\nfields={model.Fields?.Count ?? 0}\napprovals={model.Approvals?.Count ?? 0}");
        }
    }
}
