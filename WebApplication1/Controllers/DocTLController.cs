using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Localization;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.IO;
using System.Linq;
using System.Security.Claims;
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

        // ── Claims helpers ──────────────────────────────────────────────────────
        private string? CurrentUserId()
        {
            return User?.FindFirstValue(ClaimTypes.NameIdentifier)
                ?? User?.FindFirst("sub")?.Value;
        }

        private string? GetClaim(params string[] keys)
        {
            foreach (var k in keys)
            {
                var v = User?.FindFirst(k)?.Value;
                if (!string.IsNullOrWhiteSpace(v)) return v;
            }
            return null;
        }

        private bool IsAdminUser() =>
            User.HasClaim("is_admin", "1") || User.HasClaim("is_admin", "2")
            || (User?.IsInRole("Admin") ?? false) || (User?.IsInRole("Administrator") ?? false);

        private int GetAdminLevel()
        {
            if (User.HasClaim("is_admin", "2") || (User?.IsInRole("Administrator") ?? false)) return 2;
            if (User.HasClaim("is_admin", "1") || (User?.IsInRole("Admin") ?? false)) return 1;
            return 0;
        }

        // CompCd 문자열을 CompMasters에 맞춰 정규화(코드/이름 모두 허용)
        private async Task<string?> NormalizeCompCdAsync(string? compClaim)
        {
            if (string.IsNullOrWhiteSpace(compClaim)) return null;
            var cd = await _db.CompMasters
                .Where(c => c.CompCd == compClaim || c.Name == compClaim)
                .Select(c => c.CompCd)
                .FirstOrDefaultAsync();
            return cd ?? compClaim;
        }

        // DB(UserProfiles) 또는 클레임으로 사용자 사업장 결정
        private async Task<(string compCd, string compName)> ResolveUserCompAsync()
        {
            // DB 우선
            var uid = CurrentUserId();
            if (!string.IsNullOrEmpty(uid))
            {
                var up = await _db.UserProfiles.AsNoTracking()
                    .Where(p => p.UserId == uid)
                    .Select(p => p.CompCd)
                    .FirstOrDefaultAsync();

                if (!string.IsNullOrWhiteSpace(up))
                {
                    var row = await _db.CompMasters.AsNoTracking()
                                .Where(c => c.CompCd == up)
                                .Select(c => new { c.CompCd, c.Name })
                                .FirstOrDefaultAsync();
                    if (row != null) return (row.CompCd, row.Name);
                }
            }

            // 폴백: 클레임 → 정규화
            var claim = GetClaim("CompCd", "compcd", "COMP_CD", "comp", "site", "Site") ?? "";
            var norm = await NormalizeCompCdAsync(claim) ?? "";
            var found = await _db.CompMasters.AsNoTracking()
                            .Where(c => c.CompCd == norm)
                            .Select(c => new { c.CompCd, c.Name })
                            .FirstOrDefaultAsync();

            if (found != null) return (found.CompCd, found.Name);

            var first = await _db.CompMasters.AsNoTracking()
                           .OrderBy(c => c.CompCd)
                           .Select(c => new { c.CompCd, c.Name })
                           .FirstOrDefaultAsync();
            return first is null ? ("", "") : (first.CompCd, first.Name);
        }

        // 클레임/이름/코드로 부서 Id 추적
        private async Task<int?> ResolveUserDeptIdAsync(string compCd)
        {
            var rawId = GetClaim("DepartmentId", "departmentid", "DeptId", "deptid", "DEPT_ID", "DeptID");
            var rawCode = GetClaim("DepartmentCode", "departmentcode", "DeptCode", "deptcode", "DEPT_CODE");
            var rawName = GetClaim("DepartmentName", "departmentname", "DeptName", "deptname", "DEPT_NAME", "Dept");

            if (!string.IsNullOrWhiteSpace(rawId) && int.TryParse(rawId, out var idParsed))
                return idParsed;

            // ← 여기서 구체 타입 명시로 인한 CS0234가 났었습니다. var로 변경.
            var q = _db.DepartmentMasters.AsNoTracking().Where(d => d.CompCd == compCd);

            if (!string.IsNullOrWhiteSpace(rawCode))
            {
                var byCode = await q.Where(d => d.Code == rawCode)
                                    .Select(d => (int?)d.Id).FirstOrDefaultAsync();
                if (byCode.HasValue) return byCode.Value;
            }

            if (!string.IsNullOrWhiteSpace(rawName))
            {
                var byName = await q.Where(d => d.Name == rawName)
                                    .Select(d => (int?)d.Id).FirstOrDefaultAsync();
                if (byName.HasValue) return byName.Value;
            }

            if (!string.IsNullOrWhiteSpace(rawId))
            {
                var byStr = await q.Where(d => d.Code == rawId || d.Name == rawId)
                                   .Select(d => (int?)d.Id).FirstOrDefaultAsync();
                if (byStr.HasValue) return byStr.Value;
            }

            return null;
        }

        // ── Actions ────────────────────────────────────────────────────────────
        [HttpGet("", Name = "DocumentTemplates.Index")]
        public async Task<IActionResult> Index()
        {
            var vm = new DocTLViewModel();

            var adminLevel = GetAdminLevel();
            ViewBag.IsAdmin = adminLevel > 0;

            var userName = GetClaim("name", "Name", "nickname", "preferred_username") ?? (User?.Identity?.Name ?? "");

            // DB(UserProfiles) 우선
            string? dbCompCd = null; int? dbDeptId = null;
            var uid = CurrentUserId();
            if (!string.IsNullOrEmpty(uid))
            {
                var up = await _db.UserProfiles.AsNoTracking()
                    .Where(p => p.UserId == uid)
                    .Select(p => new { p.CompCd, p.DepartmentId })
                    .FirstOrDefaultAsync();
                dbCompCd = up?.CompCd;
                dbDeptId = up?.DepartmentId;
            }

            // 클레임 폴백 + 정규화
            var claimComp = GetClaim("CompCd", "compcd", "COMP_CD") ?? "";
            var normCompCd = await NormalizeCompCdAsync(dbCompCd ?? claimComp) ?? "";

            // 사업장 드롭다운
            var comps = await _db.CompMasters.OrderBy(c => c.CompCd).ToListAsync();
            var selectedComp = comps.FirstOrDefault(c => c.CompCd == normCompCd) ?? comps.FirstOrDefault();

            vm.CompOptions = comps.Select(c => new SelectListItem
            {
                Value = c.CompCd,
                Text = c.Name,
                Selected = selectedComp != null && c.CompCd == selectedComp.CompCd
            }).ToList();

            // 부서 콤보 기본 2개
            vm.DepartmentOptions.Add(new SelectListItem { Value = "__SELECT__", Text = $"--{_S["_CM_Select"]}--", Selected = true });
            vm.DepartmentOptions.Add(new SelectListItem { Value = "", Text = $"--{_S["_CM_Common"]}--" });

            ViewBag.UserName = userName;
            ViewBag.UserCompCd = selectedComp?.CompCd ?? "";
            ViewBag.CompName = selectedComp?.Name ?? "";

            // 표시용 부서명
            string? deptName = null;
            if ((dbDeptId ?? 0) > 0)
                deptName = await _db.DepartmentMasters
                            .Where(d => d.Id == dbDeptId)
                            .Select(d => d.Name)
                            .FirstOrDefaultAsync();
            ViewBag.DeptName = deptName ?? "";
            ViewBag.AdminLevel = adminLevel;

            return View("DocTL", vm);
        }

        [HttpGet("get-departments")]
        public async Task<IActionResult> GetDepartments([FromQuery] string compCd)
        {
            var admin = IsAdminUser();

            // DB(UserProfiles) 우선
            string? dbCompCd = null; int? dbDeptId = null;
            var uid = CurrentUserId();
            if (!string.IsNullOrEmpty(uid))
            {
                var up = await _db.UserProfiles.AsNoTracking()
                    .Where(p => p.UserId == uid)
                    .Select(p => new { p.CompCd, p.DepartmentId })
                    .FirstOrDefaultAsync();
                dbCompCd = up?.CompCd;
                dbDeptId = up?.DepartmentId;
            }

            var claimComp = GetClaim("CompCd", "compcd", "COMP_CD") ?? "";
            var userCompCd = await NormalizeCompCdAsync(dbCompCd ?? claimComp) ?? "";

            if (!admin) compCd = userCompCd;

            var itemsCore = Enumerable.Empty<object>();
            if (!string.IsNullOrWhiteSpace(compCd))
            {
                itemsCore = await _db.DepartmentMasters
                    .Where(d => d.CompCd == compCd)
                    .OrderBy(d => d.Name)
                    .Select(d => new { id = d.Id, text = d.Name })
                    .ToListAsync();
            }

            var items = new object[]
            {
                new { id = "__SELECT__", text = $"--{_S["_CM_Select"]}--" },
                new { id = (int?)null,   text = $"--{_S["_CM_Common"]}--" }
            }.Concat(itemsCore);

            string selectedValue;
            var effDept = dbDeptId ?? 0;
            if (!admin)
                selectedValue = effDept > 0 ? effDept.ToString() : "";
            else
                selectedValue = string.Equals(compCd, userCompCd, StringComparison.OrdinalIgnoreCase)
                    ? (effDept > 0 ? effDept.ToString() : "")
                    : "__SELECT__";

            return Ok(new { items, selectedValue });
        }

        [HttpGet("get-documents")]
        public async Task<IActionResult> GetDocuments([FromQuery] string compCd, [FromQuery] int? departmentId)
        {
            if (!IsAdminUser())
            {
                string? dbCompCd = null;
                var uid = CurrentUserId();
                if (!string.IsNullOrEmpty(uid))
                    dbCompCd = await _db.UserProfiles.AsNoTracking()
                                .Where(p => p.UserId == uid)
                                .Select(p => p.CompCd)
                                .FirstOrDefaultAsync();

                var claimComp = GetClaim("CompCd", "compcd", "COMP_CD") ?? "";
                compCd = await NormalizeCompCdAsync(dbCompCd ?? claimComp) ?? compCd;
            }
            // 아직 목록은 더미
            return Ok(new { items = Array.Empty<object>() });
        }

        [HttpGet("open")]
        public IActionResult Open([FromQuery] string compCd, [FromQuery] int? departmentId, [FromQuery] string docCode)
        {
            var dept = departmentId.HasValue ? departmentId.Value.ToString() : $" --{_S["_CM_Common"]}-- ";
            return Content($"[OPEN] compCd={compCd}, dept={dept}, doc={docCode}");
        }

        [HttpGet("new-template")]
        public IActionResult NewTemplate([FromQuery] string? compCd, [FromQuery] int? departmentId)
        {
            return View();
        }

        [HttpPost("new-template")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> NewTemplatePost([FromForm] string? compCd,
                                                         [FromForm] int? departmentId,
                                                         [FromForm] string? kind,
                                                         [FromForm] string docName,
                                                         [FromForm] IFormFile? excelFile)
        {
            if (!IsAdminUser())
            {
                var resolved = await ResolveUserCompAsync();
                compCd = resolved.compCd;
            }

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

            ViewBag.DescriptorJson = JsonSerializer.Serialize(descriptor, new JsonSerializerOptions { WriteIndented = true });

            var phys = Path.Combine(_env.ContentRootPath, "Views", "DocTL", "DocTLMap.cshtml");
            if (!System.IO.File.Exists(phys))
                return Content("VIEW FILE NOT FOUND: " + phys);

            return View("DocTLMap");
        }

        // ── Excel parser ────────────────────────────────────────────────────────
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
#pragma warning disable CS0618
            var nr = wb.NamedRanges.FirstOrDefault(n => string.Equals(n.Name, "F_Title", StringComparison.OrdinalIgnoreCase));
#pragma warning restore CS0618
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
            if (string.IsNullOrWhiteSpace(input)) return false;
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

            var a1 = first.Equals(last)
                ? first.ToStringRelative()
                : $"{first.ToStringRelative()}:{last.ToStringRelative()}";

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

        // ── POCO ────────────────────────────────────────────────────────────────
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

        // ── Map Save ───────────────────────────────────────────────────────────
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
