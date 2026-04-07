// 2026.03.23 Changed: 승인 descriptor 처리와 동일하게 협조 descriptor 파싱 저장 정규화를 추가하여 템플릿 저장 시 Cooperations 정보가 누락되지 않도록 수정함
using ClosedXML.Excel;
using DevExpress.AspNetCore.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Localization;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WebApplication1.Data;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("DocumentTemplatesDX")]
    public sealed class DocTLDXController : Controller
    {
        private readonly ApplicationDbContext _db;
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IWebHostEnvironment _env;

        public DocTLDXController(ApplicationDbContext db, IStringLocalizer<SharedResource> s, IWebHostEnvironment env)
        {
            _db = db;
            _S = s;
            _env = env;
        }

        [HttpGet("dx-callback")]
        [HttpPost("dx-callback")]
        public IActionResult DxCallback()
        {
            return SpreadsheetRequestProcessor.GetResponse(HttpContext);
        }

        private string? CurrentUserId() => User?.FindFirst(ClaimTypes.NameIdentifier)?.Value;

        private static string SafeDocId(string raw)
        {
            raw ??= string.Empty;
            raw = raw.Trim();
            if (raw.Length == 0) raw = "doc";
            raw = Regex.Replace(raw, @"[^a-zA-Z0-9_\-\.]", "_");
            if (raw.Length > 120) raw = raw.Substring(0, 120);
            return raw;
        }

        private static string? FirstNonEmpty(params string?[] xs)
            => xs.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));

        private static string SafeFilePart(string? s)
            => string.Concat((s ?? string.Empty).Where(ch => char.IsLetterOrDigit(ch) || ch is '-' or '_'));

        private static bool IsExcelOpenXml(IFormFile f)
        {
            if (f is null || f.Length == 0) return false;
            var ext = Path.GetExtension(f.FileName).ToLowerInvariant();
            return ext == ".xlsx" || ext == ".xlsm";
        }

        private async Task<(bool found, string compCd, string compName, int? deptId, string? deptName, int adminLevel, string userName)> GetUserContextAsync()
        {
            var uid = CurrentUserId();
            if (string.IsNullOrEmpty(uid))
                return (false, string.Empty, string.Empty, null, null, 0, string.Empty);

            var profile = await _db.UserProfiles.AsNoTracking().FirstOrDefaultAsync(p => p.UserId == uid);
            var userName = User?.Identity?.Name ?? profile?.DisplayName ?? string.Empty;

            string compCd = (profile?.CompCd ?? string.Empty).Trim();
            string compName = string.Empty;
            if (!string.IsNullOrWhiteSpace(compCd))
            {
                compName = await _db.CompMasters
                    .Where(c => c.CompCd == compCd)
                    .Select(c => c.Name)
                    .FirstOrDefaultAsync() ?? string.Empty;
            }

            int? deptId = profile?.DepartmentId;
            string? deptName = null;
            if (deptId.HasValue)
            {
                deptName = await _db.DepartmentMasters
                    .Where(d => d.Id == deptId.Value)
                    .Select(d => d.Name)
                    .FirstOrDefaultAsync();
            }

            var adminLevel = profile?.IsAdmin ?? 0;
            return (true, compCd, compName, deptId, deptName, adminLevel, userName);
        }

        private static bool IsJson(string? s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            try { using var _ = JsonDocument.Parse(s); return true; }
            catch { return false; }
        }

        public sealed class DocTLMapDxViewModel
        {
            public string DescriptorJson { get; set; } = "{}";
            public string ExcelPath { get; set; } = string.Empty;
            public string TemplateTitle { get; set; } = string.Empty;
            public string DocCode { get; set; } = string.Empty;
            public string DxDocumentId { get; set; } = string.Empty;
            public string DxCallbackUrl { get; set; } = "/DocumentTemplatesDX/dx-callback";
        }

        public sealed class CellRef
        {
            public string Sheet { get; set; } = string.Empty;
            public int Row { get; set; }
            public int Column { get; set; }
            public int RowSpan { get; set; } = 1;
            public int ColSpan { get; set; } = 1;
            public string A1 { get; set; } = string.Empty;
        }

        public sealed class FieldDef
        {
            public string Key { get; set; } = string.Empty;
            public string Type { get; set; } = "Text";
            public CellRef Cell { get; set; } = new();
        }

        public sealed class ApprovalDef
        {
            public int Slot { get; set; }
            public string Part { get; set; } = string.Empty;
            public CellRef Cell { get; set; } = new();
            public string ApproverType { get; set; } = "Person";
            public string? ApproverValue { get; set; } = string.Empty;
        }

        public sealed class TemplateDescriptor
        {
            public string CompCd { get; set; } = string.Empty;
            public int? DepartmentId { get; set; }
            public string? Kind { get; set; }
            public string DocName { get; set; } = string.Empty;
            public string? Title { get; set; }
            public int ApprovalCount { get; set; }
            public List<FieldDef> Fields { get; set; } = new();
            public List<ApprovalDef> Approvals { get; set; } = new();
            public List<ApprovalDef> Cooperations { get; set; } = new();
        }

        [HttpGet("SearchUser")]
        public async Task<IActionResult> SearchUser(
    string? q, int take = 50, string? id = null, string? compCd = null)
        {
            take = Math.Clamp(take, 1, 200);

            var myUserId = User.FindFirstValue(ClaimTypes.NameIdentifier);

            var myProfile = await _db.UserProfiles
     .AsNoTracking()
     .Where(x => x.UserId == myUserId)
     .Select(x => new { x.CompCd, x.IsAdmin })
     .FirstOrDefaultAsync();

            var isAdmin = (myProfile?.IsAdmin ?? 0) >= 1;

            // 관리자면 무조건 전체, 일반이면 본인 CompCd만
            string? filterCompCd = isAdmin
                ? null
                : (!string.IsNullOrWhiteSpace(compCd) ? compCd : myProfile?.CompCd);

            var baseQuery =
                from p in _db.UserProfiles
                join u in _db.Users on p.UserId equals u.Id
                join cm in _db.CompMasters on p.CompCd equals cm.CompCd into gcm
                from cm in gcm.DefaultIfEmpty()
                join d in _db.DepartmentMasters on p.DepartmentId equals d.Id into gd
                from d in gd.DefaultIfEmpty()
                join pos in _db.PositionMasters on p.PositionId equals pos.Id into gpos
                from pos in gpos.DefaultIfEmpty()
                where p.IsApprover
                   && (filterCompCd == null || p.CompCd == filterCompCd)
                   && (cm == null || cm.IsActive)
                   && (d == null || d.IsActive)
                select new
                {
                    UserId = p.UserId,
                    UserName = u.UserName,
                    DisplayName = p.DisplayName ?? u.UserName ?? string.Empty,
                    CompCd = p.CompCd,
                    CompName = cm != null ? cm.Name : string.Empty,
                    DeptName = d != null ? d.Name : string.Empty,
                    PosName = pos != null ? pos.Name : string.Empty,
                    DeptSortOrder = d != null ? d.SortOrder : int.MaxValue,
                    PosSortOrder = pos != null ? pos.SortOrder : int.MaxValue,
                };

            // 표시 텍스트 로컬 헬퍼 (EF 외부에서 사용)
            static string MakeText(bool admin, string compName, string posName,
                                   string displayName, string deptName)
            {
                // 형식: "회사명 직급 이름 (부서)"  ex) (주)ABC 부장 홍길동 (생산)
                // 관리자가 아니면 회사명 생략: "직급 이름 (부서)"
                var sb = new System.Text.StringBuilder();
                if (admin && !string.IsNullOrEmpty(compName))
                    sb.Append(compName).Append(' ');
                if (!string.IsNullOrEmpty(posName))
                    sb.Append(posName).Append(' ');
                sb.Append(displayName);
                if (!string.IsNullOrEmpty(deptName))
                    sb.Append(" (").Append(deptName).Append(')');
                return sb.ToString();
            }

            // ── 단건 조회
            if (!string.IsNullOrWhiteSpace(id))
            {
                var rows = await baseQuery
                    .Where(x => x.UserId == id)
                    .Take(1)
                    .ToListAsync();

                var one = rows.Select(x => new
                {
                    id = x.UserId,
                    text = MakeText(isAdmin, x.CompName, x.PosName, x.DisplayName, x.DeptName)
                }).ToList();

                return Json(one);
            }

            // ── 검색어 필터
            if (!string.IsNullOrWhiteSpace(q))
            {
                var qq = q.Trim();
                baseQuery = baseQuery.Where(x =>
                    x.DisplayName.Contains(qq) ||
                    (x.UserName ?? string.Empty).Contains(qq) ||
                    x.DeptName.Contains(qq) ||
                    x.CompName.Contains(qq) ||
                    x.PosName.Contains(qq));
            }

            // ── 정렬 후 목록
            var rows2 = await baseQuery
                .OrderBy(x => x.CompCd)
                .ThenBy(x => x.DeptSortOrder)
                .ThenBy(x => x.PosSortOrder)
                .ThenBy(x => x.DisplayName)
                .Take(take)
                .ToListAsync();

            var list = rows2.Select(x => new
            {
                id = x.UserId,
                text = MakeText(isAdmin, x.CompName, x.PosName, x.DisplayName, x.DeptName)
            }).ToList();

            return Json(list);
        }

        [HttpGet("", Name = "DocumentTemplatesDX.Index")]
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

            vm.CompOptions = await _db.CompMasters
                .OrderBy(c => c.CompCd)
                .Select(c => new SelectListItem
                {
                    Value = c.CompCd,
                    Text = c.Name,
                    Selected = c.CompCd == ctx.compCd
                })
                .ToListAsync();

            vm.DepartmentOptions.Add(new SelectListItem
            {
                Value = "__SELECT__",
                Text = $"-- {_S["_CM_Select"]} --",
                Selected = true
            });
            vm.DepartmentOptions.Add(new SelectListItem
            {
                Value = string.Empty,
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

                if (vm.DepartmentOptions.Any(o => o.Selected))
                {
                    foreach (var o in vm.DepartmentOptions)
                        if (o.Value == "__SELECT__") o.Selected = false;
                }
            }

            return View("~/Views/DocTL/DocTLDX.cshtml", vm);
        }

        [HttpGet("get-departments")]
        public async Task<IActionResult> GetDepartments([FromQuery] string compCd)
        {
            var ctx = await GetUserContextAsync();
            var isAdmin = ctx.adminLevel > 0;

            if (!isAdmin) compCd = ctx.compCd;
            compCd = (compCd ?? string.Empty).Trim();

            var list = new List<object>();
            if (!string.IsNullOrEmpty(compCd))
            {
                var raw = await _db.DepartmentMasters
                    .Where(d => (d.CompCd ?? string.Empty).Trim() == compCd)
                    .OrderBy(d => d.Name)
                    .Select(d => new { d.Id, d.Name })
                    .ToListAsync();

                list.AddRange(raw.Select(d => new { id = d.Id, text = d.Name ?? string.Empty }));
            }

            var items = new List<object>
            {
                new { id = "__SELECT__", text = $"-- {_S["_CM_Select"]} --" },
                new { id = (int?)null, text = $"{_S["_CM_Common"]}" }
            };
            items.AddRange(list);

            var selectedValue =
                (!isAdmin || string.Equals(compCd, ctx.compCd, StringComparison.OrdinalIgnoreCase))
                    ? (ctx.deptId?.ToString() ?? string.Empty)
                    : "__SELECT__";

            return Ok(new { items, selectedValue });
        }

        [HttpGet("get-documents")]
        public async Task<IActionResult> GetDocuments([FromQuery] string compCd, [FromQuery] int? departmentId, [FromQuery] string? kind)
        {
            var ctx = await GetUserContextAsync();
            if (ctx.adminLevel == 0) compCd = ctx.compCd;

            compCd = (compCd ?? string.Empty).Trim();
            var dep = departmentId ?? ctx.deptId ?? 0;

            if (string.IsNullOrEmpty(compCd))
                return Ok(new { items = Array.Empty<object>() });

            var query = _db.DocTemplateMasters
               .AsNoTracking()
               .Where(m => m.CompCd == compCd && (dep == 0 ? (m.DepartmentId == 0) : m.DepartmentId == dep))
               .Where(m => m.IsActive == 1);

            if (!string.IsNullOrWhiteSpace(kind))
                query = query.Where(m => m.KindCode == kind);

            var items = await query
                .OrderBy(m => m.DocName)
                .Select(m => new
                {
                    id = m.DocCode,
                    text = m.DocName + " v" + _db.DocTemplateVersions
                        .Where(v => v.TemplateId == m.Id)
                        .Max(v => (int?)v.VersionNo)!
                        .GetValueOrDefault(1)
                })
                .ToListAsync();

            return Ok(new { items });
        }

        [HttpGet("get-kinds")]
        public async Task<IActionResult> GetKinds([FromQuery] string compCd, [FromQuery] int? departmentId)
        {
            var ctx = await GetUserContextAsync();
            if (ctx.adminLevel == 0) compCd = ctx.compCd;

            compCd = (compCd ?? string.Empty).Trim();
            var deptId = departmentId ?? ctx.deptId ?? 0;
            var lang = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;

            var items = await (
                from k in _db.TemplateKindMasters.AsNoTracking()
                where k.CompCd == compCd
                      && k.IsActive
                      && (deptId == 0 ? k.DepartmentId == 0 : k.DepartmentId == deptId)
                join loc in _db.TemplateKindMasterLoc.AsNoTracking()
                           .Where(l => l.CompCd == compCd && l.LangCode == lang)
                     on new { k.Id, k.DepartmentId } equals new { loc.Id, loc.DepartmentId } into lj
                from l in lj.DefaultIfEmpty()
                select new
                {
                    id = k.Code,
                    text = (l != null && !string.IsNullOrWhiteSpace(l.Name)) ? l.Name : (k.Name ?? string.Empty)
                }
            )
            .OrderBy(x => x.text)
            .ToListAsync();

            return Ok(new { items });
        }

        private async Task<IActionResult?> GuardCompAndDeptAsync(string compCd, int? departmentId)
        {
            compCd = (compCd ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(compCd))
                return BadRequest(new { ok = false, message = "comp required" });

            var compExists = await _db.CompMasters.AnyAsync(c => c.CompCd == compCd);
            if (!compExists)
                return BadRequest(new { ok = false, message = "invalid comp" });

            if (departmentId.HasValue && departmentId.Value != 0)
            {
                var deptOk = await _db.DepartmentMasters.AnyAsync(d => d.Id == departmentId.Value && d.CompCd == compCd);
                if (!deptOk)
                    return BadRequest(new { ok = false, message = "invalid department" });
            }

            return null;
        }

        private async Task<string> GenerateNextKindCodeAsync(string compCd)
        {
            var codes = await _db.TemplateKindMasters
                .AsNoTracking()
                .Where(x => x.CompCd == compCd)
                .Select(x => x.Code)
                .ToListAsync();

            int maxN = 0;
            foreach (var c in codes)
            {
                if (!string.IsNullOrEmpty(c) && c.Length == 5 && c[0] == 'T' && int.TryParse(c.AsSpan(1), out var n))
                {
                    if (n > maxN) maxN = n;
                }
            }

            var next = Math.Max(1, maxN + 1);
            if (next > 9999) throw new InvalidOperationException("Template kind code overflow (max T9999).");
            return $"T{next:D4}";
        }

        [HttpPost("kind-add")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> AddKind(
            [FromForm] string compCd,
            [FromForm] int? departmentId,
            [FromForm] string? nameKo,
            [FromForm] string? nameEn,
            [FromForm] string? nameVi,
            [FromForm] string? nameId,
            [FromForm] string? nameZh)
        {
            var ctx = await GetUserContextAsync();
            if (ctx.adminLevel == 0) compCd = ctx.compCd;

            var guard = await GuardCompAndDeptAsync(compCd, departmentId);
            if (guard is not null) return guard;

            var displayName = FirstNonEmpty(nameKo, nameEn, nameVi, nameId, nameZh);
            if (string.IsNullOrWhiteSpace(displayName))
                return BadRequest(new { ok = false, message = "name required" });

            const int maxAttempts = 5;
            TemplateKindMaster? master = null;

            for (int attempt = 1; attempt <= maxAttempts; attempt++)
            {
                var code = await GenerateNextKindCodeAsync(compCd);

                master = new TemplateKindMaster
                {
                    CompCd = compCd.Trim(),
                    DepartmentId = departmentId ?? 0,
                    Code = code,
                    Name = displayName.Trim(),
                    IsActive = true
                };

                _db.TemplateKindMasters.Add(master);

                try
                {
                    await _db.SaveChangesAsync();
                    break;
                }
                catch (DbUpdateException ex)
                {
                    _db.Entry(master).State = EntityState.Detached;
                    var msg = (ex.InnerException?.Message ?? ex.Message).ToLowerInvariant();
                    var isUnique = msg.Contains("unique") || msg.Contains("duplicate") || msg.Contains("ix_templatekindmasters_compcd_code");
                    if (!isUnique || attempt == maxAttempts) throw;
                    await Task.Delay(20 * attempt);
                }
            }

            var locs = new List<TemplateKindMasterLoc>();

            void AddLoc(string? val, string lang)
            {
                if (!string.IsNullOrWhiteSpace(val))
                {
                    locs.Add(new TemplateKindMasterLoc
                    {
                        Id = master!.Id,
                        CompCd = master.CompCd,
                        DepartmentId = master.DepartmentId,
                        LangCode = lang,
                        Name = val.Trim()
                    });
                }
            }

            AddLoc(nameKo, "ko");
            AddLoc(nameEn, "en");
            AddLoc(nameVi, "vi");
            AddLoc(nameId, "id");
            AddLoc(nameZh, "zh");

            if (locs.Count > 0)
            {
                _db.TemplateKindMasterLoc.AddRange(locs);
                await _db.SaveChangesAsync();
            }

            var langUi = CultureInfo.CurrentUICulture.TwoLetterISOLanguageName;
            var text = langUi switch
            {
                "ko" => nameKo ?? displayName,
                "vi" => nameVi ?? displayName,
                "id" => nameId ?? displayName,
                "zh" => nameZh ?? displayName,
                "en" => nameEn ?? displayName,
                _ => displayName
            };

            return Ok(new { ok = true, item = new { id = master!.Code, text } });
        }

        [HttpGet("new-template")]
        public IActionResult NewTemplate([FromQuery] string? compCd, [FromQuery] int? departmentId)
        {
            var rv = new Microsoft.AspNetCore.Routing.RouteValueDictionary();
            if (!string.IsNullOrWhiteSpace(compCd)) rv["compCd"] = compCd;
            if (departmentId.HasValue) rv["departmentId"] = departmentId.Value;
            rv["openNewTemplate"] = "1";
            return RedirectToRoute("DocumentTemplatesDX.Index", rv);
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
            var dn = wb.DefinedNames.FirstOrDefault(n => string.Equals(n.Name, "F_Title", StringComparison.OrdinalIgnoreCase));
            if (dn != null && dn.Ranges.Any())
                return dn.Ranges.First().FirstCell().GetString().Trim();

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
            public int MaxCooperationSlot { get; set; }
            public List<FieldDef> Fields { get; set; } = new();
            public List<ApprovalDef> Approvals { get; set; } = new();
            public List<ApprovalDef> Cooperations { get; set; } = new();
        }

        private static Dictionary<string, string> ParseCommentTags(string text)
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(text)) return dict;

            foreach (var raw in text.Replace("\r", string.Empty).Split('\n'))
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
                    var text = cell.GetComment().Text?.Trim() ?? string.Empty;
                    var tags = ParseCommentTags(text);
                    if (tags.Count == 0) continue;

                    if (tags.TryGetValue("Title", out var tf) && tf.Equals("true", StringComparison.OrdinalIgnoreCase))
                        result.Title ??= cell.GetString().Trim();

                    if (tags.TryGetValue("Field", out var key) && !string.IsNullOrWhiteSpace(key))
                    {
                        var type = tags.TryGetValue("Type", out var t) && !string.IsNullOrWhiteSpace(t)
                            ? NormalizeType(t)
                            : TryInferTypeFromValidationCX(ws, cell) ?? "Text";

                        result.Fields.Add(new FieldDef
                        {
                            Key = key.Trim(),
                            Type = type,
                            Cell = ToCellRef(cell)
                        });
                        continue;
                    }

                    if (tags.TryGetValue("Approval", out var slotStr) &&
                        int.TryParse(slotStr, out int slot) &&
                        tags.TryGetValue("Part", out var part) &&
                        !string.IsNullOrWhiteSpace(part))
                    {
                        result.Approvals.Add(new ApprovalDef
                        {
                            Slot = slot,
                            Part = part.Trim(),
                            Cell = ToCellRef(cell)
                        });

                        if (slot > result.MaxApprovalSlot) result.MaxApprovalSlot = slot;
                        continue;
                    }

                    if (tags.TryGetValue("ApprovalKey", out var ak) && TryParseApprovalKey(ak, out int s, out string p))
                    {
                        result.Approvals.Add(new ApprovalDef
                        {
                            Slot = s,
                            Part = p,
                            Cell = ToCellRef(cell)
                        });

                        if (s > result.MaxApprovalSlot) result.MaxApprovalSlot = s;
                        continue;
                    }

                    if (tags.TryGetValue("Cooperation", out var coopSlotStr) &&
                        int.TryParse(coopSlotStr, out int coopSlot) &&
                        tags.TryGetValue("Part", out var coopPart) &&
                        !string.IsNullOrWhiteSpace(coopPart))
                    {
                        result.Cooperations.Add(new ApprovalDef
                        {
                            Slot = coopSlot,
                            Part = coopPart.Trim(),
                            Cell = ToCellRef(cell)
                        });

                        if (coopSlot > result.MaxCooperationSlot) result.MaxCooperationSlot = coopSlot;
                        continue;
                    }

                    if (tags.TryGetValue("CooperationKey", out var ck) && TryParseCooperationKey(ck, out int cs, out string cp))
                    {
                        result.Cooperations.Add(new ApprovalDef
                        {
                            Slot = cs,
                            Part = cp,
                            Cell = ToCellRef(cell)
                        });

                        if (cs > result.MaxCooperationSlot) result.MaxCooperationSlot = cs;
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

            result.Cooperations = result.Cooperations
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
            slot = 0;
            part = string.Empty;

            if (string.IsNullOrWhiteSpace(input)) return false;

            var m = Regex.Match(input, @"^A(\d+)_(\w+)$", RegexOptions.IgnoreCase);
            if (!m.Success) return false;

            slot = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
            part = m.Groups[2].Value;
            return true;
        }

        private static bool TryParseCooperationKey(string input, out int slot, out string part)
        {
            slot = 0;
            part = string.Empty;

            if (string.IsNullOrWhiteSpace(input)) return false;

            var m = Regex.Match(input, @"^C(\d+)_(\w+)$", RegexOptions.IgnoreCase);
            if (!m.Success) return false;

            slot = int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
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

        [HttpPost("new-template")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> NewTemplatePost([FromForm] string? compCd, [FromForm] int? departmentId, [FromForm] string? kind, [FromForm] string docName, [FromForm] IFormFile? excelFile, [FromForm] bool embed = false)
        {
            var ctx = await GetUserContextAsync();
            if (ctx.adminLevel == 0) compCd = ctx.compCd;

            if (string.IsNullOrWhiteSpace(compCd))
            {
                TempData["Alert"] = _S["_Alert_Require_ValidSite"].Value;
                return RedirectToRoute("DocumentTemplatesDX.Index");
            }

            if (string.IsNullOrWhiteSpace(docName))
            {
                TempData["Alert"] = _S["DTL_Alert_EnterDocName"].Value;
                return RedirectToRoute("DocumentTemplatesDX.Index");
            }

            if (excelFile is null || excelFile.Length == 0)
            {
                TempData["Alert"] = _S["DTL_Alert_ExcelRequired"].Value;
                return RedirectToRoute("DocumentTemplatesDX.Index");
            }

            if (!IsExcelOpenXml(excelFile))
            {
                TempData["Alert"] = _S["DTL_Alert_ExcelOpenXmlOnly"].Value;
                return RedirectToRoute("DocumentTemplatesDX.Index");
            }

            var baseDir = Path.Combine(_env.ContentRootPath, "App_Data", "DocTemplates", "files");
            Directory.CreateDirectory(baseDir);

            static string SafeFile(string s) => string.Concat((s ?? string.Empty).Where(ch => char.IsLetterOrDigit(ch) || ch is '-' or '_'));

            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var ext = Path.GetExtension(excelFile.FileName).ToLowerInvariant();
            var fileName = $"{SafeFile(compCd!)}_{SafeFile(docName)}_{stamp}_{Guid.NewGuid():N}{ext}";
            var excelPath = Path.Combine(baseDir, fileName);

            await using (var fs = System.IO.File.Create(excelPath))
            {
                await excelFile.CopyToAsync(fs);
            }

            using var wb = new XLWorkbook(excelPath);

            // ── 모든 시트의 활성 셀을 A1으로 초기화하여 저장
            foreach (var ws in wb.Worksheets)
            {
                try
                {
                    ws.SetTabActive(false);
                    ws.ActiveCell = ws.Cell("A1");
                    // 스크롤 위치도 A1으로
                    ws.SheetView.TopLeftCellAddress = ws.Cell("A1").Address;
                }
                catch { }
            }
            // 첫 번째 시트 활성화
            if (wb.Worksheets.Any())
            {
                wb.Worksheets.First().SetTabActive(true);
            }
            wb.Save();

            var meta = ReadMetaCX(wb);
            var parsed = ParseByCommentsCX(wb);

            if (meta.ApprovalCount == null)
                meta.ApprovalCount = parsed.MaxApprovalSlot;

            var title = parsed.Title ?? ResolveTitleByNameOrMetaCX(wb, meta);

            var descriptor = new TemplateDescriptor
            {
                CompCd = compCd!,
                DepartmentId = departmentId ?? 0,
                Kind = kind,
                DocName = docName,
                Title = title,
                ApprovalCount = meta.ApprovalCount ?? 0,
                Fields = parsed.Fields,
                Approvals = parsed.Approvals,
                Cooperations = parsed.Cooperations
            };

            var descriptorJsonPretty = JsonSerializer.Serialize(descriptor, new JsonSerializerOptions { WriteIndented = true });

            ViewBag.ExcelPath = excelPath;
            ViewBag.DescriptorJson = descriptorJsonPretty;
            ViewBag.TemplateTitle = docName;
            ViewBag.DocCode = string.Empty;
            ViewBag.DxDocumentId = SafeDocId($"{(CurrentUserId() ?? "anon")}_{Path.GetFileName(excelPath)}");
            ViewBag.DxCallbackUrl = "/DocumentTemplatesDX/dx-callback";

            const string viewPath = "~/Views/DocTL/DocTLMapDX.cshtml";
            var isAjax = string.Equals(Request.Headers["X-Requested-With"].ToString(), "XMLHttpRequest", StringComparison.OrdinalIgnoreCase);

            if (embed || isAjax)
                return PartialView(viewPath);

            return View(viewPath);
        }

        [HttpPost("upload-excel")]
        [ValidateAntiForgeryToken]
        [RequestSizeLimit(50L * 1024L * 1024L)]
        public async Task<IActionResult> UploadExcel([FromForm] IFormFile file, [FromForm] string? docCode)
        {
            if (file == null || file.Length <= 0)
                return BadRequest("No file.");

            var ext = (Path.GetExtension(file.FileName) ?? string.Empty).ToLowerInvariant();
            if (ext != ".xlsx" && ext != ".xlsm" && ext != ".xls")
                return BadRequest("Invalid extension.");

            var uid = CurrentUserId() ?? "anon";
            var safeBase = SafeDocId($"{uid}_{(docCode ?? "DOC")}_{DateTime.Now:yyyyMMdd_HHmmss}");
            var safeName = SafeDocId(Path.GetFileNameWithoutExtension(file.FileName));
            if (string.IsNullOrWhiteSpace(safeName)) safeName = "upload";

            var baseDir = Path.Combine(_env.ContentRootPath, "App_Data", "DocTemplates", "Uploads");
            Directory.CreateDirectory(baseDir);

            var saveAbs = Path.Combine(baseDir, $"{safeBase}_{safeName}{ext}");
            await using (var fs = new FileStream(saveAbs, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                await file.CopyToAsync(fs);
            }

            var openRel = Path.GetRelativePath(_env.ContentRootPath, saveAbs);
            return Ok(new
            {
                openRel,
                fileName = Path.GetFileName(saveAbs),
                byteSize = file.Length
            });
        }

        [HttpGet("load-template")]
        public async Task<IActionResult> LoadTemplateDX(
            [FromQuery] string? compCd,
            [FromQuery] int? departmentId,
            [FromQuery] string? kind,
            [FromQuery] string? docCode,
            [FromQuery] bool embed = false,
            [FromQuery] string? openRel = null)
        {
            var ctx = await GetUserContextAsync();

            if (ctx.adminLevel == 0)
                compCd = ctx.compCd ?? string.Empty;

            compCd = (compCd ?? string.Empty).Trim();
            docCode = (docCode ?? string.Empty).Trim();

            if (string.IsNullOrEmpty(compCd) || string.IsNullOrWhiteSpace(docCode))
                return BadRequest("Invalid request.");

            var dep = departmentId ?? ctx.deptId ?? 0;

            var master = await _db.DocTemplateMasters
                .AsNoTracking()
                .Where(m => m.CompCd == compCd && (dep == 0 ? (m.DepartmentId == 0) : m.DepartmentId == dep))
                .Where(m => m.DocCode == docCode)
                .Select(m => new { m.Id, m.DocName, m.DocCode })
                .FirstOrDefaultAsync();

            if (master == null)
                return NotFound("Template not found.");

            var latest = await _db.DocTemplateVersions
                .AsNoTracking()
                .Where(v => v.TemplateId == master.Id)
                .OrderByDescending(v => v.VersionNo)
                .ThenByDescending(v => v.Id)
                .Select(v => new { v.Id, v.VersionNo })
                .FirstOrDefaultAsync();

            if (latest == null)
                return NotFound("No version found.");

            var wantRoles = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "DescriptorJson", "DescriptorJSON", "Descriptor",
                "ExcelFile", "Excel"
            };

            var files = await _db.DocTemplateFiles
                .AsNoTracking()
                .Where(f => f.VersionId == latest.Id)
                .Where(f => wantRoles.Contains((f.FileRole ?? string.Empty).Trim()))
                .Select(f => new { f.FileRole, f.Storage, f.FilePath, f.Contents })
                .ToListAsync();

            string? ReadFileByRole(params string[] roles)
            {
                var set = roles.Select(r => (r ?? string.Empty).Trim().ToLowerInvariant()).ToHashSet();
                var meta = files.FirstOrDefault(f => set.Contains((f.FileRole ?? string.Empty).Trim().ToLowerInvariant()));
                if (meta == null) return null;

                if (!string.IsNullOrWhiteSpace(meta.Contents))
                    return meta.Contents;

                if (string.Equals(meta.Storage, "Disk", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(meta.FilePath) &&
                    System.IO.File.Exists(meta.FilePath))
                    return System.IO.File.ReadAllText(meta.FilePath);

                return null;
            }

            var descriptorJsonRaw = ReadFileByRole("DescriptorJson", "DescriptorJSON", "Descriptor");
            if (!IsJson(descriptorJsonRaw))
                return BadRequest("Descriptor parse error.");

            string? excelPath = null;

            if (!string.IsNullOrWhiteSpace(openRel))
            {
                try
                {
                    var abs = Path.GetFullPath(Path.Combine(_env.ContentRootPath, openRel));
                    var allowBase = Path.GetFullPath(Path.Combine(_env.ContentRootPath, "App_Data", "DocTemplates", "Uploads"));
                    if (abs.StartsWith(allowBase, StringComparison.OrdinalIgnoreCase) && System.IO.File.Exists(abs))
                        excelPath = abs;
                }
                catch { }
            }

            if (string.IsNullOrWhiteSpace(excelPath))
            {
                var excelMeta = files.FirstOrDefault(f =>
                    string.Equals((f.FileRole ?? string.Empty).Trim(), "ExcelFile", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals((f.FileRole ?? string.Empty).Trim(), "Excel", StringComparison.OrdinalIgnoreCase));

                if (excelMeta != null &&
                    string.Equals(excelMeta.Storage, "Disk", StringComparison.OrdinalIgnoreCase) &&
                    !string.IsNullOrWhiteSpace(excelMeta.FilePath))
                {
                    var normalized = NormalizeTemplateExcelPathLocal(excelMeta.FilePath);
                    var abs = ToContentRootAbsoluteLocal(normalized);
                    if (System.IO.File.Exists(abs))
                        excelPath = abs;
                }
            }

            if (string.IsNullOrWhiteSpace(excelPath))
            {
                await using var conn = _db.Database.GetDbConnection();
                if (conn.State != ConnectionState.Open) await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT TOP (1) ExcelFilePath FROM dbo.DocTemplateVersion WHERE Id = @vid";

                var pVid = cmd.CreateParameter();
                pVid.ParameterName = "@vid";
                pVid.Value = latest.Id;
                cmd.Parameters.Add(pVid);

                var scalar = await cmd.ExecuteScalarAsync();
                if (scalar != null && scalar != DBNull.Value)
                {
                    var normalized = NormalizeTemplateExcelPathLocal(Convert.ToString(scalar));
                    var abs = ToContentRootAbsoluteLocal(normalized);
                    if (System.IO.File.Exists(abs))
                        excelPath = abs;
                }
            }

            if (string.IsNullOrWhiteSpace(excelPath) || !System.IO.File.Exists(excelPath))
                return NotFound("Excel file not found.");

            var uid = CurrentUserId() ?? "anon";
            var dxDocId = SafeDocId($"{uid}_{master.DocCode}_{latest.Id}_{Path.GetFileName(excelPath)}");

            var vm = new DocTLMapDxViewModel
            {
                DescriptorJson = descriptorJsonRaw ?? "{}",
                ExcelPath = excelPath,
                TemplateTitle = master.DocName ?? string.Empty,
                DocCode = master.DocCode ?? string.Empty,
                DxDocumentId = dxDocId,
                DxCallbackUrl = "/DocumentTemplatesDX/dx-callback"
            };

            ViewBag.ExcelPath = vm.ExcelPath;
            ViewBag.DescriptorJson = vm.DescriptorJson;
            ViewBag.TemplateTitle = vm.TemplateTitle;
            ViewBag.DocCode = vm.DocCode;
            ViewBag.DxDocumentId = vm.DxDocumentId;
            ViewBag.DxCallbackUrl = vm.DxCallbackUrl;

            const string viewPath = "~/Views/DocTL/DocTLMapDX.cshtml";
            var isAjax = string.Equals(Request.Headers["X-Requested-With"].ToString(), "XMLHttpRequest", StringComparison.OrdinalIgnoreCase);

            if (embed || isAjax) return PartialView(viewPath, vm);
            return View(viewPath, vm);
        }

        private static readonly Regex _reA1Range = new(@"^([A-Z]+)(\d+):([A-Z]+)(\d+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static bool TryParseA1Range(string? a1, out string col, out int r1, out int r2)
        {
            col = string.Empty;
            r1 = 0;
            r2 = 0;

            var m = _reA1Range.Match((a1 ?? string.Empty).Trim().ToUpperInvariant());
            if (!m.Success) return false;
            if (!string.Equals(m.Groups[1].Value, m.Groups[3].Value, StringComparison.OrdinalIgnoreCase)) return false;

            col = m.Groups[1].Value;
            r1 = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            r2 = int.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture);
            if (r2 < r1) (r1, r2) = (r2, r1);

            return true;
        }

        private static int ColLettersToIndex(string letters)
        {
            if (string.IsNullOrWhiteSpace(letters)) return 0;

            int n = 0;
            foreach (var ch in letters.Trim().ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') break;
                n = n * 26 + (ch - 'A' + 1);
            }

            return n;
        }

        private static List<FieldDef> ExpandFieldsRange(IEnumerable<FieldDef> fields)
        {
            var outList = new List<FieldDef>();

            foreach (var f in fields ?? Enumerable.Empty<FieldDef>())
            {
                if (f is null) continue;

                var a1 = f.Cell?.A1 ?? string.Empty;
                if (!TryParseA1Range(a1, out var colLetters, out var r1, out var r2))
                {
                    outList.Add(f);
                    continue;
                }

                var tpl = (f.Key ?? string.Empty).Trim();
                var baseKey = string.IsNullOrEmpty(tpl) ? "Field" : tpl;
                var idx = 1;
                var col0 = ColLettersToIndex(colLetters) - 1;

                for (int r = r1; r <= r2; r++, idx++)
                {
                    var key = baseKey;
                    if (key.Contains("{n}", StringComparison.Ordinal)) key = key.Replace("{n}", idx.ToString(CultureInfo.InvariantCulture));
                    if (key.Contains("{row}", StringComparison.Ordinal)) key = key.Replace("{row}", r.ToString(CultureInfo.InvariantCulture));
                    if (key == baseKey) key = $"{baseKey}_{r}";

                    outList.Add(new FieldDef
                    {
                        Key = key,
                        Type = string.IsNullOrWhiteSpace(f.Type) ? "Text" : f.Type,
                        Cell = new CellRef
                        {
                            Sheet = f.Cell?.Sheet ?? string.Empty,
                            Row = r - 1,
                            Column = col0,
                            RowSpan = 1,
                            ColSpan = 1,
                            A1 = $"{colLetters}{r}"
                        }
                    });
                }
            }

            return outList;
        }

        private static (bool Ok, string? Reason, string? Key) ValidateFieldKeys(IEnumerable<FieldDef> fields)
        {
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var f in fields ?? Enumerable.Empty<FieldDef>())
            {
                var k = (f?.Key ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(k)) return (false, "empty", null);
                if (!seen.Add(k)) return (false, "dup", k);
            }

            return (true, null, null);
        }

        private string CreateRevisionExcelPath(
            string? sourceExcelPath,
            string compCd,
            string docName,
            string docCode,
            int? departmentId
        )
        {
            var now = DateTime.Now;

            var safeCompCd = SafeFilePart(string.IsNullOrWhiteSpace(compCd) ? "0000" : compCd.Trim());
            var safeDocName = SafeFilePart(string.IsNullOrWhiteSpace(docName) ? "Template" : docName.Trim());
            var safeDocCode = SafeFilePart(string.IsNullOrWhiteSpace(docCode) ? $"DOC_{Guid.NewGuid():N}".ToUpperInvariant() : docCode.Trim());

            var deptCd = (departmentId.HasValue && departmentId.Value > 0)
                ? departmentId.Value.ToString()
                : "0";

            var ext = ".xlsx";
            try
            {
                var srcExt = Path.GetExtension(sourceExcelPath ?? string.Empty);
                if (!string.IsNullOrWhiteSpace(srcExt))
                    ext = srcExt;
            }
            catch
            {
                ext = ".xlsx";
            }

            var rootDir = Path.Combine(
                _env.ContentRootPath,
                "App_Data",
                "DocTemplates",
                safeCompCd,
                now.ToString("yyyy"),
                now.ToString("MM"),
                SafeFilePart(deptCd)
            );

            Directory.CreateDirectory(rootDir);

            var fileName = $"{safeCompCd}_{safeDocName}_{safeDocCode}_{now:yyyyMMdd_HHmmss}{ext}";
            var fullPath = Path.Combine(rootDir, fileName);

            var seq = 1;
            while (System.IO.File.Exists(fullPath))
            {
                fileName = $"{safeCompCd}_{safeDocName}_{safeDocCode}_{now:yyyyMMdd_HHmmss}_{seq}{ext}";
                fullPath = Path.Combine(rootDir, fileName);
                seq++;
            }

            return fullPath;
        }

        private string? CreateRevisionExcelSnapshot(string? sourceExcelPath, string compCd, string docName, string docCode, int? departmentId)
        {
            if (string.IsNullOrWhiteSpace(sourceExcelPath)) return null;
            if (!System.IO.File.Exists(sourceExcelPath)) return null;

            var revisionPath = CreateRevisionExcelPath(
                sourceExcelPath,
                compCd,
                docName,
                docCode,
                departmentId
            );

            System.IO.File.Copy(sourceExcelPath, revisionPath, overwrite: false);
            return revisionPath;
        }

        [HttpPost("map-save")]
        [ValidateAntiForgeryToken]
        public IActionResult MapSave(
            [FromForm] string descriptor,
            [FromForm] string? excelPath,
            [FromForm] string? previewJson,
            [FromForm] string? docCode,
            [FromForm] SpreadsheetClientState? spreadsheetState
        )
        {
            var swAll = Stopwatch.StartNew();

            if (string.IsNullOrWhiteSpace(descriptor))
                return BadRequest("No descriptor");

            TemplateDescriptor? model;
            try { model = JsonSerializer.Deserialize<TemplateDescriptor>(descriptor); }
            catch { return BadRequest("Invalid descriptor"); }
            if (model == null) return BadRequest("Empty descriptor");

            var docCodeFromForm = (FirstNonEmpty(
                docCode,
                Request.Form["docCode"].ToString(),
                Request.Form["DocCode"].ToString(),
                Request.Form["templateCode"].ToString(),
                Request.Form["TemplateCode"].ToString(),
                Request.Form["docTemplateCode"].ToString(),
                Request.Form["DocTemplateCode"].ToString()
            ) ?? string.Empty).Trim();

            var docCodeToSave = !string.IsNullOrWhiteSpace(docCodeFromForm)
                ? docCodeFromForm
                : $"DOC_{Guid.NewGuid():N}".ToUpperInvariant();

            try
            {
                bool needComp = string.IsNullOrWhiteSpace(model.CompCd);
                bool needName = string.IsNullOrWhiteSpace(model.DocName);
                bool needDept = !model.DepartmentId.HasValue;
                bool needKind = string.IsNullOrWhiteSpace(model.Kind);

                if (needComp || needName || needDept || needKind)
                {
                    var m = _db.DocTemplateMasters
                        .AsNoTracking()
                        .Where(x => x.DocCode == docCodeToSave)
                        .Select(x => new { x.CompCd, x.DepartmentId, x.KindCode, x.DocName })
                        .FirstOrDefault();

                    if (m != null)
                    {
                        if (needComp) model.CompCd = (m.CompCd ?? string.Empty).Trim();
                        if (needName) model.DocName = (m.DocName ?? string.Empty).Trim();
                        if (needDept) model.DepartmentId = m.DepartmentId;
                        if (needKind) model.Kind = m.KindCode;
                    }
                }
            }
            catch
            {
            }

            if (string.IsNullOrWhiteSpace(model.CompCd) || string.IsNullOrWhiteSpace(model.DocName))
                return BadRequest("Invalid descriptor (CompCd/DocName)");

            var deptIdToSave = model.DepartmentId ?? 0;

            model.Fields = ExpandFieldsRange(model.Fields ?? new List<FieldDef>());
            var ck = ValidateFieldKeys(model.Fields);
            if (!ck.Ok)
            {
                return BadRequest(ck.Reason == "dup"
                    ? $"Duplicate field key: {ck.Key}"
                    : "Field key is required.");
            }

            var apprs = model.Approvals ??= new List<ApprovalDef>();
            model.ApprovalCount = apprs.Count;

            for (int i = 0; i < model.Approvals.Count; i++)
            {
                var a = model.Approvals[i] ?? new ApprovalDef();
                a.Cell ??= new CellRef();

                a.ApproverType = string.IsNullOrWhiteSpace(a.ApproverType) ? "Person" : a.ApproverType;
                if (a.ApproverType != "Person" && a.ApproverType != "Role" && a.ApproverType != "Rule")
                    a.ApproverType = "Person";

                a.ApproverValue ??= string.Empty;

                if (a.ApproverType == "Rule")
                {
                    try { using var _ = JsonDocument.Parse(a.ApproverValue); }
                    catch { return BadRequest($"Approvals[{i}] 규칙 JSON이 올바르지 않습니다."); }
                }

                if ((a.ApproverType == "Person" || a.ApproverType == "Role") &&
                    string.IsNullOrWhiteSpace(a.ApproverValue))
                {
                    return BadRequest($"Approvals[{i}] {(a.ApproverType == "Person" ? "사용자ID" : "역할코드")}를 입력해 주세요.");
                }

                if (a.Slot <= 0) a.Slot = 1;
                model.Approvals[i] = a;
            }

            var coops = model.Cooperations ??= new List<ApprovalDef>();
            for (int i = 0; i < model.Cooperations.Count; i++)
            {
                var c = model.Cooperations[i] ?? new ApprovalDef();
                c.Cell ??= new CellRef();

                c.ApproverType = string.IsNullOrWhiteSpace(c.ApproverType) ? "Person" : c.ApproverType;
                if (c.ApproverType != "Person" && c.ApproverType != "Role" && c.ApproverType != "Rule")
                    c.ApproverType = "Person";

                c.ApproverValue ??= string.Empty;

                if (c.ApproverType == "Rule")
                {
                    try { using var _ = JsonDocument.Parse(c.ApproverValue); }
                    catch { return BadRequest($"Cooperations[{i}] 규칙 JSON이 올바르지 않습니다."); }
                }

                if ((c.ApproverType == "Person" || c.ApproverType == "Role") &&
                    string.IsNullOrWhiteSpace(c.ApproverValue))
                {
                    return BadRequest($"Cooperations[{i}] {(c.ApproverType == "Person" ? "사용자ID" : "역할코드")}를 입력해 주세요.");
                }

                if (c.Slot <= 0) c.Slot = 1;
                model.Cooperations[i] = c;
            }

            var fileSetStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            var tempExcelPath = CreateRevisionExcelPath(
                excelPath,
                model.CompCd,
                model.DocName,
                docCodeToSave,
                deptIdToSave
            );

            try
            {
                if (spreadsheetState != null)
                {
                    var spreadsheet = SpreadsheetRequestProcessor.GetSpreadsheetFromState(spreadsheetState);
                    spreadsheet.SaveCopy(tempExcelPath);
                }
                else if (!string.IsNullOrWhiteSpace(excelPath) && System.IO.File.Exists(excelPath))
                {
                    System.IO.File.Copy(excelPath, tempExcelPath, false);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[DocTLDX][MapSave] SaveCopy failed: " + ex);
                if (!string.IsNullOrWhiteSpace(excelPath) && System.IO.File.Exists(excelPath) && !System.IO.File.Exists(tempExcelPath))
                {
                    System.IO.File.Copy(excelPath, tempExcelPath, false);
                }
            }

            if (!System.IO.File.Exists(tempExcelPath))
            {
                TempData["Alert"] = "Excel save failed: temp excel file does not exist.";
                return RedirectToAction(nameof(MapSaved), new
                {
                    path = "",
                    excelPath = tempExcelPath,
                    fields = 0,
                    approvals = 0
                });
            }

            try
            {
                previewJson = DocControllerHelper.BuildPreviewJsonFromExcel(tempExcelPath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[DocTLDX][MapSave] BuildPreviewJsonFromExcel failed: " + ex);
                previewJson = "{}";
            }

            var descriptorJson = JsonSerializer.Serialize(model, new JsonSerializerOptions { WriteIndented = false });

            var tempExcelFileName = Path.GetFileName(tempExcelPath);
            var tempExcelSize = System.IO.File.Exists(tempExcelPath)
                ? new FileInfo(tempExcelPath).Length
                : 0L;

            int outTemplateId = 0;
            long outVersionId = 0;
            int outVersionNo = 0;

            try
            {
                var pOutTemplateId = new SqlParameter("@OutTemplateId", SqlDbType.Int) { Direction = ParameterDirection.Output };
                var pOutVersionId = new SqlParameter("@OutVersionId", SqlDbType.BigInt) { Direction = ParameterDirection.Output };
                var pOutVersionNo = new SqlParameter("@OutVersionNo", SqlDbType.Int) { Direction = ParameterDirection.Output };

                var p = new[]
                {
                    new SqlParameter("@CompCd", SqlDbType.NVarChar, 10){ Value = model.CompCd },
                    new SqlParameter("@DepartmentId", SqlDbType.Int){ Value = deptIdToSave },
                    new SqlParameter("@KindCode", SqlDbType.NVarChar, 20){ Value = (object?)model.Kind ?? DBNull.Value },
                    new SqlParameter("@DocCode", SqlDbType.NVarChar, 40){ Value = docCodeToSave },
                    new SqlParameter("@DocName", SqlDbType.NVarChar, 200){ Value = model.DocName },
                    new SqlParameter("@Title", SqlDbType.NVarChar, 200){ Value = (object?)model.Title ?? DBNull.Value },
                    new SqlParameter("@ApprovalCount", SqlDbType.Int){ Value = model.ApprovalCount },
                    new SqlParameter("@DescriptorJson", SqlDbType.NVarChar, -1){ Value = descriptorJson },
                    new SqlParameter("@PreviewJson", SqlDbType.NVarChar, -1){ Value = (object?)previewJson ?? DBNull.Value },
                    new SqlParameter("@ExcelFileName", SqlDbType.NVarChar, 255){ Value = tempExcelFileName },
                    new SqlParameter("@ExcelStorage", SqlDbType.NVarChar, 20){ Value = "Disk" },
                    new SqlParameter("@ExcelFilePath", SqlDbType.NVarChar, 500){ Value = tempExcelPath },
                    new SqlParameter("@ExcelFileSize", SqlDbType.BigInt){ Value = tempExcelSize },
                    new SqlParameter("@ExcelContentType", SqlDbType.NVarChar, 100){ Value = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                    new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 100){ Value = (User?.Identity?.Name ?? "system") },
                    pOutTemplateId,
                    pOutVersionId,
                    pOutVersionNo
                };

                const string sql =
                    "EXEC dbo.sp_DocTemplate_SaveFromDescriptor " +
                    "@CompCd,@DepartmentId,@KindCode,@DocCode,@DocName,@Title,@ApprovalCount," +
                    "@DescriptorJson,@PreviewJson,@ExcelFileName,@ExcelStorage,@ExcelFilePath,@ExcelFileSize,@ExcelContentType," +
                    "@CreatedBy,@OutTemplateId OUTPUT,@OutVersionId OUTPUT,@OutVersionNo OUTPUT";

                _db.Database.ExecuteSqlRaw(sql, p);

                if (pOutTemplateId.Value != DBNull.Value) outTemplateId = Convert.ToInt32(pOutTemplateId.Value);
                if (pOutVersionId.Value != DBNull.Value) outVersionId = Convert.ToInt64(pOutVersionId.Value);
                if (pOutVersionNo.Value != DBNull.Value) outVersionNo = Convert.ToInt32(pOutVersionNo.Value);
            }
            catch (Exception ex)
            {
                TempData["Alert"] = ex.Message;
                return RedirectToAction(nameof(MapSaved), new
                {
                    path = "",
                    excelPath = tempExcelPath,
                    fields = 0,
                    approvals = 0
                });
            }

            var finalExcelPath = tempExcelPath;
            try
            {
                finalExcelPath = BuildVersionedExcelPath(
                    tempExcelPath,
                    model.CompCd,
                    model.DocName,
                    docCodeToSave,
                    outVersionNo,
                    fileSetStamp
                );

                if (!string.Equals(tempExcelPath, finalExcelPath, StringComparison.OrdinalIgnoreCase)
                    && System.IO.File.Exists(tempExcelPath))
                {
                    var finalDir = Path.GetDirectoryName(finalExcelPath);
                    if (!string.IsNullOrWhiteSpace(finalDir))
                        Directory.CreateDirectory(finalDir);

                    System.IO.File.Move(tempExcelPath, finalExcelPath);
                }

                if (!System.IO.File.Exists(finalExcelPath))
                {
                    throw new IOException("Excel rename failed. Final excel file does not exist.");
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[DocTLDX][MapSave] Excel rename failed: " + ex);
                TempData["Alert"] = "Excel finalization failed: " + ex.Message;

                return RedirectToAction(nameof(MapSaved), new
                {
                    path = "",
                    excelPath = tempExcelPath,
                    fields = 0,
                    approvals = 0
                });
            }

            var finalExcelFileName = Path.GetFileName(finalExcelPath);
            var finalExcelSize = System.IO.File.Exists(finalExcelPath)
                ? new FileInfo(finalExcelPath).Length
                : 0L;
            var finalExcelRelPath = ToAppDataRelativePath(finalExcelPath);

            try
            {
                const string sqlUpdateVersion = @"
UPDATE dbo.DocTemplateVersion
   SET ExcelFileName = @ExcelFileName,
       ExcelFilePath = @ExcelFilePath,
       ExcelFileSize = @ExcelFileSize
 WHERE Id = @VersionId;";

                _db.Database.ExecuteSqlRaw(
                    sqlUpdateVersion,
                    new SqlParameter("@ExcelFileName", SqlDbType.NVarChar, 255) { Value = finalExcelFileName },
                    new SqlParameter("@ExcelFilePath", SqlDbType.NVarChar, 500) { Value = finalExcelRelPath },
                    new SqlParameter("@ExcelFileSize", SqlDbType.BigInt) { Value = finalExcelSize },
                    new SqlParameter("@VersionId", SqlDbType.BigInt) { Value = outVersionId }
                );

                const string sqlUpdateFile = @"
UPDATE dbo.DocTemplateFile
   SET FileName = @ExcelFileName,
       FilePath = @ExcelFilePath,
       FileSize = @ExcelFileSize,
       FileSizeBytes = @ExcelFileSize
 WHERE VersionId = @VersionId
   AND FileRole = N'ExcelFile';";

                _db.Database.ExecuteSqlRaw(
                    sqlUpdateFile,
                    new SqlParameter("@ExcelFileName", SqlDbType.NVarChar, 255) { Value = finalExcelFileName },
                    new SqlParameter("@ExcelFilePath", SqlDbType.NVarChar, 500) { Value = finalExcelRelPath },
                    new SqlParameter("@ExcelFileSize", SqlDbType.BigInt) { Value = finalExcelSize },
                    new SqlParameter("@VersionId", SqlDbType.BigInt) { Value = outVersionId }
                );
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[DocTLDX][MapSave] Excel DB sync update failed: " + ex);
            }

            var jsonDir = Path.GetDirectoryName(finalExcelPath)
                          ?? Path.Combine(_env.ContentRootPath, "App_Data", "DocTemplates");
            Directory.CreateDirectory(jsonDir);

            var vtag = (outVersionNo > 0) ? $"v{outVersionNo:D4}" : "v0000";
            var jsonFileName = $"{SafeFilePart(model.CompCd)}_{SafeFilePart(model.DocName)}_{SafeFilePart(docCodeToSave)}_{vtag}_{fileSetStamp}.json";
            var path = Path.Combine(jsonDir, jsonFileName);

            var snap = new
            {
                docCode = docCodeToSave,
                outTemplateId,
                outVersionId,
                outVersionNo,
                excelPath = finalExcelPath,
                model
            };

            System.IO.File.WriteAllText(
                path,
                JsonSerializer.Serialize(snap, new JsonSerializerOptions { WriteIndented = true })
            );

            Debug.WriteLine($"[DocTLDX][MapSave] saved docCode={docCodeToSave} templateId={outTemplateId} versionId={outVersionId} versionNo={outVersionNo} fields={(model.Fields?.Count ?? 0)} approvals={(model.Approvals?.Count ?? 0)} cooperations={(model.Cooperations?.Count ?? 0)} elapsed={swAll.ElapsedMilliseconds}ms");
            Debug.WriteLine($"[DocTLDX][MapSave] spreadsheetStateNull={(spreadsheetState == null)}");
            Debug.WriteLine($"[DocTLDX][MapSave] fileSetStamp={fileSetStamp}");
            Debug.WriteLine($"[DocTLDX][MapSave] tempExcelPath={tempExcelPath}");
            Debug.WriteLine($"[DocTLDX][MapSave] finalExcelPath={finalExcelPath}");
            Debug.WriteLine($"[DocTLDX][MapSave] saved snapshot path={path}");

            return RedirectToAction(nameof(MapSaved), new
            {
                path,
                excelPath = finalExcelPath,
                fields = model.Fields?.Count ?? 0,
                approvals = model.Approvals?.Count ?? 0
            });
        }

        private string NormalizeTemplateExcelPathLocal(string? raw)
        {
            var s = (raw ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(s)) return s;

            if (s.StartsWith("App_Data\\", StringComparison.OrdinalIgnoreCase) ||
                s.StartsWith("App_Data/", StringComparison.OrdinalIgnoreCase))
                return s.Replace('/', '\\');

            var idx = s.IndexOf("App_Data\\", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) idx = s.IndexOf("App_Data/", StringComparison.OrdinalIgnoreCase);

            return idx >= 0 ? s.Substring(idx).Replace('/', '\\') : s;
        }

        private string ToContentRootAbsoluteLocal(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;

            var normalized = path.Replace('/', Path.DirectorySeparatorChar)
                                 .Replace('\\', Path.DirectorySeparatorChar);

            return Path.IsPathRooted(normalized)
                ? normalized
                : Path.Combine(_env.ContentRootPath, normalized);
        }

        private string BuildVersionedExcelPath(
            string currentExcelPath,
            string compCd,
            string docName,
            string docCode,
            int versionNo,
            string stamp
        )
        {
            var dir = Path.GetDirectoryName(currentExcelPath)
                      ?? Path.Combine(_env.ContentRootPath, "App_Data", "DocTemplates");

            Directory.CreateDirectory(dir);

            var ext = Path.GetExtension(currentExcelPath);
            if (string.IsNullOrWhiteSpace(ext)) ext = ".xlsx";

            var safeStamp = string.IsNullOrWhiteSpace(stamp)
                ? DateTime.Now.ToString("yyyyMMdd_HHmmss")
                : stamp.Trim();

            var vtag = (versionNo > 0) ? $"v{versionNo:D4}" : "v0000";

            var fileName = $"{SafeFilePart(compCd)}_{SafeFilePart(docName)}_{SafeFilePart(docCode)}_{vtag}_{safeStamp}{ext}";
            var fullPath = Path.Combine(dir, fileName);

            var seq = 1;
            while (System.IO.File.Exists(fullPath))
            {
                fileName = $"{SafeFilePart(compCd)}_{SafeFilePart(docName)}_{SafeFilePart(docCode)}_{vtag}_{safeStamp}_{seq}{ext}";
                fullPath = Path.Combine(dir, fileName);
                seq++;
            }

            return fullPath;
        }

        private string ToAppDataRelativePath(string fullPath)
        {
            var normalized = (fullPath ?? string.Empty).Replace('/', '\\');
            var token = "\\App_Data\\";
            var pos = normalized.IndexOf(token, StringComparison.OrdinalIgnoreCase);
            if (pos >= 0)
            {
                return "App_Data\\" + normalized.Substring(pos + token.Length);
            }

            return normalized;
        }

        [HttpGet("map-saved")]
        public IActionResult MapSaved(string path, string? excelPath, int fields, int approvals)
        {
            ViewBag.Path = path;
            ViewBag.ExcelPath = excelPath;
            ViewBag.Fields = fields;
            ViewBag.Approvals = approvals;
            ViewBag.RouteBase = "/DocumentTemplatesDX";
            return View("~/Views/DocTL/DocTLMapSaved.cshtml");
        }

        [HttpGet("download-descriptor")]
        public IActionResult DownloadDescriptor([FromQuery] string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path))
                return NotFound();

            var bytes = System.IO.File.ReadAllBytes(path);
            var fileName = Path.GetFileName(path);
            return File(bytes, "application/json", fileName);
        }
    }
}