// 2025.09.29 Changed: MapSave()에서 생성한 docCode를 @DocCode 파라미터에 그대로 사용하도록 통일
// 2025.09.29 Changed: ExpandFieldsRange()의 Row/Column을 0-기반으로 일치화(기존 ToCellRef와 기준 통일). A1 표기는 기존대로 1-기반 유지

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
using System.Globalization;
using Microsoft.AspNetCore.Routing;
using System.Net;
using System.Linq;
using System.Data;
using Microsoft.Data.SqlClient;

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

                if (vm.DepartmentOptions.Any(o => o.Selected))
                {
                    foreach (var o in vm.DepartmentOptions)
                        if (o.Value == "__SELECT__") o.Selected = false;
                }
            }

            return View("DocTL", vm);
        }

        [HttpGet("get-departments")]
        public async Task<IActionResult> GetDepartments([FromQuery] string compCd)
        {
            var ctx = await GetUserContextAsync();
            var isAdmin = ctx.adminLevel > 0;

            if (!isAdmin) compCd = ctx.compCd;
            compCd = (compCd ?? "").Trim();

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

            if (ctx.deptId.HasValue
                && string.Equals(compCd, ctx.compCd, StringComparison.OrdinalIgnoreCase)
                && list.All(o => (o as dynamic).id != ctx.deptId.Value))
            {
                var mine = await _db.DepartmentMasters
                    .Where(d => d.Id == ctx.deptId.Value && (d.CompCd ?? "").Trim() == compCd)
                    .Select(d => new { d.Id, d.Name })
                    .FirstOrDefaultAsync();

                if (mine != null)
                    list.Insert(0, new { id = mine.Id, text = mine.Name ?? "" });
            }

            var items = new List<object>
            {
                new { id = "__SELECT__", text = $"-- {_S["_CM_Select"]} --" },
                new { id = (int?)null,   text = $"{_S["_CM_Common"]}" }
            };
            items.AddRange(list);

            var selectedValue =
                (!isAdmin || string.Equals(compCd, ctx.compCd, StringComparison.OrdinalIgnoreCase))
                    ? (ctx.deptId?.ToString() ?? "")
                    : "__SELECT__";

            return Ok(new { items, selectedValue });
        }

        [HttpGet("get-documents")]
        public async Task<IActionResult> GetDocuments(
                    [FromQuery] string compCd,
                    [FromQuery] int? departmentId,
                    [FromQuery] string? kind)
        {
            var ctx = await GetUserContextAsync();
            if (ctx.adminLevel == 0) compCd = ctx.compCd;

            compCd = (compCd ?? "").Trim();
            var dep = departmentId ?? ctx.deptId ?? 0;

            if (string.IsNullOrEmpty(compCd))
                return Ok(new { items = Array.Empty<object>() });

            var query = _db.DocTemplateMasters
                .AsNoTracking()
                 .Where(m => m.CompCd == compCd &&
        (dep == 0
            ? (m.DepartmentId == 0)   // 2025.10.01 Removed: || m.DepartmentId == null
            : m.DepartmentId == dep));

            if (!string.IsNullOrWhiteSpace(kind))
                query = query.Where(m => m.KindCode == kind);

            var items = await query
                .OrderBy(m => m.DocName)
                .Select(m => new { id = m.DocCode, text = m.DocName })
                .ToListAsync();

            return Ok(new { items });
        }

        // =====================================================================
        // 2025.09.30 Added: UI 범위 표기를 위한 Collapse 유틸(저장 데이터는 그대로, 화면에만 적용)
        // =====================================================================
        // 0-기반 컬럼 인덱스를 A1 열 문자로 변환
        private static string ColIndexToLetters(int col0)
        {
            int n = col0 + 1;
            string s = string.Empty;
            while (n > 0)
            {
                int rem = (n - 1) % 26;
                s = (char)('A' + rem) + s;
                n = (n - 1) / 26;
            }
            return s;
        }

        private static readonly Regex _reA1Single =
            new(@"^([A-Z]+)(\d+)$", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static void NormalizeCellFromA1IfNeeded(FieldDef f)
        {
            if (f?.Cell is null) return;
            var a1 = f.Cell.A1?.Trim();
            if (string.IsNullOrWhiteSpace(a1)) return;

            // 단일 셀 & UI 병합 판단에 쓰일 값들을 '항상' A1에서 다시 세팅
            if (_reA1Single.IsMatch(a1) && f.Cell.RowSpan <= 1 && f.Cell.ColSpan <= 1)
            {
                var m = _reA1Single.Match(a1.ToUpperInvariant());
                var col0 = ColLettersToIndex(m.Groups[1].Value) - 1; // 0-기반
                var r0 = int.Parse(m.Groups[2].Value) - 1;           // 0-기반
                f.Cell.Row = r0;
                f.Cell.Column = col0;
                f.Cell.RowSpan = 1;
                f.Cell.ColSpan = 1;
            }
        }

        // 키가 base_숫자 형태인지 분리
        private static bool TrySplitKeyBaseNumber(string key, out string baseKey, out int num)
        {
            baseKey = key;
            num = 0;
            var m = Regex.Match(key ?? string.Empty, @"^(.*)_([0-9]+)$");
            if (!m.Success) return false;
            baseKey = m.Groups[1].Value;
            return int.TryParse(m.Groups[2].Value, out num);
        }

        // 연속 행(같은 시트/열, RowSpan/ColSpan=1) 묶어서 A1 범위로 표시
        private static List<FieldDef> CollapseFieldsForUI(IEnumerable<FieldDef> fields)
        {
            var list = (fields ?? Enumerable.Empty<FieldDef>())
                       .Where(f => f != null)
                       .Select(f => { NormalizeCellFromA1IfNeeded(f); return f; })
                       .OrderBy(f => f.Cell.Sheet)
                       .ThenBy(f => f.Cell.Column)
                       .ThenBy(f => f.Cell.Row)
                       .ToList();

            var result = new List<FieldDef>();
            int i = 0;
            while (i < list.Count)
            {
                var f0 = list[i];
                var sheet = f0.Cell.Sheet;
                var col0 = f0.Cell.Column;

                // 2025.09.30 Changed: 단일셀 판정 완화 (<=1 이면 단일셀로 처리)
                bool isSingle0 = (f0.Cell.ColSpan <= 1) && (f0.Cell.RowSpan <= 1);
                if (!isSingle0)
                {
                    result.Add(f0);
                    i++;
                    continue;
                }

                string baseKey = f0.Key;
                int n0 = 0;
                bool hasBase = TrySplitKeyBaseNumber(f0.Key, out var baseFromKey, out n0);
                if (hasBase) baseKey = baseFromKey;

                int start = f0.Cell.Row;
                int end = start;
                int j = i + 1;

                while (j < list.Count)
                {
                    var f = list[j];
                    if (f.Cell.Sheet != sheet || f.Cell.Column != col0) break;

                    // 2025.09.30 Changed: 단일셀 판정 완화
                    bool isSingle = (f.Cell.ColSpan <= 1) && (f.Cell.RowSpan <= 1);
                    if (!isSingle) break;

                    if (f.Cell.Row != end + 1) break;

                    if (hasBase)
                    {
                        if (!TrySplitKeyBaseNumber(f.Key, out var b2, out var n2)) break;
                        if (!string.Equals(b2, baseKey, StringComparison.OrdinalIgnoreCase)) break;
                        if (n2 != n0 + (f.Cell.Row - start)) break;
                    }

                    end = f.Cell.Row;
                    j++;
                }

                if (end > start)
                {
                    var letters = ColIndexToLetters(col0);
                    var a1 = $"{letters}{start + 1}:{letters}{end + 1}";
                    result.Add(new FieldDef
                    {
                        Key = hasBase ? $"{baseKey}_{{n}}" : f0.Key,
                        Type = f0.Type,
                        Cell = new CellRef
                        {
                            Sheet = sheet,
                            Row = start,
                            Column = col0,
                            RowSpan = (end - start + 1),
                            ColSpan = 1,
                            A1 = a1
                        }
                    });
                    i = j;
                }
                else
                {
                    result.Add(f0);
                    i++;
                }
            }
            return result;
        }

        // =====================================================================
        [HttpGet("load-template")]
        public async Task<IActionResult> LoadTemplate(
   [FromQuery] string compCd,
   [FromQuery] int? departmentId,
   [FromQuery] string docCode,
   [FromQuery] bool embed = false) // 본문 임베드용
        {
            var ctx = await GetUserContextAsync();
            if (ctx.adminLevel == 0) compCd = ctx.compCd;

            compCd = (compCd ?? "").Trim();
            if (string.IsNullOrEmpty(compCd) || string.IsNullOrWhiteSpace(docCode))
            {
                TempData["Alert"] = "Invalid request.";
                return RedirectToRoute("DocumentTemplates.Index");
            }

            var dep = departmentId ?? ctx.deptId ?? 0;
            var master = await _db.DocTemplateMasters
                .AsNoTracking()
                .Where(m => m.CompCd == compCd &&
                            (dep == 0 ? m.DepartmentId == 0 : m.DepartmentId == dep) &&
                            m.DocCode == docCode)
                .FirstOrDefaultAsync();

            if (master == null)
            {
                TempData["Alert"] = "Template not found.";
                return RedirectToRoute("DocumentTemplates.Index");
            }

            var latest = await _db.DocTemplateVersions
                .AsNoTracking()
                .Where(v => v.TemplateId == master.Id)
                .OrderByDescending(v => v.VersionNo)
                .FirstOrDefaultAsync();

            if (latest == null)
            {
                TempData["Alert"] = "No version found.";
                return RedirectToRoute("DocumentTemplates.Index");
            }

            // Excel 메타 포함 조회
            var files = await _db.DocTemplateFiles
                .AsNoTracking()
                .Where(f => f.VersionId == latest.Id)
                .Select(f => new { f.FileRole, f.Storage, f.FilePath, f.Contents })
                .ToListAsync();

            string? descriptorJsonRaw = files.FirstOrDefault(f => f.FileRole == "DescriptorJson")?.Contents;
            string? previewJson = files.FirstOrDefault(f => f.FileRole == "PreviewJson")?.Contents;

            if (string.IsNullOrWhiteSpace(descriptorJsonRaw))
            {
                TempData["Alert"] = "Descriptor not found.";
                return RedirectToRoute("DocumentTemplates.Index");
            }

            TemplateDescriptor? desc;
            try { desc = JsonSerializer.Deserialize<TemplateDescriptor>(descriptorJsonRaw); }
            catch
            {
                TempData["Alert"] = "Descriptor parse error.";
                return RedirectToRoute("DocumentTemplates.Index");
            }

            if (desc != null && desc.Fields != null && desc.Fields.Count > 0)
            {
                var collapsed = CollapseFieldsForUI(desc.Fields);
                desc.Fields = collapsed;
            }
            var descriptorJsonForUi = JsonSerializer.Serialize(desc, new JsonSerializerOptions { WriteIndented = true });

            // PreviewJson 누락 시 Excel 원본에서 생성 시도
            if (string.IsNullOrWhiteSpace(previewJson))
            {
                try
                {
                    var excelMeta = files.FirstOrDefault(f => f.FileRole == "ExcelFile");
                    if (excelMeta != null && string.Equals(excelMeta.Storage, "Disk", StringComparison.OrdinalIgnoreCase))
                    {
                        var path = excelMeta.FilePath;
                        if (!string.IsNullOrWhiteSpace(path) && System.IO.File.Exists(path))
                        {
                            using var wb = new XLWorkbook(path);
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

                                previewJson = JsonSerializer.Serialize(new
                                {
                                    sheet = ws0.Name,
                                    rows = PREV_MAX_ROWS,
                                    cols = PREV_MAX_COLS,
                                    cells,
                                    merges,
                                    colW
                                });
                            }
                        }
                    }
                }
                catch
                {
                    // 실패 시 아래 EnsurePreviewJson에서 폴백 처리
                }
            }

            // 표준화 및 최종 폴백
            previewJson = EnsurePreviewJson(previewJson);

            ViewBag.DescriptorJson = descriptorJsonForUi;
            ViewBag.PreviewJson = previewJson;
            ViewData["DescriptorJson"] = descriptorJsonForUi;
            ViewData["PreviewJson"] = previewJson;

            // 타이틀 전달
            ViewBag.TemplateTitle = master.DocName;

            // 임베드 또는 AJAX 요청이면 Partial 반환
            if (embed || string.Equals(Request.Headers["X-Requested-With"], "XMLHttpRequest", StringComparison.OrdinalIgnoreCase))
                return PartialView("DocTLMap");

            return View("DocTLMap");
        }

        [HttpGet("get-kinds")]
        public async Task<IActionResult> GetKinds([FromQuery] string compCd, [FromQuery] int? departmentId)
        {
            var ctx = await GetUserContextAsync();
            if (ctx.adminLevel == 0) compCd = ctx.compCd;

            compCd = (compCd ?? "").Trim();
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
                    text = (l != null && !string.IsNullOrWhiteSpace(l.Name)) ? l.Name : (k.Name ?? "")
                }
            )
            .OrderBy(x => x.text)
            .ToListAsync();

            return Ok(new { items });
        }

        private static string? FirstNonEmpty(params string?[] xs) => xs.FirstOrDefault(s => !string.IsNullOrWhiteSpace(s));

        private async Task<IActionResult?> GuardCompAndDeptAsync(string compCd, int? departmentId)
        {
            compCd = (compCd ?? "").Trim();
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

            const int MAX_ATTEMPTS = 5;
            TemplateKindMaster? master = null;
            for (int attempt = 1; attempt <= MAX_ATTEMPTS; attempt++)
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
                    if (!isUnique || attempt == MAX_ATTEMPTS) throw;
                    await Task.Delay(20 * attempt);
                }
            }

            var locs = new List<TemplateKindMasterLoc>();
            void AddLoc(string? val, string lang)
            {
                if (!string.IsNullOrWhiteSpace(val))
                    locs.Add(new TemplateKindMasterLoc
                    {
                        Id = master!.Id,
                        CompCd = master!.CompCd,
                        DepartmentId = master!.DepartmentId,
                        LangCode = lang,
                        Name = val.Trim()
                    });
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
                    if (n > maxN) maxN = n;
            }

            var next = Math.Max(1, maxN + 1);
            if (next > 9999) throw new InvalidOperationException("Template kind code overflow (max T9999).");
            return $"T{next:D4}";
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
                                                         [FromForm] IFormFile? excelFile,
                                                         [FromForm] bool embed = false) // ★ 추가: 임베드 지원
        {
            var ctx = await GetUserContextAsync();
            if (ctx.adminLevel == 0) compCd = ctx.compCd;

            if (string.IsNullOrWhiteSpace(compCd))
            {
                TempData["Alert"] = _S["_Alert_Require_ValidSite"].Value;
                return RedirectToRoute("DocumentTemplates.Index");
            }
            if (string.IsNullOrWhiteSpace(docName))
            {
                TempData["Alert"] = _S["DTL_Alert_EnterDocName"].Value;
                return RedirectToRoute("DocumentTemplates.Index");
            }
            if (excelFile is null || excelFile.Length == 0)
            {
                TempData["Alert"] = _S["DTL_Alert_ExcelRequired"].Value;
                return RedirectToRoute("DocumentTemplates.Index");
            }
            if (!IsExcelOpenXml(excelFile))
            {
                TempData["Alert"] = _S["DTL_Alert_ExcelOpenXmlOnly"].Value;
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

                // 2025.09.30 Added: 뷰에서 ViewData 경로도 인식하도록 동시 설정
                ViewData["PreviewJson"] = ViewBag.PreviewJson;
            }

            var descriptor = new TemplateDescriptor
            {
                CompCd = compCd!,
                DepartmentId = departmentId ?? 0,
                Kind = kind,
                DocName = docName,
                Title = title,
                ApprovalCount = meta.ApprovalCount ?? 0,
                Fields = parsed.Fields,
                Approvals = parsed.Approvals
            };

            ViewBag.DescriptorJson = JsonSerializer.Serialize(descriptor, new JsonSerializerOptions { WriteIndented = true });
            ViewData["DescriptorJson"] = ViewBag.DescriptorJson; // 2025.09.30 Added

            // ★ 타이틀 전달
            ViewBag.TemplateTitle = docName;

            // ★ 임베드 요청이면 Partial 반환
            if (embed || string.Equals(Request.Headers["X-Requested-With"], "XMLHttpRequest", StringComparison.OrdinalIgnoreCase))
                return PartialView("DocTLMap");

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
            var dn = wb.DefinedNames.FirstOrDefault(n =>
                string.Equals(n.Name, "F_Title", StringComparison.OrdinalIgnoreCase));
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

        // =====================================================================
        // 2025.09.15 [추가] A1 범위 파싱/확장 + 키 검증 유틸 (서버-사이드 안전망)
        // =====================================================================
        private static readonly Regex _reA1Range = new(@"^([A-Z]+)(\d+):([A-Z]+)(\d+)$",
                                                       RegexOptions.Compiled | RegexOptions.IgnoreCase);

        private static bool TryParseA1Range(string? a1, out string col, out int r1, out int r2)
        {
            col = string.Empty; r1 = 0; r2 = 0;
            var m = _reA1Range.Match((a1 ?? string.Empty).Trim().ToUpperInvariant());
            if (!m.Success) return false;
            if (!string.Equals(m.Groups[1].Value, m.Groups[3].Value, StringComparison.OrdinalIgnoreCase)) return false;
            col = m.Groups[1].Value;
            r1 = int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture);
            r2 = int.Parse(m.Groups[4].Value, CultureInfo.InvariantCulture);
            if (r2 < r1) (r1, r2) = (r2, r1);
            return true;
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

                int col0 = ColLettersToIndex(colLetters) - 1; // 0-기반

                for (int r = r1; r <= r2; r++, idx++)
                {
                    var key = baseKey;
                    if (key.Contains("{n}", StringComparison.Ordinal)) key = key.Replace("{n}", idx.ToString(CultureInfo.InvariantCulture));
                    if (key.Contains("{row}", StringComparison.Ordinal)) key = key.Replace("{row}", r.ToString(CultureInfo.InvariantCulture));
                    if (key == baseKey) key = $"{baseKey}_{r}";

                    outList.Add(new FieldDef
                    {
                        Key = key,
                        Type = string.IsNullOrWhiteSpace(f?.Type) ? "Text" : f!.Type,
                        Cell = new CellRef
                        {
                            Sheet = f?.Cell?.Sheet ?? string.Empty,
                            Row = r - 1,     // 0-기반으로 '항상' 재계산
                            Column = col0,   // 0-기반으로 '항상' 재계산
                            RowSpan = 1,
                            ColSpan = 1,
                            A1 = $"{colLetters}{r}" // A1 표기는 1-기반
                        }
                    });
                }
            }
            return outList;
        }

        private static int ColLettersToIndex(string letters)
        {
            if (string.IsNullOrWhiteSpace(letters))
                return 0;
            int n = 0;
            foreach (var ch in letters.Trim().ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z') break;
                n = n * 26 + (ch - 'A' + 1);
            }
            return n;
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
        private static string BuildBlankPreviewJson(int rows = 50, int cols = 26)
        {
            var cells = new List<List<string>>(rows);
            for (int r = 0; r < rows; r++)
            {
                var row = new List<string>(cols);
                for (int c = 0; c < cols; c++) row.Add(string.Empty);
                cells.Add(row);
            }
            var colW = new List<double>(cols);
            for (int c = 0; c < cols; c++) colW.Add(12.0);

            return JsonSerializer.Serialize(new
            {
                sheet = "Sheet1",
                rows,
                cols,
                cells,
                merges = Array.Empty<int[]>(),
                colW
            });
        }
        private static string EnsurePreviewJson(string? previewJson)
        {
            if (string.IsNullOrWhiteSpace(previewJson))
                return BuildBlankPreviewJson();

            try
            {
                using var doc = JsonDocument.Parse(previewJson);
                var root = doc.RootElement;

                // 키 대소문자 무시 접근용 로컬 함수
                bool TryGetProp(string name, out JsonElement value)
                {
                    foreach (var p in root.EnumerateObject())
                    {
                        if (string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase))
                        {
                            value = p.Value;
                            return true;
                        }
                    }
                    value = default;
                    return false;
                }

                // rows
                int rows = 0, cols = 0;
                if (TryGetProp("rows", out var jeRows) && jeRows.TryGetInt32(out var r)) rows = r;
                if (rows <= 0 && TryGetProp("rowCount", out var jeRowCnt) && jeRowCnt.TryGetInt32(out r)) rows = r;

                // cols
                if (TryGetProp("cols", out var jeCols) && jeCols.TryGetInt32(out var c)) cols = c;
                if (cols <= 0 && TryGetProp("colCount", out var jeColCnt) && jeColCnt.TryGetInt32(out c)) cols = c;

                // cells (2차원 배열)
                List<List<string>> cells = new();
                if (TryGetProp("cells", out var jeCells) && jeCells.ValueKind == JsonValueKind.Array)
                {
                    foreach (var row in jeCells.EnumerateArray())
                    {
                        var line = new List<string>();
                        if (row.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var cell in row.EnumerateArray())
                                line.Add(cell.ValueKind == JsonValueKind.String ? cell.GetString() ?? "" : cell.ToString());
                        }
                        cells.Add(line);
                    }
                }
                else if (TryGetProp("Cells", out var jeCellsLegacy) && jeCellsLegacy.ValueKind == JsonValueKind.Array)
                {
                    foreach (var row in jeCellsLegacy.EnumerateArray())
                    {
                        var line = new List<string>();
                        if (row.ValueKind == JsonValueKind.Array)
                        {
                            foreach (var cell in row.EnumerateArray())
                                line.Add(cell.ValueKind == JsonValueKind.String ? cell.GetString() ?? "" : cell.ToString());
                        }
                        cells.Add(line);
                    }
                }

                // rows/cols 보정 (cells에서 유도)
                if (rows <= 0) rows = cells.Count;
                if (cols <= 0) cols = cells.Count > 0 ? cells.Max(line => line.Count) : 0;

                // merges (옵션)
                List<int[]> merges = new();
                if (TryGetProp("merges", out var jeMerges) && jeMerges.ValueKind == JsonValueKind.Array)
                {
                    foreach (var m in jeMerges.EnumerateArray())
                    {
                        if (m.ValueKind == JsonValueKind.Array)
                        {
                            var arr = m.EnumerateArray().Select(x => x.TryGetInt32(out var n) ? n : 0).Take(4).ToArray();
                            if (arr.Length == 4) merges.Add(arr);
                        }
                    }
                }

                // colW (옵션)
                List<double> colW = new();
                if (TryGetProp("colW", out var jeColW) && jeColW.ValueKind == JsonValueKind.Array)
                {
                    foreach (var w in jeColW.EnumerateArray())
                    {
                        if (w.ValueKind == JsonValueKind.Number && w.TryGetDouble(out var dv)) colW.Add(dv);
                        else colW.Add(12.0);
                    }
                }

                // sheet (옵션)
                string sheet = "Sheet1";
                if (TryGetProp("sheet", out var jeSheet) && jeSheet.ValueKind == JsonValueKind.String)
                    sheet = jeSheet.GetString() ?? sheet;

                // 최소 요건 검사: cells가 없거나 행/열이 0이면 블랭크 생성
                if (rows <= 0 || cols <= 0 || cells.Count == 0)
                    return BuildBlankPreviewJson();

                // 표준 스키마로 재직렬화
                return JsonSerializer.Serialize(new { sheet, rows, cols, cells, merges, colW });
            }
            catch
            {
                // 파싱 실패 시 블랭크
                return BuildBlankPreviewJson();
            }
        }
        // =====================================================================

        [HttpPost("map-save")]
        [ValidateAntiForgeryToken]
        public IActionResult MapSave([FromForm] string descriptor,
                              [FromForm] string? excelPath,
                              [FromForm] string? previewJson)
        {
            if (string.IsNullOrWhiteSpace(descriptor))
                return BadRequest("No descriptor");

            TemplateDescriptor? model;
            try { model = JsonSerializer.Deserialize<TemplateDescriptor>(descriptor); }
            catch { return BadRequest("Invalid descriptor"); }
            if (model == null) return BadRequest("Empty descriptor");

            model.Fields = ExpandFieldsRange(model.Fields ?? new List<FieldDef>());
            var ck = ValidateFieldKeys(model.Fields);
            if (!ck.Ok)
            {
                return BadRequest(ck.Reason == "dup"
                    ? $"Duplicate field key: {ck.Key}"
                    : "Field key is required.");
            }

            var descriptorJson = JsonSerializer.Serialize(model, new JsonSerializerOptions { WriteIndented = false });

            var excelFileName = string.IsNullOrWhiteSpace(excelPath) ? "unknown.xlsx" : Path.GetFileName(excelPath);
            var excelSize = (!string.IsNullOrWhiteSpace(excelPath) && System.IO.File.Exists(excelPath))
                ? new FileInfo(excelPath).Length
                : 0L;

            var docCode = $"DOC_{Guid.NewGuid():N}".ToUpperInvariant();
            var deptIdToSave = model.DepartmentId ?? 0;
            try
            {
                var p = new[]
                {
        new SqlParameter("@CompCd", SqlDbType.NVarChar, 10){ Value = model.CompCd },
        // 2025.10.01 Changed: ?? 0 제거 (deptIdToSave 자체가 int)
        new SqlParameter("@DepartmentId", SqlDbType.Int) { Value = deptIdToSave },
        new SqlParameter("@KindCode", SqlDbType.NVarChar, 20){ Value = (object?)model.Kind ?? DBNull.Value },

        // 여기 통일 사용
        new SqlParameter("@DocCode", SqlDbType.NVarChar, 40){ Value = docCode },

        new SqlParameter("@DocName", SqlDbType.NVarChar, 200){ Value = model.DocName },
        new SqlParameter("@Title", SqlDbType.NVarChar, 200){ Value = (object?)model.Title ?? DBNull.Value },
        new SqlParameter("@ApprovalCount", SqlDbType.Int){ Value = model.ApprovalCount },

        new SqlParameter("@DescriptorJson", SqlDbType.NVarChar, -1){ Value = descriptorJson },
        new SqlParameter("@PreviewJson",    SqlDbType.NVarChar, -1){ Value = (object?)previewJson ?? DBNull.Value },

        new SqlParameter("@ExcelFileName",   SqlDbType.NVarChar, 255){ Value = excelFileName },
        new SqlParameter("@ExcelStorage",    SqlDbType.NVarChar, 20 ){ Value = "Disk" },
        new SqlParameter("@ExcelFilePath",   SqlDbType.NVarChar, 500){ Value = (object?)excelPath ?? DBNull.Value },
        new SqlParameter("@ExcelFileSize",   SqlDbType.BigInt){ Value = excelSize },
        new SqlParameter("@ExcelContentType",SqlDbType.NVarChar, 100){ Value = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },

        new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 100){ Value = (User?.Identity?.Name ?? "system") },

        new SqlParameter("@OutTemplateId", SqlDbType.Int){ Direction = ParameterDirection.Output },
        new SqlParameter("@OutVersionId", SqlDbType.BigInt){ Direction = ParameterDirection.Output },
        new SqlParameter("@OutVersionNo", SqlDbType.Int){ Direction = ParameterDirection.Output },
    };

                const string sql =
                    "EXEC dbo.sp_DocTemplate_SaveFromDescriptor " +
                    "@CompCd,@DepartmentId,@KindCode,@DocCode,@DocName,@Title,@ApprovalCount," +
                    "@DescriptorJson,@PreviewJson,@ExcelFileName,@ExcelStorage,@ExcelFilePath,@ExcelFileSize,@ExcelContentType," +
                    "@CreatedBy,@OutTemplateId OUTPUT,@OutVersionId OUTPUT,@OutVersionNo OUTPUT";

                _db.Database.ExecuteSqlRaw(sql, p);

                var outTemplateId = (int)(p[^3].Value ?? 0);
                var outVersionId = (long)(p[^2].Value ?? 0L);
                var outVersionNo = (int)(p[^1].Value ?? 0);

                TempData["Alert"] = $"Saved: TmplId={outTemplateId}, VerId={outVersionId}, VerNo={outVersionNo}";
            }
            catch (Exception ex)
            {
                TempData["Alert"] = ex.Message;
            }

            var baseDir = Path.Combine(_env.ContentRootPath, "App_Data", "DocTemplates");
            Directory.CreateDirectory(baseDir);
            static string Safe(string s) => string.Concat(s.Where(ch => char.IsLetterOrDigit(ch) || ch is '-' or '_'));
            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var name = $"{Safe(model.CompCd)}_{Safe(model.DocName)}_{stamp}_{Guid.NewGuid():N}.json";
            var path = Path.Combine(baseDir, name);

            System.IO.File.WriteAllText(path, JsonSerializer.Serialize(model, new JsonSerializerOptions { WriteIndented = true }));

            return RedirectToAction(nameof(MapSaved), new
            {
                path,
                excelPath,
                fields = model.Fields?.Count ?? 0,
                approvals = model.Approvals?.Count ?? 0
            });
        }

        [HttpGet("map-saved")]
        public IActionResult MapSaved(string path, string? excelPath, int fields, int approvals)
        {
            ViewBag.Path = path;
            ViewBag.ExcelPath = excelPath;
            ViewBag.Fields = fields;
            ViewBag.Approvals = approvals;
            return View("DocTLMapSaved");
        }

        [HttpGet("download-descriptor")]
        public IActionResult DownloadDescriptor([FromQuery] string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path)) return NotFound();
            var bytes = System.IO.File.ReadAllBytes(path);
            var fileName = Path.GetFileName(path);
            return File(bytes, "application/json", fileName);
        }
    }
}
