// 2026.06.16 Changed: 전결 Always 및 AmountLimit 저장 조회 처리를 추가함 Contents 템플릿 버전별 전결 정책을 별도 테이블에 저장하고 로드시 Descriptor에 병합
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
using System.Security.Cryptography;
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

        private const string TemplateProtectionRuleCode = "ResetLockByMappingV1";
        private const string TemplateVisualMetricRuleCode = "OpenXmlRangePxV1";
        private const string ApproverValueDrafter = "__DRAFTER__";

        // 2026.06.18 Added: 전결 허용 통화 코드 목록 추가 Contents 화면 통화 콤보와 서버 저장 검증 기준을 동일하게 유지
        private static readonly HashSet<string> AllowedDelegationCurrencyCodes = new(StringComparer.OrdinalIgnoreCase)
        {
            "KRW",
            "USD",
            "VND",
            "IDR",
            "CNY"
        };

        // 2026.06.18 Added: 기안자 예약값 판정 추가 Contents 실제 사용자 코드가 없는 기안자 항목도 템플릿 저장 검증에서 허용
        private static bool IsDrafterApproverValue(string? value)
            => string.Equals((value ?? string.Empty).Trim(), ApproverValueDrafter, StringComparison.OrdinalIgnoreCase);

        // 2026.06.18 Added: 전결 통화 코드 허용 여부 판정 추가 Contents 기준 통화 코드가 허용 목록에 있는지 서버에서 재검증
        private static bool IsAllowedDelegationCurrencyCode(string? value)
            => AllowedDelegationCurrencyCodes.Contains((value ?? string.Empty).Trim());

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
            return GetAllowedTemplateExcelExtension(f.FileName) != null;
        }

        private static string? GetAllowedTemplateExcelExtension(string? fileNameOrExtension)
        {
            var value = (fileNameOrExtension ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value)) return null;

            string ext;

            try
            {
                if (value.StartsWith(".", StringComparison.Ordinal) &&
                    !value.Contains(Path.DirectorySeparatorChar) &&
                    !value.Contains(Path.AltDirectorySeparatorChar))
                {
                    ext = value.ToLowerInvariant();
                }
                else
                {
                    ext = (Path.GetExtension(value) ?? string.Empty).ToLowerInvariant();
                }
            }
            catch
            {
                return null;
            }

            return ext == ".xlsx" || ext == ".xlsm"
                ? ext
                : null;
        }
        private string GetTemplateUploadsDirectory()
        {
            var dir = Path.Combine(_env.ContentRootPath, "App_Data", "DocTemplates", "Uploads");
            Directory.CreateDirectory(dir);
            return dir;
        }

        private string GetTemplateFinalDirectory(string? compCd, int? departmentId)
        {
            var safeCompCd = SafeFilePart(string.IsNullOrWhiteSpace(compCd) ? "0000" : compCd.Trim());

            var deptPart = departmentId.HasValue && departmentId.Value > 0
                ? departmentId.Value.ToString(CultureInfo.InvariantCulture)
                : "0";

            var safeDeptPart = SafeFilePart(deptPart);

            var dir = Path.Combine(
                _env.ContentRootPath,
                "App_Data",
                "DocTemplates",
                safeCompCd,
                safeDeptPart
            );

            Directory.CreateDirectory(dir);
            return dir;
        }

        private static string NormalizeExcelExtension(string? sourceExcelPathOrExtension)
        {
            return GetAllowedTemplateExcelExtension(sourceExcelPathOrExtension) ?? ".xlsx";
        }

        private void TryRollbackFinalExcelToTemp(string? finalExcelPath, string? tempExcelPath)
        {
            if (string.IsNullOrWhiteSpace(finalExcelPath) || string.IsNullOrWhiteSpace(tempExcelPath))
                return;

            try
            {
                var finalAbs = Path.GetFullPath(finalExcelPath);
                var tempAbs = Path.GetFullPath(tempExcelPath);

                if (string.Equals(finalAbs, tempAbs, StringComparison.OrdinalIgnoreCase))
                    return;

                if (!System.IO.File.Exists(finalAbs))
                    return;

                if (System.IO.File.Exists(tempAbs))
                    return;

                var tempDir = Path.GetDirectoryName(tempAbs);
                if (!string.IsNullOrWhiteSpace(tempDir))
                    Directory.CreateDirectory(tempDir);

                System.IO.File.Move(finalAbs, tempAbs);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[DocTLDX][MapSave] Excel rollback final-to-temp failed: " + ex);
            }
        }

        private static string BuildTemplateExcelFileName(
            string? compCd,
            string? kindCode,
            string? docCode,
            int versionNo,
            string? stamp,
            string? ext,
            int? seq = null)
        {
            var safeCompCd = SafeFilePart(string.IsNullOrWhiteSpace(compCd) ? "0000" : compCd.Trim());
            var safeKindCode = SafeFilePart(string.IsNullOrWhiteSpace(kindCode) ? "T0000" : kindCode.Trim());
            var safeDocCode = SafeFilePart(string.IsNullOrWhiteSpace(docCode) ? $"DOC_{Guid.NewGuid():N}".ToUpperInvariant() : docCode.Trim());

            var vtag = versionNo > 0
                ? $"v{versionNo:D4}"
                : "v0000";

            var safeStamp = string.IsNullOrWhiteSpace(stamp)
                ? DateTime.Now.ToString("yyyyMMdd_HHmmss")
                : SafeFilePart(stamp.Trim());

            var safeExt = string.IsNullOrWhiteSpace(ext)
                ? ".xlsx"
                : ext.Trim().ToLowerInvariant();

            if (!safeExt.StartsWith(".", StringComparison.Ordinal))
                safeExt = "." + safeExt;

            var baseName = $"{safeCompCd}_{safeKindCode}_{safeDocCode}_{vtag}_{safeStamp}";

            if (seq.HasValue && seq.Value > 0)
                baseName += $"_{seq.Value}";

            return baseName + safeExt;
        }

        private static bool IsPathUnderDirectory(string path, string directory)
        {
            if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(directory))
                return false;

            var fullPath = Path.GetFullPath(path);
            var fullDirectory = Path.GetFullPath(directory)
                .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                + Path.DirectorySeparatorChar;

            return fullPath.StartsWith(fullDirectory, StringComparison.OrdinalIgnoreCase);
        }

        private void TryDeleteUploadTempExcel(string? sourceExcelPath, string? finalExcelPath)
        {
            if (string.IsNullOrWhiteSpace(sourceExcelPath))
                return;

            try
            {
                var normalizedSource = NormalizeTemplateExcelPathLocal(sourceExcelPath);
                var sourceAbs = Path.GetFullPath(ToContentRootAbsoluteLocal(normalizedSource));
                var uploadsRoot = GetTemplateUploadsDirectory();

                if (!IsPathUnderDirectory(sourceAbs, uploadsRoot))
                    return;

                if (!string.IsNullOrWhiteSpace(finalExcelPath))
                {
                    var finalAbs = Path.GetFullPath(finalExcelPath);
                    if (string.Equals(sourceAbs, finalAbs, StringComparison.OrdinalIgnoreCase))
                        return;
                }

                if (System.IO.File.Exists(sourceAbs))
                    System.IO.File.Delete(sourceAbs);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[DocTLDX][MapSave] Upload temp file delete failed: " + ex);
            }
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

            // 2026.06.16 Added: 전결권자 표시값 추가 Contents 화면 결재라인에서 전결권자로 선택된 차수를 Descriptor에 보관
            public bool IsDelegationApprover { get; set; }
        }

        // 2026.06.16 Added: 전결 금액 기준 DTO 추가 Contents AmountLimit 조건에서 통화별 기준 금액을 Descriptor로 전달
        public sealed class DelegationAmountLimitDef
        {
            public string CurrencyCode { get; set; } = string.Empty;
            public decimal? LimitAmount { get; set; }
        }

        // 2026.06.16 Added: 전결 설정 DTO 추가 Contents 템플릿 매핑 화면의 전결 조건을 Descriptor로 전달
        public sealed class DelegationDef
        {
            public bool Enabled { get; set; }
            public string ConditionType { get; set; } = "None";
            public int DelegationStepOrder { get; set; }
            public int SkipFromStepOrder { get; set; }
            public int SkipToStepOrder { get; set; }

            public string? AmountFieldKey { get; set; }
            public string? CurrencyFieldKey { get; set; }

            // 2026.06.18 Added: 전결 금액 조건 셀 직접 참조 추가 Contents 입력 필드가 아닌 수식 셀과 통화 셀을 잠금 상태로 유지한 채 조건 비교에 사용
            public string? AmountCellA1 { get; set; }
            public string? CurrencyCellA1 { get; set; }

            public List<DelegationAmountLimitDef> AmountLimits { get; set; } = new();
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

            // 2026.06.16 Added: 전결 설정 추가 Contents 템플릿 버전별 전결 조건을 Descriptor에 포함
            public DelegationDef Delegation { get; set; } = new();
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

        private IActionResult? ValidateTemplateScopeForSave(string? compCd, int? departmentId, string? kindCode)
        {
            var saveCompCd = (compCd ?? string.Empty).Trim();
            var saveDeptId = departmentId ?? 0;
            var saveKindCode = (kindCode ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(saveCompCd))
                return BadRequest("Invalid template scope: CompCd is required.");

            if (string.IsNullOrWhiteSpace(saveKindCode))
                return BadRequest("Invalid template scope: KindCode is required.");

            var compOk = _db.CompMasters
                .AsNoTracking()
                .Any(c => c.CompCd == saveCompCd);

            if (!compOk)
                return BadRequest($"Invalid template scope: CompCd '{saveCompCd}' does not exist.");

            if (saveDeptId != 0)
            {
                var deptOk = _db.DepartmentMasters
                    .AsNoTracking()
                    .Any(d => d.Id == saveDeptId && d.CompCd == saveCompCd);

                if (!deptOk)
                    return BadRequest($"Invalid template scope: DepartmentId '{saveDeptId}' does not belong to CompCd '{saveCompCd}'.");
            }

            var kindOk = _db.TemplateKindMasters
                .AsNoTracking()
                .Any(k =>
                    k.CompCd == saveCompCd &&
                    k.DepartmentId == saveDeptId &&
                    k.Code == saveKindCode &&
                    k.IsActive);

            if (!kindOk)
            {
                return BadRequest(
                    $"Invalid template scope: KindCode '{saveKindCode}' does not belong to CompCd '{saveCompCd}', DepartmentId '{saveDeptId}'."
                );
            }

            return null;
        }

        [HttpPost("new-template")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> NewTemplatePost([FromForm] string? compCd, [FromForm] int? departmentId, [FromForm] string? kind, [FromForm] string docName, [FromForm] IFormFile? excelFile, [FromForm] bool embed = false)
        {
            var ctx = await GetUserContextAsync();

            var saveCompCd = (compCd ?? string.Empty).Trim();
            var saveDeptId = departmentId ?? 0;
            var saveKindCode = (kind ?? string.Empty).Trim();
            var saveDocName = (docName ?? string.Empty).Trim();

            // 일반 사용자가 회사 선택값 없이 접근한 경우에만 본인 회사로 보정합니다.
            // 화면에서 선택된 compCd가 넘어온 경우에는 그 값을 절대 덮어쓰지 않습니다.
            if (string.IsNullOrWhiteSpace(saveCompCd) && !string.IsNullOrWhiteSpace(ctx.compCd))
            {
                saveCompCd = ctx.compCd.Trim();
            }

            if (string.IsNullOrWhiteSpace(saveCompCd))
            {
                TempData["Alert"] = _S["_Alert_Require_ValidSite"].Value;
                return RedirectToRoute("DocumentTemplatesDX.Index");
            }

            if (string.IsNullOrWhiteSpace(saveDocName))
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

            var scopeGuard = ValidateTemplateScopeForSave(saveCompCd, saveDeptId, saveKindCode);
            if (scopeGuard is not null)
                return scopeGuard;

            var baseDir = GetTemplateUploadsDirectory();

            var stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var ext = NormalizeExcelExtension(excelFile.FileName);

            var fileName = string.Join("_", new[]
            {
        SafeFilePart(saveCompCd),
        SafeFilePart(saveKindCode),
        SafeFilePart(saveDocName),
        stamp,
        Guid.NewGuid().ToString("N")
    }) + ext;

            var excelPath = Path.Combine(baseDir, fileName);

            await using (var fs = System.IO.File.Create(excelPath))
            {
                await excelFile.CopyToAsync(fs);
            }

            using var wb = new XLWorkbook(excelPath);

            foreach (var ws in wb.Worksheets)
            {
                try
                {
                    ws.SetTabActive(false);
                    ws.ActiveCell = ws.Cell("A1");
                    ws.SheetView.TopLeftCellAddress = ws.Cell("A1").Address;
                }
                catch
                {
                }
            }

            if (wb.Worksheets.Any())
            {
                wb.Worksheets.First().SetTabActive(true);
            }

            wb.Save();

            var meta = ReadMetaCX(wb);
            var parsed = ParseByCommentsCX(wb);

            if (meta.ApprovalCount == null)
                meta.ApprovalCount = parsed.Approvals.Count;

            var title = ResolveTitleByNameOrMetaCX(wb, meta);
            if (string.IsNullOrWhiteSpace(title))
                title = saveDocName;

            var descriptor = new TemplateDescriptor
            {
                CompCd = saveCompCd,
                DepartmentId = saveDeptId,
                Kind = saveKindCode,
                DocName = saveDocName,
                Title = title,
                ApprovalCount = meta.ApprovalCount ?? 0,
                Fields = parsed.Fields,
                Approvals = parsed.Approvals,
                Cooperations = parsed.Cooperations,
                Delegation = new DelegationDef()
            };

            var descriptorJsonPretty = JsonSerializer.Serialize(
                descriptor,
                new JsonSerializerOptions { WriteIndented = true }
            );

            ViewBag.ExcelPath = excelPath;
            ViewBag.DescriptorJson = descriptorJsonPretty;
            ViewBag.TemplateTitle = saveDocName;
            ViewBag.DocCode = string.Empty;
            ViewBag.DxDocumentId = SafeDocId($"{(CurrentUserId() ?? "anon")}_{Path.GetFileName(excelPath)}");
            ViewBag.DxCallbackUrl = "/DocumentTemplatesDX/dx-callback";

            const string viewPath = "~/Views/DocTL/DocTLMapDX.cshtml";

            var isAjax = string.Equals(
                Request.Headers["X-Requested-With"].ToString(),
                "XMLHttpRequest",
                StringComparison.OrdinalIgnoreCase
            );

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

            var ext = GetAllowedTemplateExcelExtension(file.FileName);
            if (ext == null)
                return BadRequest("Invalid extension.");

            var uid = CurrentUserId() ?? "anon";
            var safeBase = SafeDocId($"{uid}_{(docCode ?? "DOC")}_{DateTime.Now:yyyyMMdd_HHmmss}");
            var safeName = SafeDocId(Path.GetFileNameWithoutExtension(file.FileName));
            if (string.IsNullOrWhiteSpace(safeName)) safeName = "upload";

            var baseDir = GetTemplateUploadsDirectory();
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
                    !string.IsNullOrWhiteSpace(meta.FilePath))
                {
                    var normalized = NormalizeTemplateExcelPathLocal(meta.FilePath);
                    var abs = ToContentRootAbsoluteLocal(normalized);

                    if (System.IO.File.Exists(abs))
                        return System.IO.File.ReadAllText(abs);
                }

                return null;
            }

            var descriptorJsonRaw = ReadFileByRole("DescriptorJson", "DescriptorJSON", "Descriptor");
            if (!IsJson(descriptorJsonRaw))
                return BadRequest("Descriptor parse error.");

            descriptorJsonRaw = await MergeDelegationRuleToDescriptorJsonAsync(descriptorJsonRaw, latest.Id);

            string? excelPath = null;

            if (!string.IsNullOrWhiteSpace(openRel))
            {
                try
                {
                    var abs = Path.GetFullPath(Path.Combine(_env.ContentRootPath, openRel));
                    var allowBase = GetTemplateUploadsDirectory();

                    if (IsPathUnderDirectory(abs, allowBase) && System.IO.File.Exists(abs))
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

        // 2026.06.16 Added: 전결 Descriptor 병합 처리 추가 Contents 저장된 전결 정책을 템플릿 로드시 Descriptor에 포함
        private async Task<string> MergeDelegationRuleToDescriptorJsonAsync(string? descriptorJson, long templateVersionId)
        {
            TemplateDescriptor? descriptor;
            try
            {
                descriptor = JsonSerializer.Deserialize<TemplateDescriptor>(
                    descriptorJson ?? "{}",
                    new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
            }
            catch
            {
                descriptor = null;
            }

            descriptor ??= new TemplateDescriptor();
            descriptor.Fields ??= new List<FieldDef>();
            descriptor.Approvals ??= new List<ApprovalDef>();
            descriptor.Cooperations ??= new List<ApprovalDef>();

            var delegation = await LoadDelegationRuleForVersionAsync(templateVersionId);
            descriptor.Delegation = delegation;

            foreach (var approval in descriptor.Approvals)
            {
                if (approval == null) continue;
                approval.IsDelegationApprover =
                    delegation.Enabled &&
                    delegation.DelegationStepOrder > 0 &&
                    approval.Slot == delegation.DelegationStepOrder;
            }

            return JsonSerializer.Serialize(
                descriptor,
                new JsonSerializerOptions { WriteIndented = true });
        }

        // 2026.06.16 Added: 전결 정책 조회 추가 Contents 템플릿 버전별 전결 마스터와 금액 조건을 조회
        private async Task<DelegationDef> LoadDelegationRuleForVersionAsync(long templateVersionId)
        {
            var empty = new DelegationDef();

            if (templateVersionId <= 0)
                return empty;

            var rule = await _db.DocTemplateDelegationRules
                .AsNoTracking()
                .Where(x => x.TemplateVersionId == templateVersionId && x.IsActive)
                .OrderBy(x => x.Priority)
                .ThenBy(x => x.Id)
                .FirstOrDefaultAsync();

            if (rule == null)
                return empty;

            var delegation = new DelegationDef
            {
                Enabled = true,
                ConditionType = NormalizeDelegationConditionType(rule.ConditionType),
                DelegationStepOrder = rule.DelegationStepOrder,
                SkipFromStepOrder = rule.SkipFromStepOrder,
                SkipToStepOrder = rule.SkipToStepOrder,
                AmountLimits = new List<DelegationAmountLimitDef>()
            };

            if (string.Equals(delegation.ConditionType, "AmountLimit", StringComparison.OrdinalIgnoreCase))
            {
                var amountRules = await _db.DocTemplateDelegationAmountRules
                    .AsNoTracking()
                    .Where(x => x.RuleId == rule.Id && x.IsActive)
                    .OrderBy(x => x.Id)
                    .ToListAsync();

                var first = amountRules.FirstOrDefault();
                if (first != null)
                {
                    delegation.AmountFieldKey = first.AmountFieldKey;
                    delegation.CurrencyFieldKey = first.CurrencyFieldKey;

                    // 2026.06.18 Added: 저장된 전결 조건 셀 주소를 Descriptor로 복원 Contents 전결 금액 조건 화면 재조회 시 금액 셀과 통화 셀을 표시
                    delegation.AmountCellA1 = first.AmountCellA1;
                    delegation.CurrencyCellA1 = first.CurrencyCellA1;
                }

                delegation.AmountLimits = amountRules
                    .Select(x => new DelegationAmountLimitDef
                    {
                        CurrencyCode = x.CurrencyCode,
                        LimitAmount = x.LimitAmount
                    })
                    .ToList();
            }

            return delegation;
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

        private string CreateRevisionExcelPath(string? sourceExcelPath, string compCd, string? kindCode, string docCode, int? departmentId)
        {
            var now = DateTime.Now;
            var stamp = now.ToString("yyyyMMdd_HHmmss");
            var ext = NormalizeExcelExtension(sourceExcelPath);

            var rootDir = GetTemplateFinalDirectory(compCd, departmentId);

            var fileName = BuildTemplateExcelFileName(
                compCd,
                kindCode,
                docCode,
                0,
                stamp,
                ext
            );

            var fullPath = Path.Combine(rootDir, fileName);

            var seq = 1;
            while (System.IO.File.Exists(fullPath))
            {
                fileName = BuildTemplateExcelFileName(
                    compCd,
                    kindCode,
                    docCode,
                    0,
                    stamp,
                    ext,
                    seq
                );

                fullPath = Path.Combine(rootDir, fileName);
                seq++;
            }

            return fullPath;
        }
        private string? CreateRevisionExcelSnapshot(string? sourceExcelPath, string compCd, string? kindCode, string docCode, int? departmentId)
        {
            if (string.IsNullOrWhiteSpace(sourceExcelPath)) return null;
            if (!System.IO.File.Exists(sourceExcelPath)) return null;

            var revisionPath = CreateRevisionExcelPath(
                sourceExcelPath,
                compCd,
                kindCode,
                docCode,
                departmentId
            );

            System.IO.File.Copy(sourceExcelPath, revisionPath, overwrite: false);
            return revisionPath;
        }

        private sealed class TemplatePrepareResult
        {
            public string ProtectionRuleCode { get; set; } = TemplateProtectionRuleCode;
            public string VisualMetricRuleCode { get; set; } = TemplateVisualMetricRuleCode;
            public string VisualSource { get; set; } = string.Empty;
            public string VisualRangeA1 { get; set; } = string.Empty;
            public int VisualWidthPx { get; set; }
            public int VisualHeightPx { get; set; }
        }

        private sealed class TemplateRangeInfo
        {
            public string SheetName { get; set; } = string.Empty;
            public int FirstRow1 { get; set; }
            public int FirstCol1 { get; set; }
            public int LastRow1 { get; set; }
            public int LastCol1 { get; set; }
            public string Source { get; set; } = string.Empty;

            public string A1 => $"{ToColumnLettersForTemplateVersion(FirstCol1)}{FirstRow1}:{ToColumnLettersForTemplateVersion(LastCol1)}{LastRow1}";
        }

        // 2026.06.11 Added: 템플릿 저장 시점에 실제 xlsx 파일 자체에 보호/잠금과 표시 메트릭을 확정 저장함.
        private static TemplatePrepareResult PrepareTemplateExcelForVersion(string excelAbsPath, TemplateDescriptor descriptor)
        {
            if (string.IsNullOrWhiteSpace(excelAbsPath) || !System.IO.File.Exists(excelAbsPath))
                throw new FileNotFoundException("Template Excel file not found.", excelAbsPath);

            TemplateRangeInfo primaryRange;
            string primarySheetName;

            using (var wb = new XLWorkbook(excelAbsPath))
            {
                var primaryWs = wb.Worksheets.FirstOrDefault(w => !string.Equals(w.Name, "EB_META", StringComparison.OrdinalIgnoreCase))
                                ?? wb.Worksheets.FirstOrDefault()
                                ?? throw new InvalidOperationException("Template workbook has no worksheet.");

                primarySheetName = primaryWs.Name;
                primaryRange = ResolveTemplateVisualRange(primaryWs, descriptor);

                foreach (var ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, "EB_META", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (ws.IsProtected)
                    {
                        try
                        {
                            ws.Unprotect();
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidOperationException($"Worksheet protection cannot be reset. sheet={ws.Name}", ex);
                        }
                    }

                    var lockRange = string.Equals(ws.Name, primarySheetName, StringComparison.OrdinalIgnoreCase)
                        ? primaryRange
                        : ResolveTemplateProtectRange(ws, descriptor);

                    if (lockRange != null)
                    {
                        ws.Range(lockRange.FirstRow1, lockRange.FirstCol1, lockRange.LastRow1, lockRange.LastCol1)
                          .Style.Protection.Locked = true;
                    }

                    UnlockMappedFieldCells(ws, descriptor);

                    ws.Protect(allowedElements:
                        XLSheetProtectionElements.SelectLockedCells |
                        XLSheetProtectionElements.SelectUnlockedCells);
                }

                wb.Save();
            }

            var widthPx = ComputeTemplateVisualWidthPx(excelAbsPath, primarySheetName, primaryRange.FirstCol1, primaryRange.LastCol1);
            var heightPx = ComputeTemplateVisualHeightPx(excelAbsPath, primarySheetName, primaryRange.FirstRow1, primaryRange.LastRow1);

            return new TemplatePrepareResult
            {
                ProtectionRuleCode = TemplateProtectionRuleCode,
                VisualMetricRuleCode = TemplateVisualMetricRuleCode,
                VisualSource = primaryRange.Source,
                VisualRangeA1 = primaryRange.A1,
                VisualWidthPx = widthPx,
                VisualHeightPx = heightPx
            };
        }

        private static TemplateRangeInfo ResolveTemplateProtectRange(IXLWorksheet ws, TemplateDescriptor descriptor)
        {
            if (TryGetDescriptorBoundsForSheet(ws.Name, descriptor, out var descriptorRange))
                return descriptorRange;

            var used = ws.RangeUsed(XLCellsUsedOptions.All);
            if (used != null)
            {
                var a = used.RangeAddress;
                return new TemplateRangeInfo
                {
                    SheetName = ws.Name,
                    FirstRow1 = a.FirstAddress.RowNumber,
                    FirstCol1 = a.FirstAddress.ColumnNumber,
                    LastRow1 = a.LastAddress.RowNumber,
                    LastCol1 = a.LastAddress.ColumnNumber,
                    Source = "UsedRange"
                };
            }

            return new TemplateRangeInfo
            {
                SheetName = ws.Name,
                FirstRow1 = 1,
                FirstCol1 = 1,
                LastRow1 = 1,
                LastCol1 = 1,
                Source = "DescriptorBounds"
            };
        }

        private static TemplateRangeInfo ResolveTemplateVisualRange(IXLWorksheet ws, TemplateDescriptor descriptor)
        {
            try
            {
                var printAreas = ws.PageSetup.PrintAreas;
                if (printAreas != null && printAreas.Any())
                {
                    var range = ExpandRangeWithMergedRanges(ws, printAreas.First());
                    range.Source = "PrintArea";
                    return range;
                }
            }
            catch
            {
            }

            var used = ws.RangeUsed(XLCellsUsedOptions.All);
            if (used != null)
            {
                var usedRange = ExpandRangeWithMergedRanges(ws, used);

                if (TryGetDescriptorBoundsForSheet(ws.Name, descriptor, out var descriptorRange))
                {
                    var usedRows = usedRange.LastRow1 - usedRange.FirstRow1 + 1;
                    var usedCols = usedRange.LastCol1 - usedRange.FirstCol1 + 1;
                    var descRows = descriptorRange.LastRow1 - descriptorRange.FirstRow1 + 1;
                    var descCols = descriptorRange.LastCol1 - descriptorRange.FirstCol1 + 1;

                    var usedLooksTooLarge =
                        usedRows > Math.Max(descRows + 50, 300) ||
                        usedCols > Math.Max(descCols + 20, 80);

                    if (usedLooksTooLarge)
                    {
                        descriptorRange.Source = "DescriptorBounds";
                        return descriptorRange;
                    }
                }

                usedRange.Source = "UsedRange";
                return usedRange;
            }

            if (TryGetDescriptorBoundsForSheet(ws.Name, descriptor, out var fallbackDescriptorRange))
            {
                fallbackDescriptorRange.Source = "DescriptorBounds";
                return fallbackDescriptorRange;
            }

            return new TemplateRangeInfo
            {
                SheetName = ws.Name,
                FirstRow1 = 1,
                FirstCol1 = 1,
                LastRow1 = 1,
                LastCol1 = 1,
                Source = "DescriptorBounds"
            };
        }

        private static TemplateRangeInfo ExpandRangeWithMergedRanges(IXLWorksheet ws, IXLRange range)
        {
            var a = range.RangeAddress;
            var firstRow = a.FirstAddress.RowNumber;
            var firstCol = a.FirstAddress.ColumnNumber;
            var lastRow = a.LastAddress.RowNumber;
            var lastCol = a.LastAddress.ColumnNumber;

            try
            {
                foreach (var mr in ws.MergedRanges)
                {
                    var ma = mr.RangeAddress;
                    var intersects =
                        ma.FirstAddress.RowNumber <= lastRow && ma.LastAddress.RowNumber >= firstRow &&
                        ma.FirstAddress.ColumnNumber <= lastCol && ma.LastAddress.ColumnNumber >= firstCol;

                    if (!intersects) continue;

                    firstRow = Math.Min(firstRow, ma.FirstAddress.RowNumber);
                    firstCol = Math.Min(firstCol, ma.FirstAddress.ColumnNumber);
                    lastRow = Math.Max(lastRow, ma.LastAddress.RowNumber);
                    lastCol = Math.Max(lastCol, ma.LastAddress.ColumnNumber);
                }
            }
            catch
            {
            }

            return new TemplateRangeInfo
            {
                SheetName = ws.Name,
                FirstRow1 = Math.Max(1, firstRow),
                FirstCol1 = Math.Max(1, firstCol),
                LastRow1 = Math.Max(1, lastRow),
                LastCol1 = Math.Max(1, lastCol),
                Source = string.Empty
            };
        }

        private static bool TryGetDescriptorBoundsForSheet(string sheetName, TemplateDescriptor descriptor, out TemplateRangeInfo range)
        {
            range = new TemplateRangeInfo { SheetName = sheetName };

            var firstRow = int.MaxValue;
            var firstCol = int.MaxValue;
            var lastRow = 0;
            var lastCol = 0;

            void AddCell(CellRef? cell)
            {
                if (cell == null) return;

                var cellSheet = (cell.Sheet ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(cellSheet) &&
                    !string.Equals(cellSheet, sheetName, StringComparison.OrdinalIgnoreCase))
                    return;

                var r1 = Math.Max(1, cell.Row + 1);
                var c1 = Math.Max(1, cell.Column + 1);
                var r2 = r1 + Math.Max(1, cell.RowSpan) - 1;
                var c2 = c1 + Math.Max(1, cell.ColSpan) - 1;

                firstRow = Math.Min(firstRow, r1);
                firstCol = Math.Min(firstCol, c1);
                lastRow = Math.Max(lastRow, r2);
                lastCol = Math.Max(lastCol, c2);
            }

            foreach (var f in descriptor.Fields ?? Enumerable.Empty<FieldDef>()) AddCell(f.Cell);
            foreach (var a in descriptor.Approvals ?? Enumerable.Empty<ApprovalDef>()) AddCell(a.Cell);
            foreach (var c in descriptor.Cooperations ?? Enumerable.Empty<ApprovalDef>()) AddCell(c.Cell);

            if (lastRow <= 0 || lastCol <= 0 || firstRow == int.MaxValue || firstCol == int.MaxValue)
                return false;

            range = new TemplateRangeInfo
            {
                SheetName = sheetName,
                FirstRow1 = firstRow,
                FirstCol1 = firstCol,
                LastRow1 = lastRow,
                LastCol1 = lastCol,
                Source = "DescriptorBounds"
            };

            return true;
        }

        private static void UnlockMappedFieldCells(IXLWorksheet ws, TemplateDescriptor descriptor)
        {
            foreach (var f in descriptor.Fields ?? Enumerable.Empty<FieldDef>())
            {
                if (f?.Cell == null) continue;

                var cellSheet = (f.Cell.Sheet ?? string.Empty).Trim();
                if (!string.IsNullOrWhiteSpace(cellSheet) &&
                    !string.Equals(cellSheet, ws.Name, StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
                    var r1 = Math.Max(1, f.Cell.Row + 1);
                    var c1 = Math.Max(1, f.Cell.Column + 1);
                    var r2 = r1 + Math.Max(1, f.Cell.RowSpan) - 1;
                    var c2 = c1 + Math.Max(1, f.Cell.ColSpan) - 1;

                    var range = ws.Range(r1, c1, r2, c2);
                    range.Style.Protection.Locked = false;
                    range.Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF6FF");
                    range.Style.Alignment.WrapText = false;
                }
                catch
                {
                }
            }
        }

        private static int ComputeTemplateVisualWidthPx(string excelAbsPath, string sheetName, int firstCol1, int lastCol1)
        {
            var openXmlColWidths = new Dictionary<int, double>();
            var openXmlRowHeights = new Dictionary<int, double>();
            double defaultColWidth = 8.43;
            double defaultRowHeightPt = 15.0;

            TryReadOpenXmlSheetMetricsForTemplateVersion(
                excelAbsPath,
                sheetName,
                openXmlColWidths,
                openXmlRowHeights,
                ref defaultColWidth,
                ref defaultRowHeightPt);

            double sum = 0;
            for (int c = Math.Max(1, firstCol1); c <= Math.Max(firstCol1, lastCol1); c++)
            {
                var width = openXmlColWidths.TryGetValue(c, out var w) && w > 0
                    ? w
                    : defaultColWidth;

                sum += ExcelColumnWidthToPixelsForTemplateVersion(width);
            }

            sum += 24;
            return Math.Max(1, (int)Math.Ceiling(sum));
        }

        private static int ComputeTemplateVisualHeightPx(string excelAbsPath, string sheetName, int firstRow1, int lastRow1)
        {
            var openXmlColWidths = new Dictionary<int, double>();
            var openXmlRowHeights = new Dictionary<int, double>();
            double defaultColWidth = 8.43;
            double defaultRowHeightPt = 15.0;

            TryReadOpenXmlSheetMetricsForTemplateVersion(
                excelAbsPath,
                sheetName,
                openXmlColWidths,
                openXmlRowHeights,
                ref defaultColWidth,
                ref defaultRowHeightPt);

            double sum = 0;
            for (int r = Math.Max(1, firstRow1); r <= Math.Max(firstRow1, lastRow1); r++)
            {
                var heightPt = openXmlRowHeights.TryGetValue(r, out var h) && h > 0
                    ? h
                    : defaultRowHeightPt;

                sum += heightPt * 96.0 / 72.0;
            }

            sum += 32;
            return Math.Max(1, (int)Math.Ceiling(sum));
        }

        private static bool TryReadOpenXmlSheetMetricsForTemplateVersion(
            string excelAbsPath,
            string sheetName,
            Dictionary<int, double> openXmlColWidths,
            Dictionary<int, double> openXmlRowHeights,
            ref double defaultColWidth,
            ref double defaultRowHeightPt)
        {
            try
            {
                using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(excelAbsPath, false);
                var wbPart = doc.WorkbookPart;
                if (wbPart?.Workbook == null) return false;

                var sheets = wbPart.Workbook.Sheets?.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()?.ToList();
                if (sheets == null || sheets.Count == 0) return false;

                var sheet = sheets.FirstOrDefault(x =>
                    string.Equals(x.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                    ?? sheets.First();

                var relId = sheet.Id?.Value;
                if (string.IsNullOrWhiteSpace(relId)) return false;

                var wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)wbPart.GetPartById(relId);
                var worksheet = wsPart.Worksheet;
                if (worksheet == null) return false;

                var sfp = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetFormatProperties>();
                if (sfp?.DefaultColumnWidth?.Value != null && sfp.DefaultColumnWidth.Value > 0)
                    defaultColWidth = sfp.DefaultColumnWidth.Value;

                if (sfp?.DefaultRowHeight?.Value != null && sfp.DefaultRowHeight.Value > 0)
                    defaultRowHeightPt = sfp.DefaultRowHeight.Value;

                foreach (var cols in worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Columns>())
                {
                    foreach (var col in cols.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>())
                    {
                        var min = (int)(col.Min?.Value ?? 0U);
                        var max = (int)(col.Max?.Value ?? 0U);
                        var width = col.Width?.Value;
                        var hidden = col.Hidden?.Value ?? false;

                        if (hidden || min <= 0 || max <= 0 || width == null || width.Value <= 0) continue;

                        for (int i = min; i <= max; i++)
                            openXmlColWidths[i] = width.Value;
                    }
                }

                var sheetData = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                if (sheetData != null)
                {
                    foreach (var row in sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>())
                    {
                        var rowIndex = (int)(row.RowIndex?.Value ?? 0U);
                        if (rowIndex <= 0) continue;
                        if (row.Hidden?.Value == true) continue;

                        if (row.Height?.Value != null && row.Height.Value > 0)
                            openXmlRowHeights[rowIndex] = row.Height.Value;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private static double ExcelColumnWidthToPixelsForTemplateVersion(double excelWidth)
        {
            if (excelWidth <= 0) return 0;
            var px = Math.Round(excelWidth * 7.0 + 5.0, MidpointRounding.AwayFromZero);
            return px < 1 ? 1 : px;
        }

        private static string ToColumnLettersForTemplateVersion(int col)
        {
            col = Math.Max(1, col);
            var s = string.Empty;
            var n = col;
            while (n > 0)
            {
                var mod = (n - 1) % 26;
                s = ((char)('A' + mod)) + s;
                n = (n - 1) / 26;
            }
            return s;
        }

        private static string ComputeSha256Hex(string filePath)
        {
            using var sha = SHA256.Create();
            using var stream = System.IO.File.OpenRead(filePath);
            return Convert.ToHexString(sha.ComputeHash(stream));
        }

        private sealed class TemplateFieldCellInfo
        {
            public string Key { get; set; } = string.Empty;
            public string Type { get; set; } = "Text";
            public string Sheet { get; set; } = string.Empty;
            public string A1 { get; set; } = string.Empty;
        }

        // 2026.06.12 Added: 템플릿 저장 프로시저 실행 후 확정된 DocTemplateField 좌표를 기준으로
        // 최종 xlsx에 보호/Unlock/입력색/표시 메트릭을 적용한다.
        private TemplatePrepareResult PrepareTemplateExcelForVersionByDbFields(string excelAbsPath, long versionId)
        {
            if (versionId <= 0)
                throw new InvalidOperationException("Template version id is invalid.");

            var fields = LoadTemplateFieldCellsForVersion(versionId);
            return PrepareTemplateExcelForVersionByDbFields(excelAbsPath, fields);
        }

        private List<TemplateFieldCellInfo> LoadTemplateFieldCellsForVersion(long versionId)
        {
            var list = new List<TemplateFieldCellInfo>();
            var cs = _db.Database.GetConnectionString();
            if (string.IsNullOrWhiteSpace(cs))
                throw new InvalidOperationException("Default database connection string is empty.");

            using var conn = new SqlConnection(cs);
            conn.Open();

            using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT
    Id,
    [Key],
    [Type],
    Sheet,
    A1,
    CellA1,
    [Row],
    [Column],
    CellRow,
    CellColumn
FROM dbo.DocTemplateField
WHERE VersionId = @VersionId
ORDER BY Id;";
            cmd.Parameters.Add(new SqlParameter("@VersionId", SqlDbType.BigInt) { Value = versionId });

            using var rd = cmd.ExecuteReader();
            while (rd.Read())
            {
                var key = ReadString(rd, "Key");
                var type = NormalizeType(ReadString(rd, "Type"));
                var sheet = ReadString(rd, "Sheet");
                var a1 = FirstNonEmpty(ReadString(rd, "A1"), ReadString(rd, "CellA1")) ?? string.Empty;

                if (string.IsNullOrWhiteSpace(a1))
                {
                    var row = FirstPositive(ReadInt(rd, "CellRow"), ReadInt(rd, "Row"));
                    var col = FirstPositive(ReadInt(rd, "CellColumn"), ReadInt(rd, "Column"));
                    if (row > 0 && col > 0)
                        a1 = ToA1FromRowColForTemplateVersion(row, col);
                }

                if (string.IsNullOrWhiteSpace(a1))
                    continue;

                var (sheetFromA1, localA1) = SplitSheetAndA1ForTemplateVersion(a1);
                if (string.IsNullOrWhiteSpace(sheet)) sheet = sheetFromA1;
                if (string.IsNullOrWhiteSpace(localA1)) continue;

                list.Add(new TemplateFieldCellInfo
                {
                    Key = key,
                    Type = type,
                    Sheet = sheet,
                    A1 = localA1
                });
            }

            return list;
        }

        private static TemplatePrepareResult PrepareTemplateExcelForVersionByDbFields(string excelAbsPath, IReadOnlyList<TemplateFieldCellInfo> fields)
        {
            if (string.IsNullOrWhiteSpace(excelAbsPath) || !System.IO.File.Exists(excelAbsPath))
                throw new FileNotFoundException("Template Excel file not found.", excelAbsPath);

            fields ??= Array.Empty<TemplateFieldCellInfo>();

            TemplateRangeInfo primaryRange;
            string primarySheetName;

            using (var wb = new XLWorkbook(excelAbsPath))
            {
                var primaryWs = wb.Worksheets.FirstOrDefault(w => !string.Equals(w.Name, "EB_META", StringComparison.OrdinalIgnoreCase))
                                ?? wb.Worksheets.FirstOrDefault()
                                ?? throw new InvalidOperationException("Template workbook has no worksheet.");

                primarySheetName = primaryWs.Name;
                primaryRange = ResolveTemplateVisualRangeByDbFields(primaryWs, fields, primarySheetName);

                foreach (var ws in wb.Worksheets)
                {
                    if (string.Equals(ws.Name, "EB_META", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (ws.IsProtected)
                    {
                        try
                        {
                            ws.Unprotect();
                        }
                        catch (Exception ex)
                        {
                            throw new InvalidOperationException($"Worksheet protection cannot be reset. sheet={ws.Name}", ex);
                        }
                    }

                    var lockRanges = ResolveTemplateLockRangesByDbFields(ws, fields, primarySheetName);
                    foreach (var r in lockRanges)
                    {
                        r.Style.Protection.Locked = true;
                    }

                    ClearOldTemplateInputHighlights(ws, lockRanges, fields, primarySheetName);
                    UnlockMappedFieldCellsByDbFields(ws, fields, primarySheetName);

                    ws.Protect(allowedElements:
                        XLSheetProtectionElements.SelectLockedCells |
                        XLSheetProtectionElements.SelectUnlockedCells);
                }

                wb.Save();
            }

            var widthPx = ComputeTemplateVisualWidthPx(excelAbsPath, primarySheetName, primaryRange.FirstCol1, primaryRange.LastCol1);
            var heightPx = ComputeTemplateVisualHeightPx(excelAbsPath, primarySheetName, primaryRange.FirstRow1, primaryRange.LastRow1);

            return new TemplatePrepareResult
            {
                ProtectionRuleCode = TemplateProtectionRuleCode,
                VisualMetricRuleCode = TemplateVisualMetricRuleCode,
                VisualSource = primaryRange.Source,
                VisualRangeA1 = primaryRange.A1,
                VisualWidthPx = widthPx,
                VisualHeightPx = heightPx
            };
        }

        private static TemplateRangeInfo ResolveTemplateVisualRangeByDbFields(IXLWorksheet ws, IReadOnlyList<TemplateFieldCellInfo> fields, string primarySheetName)
        {
            try
            {
                var printAreas = ws.PageSetup.PrintAreas;
                if (printAreas != null && printAreas.Any())
                {
                    var range = ExpandRangeWithMergedRanges(ws, printAreas.First());
                    range.Source = "PrintArea";
                    return range;
                }
            }
            catch
            {
            }

            var used = ws.RangeUsed(XLCellsUsedOptions.All);
            if (used != null)
            {
                var usedRange = ExpandRangeWithMergedRanges(ws, used);

                if (TryGetDbFieldBoundsForSheet(ws, fields, primarySheetName, out var fieldRange))
                {
                    var usedRows = usedRange.LastRow1 - usedRange.FirstRow1 + 1;
                    var usedCols = usedRange.LastCol1 - usedRange.FirstCol1 + 1;
                    var fieldRows = fieldRange.LastRow1 - fieldRange.FirstRow1 + 1;
                    var fieldCols = fieldRange.LastCol1 - fieldRange.FirstCol1 + 1;

                    var usedLooksTooLarge =
                        usedRows > Math.Max(fieldRows + 50, 300) ||
                        usedCols > Math.Max(fieldCols + 20, 80);

                    if (usedLooksTooLarge)
                    {
                        fieldRange.Source = "DescriptorBounds";
                        return fieldRange;
                    }
                }

                usedRange.Source = "UsedRange";
                return usedRange;
            }

            if (TryGetDbFieldBoundsForSheet(ws, fields, primarySheetName, out var fallbackRange))
            {
                fallbackRange.Source = "DescriptorBounds";
                return fallbackRange;
            }

            return new TemplateRangeInfo
            {
                SheetName = ws.Name,
                FirstRow1 = 1,
                FirstCol1 = 1,
                LastRow1 = 1,
                LastCol1 = 1,
                Source = "DescriptorBounds"
            };
        }

        private static List<IXLRange> ResolveTemplateLockRangesByDbFields(IXLWorksheet ws, IReadOnlyList<TemplateFieldCellInfo> fields, string primarySheetName)
        {
            var ranges = new List<IXLRange>();

            try
            {
                var used = ws.RangeUsed(XLCellsUsedOptions.All);
                if (used != null) ranges.Add(used);
            }
            catch
            {
            }

            try
            {
                var printAreas = ws.PageSetup.PrintAreas;
                if (printAreas != null)
                {
                    foreach (var pa in printAreas)
                    {
                        if (pa != null) ranges.Add(pa);
                    }
                }
            }
            catch
            {
            }

            if (TryGetDbFieldBoundsForSheet(ws, fields, primarySheetName, out var fieldBounds))
            {
                try
                {
                    ranges.Add(ws.Range(fieldBounds.FirstRow1, fieldBounds.FirstCol1, fieldBounds.LastRow1, fieldBounds.LastCol1));
                }
                catch
                {
                }
            }

            if (ranges.Count == 0)
                ranges.Add(ws.Range(1, 1, 1, 1));

            return ranges;
        }

        private static bool TryGetDbFieldBoundsForSheet(IXLWorksheet ws, IReadOnlyList<TemplateFieldCellInfo> fields, string primarySheetName, out TemplateRangeInfo range)
        {
            range = new TemplateRangeInfo { SheetName = ws.Name };

            var firstRow = int.MaxValue;
            var firstCol = int.MaxValue;
            var lastRow = 0;
            var lastCol = 0;

            foreach (var f in fields ?? Array.Empty<TemplateFieldCellInfo>())
            {
                if (!IsTemplateFieldForWorksheet(f, ws.Name, primarySheetName))
                    continue;

                if (!TryParseA1OrRangeForTemplateVersion(f.A1, out var r1, out var c1, out var r2, out var c2))
                    continue;

                try
                {
                    var xlRange = ws.Range(r1, c1, r2, c2);
                    var effective = xlRange.FirstCell().IsMerged()
                        ? xlRange.FirstCell().MergedRange()
                        : xlRange;

                    var a = effective.RangeAddress;
                    r1 = a.FirstAddress.RowNumber;
                    c1 = a.FirstAddress.ColumnNumber;
                    r2 = a.LastAddress.RowNumber;
                    c2 = a.LastAddress.ColumnNumber;
                }
                catch
                {
                }

                firstRow = Math.Min(firstRow, r1);
                firstCol = Math.Min(firstCol, c1);
                lastRow = Math.Max(lastRow, r2);
                lastCol = Math.Max(lastCol, c2);
            }

            if (lastRow <= 0 || lastCol <= 0 || firstRow == int.MaxValue || firstCol == int.MaxValue)
                return false;

            range = new TemplateRangeInfo
            {
                SheetName = ws.Name,
                FirstRow1 = firstRow,
                FirstCol1 = firstCol,
                LastRow1 = lastRow,
                LastCol1 = lastCol,
                Source = "DescriptorBounds"
            };

            return true;
        }

        private static void UnlockMappedFieldCellsByDbFields(IXLWorksheet ws, IReadOnlyList<TemplateFieldCellInfo> fields, string primarySheetName)
        {
            foreach (var f in fields ?? Array.Empty<TemplateFieldCellInfo>())
            {
                if (!IsTemplateFieldForWorksheet(f, ws.Name, primarySheetName))
                    continue;

                if (!TryParseA1OrRangeForTemplateVersion(f.A1, out var r1, out var c1, out var r2, out var c2))
                    continue;

                try
                {
                    var range = ws.Range(r1, c1, r2, c2);
                    var targetRange = range.FirstCell().IsMerged()
                        ? range.FirstCell().MergedRange()
                        : range;

                    targetRange.Style.Protection.Locked = false;
                    targetRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF6FF");
                    targetRange.Style.Alignment.WrapText = false;
                }
                catch
                {
                }
            }
        }

        private static void ClearOldTemplateInputHighlights(IXLWorksheet ws, IReadOnlyList<IXLRange> lockRanges, IReadOnlyList<TemplateFieldCellInfo> fields, string primarySheetName)
        {
            var editableCells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var f in fields ?? Array.Empty<TemplateFieldCellInfo>())
            {
                if (!IsTemplateFieldForWorksheet(f, ws.Name, primarySheetName))
                    continue;

                if (!TryParseA1OrRangeForTemplateVersion(f.A1, out var r1, out var c1, out var r2, out var c2))
                    continue;

                try
                {
                    var range = ws.Range(r1, c1, r2, c2);
                    var targetRange = range.FirstCell().IsMerged()
                        ? range.FirstCell().MergedRange()
                        : range;

                    foreach (var cell in targetRange.Cells())
                        editableCells.Add(cell.Address.ToStringRelative());
                }
                catch
                {
                }
            }

            foreach (var range in lockRanges ?? Array.Empty<IXLRange>())
            {
                foreach (var cell in range.Cells())
                {
                    if (editableCells.Contains(cell.Address.ToStringRelative()))
                        continue;

                    if (IsTemplateInputHighlight(cell))
                    {
                        cell.Style.Fill.BackgroundColor = XLColor.NoColor;
                    }
                }
            }
        }

        private static bool IsTemplateInputHighlight(IXLCell cell)
        {
            try
            {
                var bg = cell.Style.Fill.BackgroundColor;
                if (bg.ColorType != XLColorType.Color)
                    return false;

                var c = bg.Color;
                return c.R == 0xEA && c.G == 0xF6 && c.B == 0xFF;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsTemplateFieldForWorksheet(TemplateFieldCellInfo field, string worksheetName, string primarySheetName)
        {
            var sheet = (field?.Sheet ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(sheet))
                return string.Equals(worksheetName, primarySheetName, StringComparison.OrdinalIgnoreCase);

            return string.Equals(sheet, worksheetName, StringComparison.OrdinalIgnoreCase);
        }

        private static (string sheetName, string localA1) SplitSheetAndA1ForTemplateVersion(string? a1)
        {
            var s = (a1 ?? string.Empty).Trim().Replace("$", string.Empty).Replace("'", string.Empty);
            if (string.IsNullOrWhiteSpace(s))
                return (string.Empty, string.Empty);

            var bang = s.LastIndexOf('!');
            if (bang >= 0)
                return (s[..bang].Trim(), s[(bang + 1)..].Trim());

            return (string.Empty, s);
        }

        private static bool TryParseA1OrRangeForTemplateVersion(string? a1, out int firstRow, out int firstCol, out int lastRow, out int lastCol)
        {
            firstRow = 0;
            firstCol = 0;
            lastRow = 0;
            lastCol = 0;

            var (_, localA1) = SplitSheetAndA1ForTemplateVersion(a1);
            localA1 = (localA1 ?? string.Empty).Trim().ToUpperInvariant();
            if (string.IsNullOrWhiteSpace(localA1))
                return false;

            var parts = localA1.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            if (parts.Length == 0 || parts.Length > 2)
                return false;

            if (!TryParseSingleA1ForTemplateVersion(parts[0], out firstRow, out firstCol))
                return false;

            if (parts.Length == 1)
            {
                lastRow = firstRow;
                lastCol = firstCol;
                return true;
            }

            if (!TryParseSingleA1ForTemplateVersion(parts[1], out lastRow, out lastCol))
                return false;

            if (lastRow < firstRow) (firstRow, lastRow) = (lastRow, firstRow);
            if (lastCol < firstCol) (firstCol, lastCol) = (lastCol, firstCol);
            return true;
        }

        private static bool TryParseSingleA1ForTemplateVersion(string? a1, out int row, out int col)
        {
            row = 0;
            col = 0;

            var m = Regex.Match((a1 ?? string.Empty).Trim().ToUpperInvariant(), @"^([A-Z]{1,3})(\d{1,7})$", RegexOptions.CultureInvariant);
            if (!m.Success)
                return false;

            col = ColLettersToIndex(m.Groups[1].Value);
            return col > 0 && int.TryParse(m.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out row) && row > 0;
        }

        private static string ToA1FromRowColForTemplateVersion(int row, int col)
        {
            if (row < 1 || col < 1) return string.Empty;
            return ToColumnLettersForTemplateVersion(col) + row.ToString(CultureInfo.InvariantCulture);
        }

        private static int ReadOrdinal(System.Data.Common.DbDataReader rd, string name)
        {
            for (var i = 0; i < rd.FieldCount; i++)
            {
                if (string.Equals(rd.GetName(i), name, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        private static string ReadString(System.Data.Common.DbDataReader rd, string name)
        {
            var ord = ReadOrdinal(rd, name);
            if (ord < 0 || rd.IsDBNull(ord)) return string.Empty;
            return Convert.ToString(rd.GetValue(ord), CultureInfo.InvariantCulture) ?? string.Empty;
        }

        private static int ReadInt(System.Data.Common.DbDataReader rd, string name)
        {
            var ord = ReadOrdinal(rd, name);
            if (ord < 0 || rd.IsDBNull(ord)) return 0;
            try { return Convert.ToInt32(rd.GetValue(ord), CultureInfo.InvariantCulture); }
            catch { return 0; }
        }

        private static int FirstPositive(params int[] values)
        {
            foreach (var v in values)
                if (v > 0) return v;
            return 0;
        }

        // 2026.06.16 Added: 전결 조건 유형 정규화 추가 Contents 화면 입력값을 DB 저장 조건 유형으로 변환
        private static string NormalizeDelegationConditionType(string? value)
        {
            var v = (value ?? string.Empty).Trim();

            if (v.Equals("Always", StringComparison.OrdinalIgnoreCase))
                return "Always";

            if (v.Equals("AmountLimit", StringComparison.OrdinalIgnoreCase))
                return "AmountLimit";

            return "None";
        }

        // 2026.06.16 Added: 전결 저장값 정규화 추가 Contents 전결권자 체크와 조건 패널 값을 저장 가능한 형태로 보정
        private static void NormalizeDelegationForSave(TemplateDescriptor model)
        {
            model.Delegation ??= new DelegationDef();
            model.Approvals ??= new List<ApprovalDef>();
            model.Fields ??= new List<FieldDef>();

            var d = model.Delegation;
            d.ConditionType = NormalizeDelegationConditionType(d.ConditionType);

            if (!d.Enabled || string.Equals(d.ConditionType, "None", StringComparison.OrdinalIgnoreCase))
            {
                d.Enabled = false;
                d.ConditionType = "None";
                d.DelegationStepOrder = 0;
                d.SkipFromStepOrder = 0;
                d.SkipToStepOrder = 0;
                d.AmountFieldKey = null;
                d.CurrencyFieldKey = null;
                d.AmountCellA1 = null;
                d.CurrencyCellA1 = null;
                d.AmountLimits = new List<DelegationAmountLimitDef>();
                foreach (var a in model.Approvals)
                    if (a != null) a.IsDelegationApprover = false;
                return;
            }

            if (d.DelegationStepOrder <= 0)
            {
                var marked = model.Approvals
                    .Where(x => x != null && x.IsDelegationApprover)
                    .OrderBy(x => x.Slot)
                    .FirstOrDefault();

                if (marked != null)
                    d.DelegationStepOrder = marked.Slot;
            }

            foreach (var a in model.Approvals)
            {
                if (a == null) continue;
                a.IsDelegationApprover = d.DelegationStepOrder > 0 && a.Slot == d.DelegationStepOrder;
            }

            var maxSlot = model.Approvals
                .Where(x => x != null)
                .Select(x => x.Slot)
                .DefaultIfEmpty(0)
                .Max();

            if (d.SkipFromStepOrder <= 0 && d.DelegationStepOrder > 0)
                d.SkipFromStepOrder = d.DelegationStepOrder + 1;

            if (d.SkipToStepOrder <= 0 && d.SkipFromStepOrder > 0)
                d.SkipToStepOrder = Math.Max(d.SkipFromStepOrder, maxSlot);

            d.AmountFieldKey = string.IsNullOrWhiteSpace(d.AmountFieldKey) ? null : d.AmountFieldKey.Trim();
            d.CurrencyFieldKey = string.IsNullOrWhiteSpace(d.CurrencyFieldKey) ? null : d.CurrencyFieldKey.Trim();

            // 2026.06.18 Added: 전결 조건 셀 주소 정규화 추가 Contents 금액 조건 비교에 사용할 금액 셀과 통화 셀 주소를 저장 전 정리
            d.AmountCellA1 = string.IsNullOrWhiteSpace(d.AmountCellA1) ? null : d.AmountCellA1.Trim().ToUpperInvariant();
            d.CurrencyCellA1 = string.IsNullOrWhiteSpace(d.CurrencyCellA1) ? null : d.CurrencyCellA1.Trim().ToUpperInvariant();

            d.AmountLimits ??= new List<DelegationAmountLimitDef>();
            d.AmountLimits = d.AmountLimits
                .Where(x => x != null)
                .Select(x => new DelegationAmountLimitDef
                {
                    CurrencyCode = (x.CurrencyCode ?? string.Empty).Trim().ToUpperInvariant(),
                    LimitAmount = x.LimitAmount
                })
                .Where(x => !string.IsNullOrWhiteSpace(x.CurrencyCode) || x.LimitAmount.HasValue)
                .ToList();

            if (!string.Equals(d.ConditionType, "AmountLimit", StringComparison.OrdinalIgnoreCase))
            {
                d.AmountFieldKey = null;
                d.CurrencyFieldKey = null;
                d.AmountCellA1 = null;
                d.CurrencyCellA1 = null;
                d.AmountLimits = new List<DelegationAmountLimitDef>();
            }
        }

        // 2026.06.16 Added: 전결 설정 검증 추가 Contents Always 및 AmountLimit 저장 전 필수값을 확인
        private IActionResult? ValidateDelegationForSave(TemplateDescriptor model)
        {
            var d = model.Delegation ?? new DelegationDef();

            if (!d.Enabled || string.Equals(d.ConditionType, "None", StringComparison.OrdinalIgnoreCase))
                return null;

            if (!string.Equals(d.ConditionType, "Always", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(d.ConditionType, "AmountLimit", StringComparison.OrdinalIgnoreCase))
            {
                return BadRequest(_S["DTL_V_DelegationConditionRequired"].Value);
            }

            if (d.DelegationStepOrder <= 0 ||
                !(model.Approvals ?? new List<ApprovalDef>()).Any(x => x != null && x.Slot == d.DelegationStepOrder))
            {
                return BadRequest(_S["DTL_V_DelegationStepRequired"].Value);
            }

            if (d.SkipFromStepOrder <= d.DelegationStepOrder ||
                d.SkipToStepOrder < d.SkipFromStepOrder ||
                !(model.Approvals ?? new List<ApprovalDef>()).Any(x => x != null && x.Slot == d.SkipFromStepOrder))
            {
                return BadRequest(_S["DTL_V_DelegationSkipStepRequired"].Value);
            }

            if (string.Equals(d.ConditionType, "AmountLimit", StringComparison.OrdinalIgnoreCase))
            {
                // 2026.06.18 Changed: 금액 셀은 항상 필수로 유지 Contents 전결 금액 조건은 금액 셀 값을 기준으로 판단
                if (string.IsNullOrWhiteSpace(d.AmountCellA1))
                    return BadRequest(_S["DTL_V_DelegationAmountCellRequired"].Value);

                if (d.AmountLimits == null || d.AmountLimits.Count == 0)
                    return BadRequest(_S["DTL_V_DelegationAmountLimitRequired"].Value);

                var currencyCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var item in d.AmountLimits)
                {
                    if (item == null || string.IsNullOrWhiteSpace(item.CurrencyCode))
                        return BadRequest(_S["DTL_V_DelegationCurrencyCodeRequired"].Value);

                    var currencyCode = item.CurrencyCode.Trim().ToUpperInvariant();

                    // 2026.06.18 Added: 허용 통화 코드 서버 검증 추가 Contents 화면 우회 저장 시 허용되지 않은 통화 코드 저장을 차단
                    if (!IsAllowedDelegationCurrencyCode(currencyCode))
                        return BadRequest(_S["DTL_V_DelegationCurrencyCodeRequired"].Value);

                    if (!currencyCodes.Add(currencyCode))
                        return BadRequest(_S["DTL_V_DelegationCurrencyCodeDup"].Value);

                    if (!item.LimitAmount.HasValue || item.LimitAmount.Value <= 0)
                        return BadRequest(_S["DTL_V_DelegationAmountInvalid"].Value);
                }

                // 2026.06.18 Changed: 통화 셀 필수 조건을 기준 통화 개수에 따라 분기 Contents 기준 통화가 1개이면 고정 통화로 보고 통화 셀 없이 저장 허용
                if (string.IsNullOrWhiteSpace(d.CurrencyCellA1) && currencyCodes.Count > 1)
                    return BadRequest(_S["DTL_V_DelegationCurrencyCellRequired"].Value);
            }

            return null;
        }

        private static string TrimMaxOrDefault(string? value, int max, string fallback)
        {
            var v = string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
            if (string.IsNullOrWhiteSpace(v)) v = fallback;
            return v.Length > max ? v.Substring(0, max) : v;
        }

        // 2026.06.16 Added: 전결 정책 저장 추가 Contents 템플릿 버전별 기존 전결 정책을 삭제 후 현재 화면 설정으로 재저장
        private void SaveDelegationRulesForVersion(int templateId, long versionId, TemplateDescriptor model)
        {
            if (versionId <= 0)
                return;

            model.Delegation ??= new DelegationDef();
            var d = model.Delegation;

            var actor = TrimMaxOrDefault(User?.Identity?.Name ?? CurrentUserId(), 100, "system");
            var now = DateTime.UtcNow;

            using var tx = _db.Database.BeginTransaction();

            var oldRules = _db.DocTemplateDelegationRules
                .Where(x => x.TemplateVersionId == versionId)
                .ToList();

            if (oldRules.Count > 0)
            {
                var oldIds = oldRules.Select(x => x.Id).ToList();

                var oldAmounts = _db.DocTemplateDelegationAmountRules
                    .Where(x => oldIds.Contains(x.RuleId))
                    .ToList();

                if (oldAmounts.Count > 0)
                    _db.DocTemplateDelegationAmountRules.RemoveRange(oldAmounts);

                _db.DocTemplateDelegationRules.RemoveRange(oldRules);
                _db.SaveChanges();
            }

            if (!d.Enabled || string.Equals(d.ConditionType, "None", StringComparison.OrdinalIgnoreCase))
            {
                tx.Commit();
                return;
            }

            var rule = new DocTemplateDelegationRule
            {
                TemplateId = templateId,
                TemplateVersionId = versionId,
                RuleName = null,
                ConditionType = d.ConditionType,
                DelegationStepOrder = d.DelegationStepOrder,
                SkipFromStepOrder = d.SkipFromStepOrder,
                SkipToStepOrder = d.SkipToStepOrder,
                Priority = 100,
                IsActive = true,
                Note = null,
                CreatedBy = actor,
                CreatedAt = now
            };

            _db.DocTemplateDelegationRules.Add(rule);
            _db.SaveChanges();

            if (string.Equals(d.ConditionType, "AmountLimit", StringComparison.OrdinalIgnoreCase))
            {
                foreach (var item in d.AmountLimits ?? new List<DelegationAmountLimitDef>())
                {
                    if (item == null) continue;

                    _db.DocTemplateDelegationAmountRules.Add(new DocTemplateDelegationAmountRule
                    {
                        RuleId = rule.Id,

                        // 2026.06.18 Changed: 필드 키는 호환용으로 빈 값 저장 Contents 실제 금액 조건 비교는 셀 주소 기준으로 처리
                        AmountFieldKey = d.AmountFieldKey ?? string.Empty,
                        CurrencyFieldKey = d.CurrencyFieldKey ?? string.Empty,

                        // 2026.06.18 Added: 전결 금액 조건 셀 주소 저장 Contents 입력 필드가 아닌 수식 셀과 통화 셀을 비교 기준으로 저장
                        AmountCellA1 = d.AmountCellA1,
                        CurrencyCellA1 = d.CurrencyCellA1,

                        CurrencyCode = (item.CurrencyCode ?? string.Empty).Trim().ToUpperInvariant(),
                        LimitAmount = item.LimitAmount ?? 0,
                        IsActive = true,
                        CreatedBy = actor,
                        CreatedAt = now
                    });
                }

                _db.SaveChanges();
            }

            tx.Commit();
        }

        [HttpPost("map-save")]
        [ValidateAntiForgeryToken]
        public IActionResult MapSave([FromForm] string descriptor, [FromForm] string? excelPath, [FromForm] string? previewJson, [FromForm] string? docCode, [FromForm] SpreadsheetClientState? spreadsheetState)
        {
            var swAll = Stopwatch.StartNew();

            if (string.IsNullOrWhiteSpace(descriptor))
                return BadRequest("No descriptor");

            TemplateDescriptor? model;
            try { model = JsonSerializer.Deserialize<TemplateDescriptor>(descriptor); }
            catch { return BadRequest("Invalid descriptor"); }
            if (model == null) return BadRequest("Empty descriptor");

            model.Delegation ??= new DelegationDef();

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

            var mapScopeGuard = ValidateTemplateScopeForSave(model.CompCd, deptIdToSave, model.Kind);
            if (mapScopeGuard is not null)
                return mapScopeGuard;

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

                a.ApproverValue = (a.ApproverValue ?? string.Empty).Trim();

                if (a.ApproverType == "Person")
                {
                    if (string.IsNullOrWhiteSpace(a.ApproverValue))
                        return BadRequest($"Approvals[{i}] 사용자ID를 입력해 주세요.");

                    // 2026.06.18 Added: 기안자 예약값 허용 추가 Contents 기안자는 실제 사용자 코드 없이 문서 작성 시 작성자 사용자 ID로 해석할 예약값으로 저장
                    if (IsDrafterApproverValue(a.ApproverValue))
                    {
                        model.Approvals[i] = a;
                    }
                }
                else if (a.ApproverType == "Role")
                {
                    if (string.IsNullOrWhiteSpace(a.ApproverValue))
                        return BadRequest($"Approvals[{i}] 역할코드를 입력해 주세요.");
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

                c.ApproverValue = (c.ApproverValue ?? string.Empty).Trim();

                if (c.ApproverType == "Person")
                {
                    if (string.IsNullOrWhiteSpace(c.ApproverValue))
                        return BadRequest($"Cooperations[{i}] 사용자ID를 입력해 주세요.");

                    // 2026.06.18 Added: 기안자 예약값 허용 추가 Contents 협조란에서도 실제 사용자 코드 없는 기안자 예약값을 템플릿 저장 단계에서 허용
                    if (IsDrafterApproverValue(c.ApproverValue))
                    {
                        model.Cooperations[i] = c;
                    }
                }
                else if (c.ApproverType == "Role")
                {
                    if (string.IsNullOrWhiteSpace(c.ApproverValue))
                        return BadRequest($"Cooperations[{i}] 역할코드를 입력해 주세요.");
                }

                if (c.Slot <= 0) c.Slot = 1;
                model.Cooperations[i] = c;
            }

            // 2026.06.16 Added: 전결 저장 전 정규화 및 검증 추가 Contents 전결권자 차수와 조건 값을 DB 저장 전 확인
            NormalizeDelegationForSave(model);
            var delegationGuard = ValidateDelegationForSave(model);
            if (delegationGuard is not null)
                return delegationGuard;

            var fileSetStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            var tempExcelPath = CreateRevisionExcelPath(excelPath, model.CompCd, model.Kind, docCodeToSave, deptIdToSave);

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
                finalExcelPath = BuildVersionedExcelPath(tempExcelPath, model.CompCd, model.Kind, docCodeToSave, outVersionNo, fileSetStamp, deptIdToSave);

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

            TemplatePrepareResult templatePrepare;
            try
            {
                // 2026.06.12 Changed: 보호/파란 배경/표시 메트릭은 프로시저 실행 후
                // 확정된 outVersionId 기준 dbo.DocTemplateField 좌표를 다시 조회해서 적용한다.
                // descriptor/model의 Row/Column 기본값(0)이 A1로 오염되는 문제를 방지한다.
                templatePrepare = PrepareTemplateExcelForVersionByDbFields(finalExcelPath, outVersionId);
                Debug.WriteLine($"[DocTLDX][MapSave] template prepared by DB fields source={templatePrepare.VisualSource} range={templatePrepare.VisualRangeA1} width={templatePrepare.VisualWidthPx} height={templatePrepare.VisualHeightPx}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[DocTLDX][MapSave] PrepareTemplateExcelForVersionByDbFields failed: " + ex);
                TempData["Alert"] = "Template prepare failed: " + ex.Message;

                return RedirectToAction(nameof(MapSaved), new
                {
                    path = "",
                    excelPath = finalExcelPath,
                    fields = model.Fields?.Count ?? 0,
                    approvals = model.Approvals?.Count ?? 0
                });
            }

            try
            {
                // 저장 프로시저에는 임시 저장본 기준 previewJson이 먼저 들어가므로,
                // 보호/표시색/그리드 상태가 반영된 최종 xlsx 기준으로 다시 생성하여 DB에 갱신한다.
                previewJson = DocControllerHelper.BuildPreviewJsonFromExcel(finalExcelPath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[DocTLDX][MapSave] BuildPreviewJsonFromExcel(final) failed: " + ex);
                previewJson ??= "{}";
            }

            var finalExcelFileName = Path.GetFileName(finalExcelPath);
            var finalExcelSize = System.IO.File.Exists(finalExcelPath)
                ? new FileInfo(finalExcelPath).Length
                : 0L;
            var finalExcelRelPath = ToAppDataRelativePath(finalExcelPath);
            var preparedAt = DateTime.Now;
            var templateFileHash = ComputeSha256Hex(finalExcelPath);

            var excelDbSyncOk = false;

            try
            {
                using var tx = _db.Database.BeginTransaction();

                const string sqlUpdateVersion = @"
UPDATE dbo.DocTemplateVersion
   SET ExcelFileName = @ExcelFileName,
       ExcelFilePath = @ExcelFilePath,
       ExcelFileSize = @ExcelFileSize,
       PreviewJson = @PreviewJson,
       PreparedAt = @PreparedAt,
       TemplateFileHash = @TemplateFileHash,
       ProtectionRuleCode = @ProtectionRuleCode,
       VisualMetricRuleCode = @VisualMetricRuleCode,
       VisualSource = @VisualSource,
       VisualRangeA1 = @VisualRangeA1,
       VisualWidthPx = @VisualWidthPx,
       VisualHeightPx = @VisualHeightPx
 WHERE Id = @VersionId;";

                _db.Database.ExecuteSqlRaw(
                    sqlUpdateVersion,
                    new SqlParameter("@ExcelFileName", SqlDbType.NVarChar, 255) { Value = finalExcelFileName },
                    new SqlParameter("@ExcelFilePath", SqlDbType.NVarChar, 500) { Value = finalExcelRelPath },
                    new SqlParameter("@ExcelFileSize", SqlDbType.BigInt) { Value = finalExcelSize },
                    new SqlParameter("@PreviewJson", SqlDbType.NVarChar, -1) { Value = (object?)previewJson ?? "{}" },
                    new SqlParameter("@PreparedAt", SqlDbType.DateTime2) { Value = preparedAt },
                    new SqlParameter("@TemplateFileHash", SqlDbType.Char, 64) { Value = templateFileHash },
                    new SqlParameter("@ProtectionRuleCode", SqlDbType.NVarChar, 50) { Value = templatePrepare.ProtectionRuleCode },
                    new SqlParameter("@VisualMetricRuleCode", SqlDbType.NVarChar, 50) { Value = templatePrepare.VisualMetricRuleCode },
                    new SqlParameter("@VisualSource", SqlDbType.NVarChar, 50) { Value = templatePrepare.VisualSource },
                    new SqlParameter("@VisualRangeA1", SqlDbType.NVarChar, 100) { Value = templatePrepare.VisualRangeA1 },
                    new SqlParameter("@VisualWidthPx", SqlDbType.Int) { Value = templatePrepare.VisualWidthPx },
                    new SqlParameter("@VisualHeightPx", SqlDbType.Int) { Value = templatePrepare.VisualHeightPx },
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

                const string sqlUpsertPreviewFile = @"
IF EXISTS (SELECT 1 FROM dbo.DocTemplateFile WHERE VersionId = @VersionId AND FileRole = N'PreviewJson')
BEGIN
    UPDATE dbo.DocTemplateFile
       SET Contents = @PreviewJson,
           FileSize = DATALENGTH(@PreviewJson),
           FileSizeBytes = DATALENGTH(@PreviewJson),
           ContentType = N'application/json'
     WHERE VersionId = @VersionId
       AND FileRole = N'PreviewJson';
END
ELSE
BEGIN
    INSERT dbo.DocTemplateFile
           (TemplateId, VersionId, FileRole, Storage, FileName, FilePath,
            FileSize, FileSizeBytes, ContentType, Contents, CreatedAt, CreatedBy)
    VALUES (@TemplateId, @VersionId, N'PreviewJson', N'Db', N'preview.json', NULL,
            DATALENGTH(@PreviewJson), DATALENGTH(@PreviewJson), N'application/json', @PreviewJson,
            SYSUTCDATETIME(), @CreatedBy);
END";

                _db.Database.ExecuteSqlRaw(
                    sqlUpsertPreviewFile,
                    new SqlParameter("@TemplateId", SqlDbType.Int) { Value = outTemplateId },
                    new SqlParameter("@VersionId", SqlDbType.BigInt) { Value = outVersionId },
                    new SqlParameter("@PreviewJson", SqlDbType.NVarChar, -1) { Value = (object?)previewJson ?? "{}" },
                    new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 100) { Value = User?.Identity?.Name ?? "system" }
                );

                tx.Commit();
                excelDbSyncOk = true;
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[DocTLDX][MapSave] Excel DB sync update failed: " + ex);

                TryRollbackFinalExcelToTemp(finalExcelPath, tempExcelPath);

                TempData["Alert"] = "Excel DB path sync failed: " + ex.Message;

                return RedirectToAction(nameof(MapSaved), new
                {
                    path = "",
                    excelPath = System.IO.File.Exists(tempExcelPath) ? tempExcelPath : finalExcelPath,
                    fields = model.Fields?.Count ?? 0,
                    approvals = model.Approvals?.Count ?? 0
                });
            }

            // 2026.06.16 Added: 전결 정책 DB 저장 추가 Contents 템플릿 파일 및 버전 정보 동기화 성공 후 전결 규칙을 저장
            try
            {
                SaveDelegationRulesForVersion(outTemplateId, outVersionId, model);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("[DocTLDX][MapSave] Delegation rule save failed: " + ex);
                TempData["Alert"] = "Delegation rule save failed: " + ex.Message;

                return RedirectToAction(nameof(MapSaved), new
                {
                    path = "",
                    excelPath = finalExcelPath,
                    fields = model.Fields?.Count ?? 0,
                    approvals = model.Approvals?.Count ?? 0
                });
            }

            if (excelDbSyncOk)
            {
                TryDeleteUploadTempExcel(excelPath, finalExcelPath);
            }

            // 2026.06.12 Changed: 템플릿 저장 시 물리 JSON 스냅샷 파일을 생성하지 않는다.
            // 실제 운영 기준 데이터는 dbo.DocTemplateVersion / dbo.DocTemplateField / dbo.DocTemplateApproval 및 xlsx 파일에 저장된다.
            Debug.WriteLine($"[DocTLDX][MapSave] saved docCode={docCodeToSave} templateId={outTemplateId} versionId={outVersionId} versionNo={outVersionNo} fields={(model.Fields?.Count ?? 0)} approvals={(model.Approvals?.Count ?? 0)} cooperations={(model.Cooperations?.Count ?? 0)} elapsed={swAll.ElapsedMilliseconds}ms");
            Debug.WriteLine($"[DocTLDX][MapSave] spreadsheetStateNull={(spreadsheetState == null)}");
            Debug.WriteLine($"[DocTLDX][MapSave] fileSetStamp={fileSetStamp}");
            Debug.WriteLine($"[DocTLDX][MapSave] tempExcelPath={tempExcelPath}");
            Debug.WriteLine($"[DocTLDX][MapSave] finalExcelPath={finalExcelPath}");
            Debug.WriteLine($"[DocTLDX][MapSave] excelDbSyncOk={excelDbSyncOk}");
            Debug.WriteLine($"[DocTLDX][MapSave] preparedAt={preparedAt:O}");
            Debug.WriteLine($"[DocTLDX][MapSave] templateFileHash={templateFileHash}");
            Debug.WriteLine($"[DocTLDX][MapSave] protectionRule={templatePrepare.ProtectionRuleCode} visualRule={templatePrepare.VisualMetricRuleCode} visualSource={templatePrepare.VisualSource} visualRange={templatePrepare.VisualRangeA1} visualWidth={templatePrepare.VisualWidthPx} visualHeight={templatePrepare.VisualHeightPx}");
            Debug.WriteLine($"[DocTLDX][MapSave] delegationEnabled={(model.Delegation?.Enabled ?? false)} delegationCondition={model.Delegation?.ConditionType ?? "None"} delegationStep={model.Delegation?.DelegationStepOrder ?? 0}");

            return RedirectToAction(nameof(MapSaved), new
            {
                path = string.Empty,
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

        private string BuildVersionedExcelPath(string currentExcelPath, string compCd, string? kindCode, string docCode, int versionNo, string stamp, int? departmentId)
        {
            var dir = GetTemplateFinalDirectory(compCd, departmentId);

            var ext = Path.GetExtension(currentExcelPath);
            if (string.IsNullOrWhiteSpace(ext))
                ext = ".xlsx";

            ext = NormalizeExcelExtension(ext);

            var safeStamp = string.IsNullOrWhiteSpace(stamp)
                ? DateTime.Now.ToString("yyyyMMdd_HHmmss")
                : stamp.Trim();

            var fileName = BuildTemplateExcelFileName(compCd, kindCode, docCode, versionNo, safeStamp, ext);

            var fullPath = Path.Combine(dir, fileName);

            var seq = 1;
            while (System.IO.File.Exists(fullPath))
            {
                fileName = BuildTemplateExcelFileName(compCd, kindCode, docCode, versionNo, safeStamp, ext, seq);

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

        // 2026.06.12 Removed: 물리 Descriptor JSON 파일 저장을 중단했으므로 download-descriptor 기능도 사용하지 않는다.
    }
}