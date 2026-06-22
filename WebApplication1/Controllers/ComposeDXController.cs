using ClosedXML.Excel;
using DevExpress.AspNetCore.Spreadsheet;
using DevExpress.Spreadsheet;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Antiforgery;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Services;
using static WebApplication1.Controllers.DocControllerHelper;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class ComposeDXController : Controller
    {
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IConfiguration _cfg;
        private readonly IAntiforgery _antiforgery;
        private readonly IWebHostEnvironment _env;
        private readonly ApplicationDbContext _db;
        private readonly IDocTemplateService _tpl;
        private readonly ILogger _log;
        private readonly IEmailSender _emailSender;
        private readonly SmtpOptions _smtpOpt;
        private static readonly object _attachSeqLock = new();
        private static readonly string[] _msgUploadFailed = new[] { "DOC_Err_UploadFailed" };
        private readonly IWebPushNotifier _webPushNotifier;
        private readonly DocControllerHelper _helper;

        // 2026.06.18 Added: 기안자 예약값 상수 추가 Contents 공용 템플릿에서 문서 작성자를 결재자로 지정하기 위한 예약값
        private const string ApproverValueDrafter = "__DRAFTER__";

        // 2026.06.18 Added: 기안자 예약값 판정 추가 Contents 템플릿 원본은 유지하고 문서 생성 시점에 현재 작성자로 치환
        private static bool IsDrafterApproverValue(string? value)
            => string.Equals((value ?? string.Empty).Trim(), ApproverValueDrafter, StringComparison.OrdinalIgnoreCase);

        // 2026.06.18 Added: 기안자 예약값 치환 추가 Contents 문서 생성 시점에 현재 작성자 UserId를 결재자 값으로 사용
        private static string ResolveDrafterApproverValue(string? value, string? drafterUserId)
        {
            var v = (value ?? string.Empty).Trim();
            if (!IsDrafterApproverValue(v))
                return v;

            return (drafterUserId ?? string.Empty).Trim();
        }

        // 2026.06.18 Added: Descriptor JSON 내부 기안자 예약값 치환 추가 Contents 공용 템플릿 원본은 유지하고 생성 문서 JSON만 현재 작성자 기준으로 확정
        private static string ReplaceDrafterApproverValuesInJson(string? json, string? drafterUserId)
        {
            if (string.IsNullOrWhiteSpace(json))
                return "{}";

            var resolvedUserId = (drafterUserId ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(resolvedUserId))
                return json!;

            try
            {
                var node = System.Text.Json.Nodes.JsonNode.Parse(json);
                if (node == null)
                    return json!;

                static bool IsApproverValueKey(string key)
                    => string.Equals(key, "value", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(key, "ApproverValue", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(key, "approverValue", StringComparison.OrdinalIgnoreCase);

                static void Walk(System.Text.Json.Nodes.JsonNode? current, string resolved)
                {
                    if (current is System.Text.Json.Nodes.JsonObject obj)
                    {
                        var keys = obj.Select(x => x.Key).ToList();
                        foreach (var key in keys)
                        {
                            var child = obj[key];

                            if (IsApproverValueKey(key) &&
                                child is System.Text.Json.Nodes.JsonValue valueNode &&
                                valueNode.TryGetValue<string>(out var stringValue) &&
                                IsDrafterApproverValue(stringValue))
                            {
                                obj[key] = resolved;
                                continue;
                            }

                            Walk(child, resolved);
                        }

                        return;
                    }

                    if (current is System.Text.Json.Nodes.JsonArray arr)
                    {
                        foreach (var child in arr)
                            Walk(child, resolved);
                    }
                }

                Walk(node, resolvedUserId);

                return node.ToJsonString(new JsonSerializerOptions
                {
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                    WriteIndented = false
                });
            }
            catch
            {
                return json!;
            }
        }

        // 2026.06.11 Added: ComposeDX 첫 진입/반복 진입 성능 개선용 캐시.
        // OrgTree는 화면 표시용 동일 데이터이므로 짧은 TTL로 재사용한다.
        private static readonly object _composeOrgTreeCacheLock = new();
        private static DateTimeOffset _composeOrgTreeCacheUntilUtc = DateTimeOffset.MinValue;
        private static IReadOnlyList<OrgTreeNode>? _composeOrgTreeCacheKo;


        public ComposeDXController(IStringLocalizer<SharedResource> S, IConfiguration cfg, IAntiforgery antiforgery, IWebHostEnvironment env, ApplicationDbContext db, IDocTemplateService tpl, IEmailSender emailSender, IOptions<SmtpOptions> smtpOptions, ILogger<ComposeDXController> log, IWebPushNotifier webPushNotifier)
        {
            _S = S;
            _cfg = cfg;
            _antiforgery = antiforgery;
            _env = env;
            _db = db;
            _tpl = tpl;
            _emailSender = emailSender;
            _smtpOpt = smtpOptions?.Value ?? new SmtpOptions();
            _log = log;
            _webPushNotifier = webPushNotifier;
            _helper = new DocControllerHelper(_cfg, _env, () => User);
        }

        [HttpGet("CreateDX")]
        public async Task<IActionResult> CreateDX(string templateCode)
        {
            var perfSw = System.Diagnostics.Stopwatch.StartNew();
            var lastMs = 0.0;

            void Mark(string stage)
            {
                var now = perfSw.Elapsed.TotalMilliseconds;
                _log.LogInformation(
                    "COMPOSEDX-PERF {Stage} elapsedMs={ElapsedMs} deltaMs={DeltaMs} templateCode={TemplateCode}",
                    stage,
                    now,
                    now - lastMs,
                    templateCode ?? string.Empty);
                lastMs = now;
            }

            Mark("start");

            var (loadedDescriptorJson, loadedPreviewJson, templateTitle, templateVersionId, excelPathRaw) = await _tpl.LoadMetaAsync(templateCode);
            Mark("meta-load");

            var descriptorJson = string.IsNullOrWhiteSpace(loadedDescriptorJson)
                ? "{}"
                : loadedDescriptorJson;

            var previewJson = string.IsNullOrWhiteSpace(loadedPreviewJson)
                ? "{}"
                : loadedPreviewJson;

            var normalizedExcelPath = NormalizeTemplateExcelPath(excelPathRaw);
            var excelAbsPath = _helper.ToContentRootAbsolute(normalizedExcelPath);
            Mark("excel-path-normalize");

            try
            {
                var nowUtc = DateTimeOffset.UtcNow;
                IReadOnlyList<OrgTreeNode>? cachedOrg = null;

                lock (_composeOrgTreeCacheLock)
                {
                    if (_composeOrgTreeCacheKo != null && _composeOrgTreeCacheUntilUtc > nowUtc)
                        cachedOrg = _composeOrgTreeCacheKo;
                }

                if (cachedOrg == null)
                {
                    var orgNodes = (await _helper.BuildOrgTreeNodesAsync("ko")).ToList();

                    lock (_composeOrgTreeCacheLock)
                    {
                        _composeOrgTreeCacheKo = orgNodes;
                        _composeOrgTreeCacheUntilUtc = DateTimeOffset.UtcNow.AddMinutes(5);
                    }

                    ViewBag.OrgTreeNodes = orgNodes;
                    Mark("orgtree-load-db");
                }
                else
                {
                    ViewBag.OrgTreeNodes = cachedOrg;
                    Mark("orgtree-load-cache");
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "COMPOSEDX-PERF orgtree load failed templateCode={TemplateCode}", templateCode);
                ViewBag.OrgTreeNodes = Array.Empty<OrgTreeNode>();
                Mark("orgtree-load-failed");
            }

            var editableCellsFromDb = await LoadComposeEditableCellsAsync(templateVersionId);
            Mark("editable-cells-load-db");

            descriptorJson = BuildDescriptorJsonWithFlowGroups(templateCode, descriptorJson);
            descriptorJson = UpsertComposeInputsIntoDescriptorJson(descriptorJson, editableCellsFromDb);
            Mark("descriptor-flowgroup-build");

            var dxDocumentId = $"compose_{Guid.NewGuid():N}";
            var composeDxExcelPath = excelAbsPath;

            try
            {
                if (!string.IsNullOrWhiteSpace(excelAbsPath) && System.IO.File.Exists(excelAbsPath))
                {
                    composeDxExcelPath = await CreateComposeDxWorkbookCopyAsync(templateCode, excelAbsPath, dxDocumentId, descriptorJson);
                    Mark("workbook-copy");
                }
                else
                {
                    Mark("workbook-copy-skipped-no-file");
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "CreateDX workbook copy failed. templateCode={TemplateCode}", templateCode);
                composeDxExcelPath = excelAbsPath;
                Mark("workbook-copy-failed");
            }

            var visual = await LoadComposeTemplateVisualMetricsAsync(templateVersionId, templateCode);
            Mark("visual-metrics-load-db");

            ViewBag.TargetHeightPx = visual.VisualHeightPx ?? 0;
            ViewBag.TargetWidthPx = visual.VisualWidthPx ?? 0;
            ViewBag.VisualSource = visual.VisualSource ?? string.Empty;
            ViewBag.VisualRangeA1 = visual.VisualRangeA1 ?? string.Empty;
            ViewBag.VisualMetricRuleCode = visual.VisualMetricRuleCode ?? string.Empty;

            var targetA1 = TryGetLastCellFromRangeA1(visual.VisualRangeA1, out var lastA1, out var targetRow1, out var targetCol1)
                ? lastA1
                : "A1";

            ViewBag.LastRow = Math.Max(0, targetRow1 - 1);
            ViewBag.LastCol = Math.Max(0, targetCol1 - 1);
            ViewBag.LastA1 = targetA1;
            ViewBag.TargetA1 = targetA1;
            ViewBag.TargetRow1 = targetRow1 <= 0 ? 1 : targetRow1;
            ViewBag.TargetCol1 = targetCol1 <= 0 ? 1 : targetCol1;
            ViewBag.DxVisualMode = string.IsNullOrWhiteSpace(visual.VisualMetricRuleCode)
                ? "DbVisualMetricMissing"
                : $"{visual.VisualMetricRuleCode}:{visual.VisualSource}";

            _log.LogInformation(
                "COMPOSEDX-PERF visual-metrics-db templateCode={TemplateCode} versionId={VersionId} range={Range} widthPx={WidthPx} heightPx={HeightPx} source={Source} rule={Rule}",
                templateCode ?? string.Empty,
                templateVersionId,
                visual.VisualRangeA1 ?? string.Empty,
                visual.VisualWidthPx ?? 0,
                visual.VisualHeightPx ?? 0,
                visual.VisualSource ?? string.Empty,
                visual.VisualMetricRuleCode ?? string.Empty);

            ViewBag.DescriptorJson = descriptorJson;
            ViewBag.PreviewJson = previewJson;
            ViewBag.TemplateTitle = templateTitle ?? string.Empty;
            ViewBag.TemplateCode = templateCode;
            ViewBag.ExcelPath = composeDxExcelPath;
            ViewBag.DxCallbackUrl = "/Doc/dx-callback";
            ViewBag.DxDocumentId = dxDocumentId;
            ViewBag.HideCellPicker = true;

            Mark("viewbag-assign");
            Mark("before-return");

            return View("~/Views/Doc/ComposeDX.cshtml");
        }



        private sealed class ComposeTemplateVisualMetrics
        {
            public int? VisualWidthPx { get; set; }
            public int? VisualHeightPx { get; set; }
            public string? VisualSource { get; set; }
            public string? VisualRangeA1 { get; set; }
            public string? VisualMetricRuleCode { get; set; }
        }

        // 2026.06.12 Changed: ComposeDX 진입 시 보호/잠금/폭높이 재계산을 하지 않는다.
        // 템플릿 저장 시점에 준비된 실제 xlsx를 사용자별 임시 파일로 복사해서 연다.
        private async Task<string> CreateComposeDxWorkbookCopyAsync(string? templateCode, string sourceExcelFullPath, string dxDocumentId, string descriptorJson)
        {
            if (string.IsNullOrWhiteSpace(sourceExcelFullPath) || !System.IO.File.Exists(sourceExcelFullPath))
                return sourceExcelFullPath;

            var outRoot = Path.Combine(_env.ContentRootPath, "App_Data", "DocDxCompose");
            Directory.CreateDirectory(outRoot);

            var safeTemplateCode = SanitizeComposeDxFilePart(templateCode);
            var outPath = Path.Combine(
                outRoot,
                $"{safeTemplateCode}_{DateTime.Now:yyyyMMddHHmmssfff}_{dxDocumentId}.xlsx");

            await Task.Run(() =>
            {
                System.IO.File.Copy(sourceExcelFullPath, outPath, overwrite: true);

                // 2026.06.12 Changed: 폭/높이는 DB 저장값을 사용하되,
                // 작성 가능 셀의 Unlock/표시색은 작성용 임시 복사본에서 보강한다.
                // 원본 템플릿 파일은 변경하지 않는다.
                try
                {
                    ApplyComposeDxEditableCells(outPath, descriptorJson);
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "COMPOSEDX-PERF editable-cell-apply failed templateCode={TemplateCode} path={Path}", templateCode ?? string.Empty, outPath);
                    // 편집 가능 셀 보강 실패는 문서 오픈 자체를 막지 않는다.
                }

                try
                {
                    HideComposeDxWorkbookGridLines(outPath);
                }
                catch
                {
                    // 격자선 숨김 실패는 문서 오픈 자체를 막지 않는다.
                }
            });

            return outPath;
        }


        private sealed class ComposeDxEditableCell
        {
            public string Sheet { get; set; } = string.Empty;
            public string Key { get; set; } = string.Empty;
            public string Type { get; set; } = "Text";
            public string A1 { get; set; } = string.Empty;
        }

        // 2026.06.12 Changed: ComposeDX 작성용 임시 복사본의 매핑 셀을 Unlock/표시색 보강.
        // 기존 보호/잠금 상태는 신뢰하지 않고, 작성용 복사본에서만 다시 구성한다.
        // 폭/높이 재계산은 하지 않으며, 저장된 원본 템플릿 파일은 변경하지 않는다.
        private static void ApplyComposeDxEditableCells(string workbookPath, string? descriptorJson)
        {
            if (string.IsNullOrWhiteSpace(workbookPath) || !System.IO.File.Exists(workbookPath))
                return;

            var editableCells = ParseComposeDxEditableCells(descriptorJson);
            if (editableCells.Count == 0)
                return;

            // 기존 엑셀 보호 설정은 비밀번호/상태를 신뢰하지 않는다.
            // 먼저 OpenXml에서 sheetProtection 노드를 제거해서 ClosedXML이 확실히 수정할 수 있게 한다.
            TryRemoveWorksheetProtectionForCompose(workbookPath);

            var tempPath = workbookPath + ".editable.tmp.xlsx";
            if (System.IO.File.Exists(tempPath))
                System.IO.File.Delete(tempPath);

            using (var wb = new XLWorkbook(workbookPath))
            {
                var firstSheet = wb.Worksheets
                    .FirstOrDefault(w => !string.Equals(w.Name, "EB_META", StringComparison.OrdinalIgnoreCase))
                    ?? wb.Worksheets.FirstOrDefault();

                if (firstSheet == null)
                    return;

                var resolved = new List<(IXLWorksheet Ws, ComposeDxEditableCell Field, string LocalA1)>();

                foreach (var f in editableCells)
                {
                    if (string.IsNullOrWhiteSpace(f.A1))
                        continue;

                    var (sheetNameFromA1, localA1) = SplitSheetAndA1ForCompose(f.A1);
                    var sheetName = !string.IsNullOrWhiteSpace(f.Sheet) ? f.Sheet : sheetNameFromA1;

                    var ws = !string.IsNullOrWhiteSpace(sheetName)
                        ? wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, sheetName, StringComparison.OrdinalIgnoreCase)) ?? firstSheet
                        : firstSheet;

                    if (ws == null || string.IsNullOrWhiteSpace(localA1))
                        continue;

                    resolved.Add((ws, f, localA1));
                }

                if (resolved.Count == 0)
                    return;

                // 시트별로 문서 범위는 기본 Lock으로 초기화한다.
                // 이렇게 해야 기존 파일에서 임의로 Unlock 되어 있던 A1 같은 셀이 작성 화면에서 열리지 않는다.
                foreach (var g in resolved.GroupBy(x => x.Ws.Name, StringComparer.OrdinalIgnoreCase))
                {
                    var ws = wb.Worksheet(g.Key);
                    try { if (ws.IsProtected) ws.Unprotect(); } catch { }

                    var baseRange = ResolveComposeDxProtectionRange(ws, g.Select(x => x.LocalA1));
                    if (baseRange != null)
                    {
                        baseRange.Style.Protection.Locked = true;
                    }
                }

                foreach (var item in resolved)
                {
                    var ws = item.Ws;
                    var f = item.Field;
                    var localA1 = item.LocalA1;

                    try
                    {
                        var range = ws.Range(localA1);
                        var targetRange = range.FirstCell().IsMerged()
                            ? range.FirstCell().MergedRange()
                            : range;

                        // 병합셀은 좌상단 1개만이 아니라 병합 범위 전체를 Unlock 해야 DX/Excel 양쪽에서 안전하다.
                        targetRange.Style.Protection.Locked = false;
                        targetRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF6FF");
                        targetRange.Style.Alignment.WrapText = false;

                        if (string.Equals(f.Type, "Date", StringComparison.OrdinalIgnoreCase))
                        {
                            var firstCell = targetRange.FirstCell();
                            if (firstCell.IsEmpty())
                                firstCell.Value = DateTime.Today;
                            firstCell.Style.DateFormat.Format = GetExcelDateFormatByCulture(CultureInfo.CurrentUICulture);
                        }
                    }
                    catch
                    {
                    }
                }

                foreach (var sheetName in resolved.Select(x => x.Ws.Name).Distinct(StringComparer.OrdinalIgnoreCase))
                {
                    try
                    {
                        var ws = wb.Worksheet(sheetName);
                        if (ws.IsProtected)
                        {
                            try { ws.Unprotect(); } catch { }
                        }

                        ws.Protect(allowedElements:
                            XLSheetProtectionElements.SelectLockedCells |
                            XLSheetProtectionElements.SelectUnlockedCells);
                    }
                    catch
                    {
                    }
                }

                wb.SaveAs(tempPath);
            }

            if (System.IO.File.Exists(tempPath))
            {
                System.IO.File.Delete(workbookPath);
                System.IO.File.Move(tempPath, workbookPath);
            }
        }

        private static void TryRemoveWorksheetProtectionForCompose(string workbookPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath) || !System.IO.File.Exists(workbookPath))
                return;

            try
            {
                using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(workbookPath, true);
                var wbPart = doc.WorkbookPart;
                if (wbPart == null)
                    return;

                foreach (var wsPart in wbPart.WorksheetParts)
                {
                    var worksheet = wsPart.Worksheet;
                    if (worksheet == null)
                        continue;

                    var protections = worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetProtection>().ToList();
                    foreach (var p in protections)
                        p.Remove();

                    worksheet.Save();
                }
            }
            catch
            {
            }
        }

        private static IXLRange? ResolveComposeDxProtectionRange(IXLWorksheet ws, IEnumerable<string> editableLocalA1s)
        {
            int firstRow = int.MaxValue;
            int firstCol = int.MaxValue;
            int lastRow = 0;
            int lastCol = 0;

            void IncludeRange(IXLRange? r)
            {
                if (r == null) return;
                var a = r.RangeAddress;
                firstRow = Math.Min(firstRow, a.FirstAddress.RowNumber);
                firstCol = Math.Min(firstCol, a.FirstAddress.ColumnNumber);
                lastRow = Math.Max(lastRow, a.LastAddress.RowNumber);
                lastCol = Math.Max(lastCol, a.LastAddress.ColumnNumber);
            }

            // 2026.06.12 Fixed: PrintArea가 있더라도 UsedRange를 함께 포함한다.
            // PrintArea가 B1:Q40이면 A1 같은 UsedRange 영역이 잠금 초기화에서 빠질 수 있으므로 둘 다 포함한다.
            try
            {
                var pas = ws.PageSetup.PrintAreas;
                if (pas != null && pas.Any())
                {
                    foreach (var pa in pas)
                        IncludeRange(pa);
                }
            }
            catch
            {
            }

            try
            {
                IncludeRange(ws.RangeUsed(XLCellsUsedOptions.All));
            }
            catch
            {
            }

            foreach (var a1 in editableLocalA1s ?? Enumerable.Empty<string>())
            {
                if (string.IsNullOrWhiteSpace(a1))
                    continue;

                try
                {
                    var r = ws.Range(a1);
                    var effective = r.FirstCell().IsMerged()
                        ? r.FirstCell().MergedRange()
                        : r;
                    IncludeRange(effective);
                }
                catch
                {
                }
            }

            if (firstRow == int.MaxValue || firstCol == int.MaxValue || lastRow <= 0 || lastCol <= 0)
                return null;

            firstRow = Math.Max(1, firstRow);
            firstCol = Math.Max(1, firstCol);
            lastRow = Math.Max(firstRow, lastRow);
            lastCol = Math.Max(firstCol, lastCol);

            return ws.Range(firstRow, firstCol, lastRow, lastCol);
        }

        private static List<ComposeDxEditableCell> ParseComposeDxEditableCells(string? descriptorJson)
        {
            var result = new List<ComposeDxEditableCell>();
            if (string.IsNullOrWhiteSpace(descriptorJson))
                return result;

            try
            {
                using var doc = JsonDocument.Parse(descriptorJson);
                var root = doc.RootElement;

                if (TryGetJsonArrayProperty(root, out var inputs, "inputs", "Inputs"))
                {
                    foreach (var item in inputs.EnumerateArray())
                    {
                        if (item.ValueKind != JsonValueKind.Object)
                            continue;

                        var key = GetJsonStringProperty(item, "key", "Key") ?? string.Empty;
                        var type = GetJsonStringProperty(item, "type", "Type") ?? "Text";
                        var a1 = GetJsonStringProperty(item, "a1", "A1", "cellA1", "CellA1") ?? string.Empty;
                        var sheet = GetJsonStringProperty(item, "sheet", "Sheet") ?? string.Empty;

                        if (!string.IsNullOrWhiteSpace(a1))
                        {
                            result.Add(new ComposeDxEditableCell
                            {
                                Sheet = sheet.Trim(),
                                Key = key.Trim(),
                                Type = NormalizeComposeDxFieldType(type),
                                A1 = a1.Trim()
                            });
                        }
                    }
                }

                if (TryGetJsonArrayProperty(root, out var fields, "Fields", "fields"))
                {
                    foreach (var item in fields.EnumerateArray())
                    {
                        if (item.ValueKind != JsonValueKind.Object)
                            continue;

                        var key = GetJsonStringProperty(item, "Key", "key") ?? string.Empty;
                        var type = GetJsonStringProperty(item, "Type", "type") ?? "Text";
                        var a1 = GetJsonStringProperty(item, "A1", "a1", "CellA1", "cellA1") ?? string.Empty;
                        var sheet = GetJsonStringProperty(item, "Sheet", "sheet") ?? string.Empty;

                        if (string.IsNullOrWhiteSpace(a1) && TryGetJsonObjectProperty(item, out var cell, "Cell", "cell"))
                        {
                            a1 = GetJsonStringProperty(cell, "A1", "a1", "CellA1", "cellA1") ?? string.Empty;
                            sheet = GetJsonStringProperty(cell, "Sheet", "sheet") ?? sheet;
                        }

                        if (!string.IsNullOrWhiteSpace(a1))
                        {
                            result.Add(new ComposeDxEditableCell
                            {
                                Sheet = (sheet ?? string.Empty).Trim(),
                                Key = key.Trim(),
                                Type = NormalizeComposeDxFieldType(type),
                                A1 = a1.Trim()
                            });
                        }
                    }
                }
            }
            catch
            {
            }

            return result
                .Where(x => !string.IsNullOrWhiteSpace(x.A1))
                .GroupBy(x => ((x.Sheet ?? string.Empty).Trim() + "!" + NormalizeA1ForComposeDxCell(x.A1)), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.Last())
                .ToList();
        }

        private static bool TryGetJsonArrayProperty(JsonElement obj, out JsonElement value, params string[] names)
        {
            foreach (var name in names)
            {
                if (obj.TryGetProperty(name, out value) && value.ValueKind == JsonValueKind.Array)
                    return true;
            }

            value = default;
            return false;
        }

        private static bool TryGetJsonObjectProperty(JsonElement obj, out JsonElement value, params string[] names)
        {
            foreach (var name in names)
            {
                if (obj.TryGetProperty(name, out value) && value.ValueKind == JsonValueKind.Object)
                    return true;
            }

            value = default;
            return false;
        }

        private static string? GetJsonStringProperty(JsonElement obj, params string[] names)
        {
            foreach (var name in names)
            {
                if (!obj.TryGetProperty(name, out var value))
                    continue;

                if (value.ValueKind == JsonValueKind.String)
                    return value.GetString();

                if (value.ValueKind == JsonValueKind.Number || value.ValueKind == JsonValueKind.True || value.ValueKind == JsonValueKind.False)
                    return value.ToString();
            }

            return null;
        }

        private static string NormalizeComposeDxFieldType(string? type)
        {
            var s = (type ?? string.Empty).Trim().ToLowerInvariant();
            if (s.StartsWith("date")) return "Date";
            if (s.StartsWith("num") || s.Contains("number") || s.Contains("decimal") || s.Contains("integer")) return "Num";
            return "Text";
        }

        private static (string sheetName, string localA1) SplitSheetAndA1ForCompose(string? a1)
        {
            var s = (a1 ?? string.Empty).Trim().Replace("$", string.Empty).Replace("'", string.Empty);
            if (string.IsNullOrWhiteSpace(s))
                return (string.Empty, string.Empty);

            var bang = s.LastIndexOf('!');
            if (bang >= 0)
                return (s[..bang].Trim(), s[(bang + 1)..].Trim());

            return (string.Empty, s);
        }

        private static string NormalizeA1ForComposeDxCell(string? a1)
        {
            var (_, localA1) = SplitSheetAndA1ForCompose(a1);
            return (localA1 ?? string.Empty).Trim().ToUpperInvariant();
        }

        private static string ComposeDxA1FromRowCol(int row, int col)
        {
            if (row < 1 || col < 1)
                return string.Empty;

            var letters = string.Empty;
            var n = col;
            while (n > 0)
            {
                var r = (n - 1) % 26;
                letters = (char)('A' + r) + letters;
                n = (n - 1) / 26;
            }

            return letters + row.ToString(CultureInfo.InvariantCulture);
        }

        private static void HideComposeDxWorkbookGridLines(string workbookPath)
        {
            if (string.IsNullOrWhiteSpace(workbookPath) || !System.IO.File.Exists(workbookPath))
                return;

            using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(workbookPath, true);
            var wbPart = doc.WorkbookPart;
            if (wbPart == null)
                return;

            foreach (var wsPart in wbPart.WorksheetParts)
            {
                var worksheet = wsPart.Worksheet;
                if (worksheet == null)
                    continue;

                var sheetViews = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetViews>();
                if (sheetViews == null)
                {
                    sheetViews = new DocumentFormat.OpenXml.Spreadsheet.SheetViews();
                    worksheet.InsertAt(sheetViews, 0);
                }

                var sheetView = sheetViews.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetView>().FirstOrDefault();
                if (sheetView == null)
                {
                    sheetView = new DocumentFormat.OpenXml.Spreadsheet.SheetView
                    {
                        WorkbookViewId = 0U
                    };
                    sheetViews.Append(sheetView);
                }

                sheetView.ShowGridLines = false;
                worksheet.Save();
            }
        }

        private async Task<ComposeTemplateVisualMetrics> LoadComposeTemplateVisualMetricsAsync(long versionId, string? templateCode)
        {
            var result = new ComposeTemplateVisualMetrics();
            var cs = _cfg.GetConnectionString("DefaultConnection");
            if (string.IsNullOrWhiteSpace(cs))
                return result;

            static void FillMetrics(System.Data.Common.DbDataReader rd, ComposeTemplateVisualMetrics target)
            {
                target.VisualWidthPx = rd["VisualWidthPx"] == DBNull.Value ? null : Convert.ToInt32(rd["VisualWidthPx"], CultureInfo.InvariantCulture);
                target.VisualHeightPx = rd["VisualHeightPx"] == DBNull.Value ? null : Convert.ToInt32(rd["VisualHeightPx"], CultureInfo.InvariantCulture);
                target.VisualSource = rd["VisualSource"] as string;
                target.VisualRangeA1 = rd["VisualRangeA1"] as string;
                target.VisualMetricRuleCode = rd["VisualMetricRuleCode"] as string;
            }

            try
            {
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                if (versionId > 0)
                {
                    await using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
SELECT TOP (1)
       VisualWidthPx,
       VisualHeightPx,
       VisualSource,
       VisualRangeA1,
       VisualMetricRuleCode
FROM dbo.DocTemplateVersion
WHERE Id = @VersionId;";
                        cmd.Parameters.Add(new SqlParameter("@VersionId", SqlDbType.BigInt) { Value = versionId });

                        await using var rd = await cmd.ExecuteReaderAsync();
                        if (await rd.ReadAsync())
                        {
                            FillMetrics(rd, result);
                            return result;
                        }
                    }
                }

                var code = (templateCode ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(code))
                    return result;

                await using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
SELECT TOP (1)
       v.VisualWidthPx,
       v.VisualHeightPx,
       v.VisualSource,
       v.VisualRangeA1,
       v.VisualMetricRuleCode
FROM dbo.DocTemplateMaster m
JOIN dbo.DocTemplateVersion v ON v.TemplateId = m.Id
WHERE m.DocCode = @DocCode
ORDER BY v.VersionNo DESC, v.Id DESC;";
                    cmd.Parameters.Add(new SqlParameter("@DocCode", SqlDbType.NVarChar, 100) { Value = code });

                    await using var rd = await cmd.ExecuteReaderAsync();
                    if (await rd.ReadAsync())
                        FillMetrics(rd, result);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "COMPOSEDX-PERF visual metrics direct query failed templateCode={TemplateCode} versionId={VersionId}", templateCode ?? string.Empty, versionId);
            }

            return result;
        }

        private async Task<List<ComposeDxEditableCell>> LoadComposeEditableCellsAsync(long versionId)
        {
            var result = new List<ComposeDxEditableCell>();
            if (versionId <= 0)
                return result;

            var cs = _cfg.GetConnectionString("DefaultConnection");
            if (string.IsNullOrWhiteSpace(cs))
                return result;

            try
            {
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
SELECT [Key], [Type], Sheet, A1, CellA1, CellRow, CellColumn
FROM dbo.DocTemplateField
WHERE VersionId = @VersionId
ORDER BY Id;";
                cmd.Parameters.Add(new SqlParameter("@VersionId", SqlDbType.BigInt) { Value = versionId });

                await using var rd = await cmd.ExecuteReaderAsync();
                while (await rd.ReadAsync())
                {
                    var key = rd["Key"] as string ?? string.Empty;
                    var type = NormalizeComposeDxFieldType(rd["Type"] as string);
                    var sheet = rd["Sheet"] as string ?? string.Empty;
                    var a1 = rd["A1"] as string ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(a1))
                        a1 = rd["CellA1"] as string ?? string.Empty;

                    if (string.IsNullOrWhiteSpace(a1))
                    {
                        var row = rd["CellRow"] == DBNull.Value ? 0 : Convert.ToInt32(rd["CellRow"], CultureInfo.InvariantCulture);
                        var col = rd["CellColumn"] == DBNull.Value ? 0 : Convert.ToInt32(rd["CellColumn"], CultureInfo.InvariantCulture);
                        if (row > 0 && col > 0)
                            a1 = ComposeDxA1FromRowCol(row, col);
                    }

                    if (string.IsNullOrWhiteSpace(a1))
                        continue;

                    result.Add(new ComposeDxEditableCell
                    {
                        Sheet = sheet.Trim(),
                        Key = key.Trim(),
                        Type = type,
                        A1 = a1.Trim()
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "COMPOSEDX-PERF editable cells direct query failed versionId={VersionId}", versionId);
            }

            return result
                .Where(x => !string.IsNullOrWhiteSpace(x.A1))
                .GroupBy(x => ((x.Sheet ?? string.Empty).Trim() + "!" + NormalizeA1ForComposeDxCell(x.A1)), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.Last())
                .ToList();
        }

        private static string UpsertComposeInputsIntoDescriptorJson(string? descriptorJson, IReadOnlyList<ComposeDxEditableCell> editableCells)
        {
            if (editableCells == null || editableCells.Count == 0)
                return string.IsNullOrWhiteSpace(descriptorJson) ? "{}" : descriptorJson!;

            try
            {
                var root = System.Text.Json.Nodes.JsonNode.Parse(string.IsNullOrWhiteSpace(descriptorJson) ? "{}" : descriptorJson) as System.Text.Json.Nodes.JsonObject
                    ?? new System.Text.Json.Nodes.JsonObject();

                var inputs = new System.Text.Json.Nodes.JsonArray();
                foreach (var f in editableCells)
                {
                    var obj = new System.Text.Json.Nodes.JsonObject
                    {
                        ["key"] = f.Key ?? string.Empty,
                        ["type"] = NormalizeComposeDxFieldType(f.Type),
                        ["required"] = false,
                        ["a1"] = f.A1 ?? string.Empty
                    };

                    if (!string.IsNullOrWhiteSpace(f.Sheet))
                        obj["sheet"] = f.Sheet;

                    inputs.Add(obj);
                }

                root["inputs"] = inputs;

                // 혼선을 막기 위해 좌표 없는 legacy Fields는 작성 화면용 descriptor에서 제거한다.
                root.Remove("Fields");
                root.Remove("fields");

                return root.ToJsonString(new JsonSerializerOptions
                {
                    Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                    WriteIndented = false
                });
            }
            catch
            {
                return string.IsNullOrWhiteSpace(descriptorJson) ? "{}" : descriptorJson!;
            }
        }

        private static string SanitizeComposeDxFilePart(string? value)
        {
            var s = string.IsNullOrWhiteSpace(value) ? "template" : value.Trim();
            foreach (var ch in Path.GetInvalidFileNameChars())
                s = s.Replace(ch, '_');

            s = Regex.Replace(s, @"[^a-zA-Z0-9_\-.]", "_");
            if (s.Length > 80) s = s[..80];
            return string.IsNullOrWhiteSpace(s) ? "template" : s;
        }

        private static bool TryGetLastCellFromRangeA1(string? rangeA1, out string lastA1, out int lastRow1, out int lastCol1)
        {
            lastA1 = "A1";
            lastRow1 = 1;
            lastCol1 = 1;

            var s = (rangeA1 ?? string.Empty).Trim().Replace("$", string.Empty).Replace("'", string.Empty);
            if (string.IsNullOrWhiteSpace(s))
                return false;

            var bang = s.LastIndexOf('!');
            if (bang >= 0)
                s = s[(bang + 1)..];

            var parts = s.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
            var target = parts.Length >= 2 ? parts[^1] : parts[0];

            var m = Regex.Match(target, @"^([A-Z]{1,3})(\d{1,7})$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            if (!m.Success)
                return false;

            var colText = m.Groups[1].Value.ToUpperInvariant();
            var col = 0;
            foreach (var ch in colText)
                col = (col * 26) + (ch - 'A' + 1);

            if (!int.TryParse(m.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var row) || row <= 0 || col <= 0)
                return false;

            lastRow1 = row;
            lastCol1 = col;
            lastA1 = colText + row.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        [HttpPost("DeleteDxTemp")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public IActionResult DeleteDxTemp([FromBody] DeleteDxTempRequest request)
        {
            if (string.IsNullOrWhiteSpace(request?.DxDocumentId))
                return Json(new { ok = false, detail = "dxDocumentId missing" });

            try
            {
                var dir = Path.Combine(_env.ContentRootPath, "App_Data", "DocDxCompose");
                if (!Directory.Exists(dir))
                    return Json(new { ok = true, detail = "dir not found" });

                // dxDocumentId가 포함된 파일만 삭제 (다른 사용자 파일 건드리지 않음)
                var files = Directory.GetFiles(dir, $"*_{request.DxDocumentId}.*")
                    .Where(f => string.Equals(Path.GetExtension(f), ".xlsx", StringComparison.OrdinalIgnoreCase)
                             || string.Equals(Path.GetExtension(f), ".xlsm", StringComparison.OrdinalIgnoreCase))
                    .ToArray();

                foreach (var file in files)
                {
                    try { System.IO.File.Delete(file); }
                    catch (Exception ex)
                    { _log.LogWarning(ex, "DeleteDxTemp 파일 삭제 실패: {file}", file); }
                }

                _log.LogInformation("DeleteDxTemp 완료 dxDocumentId={id} 삭제={count}",
                    request.DxDocumentId, files.Length);

                return Json(new { ok = true, deleted = files.Length });
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "DeleteDxTemp 오류 dxDocumentId={id}", request?.DxDocumentId);
                return Json(new { ok = false, detail = ex.Message });
            }
        }

        [HttpGet("dx-callback")]
        [HttpPost("dx-callback")]
        public IActionResult DxCallback()
        {
            // 2026.06.10 Changed: DevExpress Spreadsheet callback 지연 원인 분리를 위한 계측 강화.
            // Request.Form / ReadFormAsync는 DevExpress 내부 처리에 영향을 줄 수 있으므로 여기서는 읽지 않는다.
            var totalSw = System.Diagnostics.Stopwatch.StartNew();
            var processorSw = new System.Diagnostics.Stopwatch();
            var req = HttpContext.Request;
            var traceId = HttpContext.TraceIdentifier ?? string.Empty;
            var callbackId = Guid.NewGuid().ToString("N").Substring(0, 8);
            var method = req.Method;
            var path = req.Path.Value ?? string.Empty;
            var query = req.QueryString.Value ?? string.Empty;
            var dxwsid = req.Query["dxwsid"].ToString();
            var contentLength = req.ContentLength;
            var contentType = req.ContentType ?? string.Empty;
            var hasFormContentType = req.HasFormContentType;
            var referer = req.Headers["Referer"].ToString();
            var userAgent = req.Headers["User-Agent"].ToString();

            HttpContext.Response.OnStarting(() =>
            {
                try
                {
                    var timing = $"dx-callback-total;dur={totalSw.Elapsed.TotalMilliseconds:0.##}";
                    var existing = HttpContext.Response.Headers["Server-Timing"].ToString();
                    HttpContext.Response.Headers["Server-Timing"] = string.IsNullOrWhiteSpace(existing)
                        ? timing
                        : existing + ", " + timing;
                }
                catch
                {
                }

                return Task.CompletedTask;
            });

            _log.LogInformation(
                "DX-CALLBACK start id={CallbackId} traceId={TraceId} method={Method} path={Path} query={Query} dxwsid={Dxwsid} contentLength={ContentLength} contentType={ContentType} hasFormContentType={HasFormContentType} referer={Referer} userAgent={UserAgent}",
                callbackId,
                traceId,
                method,
                path,
                query,
                dxwsid,
                contentLength,
                contentType,
                hasFormContentType,
                referer,
                userAgent
            );

            try
            {
                _log.LogInformation(
                    "DX-CALLBACK before-getresponse id={CallbackId} elapsedMs={ElapsedMs} dxwsid={Dxwsid}",
                    callbackId,
                    totalSw.Elapsed.TotalMilliseconds,
                    dxwsid
                );

                processorSw.Start();
                var result = SpreadsheetRequestProcessor.GetResponse(HttpContext);
                processorSw.Stop();

                _log.LogInformation(
                    "DX-CALLBACK after-getresponse id={CallbackId} processorMs={ProcessorMs} totalMs={TotalMs} resultType={ResultType} dxwsid={Dxwsid}",
                    callbackId,
                    processorSw.Elapsed.TotalMilliseconds,
                    totalSw.Elapsed.TotalMilliseconds,
                    result?.GetType().FullName ?? string.Empty,
                    dxwsid
                );

                return new DxCallbackTimedActionResult(
                    result,
                    _log,
                    callbackId,
                    traceId,
                    method,
                    path,
                    query,
                    dxwsid,
                    processorSw.Elapsed.TotalMilliseconds,
                    totalSw
                );
            }
            catch (Microsoft.AspNetCore.Http.BadHttpRequestException ex) when ((ex.Message ?? string.Empty).IndexOf("Unexpected end of request content", StringComparison.OrdinalIgnoreCase) >= 0 || HttpContext.RequestAborted.IsCancellationRequested)
            {
                // 2026.06.15 Changed: ASP.NET Core 8에서 이동된 BadHttpRequestException 네임스페이스로 변경한다.
                // 브라우저 이동/새로고침/탭 닫기 등으로 DevExpress Spreadsheet callback 본문이
                // 중간에 끊긴 경우 사용자가 취소한 요청으로 보고 499로 종료한다.
                if (processorSw.IsRunning) processorSw.Stop();
                totalSw.Stop();

                _log.LogInformation(
                    ex,
                    "DX-CALLBACK aborted id={CallbackId} processorMs={ProcessorMs} totalMs={TotalMs} method={Method} path={Path} query={Query} dxwsid={Dxwsid} reason=UnexpectedEndOfRequestContent",
                    callbackId,
                    processorSw.Elapsed.TotalMilliseconds,
                    totalSw.Elapsed.TotalMilliseconds,
                    method,
                    path,
                    query,
                    dxwsid
                );

                return StatusCode(499);
            }

            catch (OperationCanceledException ex) when (HttpContext.RequestAborted.IsCancellationRequested)
            {
                // 2026.06.11 Changed: 클라이언트가 callback 요청을 취소한 경우 정상 종료 처리.
                if (processorSw.IsRunning) processorSw.Stop();
                totalSw.Stop();

                _log.LogInformation(
                    ex,
                    "DX-CALLBACK aborted id={CallbackId} processorMs={ProcessorMs} totalMs={TotalMs} method={Method} path={Path} query={Query} dxwsid={Dxwsid} reason=RequestAborted",
                    callbackId,
                    processorSw.Elapsed.TotalMilliseconds,
                    totalSw.Elapsed.TotalMilliseconds,
                    method,
                    path,
                    query,
                    dxwsid
                );

                return StatusCode(499);
            }
            catch (Exception ex)
            {
                if (processorSw.IsRunning) processorSw.Stop();
                totalSw.Stop();

                _log.LogError(
                    ex,
                    "DX-CALLBACK failed id={CallbackId} processorMs={ProcessorMs} totalMs={TotalMs} method={Method} path={Path} query={Query} dxwsid={Dxwsid}",
                    callbackId,
                    processorSw.Elapsed.TotalMilliseconds,
                    totalSw.Elapsed.TotalMilliseconds,
                    method,
                    path,
                    query,
                    dxwsid
                );

                throw;
            }
        }

        private sealed class DxCallbackTimedActionResult : IActionResult
        {
            private readonly IActionResult? _inner;
            private readonly ILogger _log;
            private readonly string _callbackId;
            private readonly string _traceId;
            private readonly string _method;
            private readonly string _path;
            private readonly string _query;
            private readonly string _dxwsid;
            private readonly double _processorMs;
            private readonly System.Diagnostics.Stopwatch _totalSw;

            public DxCallbackTimedActionResult(
                IActionResult? inner,
                ILogger log,
                string callbackId,
                string traceId,
                string method,
                string path,
                string query,
                string dxwsid,
                double processorMs,
                System.Diagnostics.Stopwatch totalSw)
            {
                _inner = inner;
                _log = log;
                _callbackId = callbackId;
                _traceId = traceId;
                _method = method;
                _path = path;
                _query = query;
                _dxwsid = dxwsid;
                _processorMs = processorMs;
                _totalSw = totalSw;
            }

            public async Task ExecuteResultAsync(ActionContext context)
            {
                var executeSw = System.Diagnostics.Stopwatch.StartNew();

                _log.LogInformation(
                    "DX-CALLBACK before-execute id={CallbackId} traceId={TraceId} elapsedMs={ElapsedMs} processorMs={ProcessorMs} dxwsid={Dxwsid}",
                    _callbackId,
                    _traceId,
                    _totalSw.Elapsed.TotalMilliseconds,
                    _processorMs,
                    _dxwsid
                );

                try
                {
                    if (_inner == null)
                        throw new InvalidOperationException("SpreadsheetRequestProcessor.GetResponse returned null IActionResult.");

                    await _inner.ExecuteResultAsync(context);
                    executeSw.Stop();
                    _totalSw.Stop();

                    _log.LogInformation(
                        "DX-CALLBACK after-execute id={CallbackId} traceId={TraceId} executeMs={ExecuteMs} processorMs={ProcessorMs} totalMs={TotalMs} statusCode={StatusCode} method={Method} path={Path} query={Query} dxwsid={Dxwsid}",
                        _callbackId,
                        _traceId,
                        executeSw.Elapsed.TotalMilliseconds,
                        _processorMs,
                        _totalSw.Elapsed.TotalMilliseconds,
                        context.HttpContext.Response.StatusCode,
                        _method,
                        _path,
                        _query,
                        _dxwsid
                    );
                }
                catch (Exception ex)
                {
                    executeSw.Stop();
                    _totalSw.Stop();

                    _log.LogError(
                        ex,
                        "DX-CALLBACK execute-failed id={CallbackId} traceId={TraceId} executeMs={ExecuteMs} processorMs={ProcessorMs} totalMs={TotalMs} statusCode={StatusCode} method={Method} path={Path} query={Query} dxwsid={Dxwsid}",
                        _callbackId,
                        _traceId,
                        executeSw.Elapsed.TotalMilliseconds,
                        _processorMs,
                        _totalSw.Elapsed.TotalMilliseconds,
                        context.HttpContext.Response.StatusCode,
                        _method,
                        _path,
                        _query,
                        _dxwsid
                    );

                    throw;
                }
            }
        }

        [HttpPost("BatchEditDX")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public IActionResult BatchEditDX([FromBody] DxBatchEditRequest request)
        {
            if (request == null || request.SpreadsheetState == null)
                return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "batch-arg", detail = "spreadsheetState missing" });

            var ops = request.Operations?
                .Where(x => x != null && !string.IsNullOrWhiteSpace(x.A1))
                .ToList() ?? new List<DxBatchEditOperation>();

            if (ops.Count == 0)
                return Json(new { ok = true, changed = 0 });

            try
            {
                var spreadsheet = SpreadsheetRequestProcessor.GetSpreadsheetFromState(request.SpreadsheetState);
                if (spreadsheet == null)
                    return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "batch-state", detail = "spreadsheet session not found" });

                var workbook = spreadsheet.Document;
                if (workbook == null || workbook.Worksheets.Count == 0)
                    return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "batch-doc", detail = "workbook not available" });

                DevExpress.Spreadsheet.Worksheet ws;
                try { ws = workbook.Worksheets.ActiveWorksheet; }
                catch { ws = workbook.Worksheets[0]; }

                var changed = 0;
                foreach (var op in ops)
                {
                    var a1 = NormalizeBatchA1(op.A1);
                    if (string.IsNullOrWhiteSpace(a1)) continue;
                    try
                    {
                        var cell = ws.Cells[a1];
                        if (op.Clear) { cell.Value = string.Empty; changed++; continue; }

                        var rawType = (op.Type ?? string.Empty).Trim();
                        var rawValue = op.Value ?? string.Empty;

                        if (string.Equals(rawType, "Date", StringComparison.OrdinalIgnoreCase))
                        {
                            if (DateTime.TryParse(rawValue, out var dt))
                            {
                                cell.Value = dt;
                                cell.NumberFormat = GetExcelDateFormatByCulture(CultureInfo.CurrentUICulture);
                            }
                            else cell.Value = rawValue;
                        }
                        else if (string.Equals(rawType, "Num", StringComparison.OrdinalIgnoreCase)
                              || string.Equals(rawType, "Number", StringComparison.OrdinalIgnoreCase))
                        {
                            var normalized = rawValue.Replace(",", string.Empty);
                            if (decimal.TryParse(normalized, NumberStyles.Any, CultureInfo.InvariantCulture, out var dec))
                                cell.Value = dec;
                            else cell.Value = rawValue;
                        }
                        else
                        {
                            cell.Value = rawValue;
                        }

                        if (op.Wrap.HasValue) cell.Alignment.WrapText = op.Wrap.Value;
                        changed++;
                    }
                    catch (Exception exCell)
                    {
                        _log.LogDebug(exCell, "BatchEditDX cell apply skipped a1={a1}", a1);
                    }
                }

                spreadsheet.Save();
                return Json(new { ok = true, changed });
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "BatchEditDX failed");
                return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "batch", detail = ex.Message });
            }
        }

        private static string NormalizeBatchA1(string? raw)
        {
            var s = (raw ?? string.Empty).Trim().ToUpperInvariant();
            if (string.IsNullOrWhiteSpace(s)) return string.Empty;
            var m = Regex.Match(s, @"^([A-Z]+)(\d+)$", RegexOptions.CultureInvariant);
            if (!m.Success) return string.Empty;
            return $"{m.Groups[1].Value}{int.Parse(m.Groups[2].Value, CultureInfo.InvariantCulture)}";
        }

        private static void HideComposeDxSheetGridLines(string workbookPath, string sheetName)
        {
            using var doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(workbookPath, true);
            var wbPart = doc.WorkbookPart;
            if (wbPart?.Workbook?.Sheets == null) return;

            var sheets = wbPart.Workbook.Sheets.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>().ToList();
            if (sheets.Count == 0) return;

            var sheet = sheets.FirstOrDefault(s =>
                string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                ?? sheets.First();

            if (sheet.Id?.Value == null) return;

            var wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)wbPart.GetPartById(sheet.Id.Value);
            var worksheet = wsPart.Worksheet;
            if (worksheet == null) return;

            var sheetViews = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetViews>();
            if (sheetViews == null)
            {
                sheetViews = new DocumentFormat.OpenXml.Spreadsheet.SheetViews();
                worksheet.InsertAt(sheetViews, 0);
            }

            var sheetView = sheetViews.Elements<DocumentFormat.OpenXml.Spreadsheet.SheetView>().FirstOrDefault();
            if (sheetView == null)
            {
                sheetView = new DocumentFormat.OpenXml.Spreadsheet.SheetView
                {
                    WorkbookViewId = 0U
                };
                sheetViews.Append(sheetView);
            }

            sheetView.ShowGridLines = false;
            worksheet.Save();
        }

        private static string ResolveApproverValueToEmail(SqlConnection conn, string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return value;
            if (value.Contains('@')) return value;
            try
            {
                using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
SELECT TOP 1 Email FROM dbo.AspNetUsers
WHERE Id = @v OR UserName = @v OR NormalizedUserName = UPPER(@v);";
                cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
                var email = cmd.ExecuteScalar() as string;
                return !string.IsNullOrWhiteSpace(email) ? email.Trim() : value;
            }
            catch { return value; }
        }

        private string LoadCooperationsDirectFromLegacyDb(string templateCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cs)) return "[]";
            try
            {
                using var conn = new SqlConnection(cs);
                conn.Open();
                using var cmd = new SqlCommand(@"
SELECT TOP 1 v.DescriptorJson
FROM dbo.DocTemplateVersion v
JOIN dbo.DocTemplateMaster m ON m.Id = v.TemplateId
WHERE m.DocCode = @code
ORDER BY v.VersionNo DESC;", conn);
                cmd.Parameters.Add(new SqlParameter("@code", SqlDbType.NVarChar, 100) { Value = templateCode });
                var raw = cmd.ExecuteScalar() as string;
                if (string.IsNullOrWhiteSpace(raw)) return "[]";

                using var doc = JsonDocument.Parse(raw);
                var root = doc.RootElement;

                JsonElement coopEl;
                bool found = root.TryGetProperty("Cooperations", out coopEl)
                          || root.TryGetProperty("cooperations", out coopEl);

                if (!found || coopEl.ValueKind != JsonValueKind.Array || coopEl.GetArrayLength() == 0)
                    return "[]";

                var list = new List<Dictionary<string, object>>();
                int seq = 1;
                foreach (var c in coopEl.EnumerateArray())
                {
                    string? val = null;
                    if (c.TryGetProperty("ApproverValue", out var av) && av.ValueKind == JsonValueKind.String)
                        val = av.GetString();
                    if (string.IsNullOrWhiteSpace(val) && c.TryGetProperty("value", out var vv) && vv.ValueKind == JsonValueKind.String)
                        val = vv.GetString();
                    if (string.IsNullOrWhiteSpace(val)) continue;

                    val = val.Trim();
                    val = ResolveApproverValueToEmail(conn, val);

                    string typ = "Person";
                    if (c.TryGetProperty("ApproverType", out var at) && at.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(at.GetString()))
                        typ = at.GetString()!;
                    else if (c.TryGetProperty("approverType", out var at2) && at2.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(at2.GetString()))
                        typ = at2.GetString()!;

                    // ── A1 탐색: Cell.A1 → A1 → cellA1 ──
                    string cellA1 = "";
                    if (c.TryGetProperty("Cell", out var cellEl) && cellEl.ValueKind == JsonValueKind.Object)
                    {
                        if (cellEl.TryGetProperty("A1", out var a1El)) cellA1 = a1El.GetString() ?? "";
                        if (string.IsNullOrWhiteSpace(cellA1) && cellEl.TryGetProperty("a1", out var a1El2)) cellA1 = a1El2.GetString() ?? "";
                    }
                    if (string.IsNullOrWhiteSpace(cellA1) && c.TryGetProperty("A1", out var directA1)) cellA1 = directA1.GetString() ?? "";
                    if (string.IsNullOrWhiteSpace(cellA1) && c.TryGetProperty("a1", out var directA1L)) cellA1 = directA1L.GetString() ?? "";
                    if (string.IsNullOrWhiteSpace(cellA1) && c.TryGetProperty("cellA1", out var ca1)) cellA1 = ca1.GetString() ?? "";
                    if (string.IsNullOrWhiteSpace(cellA1) && c.TryGetProperty("CellA1", out var ca2)) cellA1 = ca2.GetString() ?? "";

                    list.Add(new Dictionary<string, object>
                    {
                        ["roleKey"] = $"C{seq}",
                        ["approverType"] = (typ == "Person" || typ == "Role" || typ == "Rule") ? typ : "Person",
                        ["lineType"] = "Cooperation",
                        ["value"] = val,
                        ["cellA1"] = cellA1
                    });
                    seq++;
                }

                return list.Count > 0 ? JsonSerializer.Serialize(list) : "[]";
            }
            catch { return "[]"; }
        }

        [HttpPost("CreateDX")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> CreateDX([FromBody] ComposePostDto dto)
        {
            ComposePostDto? resolvedDto = dto;

            if ((resolvedDto == null || string.IsNullOrWhiteSpace(resolvedDto.TemplateCode)) && Request.HasFormContentType)
            {
                try
                {
                    var form = await Request.ReadFormAsync();
                    var json = form["payload"].FirstOrDefault() ?? form["dto"].FirstOrDefault() ?? form["data"].FirstOrDefault();
                    if (!string.IsNullOrWhiteSpace(json))
                        resolvedDto = JsonSerializer.Deserialize<ComposePostDto>(json, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });
                }
                catch (Exception ex) { _log.LogWarning(ex, "CreateDX: form-data dto parse failed"); }
            }

            dto = resolvedDto!;

            if (dto is null || string.IsNullOrWhiteSpace(dto.TemplateCode))
                return BadRequest(new { messages = new[] { "DOC_Val_TemplateRequired" }, stage = "arg", detail = "templateCode null/empty" });

            var tc = dto.TemplateCode!.Trim();
            var inputsMap = dto.Inputs ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var userCompDept = _helper.GetUserCompDept();
            var compCd = string.IsNullOrWhiteSpace(userCompDept.compCd) ? "0000" : userCompDept.compCd!;
            var deptIdPart = string.IsNullOrWhiteSpace(userCompDept.departmentId) ? "0" : userCompDept.departmentId!;

            var (descriptorJson, _previewJsonFromTpl, title, versionId, excelPathRaw) = await _tpl.LoadMetaAsync(tc);

            excelPathRaw = NormalizeTemplateExcelPath(excelPathRaw);
            var excelPath = _helper.ToContentRootAbsolute(excelPathRaw);

            if (versionId <= 0 || string.IsNullOrWhiteSpace(excelPath) || !System.IO.File.Exists(excelPath))
                return BadRequest(new { messages = new[] { "DOC_Err_TemplateNotReady" }, stage = "generate", detail = $"Excel not found. versionId={versionId}, path='{excelPath}'" });

            string tempExcelFullPath;
            try
            {
                tempExcelFullPath = await GenerateExcelFromInputsAsync(versionId, excelPath, inputsMap, dto.DescriptorVersion);
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "CreateDX Generate failed tc={tc}", tc);
                return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "generate", detail = ex.Message });
            }

            var now = DateTime.Now;
            var destDirRel = Path.Combine("App_Data", "Docs", compCd, now.Year.ToString("D4"), now.Month.ToString("D2"), deptIdPart);
            var docId = Path.GetFileNameWithoutExtension(tempExcelFullPath);
            var excelExt = Path.GetExtension(tempExcelFullPath);
            if (string.IsNullOrWhiteSpace(excelExt)) excelExt = ".xlsx";
            var outputPathForDb = Path.Combine(destDirRel, $"{docId}{excelExt}");

            var filledPreviewJson = BuildPreviewJsonFromExcel(tempExcelFullPath);
            var normalizedDesc = BuildDescriptorJsonWithFlowGroups(tc, descriptorJson);

            // 2026.06.18 Added: 기안자 예약값 문서 생성 시점 치환 Contents 공용 템플릿의 기안자 예약값을 현재 작성자 UserId로 확정
            var currentUserId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
            normalizedDesc = ReplaceDrafterApproverValuesInJson(normalizedDesc, currentUserId);

            List<string> toEmails;
            List<string> diag;
            try
            {
                var recipients = GetInitialRecipients(dto, normalizedDesc, _cfg, docId, tc);
                toEmails = recipients.emails ?? new List<string>();
                diag = recipients.diag ?? new List<string>();
            }
            catch { toEmails = new List<string>(); diag = new List<string> { "recipients resolve error" }; }

            // ── approvals ─────────────────────────────────────────────────
            var approvalsJson = ExtractApprovalsJson(normalizedDesc);
            bool approvalsHasAnyValue = false;
            bool approvalsParsedOk = false;

            if (!string.IsNullOrWhiteSpace(approvalsJson))
            {
                try
                {
                    using var aj = JsonDocument.Parse(approvalsJson);
                    if (aj.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        approvalsParsedOk = true;
                        foreach (var a in aj.RootElement.EnumerateArray())
                        {
                            string? v = null;
                            if (a.TryGetProperty("value", out var v1) && v1.ValueKind == JsonValueKind.String) v = v1.GetString();
                            else if (a.TryGetProperty("approverValue", out var v2) && v2.ValueKind == JsonValueKind.String) v = v2.GetString();
                            else if (a.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == JsonValueKind.String) v = v3.GetString();
                            if (!string.IsNullOrWhiteSpace(v)) { approvalsHasAnyValue = true; break; }
                        }
                    }
                }
                catch { approvalsParsedOk = false; approvalsHasAnyValue = false; }
            }

            if (string.IsNullOrWhiteSpace(approvalsJson) || approvalsJson.Trim() == "[]" || !approvalsParsedOk || !approvalsHasAnyValue)
            {
                if (toEmails.Count > 0)
                {
                    var built = new List<Dictionary<string, object>>();
                    int seq = 1;
                    foreach (var raw in toEmails)
                    {
                        var email = (raw ?? string.Empty).Trim();
                        if (string.IsNullOrWhiteSpace(email)) continue;
                        built.Add(new Dictionary<string, object> { ["roleKey"] = $"A{seq}", ["approverType"] = "Person", ["required"] = true, ["value"] = email });
                        seq++;
                    }
                    if (built.Count == 0) return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "approvals", detail = "No valid approver emails." });
                    approvalsJson = JsonSerializer.Serialize(built);
                }
                else
                {
                    return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "approvals", detail = "Approver not resolved (toEmails empty)." });
                }
            }
            else
            {
                try
                {
                    using var aj = JsonDocument.Parse(approvalsJson);
                    if (aj.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        var list = new List<Dictionary<string, object>>();
                        int seq = 1;
                        foreach (var a in aj.RootElement.EnumerateArray())
                        {
                            var item = new Dictionary<string, object>();
                            string at = "Person";
                            if (a.TryGetProperty("approverType", out var at1) && at1.ValueKind == JsonValueKind.String) at = string.IsNullOrWhiteSpace(at1.GetString()) ? at : at1.GetString()!;
                            else if (a.TryGetProperty("ApproverType", out var at2) && at2.ValueKind == JsonValueKind.String) at = string.IsNullOrWhiteSpace(at2.GetString()) ? at : at2.GetString()!;
                            bool required = false;
                            if (a.TryGetProperty("required", out var r1) && (r1.ValueKind == JsonValueKind.True || r1.ValueKind == JsonValueKind.False)) required = r1.GetBoolean();
                            string? v = null;
                            if (a.TryGetProperty("value", out var v1) && v1.ValueKind == JsonValueKind.String) v = v1.GetString();
                            else if (a.TryGetProperty("approverValue", out var v2) && v2.ValueKind == JsonValueKind.String) v = v2.GetString();
                            else if (a.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == JsonValueKind.String) v = v3.GetString();
                            item["roleKey"] = $"A{seq}"; item["approverType"] = at; item["required"] = required; item["value"] = v ?? string.Empty;
                            list.Add(item); seq++;
                        }
                        approvalsJson = JsonSerializer.Serialize(list);
                    }
                }
                catch { }
            }

            // ── cooperations ──────────────────────────────────────────────
            var cooperationsJson = "[]";

            if (dto.Cooperations?.Steps is { Count: > 0 })
            {
                var coopList = new List<Dictionary<string, object>>();
                int seq = 1;
                foreach (var st in dto.Cooperations.Steps)
                {
                    if (string.IsNullOrWhiteSpace(st?.Value)) continue;
                    coopList.Add(new Dictionary<string, object>
                    {
                        ["roleKey"] = !string.IsNullOrWhiteSpace(st.RoleKey) ? st.RoleKey : $"C{seq}",
                        ["approverType"] = string.IsNullOrWhiteSpace(st.ApproverType) ? "Person" : st.ApproverType,
                        ["lineType"] = string.IsNullOrWhiteSpace(st.LineType) ? "Cooperation" : st.LineType,
                        ["value"] = st.Value.Trim()
                    });
                    seq++;
                }
                if (coopList.Count > 0) cooperationsJson = JsonSerializer.Serialize(coopList);
            }

            if (cooperationsJson == "[]") cooperationsJson = ExtractCooperationsJson(normalizedDesc);
            if (cooperationsJson == "[]" || cooperationsJson.Trim() == "[]") cooperationsJson = LoadCooperationsDirectFromLegacyDb(tc);

            bool cooperationsHasAnyValue = false;
            bool cooperationsParsedOk = false;
            if (!string.IsNullOrWhiteSpace(cooperationsJson) && cooperationsJson.Trim() != "[]")
            {
                try
                {
                    using var cj = JsonDocument.Parse(cooperationsJson);
                    if (cj.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        cooperationsParsedOk = true;
                        foreach (var c in cj.RootElement.EnumerateArray())
                            if (EnumerateCooperationValues(c).Any(v => !string.IsNullOrWhiteSpace(v))) { cooperationsHasAnyValue = true; break; }
                    }
                }
                catch { cooperationsParsedOk = false; cooperationsHasAnyValue = false; }
            }

            if (!string.IsNullOrWhiteSpace(cooperationsJson) && cooperationsJson.Trim() != "[]" && cooperationsParsedOk && cooperationsHasAnyValue)
            {
                try
                {
                    using var cj = JsonDocument.Parse(cooperationsJson);
                    if (cj.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        var list = new List<Dictionary<string, object>>();
                        int seq = 1;
                        foreach (var c in cj.RootElement.EnumerateArray())
                        {
                            var approverValue = EnumerateCooperationValues(c).Select(v => (v ?? string.Empty).Trim()).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
                            if (string.IsNullOrWhiteSpace(approverValue)) continue;

                            var item = new Dictionary<string, object>();
                            string at = "Person";
                            if (c.TryGetProperty("approverType", out var at1) && at1.ValueKind == JsonValueKind.String) at = string.IsNullOrWhiteSpace(at1.GetString()) ? at : at1.GetString()!;
                            else if (c.TryGetProperty("ApproverType", out var at2) && at2.ValueKind == JsonValueKind.String) at = string.IsNullOrWhiteSpace(at2.GetString()) ? at : at2.GetString()!;

                            string roleKey = string.Empty;
                            if (c.TryGetProperty("roleKey", out var rk1) && rk1.ValueKind == JsonValueKind.String) roleKey = rk1.GetString() ?? string.Empty;
                            else if (c.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? string.Empty;

                            string lineType = NormalizeCooperationLineType(null);
                            if (c.TryGetProperty("lineType", out var lt1) && lt1.ValueKind == JsonValueKind.String) lineType = NormalizeCooperationLineType(lt1.GetString());
                            else if (c.TryGetProperty("LineType", out var lt2) && lt2.ValueKind == JsonValueKind.String) lineType = NormalizeCooperationLineType(lt2.GetString());

                            // ★ cellA1 유지 — 탐색 순서: cellA1 → CellA1 → Cell.A1
                            string cellA1 = "";
                            if (c.TryGetProperty("cellA1", out var ca1) && ca1.ValueKind == JsonValueKind.String) cellA1 = ca1.GetString() ?? "";
                            else if (c.TryGetProperty("CellA1", out var ca2) && ca2.ValueKind == JsonValueKind.String) cellA1 = ca2.GetString() ?? "";
                            else if (c.TryGetProperty("Cell", out var cellEl) && cellEl.ValueKind == JsonValueKind.Object && cellEl.TryGetProperty("A1", out var a1El)) cellA1 = a1El.GetString() ?? "";

                            item["roleKey"] = string.IsNullOrWhiteSpace(roleKey) ? $"C{seq}" : roleKey.Trim();
                            item["approverType"] = at;
                            item["lineType"] = lineType;
                            item["value"] = approverValue;
                            item["cellA1"] = cellA1;
                            list.Add(item); seq++;
                        }
                        cooperationsJson = JsonSerializer.Serialize(list);
                    }
                }
                catch (Exception ex) { _log.LogWarning(ex, "CreateDX cooperation json normalize failed tc={tc}", tc); cooperationsJson = "[]"; }
            }
            else cooperationsJson = "[]";

            normalizedDesc = UpsertDescriptorApprovalAndCooperationArrays(normalizedDesc, approvalsJson, cooperationsJson);

            try
            {
                var pair = await SaveDocumentWithCollisionGuardAsync(
                    docId, tc, string.IsNullOrWhiteSpace(title) ? tc : title!, "PendingA1",
                    outputPathForDb, inputsMap, approvalsJson, cooperationsJson, normalizedDesc,
                    currentUserId,
                    User?.Identity?.Name ?? string.Empty, compCd, deptIdPart);
                docId = pair.docId;
                outputPathForDb = pair.outputPath;
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "CreateDX DB save failed docId={docId} tc={tc}", docId, tc);
                object? sqlDiag = null;
                if (ex is SqlException se) sqlDiag = new { se.Number, se.State, se.Procedure, se.LineNumber };
                return BadRequest(new { messages = new[] { "DOC_Err_SaveFailed" }, stage = "db", detail = ex.Message, sql = sqlDiag });
            }

            string finalExcelFullPath = string.Empty;
            try
            {
                finalExcelFullPath = _helper.ToContentRootAbsolute(outputPathForDb);
                var finalDir = Path.GetDirectoryName(finalExcelFullPath);
                if (!string.IsNullOrWhiteSpace(finalDir)) Directory.CreateDirectory(finalDir);

                if (!string.Equals(tempExcelFullPath, finalExcelFullPath, StringComparison.OrdinalIgnoreCase))
                {
                    if (!System.IO.File.Exists(finalExcelFullPath))
                    { System.IO.File.Move(tempExcelFullPath, finalExcelFullPath, overwrite: false); tempExcelFullPath = finalExcelFullPath; }
                    else _log.LogWarning("CreateDX final excel already exists. skip move. docId={docId}", docId);
                }

                if (!string.IsNullOrWhiteSpace(finalExcelFullPath) && System.IO.File.Exists(finalExcelFullPath))
                    filledPreviewJson = BuildPreviewJsonFromExcel(finalExcelFullPath);
            }
            catch (Exception ex) { _log.LogError(ex, "CreateDX move excel failed docId={docId}", docId); }

            try
            {
                await EnsureApprovalsAndSyncAsync(docId, tc);
                await FillDocumentApprovalsFromEmailsAsync(docId, toEmails);
                await EnsureCooperationsAndSyncAsync(docId, cooperationsJson);
            }
            catch (Exception ex) { _log.LogWarning(ex, "CreateDX EnsureSync failed docId={docId}", docId); }

            // UserId post-fix for approvals
            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection");
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();
                await using var fix = conn.CreateCommand();
                fix.CommandText = @"UPDATE a SET a.UserId = u.Id FROM dbo.DocumentApprovals a JOIN dbo.AspNetUsers u ON (a.ApproverValue = u.Email OR a.ApproverValue = u.UserName OR a.ApproverValue = u.Id) WHERE a.DocId = @DocId AND a.UserId IS NULL AND a.ApproverValue IS NOT NULL AND LTRIM(RTRIM(a.ApproverValue)) <> N'';";
                fix.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                await fix.ExecuteNonQueryAsync();
            }
            catch (Exception ex) { _log.LogWarning(ex, "CreateDX Post-fix UserId update failed docId={docId}", docId); }

            // UserId post-fix for cooperations
            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection");
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();
                await using var fix = conn.CreateCommand();
                fix.CommandText = @"UPDATE c SET c.UserId = u.Id FROM dbo.DocumentCooperations c JOIN dbo.AspNetUsers u ON (c.ApproverValue = u.Email OR c.ApproverValue = u.UserName OR c.ApproverValue = u.Id) WHERE c.DocId = @DocId AND c.UserId IS NULL AND c.ApproverValue IS NOT NULL AND LTRIM(RTRIM(c.ApproverValue)) <> N'';";
                fix.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                await fix.ExecuteNonQueryAsync();
            }
            catch (Exception ex) { _log.LogWarning(ex, "CreateDX Post-fix cooperation UserId failed docId={docId}", docId); }


            try
            {
                var actorId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
                var selected = (dto.SelectedRecipientUserIds ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(x => !string.Equals(x, actorId, StringComparison.OrdinalIgnoreCase)).ToList();

                if (selected.Count > 0)
                {
                    var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                    await using var conn = new SqlConnection(cs);
                    await conn.OpenAsync();
                    await using var tx = await conn.BeginTransactionAsync();
                    try
                    {
                        foreach (var uid in selected)
                        {
                            await using (var cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = (SqlTransaction)tx;
                                cmd.CommandText = @"IF EXISTS (SELECT 1 FROM dbo.DocumentShares WHERE DocId=@DocId AND UserId=@UserId) BEGIN UPDATE dbo.DocumentShares SET IsRevoked = 0, ExpireAt = NULL WHERE DocId=@DocId AND UserId=@UserId; END ELSE BEGIN INSERT INTO dbo.DocumentShares (DocId, UserId, AccessRole, ExpireAt, IsRevoked, CreatedBy, CreatedAt) VALUES (@DocId, @UserId, 'Commenter', NULL, 0, @CreatedBy, SYSUTCDATETIME()); END";
                                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = uid });
                                cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 64) { Value = actorId });
                                await cmd.ExecuteNonQueryAsync();
                            }
                            await using (var logCmd = conn.CreateCommand())
                            {
                                logCmd.Transaction = (SqlTransaction)tx;
                                logCmd.CommandText = @"INSERT INTO dbo.DocumentShareLogs (DocId,ActorId,ChangeCode,TargetUserId,BeforeJson,AfterJson,ChangedAt) VALUES (@DocId,@ActorId,@ChangeCode,@TargetUserId,NULL,@AfterJson,SYSUTCDATETIME());";
                                logCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                                logCmd.Parameters.Add(new SqlParameter("@ActorId", SqlDbType.NVarChar, 64) { Value = actorId });
                                logCmd.Parameters.Add(new SqlParameter("@ChangeCode", SqlDbType.NVarChar, 50) { Value = "ShareAdded" });
                                logCmd.Parameters.Add(new SqlParameter("@TargetUserId", SqlDbType.NVarChar, 64) { Value = uid });
                                logCmd.Parameters.Add(new SqlParameter("@AfterJson", SqlDbType.NVarChar, -1) { Value = "{\"accessRole\":\"Commenter\"}" });
                                await logCmd.ExecuteNonQueryAsync();
                            }
                        }
                        await ((SqlTransaction)tx).CommitAsync();
                    }
                    catch { await ((SqlTransaction)tx).RollbackAsync(); throw; }
                }
            }
            catch (Exception ex) { _log.LogWarning(ex, "CreateDX Share save failed docId={docId}", docId); }

            // WebPush
            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                // A1 결재자 알림 (기존)
                var a1UserId = await DocControllerHelper.GetA1ApproverUserIdByDocAsync(conn, docId);
                if (!string.IsNullOrWhiteSpace(a1UserId))
                    await DocControllerHelper.SendApprovalPendingBadgeAsync(
                        _webPushNotifier, _S, conn,
                        new List<string> { a1UserId.Trim() },
                        "/", "badge-approval-pending");
                else _log.LogWarning("CreateDX WebPush A1 userId not found docId={docId}", docId);

                var actorId2 = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;

                // 공유 수신자 알림 (기존)
                var shareIds = (dto.SelectedRecipientUserIds ?? new List<string>())
                    .Where(x => !string.IsNullOrWhiteSpace(x)).Select(x => x.Trim())
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(x => !string.Equals(x, actorId2, StringComparison.OrdinalIgnoreCase)).ToList();
                if (shareIds.Count > 0)
                    await DocControllerHelper.SendSharedUnreadBadgeAsync(
                        _webPushNotifier, _S, conn,
                        shareIds, "/", "badge-shared");

                // ★ 협조자 알림
                var coopUserIds = new List<string>();
                await using (var coopCmd = conn.CreateCommand())
                {
                    coopCmd.CommandText = @"
SELECT DISTINCT UserId
FROM dbo.DocumentCooperations
WHERE DocId = @DocId
  AND ISNULL(UserId, N'') <> N''
  AND ISNULL(Status, N'Pending') = N'Pending';";
                    coopCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    await using var rdr = await coopCmd.ExecuteReaderAsync();
                    while (await rdr.ReadAsync())
                    {
                        var uid = rdr[0] as string;
                        if (!string.IsNullOrWhiteSpace(uid)) coopUserIds.Add(uid.Trim());
                    }
                } // ← rdr 및 coopCmd 완전히 Dispose된 후

                // rdr 닫힌 후 conn 재사용 가능
                coopUserIds = coopUserIds
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .Where(x => !string.Equals(x, actorId2, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (coopUserIds.Count > 0)
                    await DocControllerHelper.SendCooperationPendingBadgeAsync(
                        _webPushNotifier, _S, conn,
                        coopUserIds,
                        docId,
                        "/Doc/DetailDX?id=" + Uri.EscapeDataString(docId),
                        "badge-cooperation-pending");
            }
            catch (Exception ex) { _log.LogWarning(ex, "CreateDX WebPush notify failed docId={docId}", docId); }

            var uploadUrl = $"/Doc/Upload?docId={Uri.EscapeDataString(docId)}";
            return Json(new { ok = true, docId, title = title ?? tc, status = "PendingA1", previewJson = filledPreviewJson, approvalsJson, attachments = Array.Empty<object>(), mailInfo = (object?)null, uploadUrl });
        }

        // =========================================================
        // Rebuild helpers — 각각 1개씩만 정의
        // =========================================================

        /// <summary>
        /// legacyJson의 Approvals 배열을 현재 포맷으로 재구성.
        /// </summary>
        private static List<object>? RebuildApprovalsFromLegacyDescriptor(string? legacyJson, List<object>? current)
        {
            if (!TryParseJsonFlexible(legacyJson, out var doc)) return current;
            try
            {
                var root = doc.RootElement;
                if (!root.TryGetProperty("Approvals", out var apprEl) || apprEl.ValueKind != JsonValueKind.Array)
                    return current;

                var list = new List<object>();
                int index = 1;
                foreach (var a in apprEl.EnumerateArray())
                {
                    var typ = a.TryGetProperty("ApproverType", out var at) ? (at.GetString() ?? "Person") : "Person";
                    var val = a.TryGetProperty("ApproverValue", out var av) ? av.GetString() : null;
                    var mappedType = (typ == "Person" || typ == "Role" || typ == "Rule") ? typ : "Person";

                    // ── A1 탐색 ──
                    string? cellA1 = null;
                    if (a.TryGetProperty("Cell", out var cellEl) && cellEl.ValueKind == JsonValueKind.Object)
                    {
                        if (cellEl.TryGetProperty("A1", out var a1El) && a1El.ValueKind == JsonValueKind.String) cellA1 = a1El.GetString();
                        if (string.IsNullOrWhiteSpace(cellA1) && cellEl.TryGetProperty("a1", out var a1El2) && a1El2.ValueKind == JsonValueKind.String) cellA1 = a1El2.GetString();
                    }
                    if (string.IsNullOrWhiteSpace(cellA1) && a.TryGetProperty("A1", out var dA1) && dA1.ValueKind == JsonValueKind.String) cellA1 = dA1.GetString();
                    if (string.IsNullOrWhiteSpace(cellA1) && a.TryGetProperty("cellA1", out var ca1) && ca1.ValueKind == JsonValueKind.String) cellA1 = ca1.GetString();
                    if (string.IsNullOrWhiteSpace(cellA1) && a.TryGetProperty("CellA1", out var ca2) && ca2.ValueKind == JsonValueKind.String) cellA1 = ca2.GetString();

                    list.Add(new Dictionary<string, object>
                    {
                        ["roleKey"] = $"A{index}",
                        ["approverType"] = mappedType,
                        ["required"] = false,
                        ["value"] = val ?? string.Empty,
                        ["cellA1"] = cellA1 ?? string.Empty
                    });
                    index++;
                }
                return list.Count == 0 ? current : list;
            }
            catch { return current; }
            finally { doc.Dispose(); }
        }

        /// <summary>
        /// legacyJson의 Cooperations/cooperations 배열을 현재 포맷으로 재구성.
        /// </summary>
        private static List<object>? RebuildCooperationsFromLegacyDescriptor(string? legacyJson, List<object>? current)
        {
            if (!TryParseJsonFlexible(legacyJson, out var doc)) return current;
            try
            {
                var root = doc.RootElement;
                JsonElement coopsEl;
                var found = root.TryGetProperty("Cooperations", out coopsEl) || root.TryGetProperty("cooperations", out coopsEl);
                if (!found || coopsEl.ValueKind != JsonValueKind.Array) return current;

                var list = new List<object>();
                int index = 1;
                foreach (var a in coopsEl.EnumerateArray())
                {
                    var typ = a.TryGetProperty("ApproverType", out var at) ? (at.GetString() ?? "Person") : "Person";
                    var val = a.TryGetProperty("ApproverValue", out var av) ? av.GetString() : null;
                    if (string.IsNullOrWhiteSpace(val) && a.TryGetProperty("value", out var v2)) val = v2.GetString();
                    var mappedType = (typ == "Person" || typ == "Role" || typ == "Rule") ? typ : "Person";

                    // ── A1 탐색: Cell.A1 → cell.A1 → A1 → a1 → cellA1 → CellA1 ──
                    string? cellA1 = null;
                    if (a.TryGetProperty("Cell", out var cellEl) && cellEl.ValueKind == JsonValueKind.Object)
                    {
                        if (cellEl.TryGetProperty("A1", out var a1El) && a1El.ValueKind == JsonValueKind.String) cellA1 = a1El.GetString();
                        if (string.IsNullOrWhiteSpace(cellA1) && cellEl.TryGetProperty("a1", out var a1El2) && a1El2.ValueKind == JsonValueKind.String) cellA1 = a1El2.GetString();
                    }
                    if (string.IsNullOrWhiteSpace(cellA1) && a.TryGetProperty("cell", out var cellEl2) && cellEl2.ValueKind == JsonValueKind.Object)
                    {
                        if (cellEl2.TryGetProperty("A1", out var a1El3) && a1El3.ValueKind == JsonValueKind.String) cellA1 = a1El3.GetString();
                        if (string.IsNullOrWhiteSpace(cellA1) && cellEl2.TryGetProperty("a1", out var a1El4) && a1El4.ValueKind == JsonValueKind.String) cellA1 = a1El4.GetString();
                    }
                    if (string.IsNullOrWhiteSpace(cellA1) && a.TryGetProperty("A1", out var dA1) && dA1.ValueKind == JsonValueKind.String) cellA1 = dA1.GetString();
                    if (string.IsNullOrWhiteSpace(cellA1) && a.TryGetProperty("a1", out var dA1L) && dA1L.ValueKind == JsonValueKind.String) cellA1 = dA1L.GetString();
                    if (string.IsNullOrWhiteSpace(cellA1) && a.TryGetProperty("cellA1", out var ca1) && ca1.ValueKind == JsonValueKind.String) cellA1 = ca1.GetString();
                    if (string.IsNullOrWhiteSpace(cellA1) && a.TryGetProperty("CellA1", out var ca2) && ca2.ValueKind == JsonValueKind.String) cellA1 = ca2.GetString();

                    list.Add(new Dictionary<string, object>
                    {
                        ["roleKey"] = $"C{index}",
                        ["approverType"] = mappedType,
                        ["lineType"] = "Cooperation",
                        ["value"] = val ?? string.Empty,
                        ["cellA1"] = cellA1 ?? string.Empty
                    });
                    index++;
                }
                return list.Count == 0 ? current : list;
            }
            catch { return current; }
            finally { doc.Dispose(); }
        }

        // =========================================================
        // BuildDescriptorJsonWithFlowGroups
        // =========================================================
        private string BuildDescriptorJsonWithFlowGroups(string templateCode, string rawDescriptorJson)
        {
            var normalized = ConvertDescriptorIfNeeded(rawDescriptorJson);
            var versionId = GetLatestVersionId(templateCode);
            if (versionId <= 0) return normalized;

            var groups = BuildFlowGroupsForTemplate(versionId);

            DescriptorDto desc;
            try { desc = JsonSerializer.Deserialize<DescriptorDto>(normalized, new JsonSerializerOptions { PropertyNameCaseInsensitive = true }) ?? new DescriptorDto(); }
            catch { desc = new DescriptorDto(); }

            desc.Approvals = RebuildApprovalsFromLegacyDescriptor(rawDescriptorJson, desc.Approvals);
            desc.Cooperations = RebuildCooperationsFromLegacyDescriptor(rawDescriptorJson, desc.Cooperations);

            if (desc.Cooperations == null || desc.Cooperations.Count == 0)
            {
                var dbOriginal = LoadRawDescriptorFromDb(templateCode);
                if (!string.IsNullOrWhiteSpace(dbOriginal))
                    desc.Cooperations = RebuildCooperationsFromLegacyDescriptor(dbOriginal, null);
            }

            desc.FlowGroups = groups;

            var json = JsonSerializer.Serialize(desc, new JsonSerializerOptions
            {
                Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
                WriteIndented = false
            });

            var dedup = DedupApprovalsJsonByStep(json);
            return UpsertDescriptorApprovalAndCooperationArrays(dedup, ExtractApprovalsJson(dedup), ExtractCooperationsJson(dedup));
        }

        private string? LoadRawDescriptorFromDb(string templateCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cs) || string.IsNullOrWhiteSpace(templateCode)) return null;
            try
            {
                using var conn = new SqlConnection(cs);
                conn.Open();
                using var cmd = new SqlCommand(@"SELECT TOP 1 v.DescriptorJson FROM dbo.DocTemplateVersion v JOIN dbo.DocTemplateMaster m ON m.Id = v.TemplateId WHERE m.DocCode = @code ORDER BY v.VersionNo DESC;", conn);
                cmd.Parameters.Add(new SqlParameter("@code", SqlDbType.NVarChar, 100) { Value = templateCode });
                return cmd.ExecuteScalar() as string;
            }
            catch { return null; }
        }

        private static string ConvertDescriptorIfNeeded(string? json)
        {
            const string EMPTY = "{\"inputs\":[],\"approvals\":[],\"cooperations\":[],\"version\":\"converted\"}";
            if (!TryParseJsonFlexible(json, out var doc)) return EMPTY;
            try
            {
                var root = doc.RootElement;
                if (root.TryGetProperty("inputs", out _))
                    return UpsertDescriptorApprovalAndCooperationArrays(json!, ExtractApprovalsJson(json!), ExtractCooperationsJson(json!));

                var inputs = new List<(string key, string type, string a1)>();
                if (root.TryGetProperty("Fields", out var fieldsEl) && fieldsEl.ValueKind == JsonValueKind.Array)
                {
                    foreach (var f in fieldsEl.EnumerateArray())
                    {
                        var key = f.TryGetProperty("Key", out var k) ? (k.GetString() ?? string.Empty).Trim() : string.Empty;
                        if (string.IsNullOrEmpty(key)) continue;
                        var type = f.TryGetProperty("Type", out var t) ? (t.GetString() ?? "Text") : "Text";
                        string a1 = string.Empty;
                        if (f.TryGetProperty("Cell", out var cell) && cell.ValueKind == JsonValueKind.Object)
                            a1 = cell.TryGetProperty("A1", out var a) ? (a.GetString() ?? string.Empty) : string.Empty;
                        inputs.Add((key, MapType(type), a1));
                    }
                }

                var approvals = new List<object>();
                if (root.TryGetProperty("Approvals", out var apprEl) && apprEl.ValueKind == JsonValueKind.Array)
                {
                    int index = 0;
                    foreach (var a in apprEl.EnumerateArray())
                    {
                        var typ = a.TryGetProperty("ApproverType", out var at) ? (at.GetString() ?? "Person") : "Person";
                        var val = a.TryGetProperty("ApproverValue", out var av) ? av.GetString() : null;
                        approvals.Add(new { roleKey = $"A{index + 1}", approverType = MapApproverType(typ), required = false, value = val ?? string.Empty });
                        index++;
                    }
                }

                var cooperations = new List<object>();
                if (root.TryGetProperty("Cooperations", out var coopEl) && coopEl.ValueKind == JsonValueKind.Array)
                {
                    int index = 0;
                    foreach (var c in coopEl.EnumerateArray())
                    {
                        var typ = c.TryGetProperty("ApproverType", out var at) ? (at.GetString() ?? "Person") : "Person";
                        var val = c.TryGetProperty("ApproverValue", out var av) ? av.GetString() : null;
                        cooperations.Add(new { roleKey = $"C{index + 1}", approverType = MapApproverType(typ), lineType = "Cooperation", value = val ?? string.Empty });
                        index++;
                    }
                }

                var obj = new
                {
                    inputs = inputs.Select(x => new { key = x.key, type = x.type, required = false, a1 = x.a1 }).ToList(),
                    approvals,
                    cooperations,
                    version = "converted"
                };
                return JsonSerializer.Serialize(obj);
            }
            catch { return EMPTY; }
            finally { doc.Dispose(); }

            static string MapType(string t)
            {
                t = (t ?? string.Empty).Trim().ToLowerInvariant();
                if (t.StartsWith("date")) return "Date";
                if (t.StartsWith("num") || t.Contains("number") || t.Contains("decimal") || t.Contains("integer")) return "Num";
                return "Text";
            }
            static string MapApproverType(string t)
            {
                t = (t ?? string.Empty).Trim();
                return (t == "Person" || t == "Role" || t == "Rule") ? t : "Person";
            }
        }

        private static string DedupApprovalsJsonByStep(string json)
        {
            try
            {
                var root = System.Text.Json.Nodes.JsonNode.Parse(json) as System.Text.Json.Nodes.JsonObject;
                if (root == null) return json;
                var approvalsNode = root["approvals"] ?? root["Approvals"];
                if (approvalsNode is not System.Text.Json.Nodes.JsonArray arr) return json;

                static int GetStep(System.Text.Json.Nodes.JsonNode? n)
                {
                    if (n is not System.Text.Json.Nodes.JsonObject o) return 0;
                    if (o.TryGetPropertyValue("Slot", out var slot) && slot != null)
                    {
                        if (slot is System.Text.Json.Nodes.JsonValue v1 && v1.TryGetValue<int>(out var si)) return si;
                        if (slot is System.Text.Json.Nodes.JsonValue v2 && v2.TryGetValue<string>(out var ss) && int.TryParse(ss, out var s2)) return s2;
                    }
                    if (o.TryGetPropertyValue("order", out var ord) && ord != null)
                    {
                        if (ord is System.Text.Json.Nodes.JsonValue v1 && v1.TryGetValue<int>(out var oi)) return oi;
                        if (ord is System.Text.Json.Nodes.JsonValue v2 && v2.TryGetValue<string>(out var os) && int.TryParse(os, out var o2)) return o2;
                    }
                    string? rk = null;
                    if (o.TryGetPropertyValue("roleKey", out var rkn) && rkn is System.Text.Json.Nodes.JsonValue rv && rv.TryGetValue<string>(out var rks)) rk = rks;
                    if (string.IsNullOrWhiteSpace(rk) && o.TryGetPropertyValue("RoleKey", out var rkn2) && rkn2 is System.Text.Json.Nodes.JsonValue rv2 && rv2.TryGetValue<string>(out var rks2)) rk = rks2;
                    if (!string.IsNullOrWhiteSpace(rk))
                    {
                        var m = Regex.Match(rk.Trim(), @"(?i)^A(\d+)$");
                        if (m.Success && int.TryParse(m.Groups[1].Value, out var rs)) return rs;
                    }
                    return 0;
                }

                static bool HasAnyValue(System.Text.Json.Nodes.JsonNode? n)
                {
                    if (n is not System.Text.Json.Nodes.JsonObject o) return false;
                    static string? GetStr(System.Text.Json.Nodes.JsonObject o2, string key)
                    {
                        if (o2.TryGetPropertyValue(key, out var v) && v is System.Text.Json.Nodes.JsonValue jv && jv.TryGetValue<string>(out var s)) return s;
                        return null;
                    }
                    if (!string.IsNullOrWhiteSpace(GetStr(o, "ApproverValue"))) return true;
                    if (!string.IsNullOrWhiteSpace(GetStr(o, "value"))) return true;
                    if (!string.IsNullOrWhiteSpace(GetStr(o, "email"))) return true;
                    if (o.TryGetPropertyValue("emails", out var e) && e is System.Text.Json.Nodes.JsonArray ea && ea.Count > 0) return true;
                    if (o.TryGetPropertyValue("users", out var us) && us is System.Text.Json.Nodes.JsonArray ua && ua.Count > 0) return true;
                    if (o.TryGetPropertyValue("user", out var uo) && uo is System.Text.Json.Nodes.JsonObject) return true;
                    return false;
                }

                var bestByStep = new Dictionary<int, System.Text.Json.Nodes.JsonNode?>();
                foreach (var item in arr)
                {
                    var step = GetStep(item);
                    if (step <= 0) continue;
                    if (!bestByStep.TryGetValue(step, out var existing)) bestByStep[step] = item;
                    else if (!HasAnyValue(existing) && HasAnyValue(item)) bestByStep[step] = item;
                }

                var rebuilt = new System.Text.Json.Nodes.JsonArray();
                foreach (var kv in bestByStep.OrderBy(k => k.Key)) rebuilt.Add(kv.Value?.DeepClone());
                if (root.ContainsKey("approvals")) root["approvals"] = rebuilt;
                else if (root.ContainsKey("Approvals")) root["Approvals"] = rebuilt;

                return root.ToJsonString(new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, WriteIndented = false });
            }
            catch { return json; }
        }

        private long GetLatestVersionId(string templateCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cs))
                return 0;

            try
            {
                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = new SqlCommand(@"
SELECT TOP 1
       CAST(v.Id AS BIGINT) AS VersionId
FROM dbo.DocTemplateMaster m
JOIN dbo.DocTemplateVersion v
     ON v.TemplateId = m.Id
WHERE m.DocCode = @code
ORDER BY v.VersionNo DESC, v.Id DESC;", conn);

                cmd.Parameters.Add(new SqlParameter("@code", SqlDbType.NVarChar, 100)
                {
                    Value = templateCode ?? string.Empty
                });

                var obj = cmd.ExecuteScalar();

                if (obj is long l)
                    return l;

                if (obj is int i)
                    return i;

                if (obj != null && obj != DBNull.Value && long.TryParse(obj.ToString(), out var p))
                    return p;
            }
            catch (Exception ex)
            {
                _log.LogWarning(
                    ex,
                    "GetLatestVersionId failed templateCode={TemplateCode}",
                    templateCode ?? string.Empty);
            }

            return 0;
        }

        private List<FlowGroupDto> BuildFlowGroupsForTemplate(long versionId)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            var rows = new List<(string Key, string? A1, int Row, int Col, string BaseKey, int? Index)>();
            using (var conn = new SqlConnection(cs))
            {
                conn.Open();
                using var cmd = new SqlCommand(@"SELECT [Key], A1, CellRow, CellColumn FROM dbo.DocTemplateField WHERE VersionId = @vid AND [Key] IS NOT NULL ORDER BY [Key], A1;", conn);
                cmd.Parameters.Add(new SqlParameter("@vid", SqlDbType.BigInt) { Value = versionId });
                using var rd = cmd.ExecuteReader();
                while (rd.Read())
                {
                    var key = rd["Key"] as string ?? string.Empty;
                    var a1 = rd["A1"] as string;
                    var r = rd["CellRow"] is DBNull ? 0 : Convert.ToInt32(rd["CellRow"]);
                    var c = rd["CellColumn"] is DBNull ? 0 : Convert.ToInt32(rd["CellColumn"]);
                    var (baseKey, idx) = ParseKey(key);
                    rows.Add((key, a1, r, c, baseKey, idx));
                }
            }
            return rows.GroupBy(x => x.BaseKey).Select(g =>
            {
                var ordered = g.OrderBy(x => x.Index.HasValue ? 0 : 1).ThenBy(x => x.Index).ThenBy(x => x.A1).ThenBy(x => x.Row).ThenBy(x => x.Col).Select(x => x.Key).ToList();
                return new FlowGroupDto { ID = g.Key, Keys = ordered };
            }).Where(g => g.Keys.Count > 1).ToList();
        }

        private static string ExtractApprovalsJson(string descriptorJson)
        {
            try
            {
                using var doc = JsonDocument.Parse(descriptorJson);
                if (doc.RootElement.TryGetProperty("approvals", out var arr)) return arr.GetRawText();
                if (doc.RootElement.TryGetProperty("Approvals", out var arr2)) return arr2.GetRawText();
            }
            catch { }
            return "[]";
        }

        private static string ExtractCooperationsJson(string descriptorJson)
        {
            try
            {
                using var doc = JsonDocument.Parse(descriptorJson);
                if (doc.RootElement.TryGetProperty("cooperations", out var arr)) return arr.GetRawText();
                if (doc.RootElement.TryGetProperty("Cooperations", out var arr2)) return arr2.GetRawText();
            }
            catch { }
            return "[]";
        }

        private static string NormalizeCooperationLineType(string? raw)
        {
            var s = (raw ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(s) ? "Cooperation" : s;
        }

        private static IEnumerable<string> EnumerateCooperationValues(JsonElement c)
        {
            if (c.ValueKind != JsonValueKind.Object) yield break;
            if (c.TryGetProperty("value", out var v1) && v1.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(v1.GetString())) yield return v1.GetString()!.Trim();
            if (c.TryGetProperty("approverValue", out var v2) && v2.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(v2.GetString())) yield return v2.GetString()!.Trim();
            if (c.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(v3.GetString())) yield return v3.GetString()!.Trim();
            if (c.TryGetProperty("email", out var e1) && e1.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(e1.GetString())) yield return e1.GetString()!.Trim();
            if (c.TryGetProperty("mail", out var e2) && e2.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(e2.GetString())) yield return e2.GetString()!.Trim();
            if (c.TryGetProperty("Email", out var e3) && e3.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(e3.GetString())) yield return e3.GetString()!.Trim();
            if (c.TryGetProperty("emails", out var emails) && emails.ValueKind == JsonValueKind.Array)
                foreach (var e in emails.EnumerateArray())
                    if (e.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(e.GetString())) yield return e.GetString()!.Trim();
            if (c.TryGetProperty("user", out var user) && user.ValueKind == JsonValueKind.Object)
            {
                if (user.TryGetProperty("email", out var ue1) && ue1.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(ue1.GetString())) yield return ue1.GetString()!.Trim();
                else if (user.TryGetProperty("mail", out var ue2) && ue2.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(ue2.GetString())) yield return ue2.GetString()!.Trim();
                else if (user.TryGetProperty("Email", out var ue3) && ue3.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(ue3.GetString())) yield return ue3.GetString()!.Trim();
                else if (user.TryGetProperty("id", out var uid) && uid.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(uid.GetString())) yield return uid.GetString()!.Trim();
            }
            if (c.TryGetProperty("users", out var users) && users.ValueKind == JsonValueKind.Array)
                foreach (var u in users.EnumerateArray())
                {
                    if (u.ValueKind != JsonValueKind.Object) continue;
                    if (u.TryGetProperty("email", out var ue1) && ue1.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(ue1.GetString())) yield return ue1.GetString()!.Trim();
                    else if (u.TryGetProperty("mail", out var ue2) && ue2.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(ue2.GetString())) yield return ue2.GetString()!.Trim();
                    else if (u.TryGetProperty("Email", out var ue3) && ue3.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(ue3.GetString())) yield return ue3.GetString()!.Trim();
                    else if (u.TryGetProperty("id", out var uid) && uid.ValueKind == JsonValueKind.String && !string.IsNullOrWhiteSpace(uid.GetString())) yield return uid.GetString()!.Trim();
                }
        }

        private static string UpsertDescriptorApprovalAndCooperationArrays(string? json, string? approvalsJson, string? cooperationsJson)
        {
            try
            {
                var root = System.Text.Json.Nodes.JsonNode.Parse(
                    string.IsNullOrWhiteSpace(json) ? "{\"inputs\":[],\"approvals\":[],\"cooperations\":[],\"version\":\"converted\"}" : json!)
                    as System.Text.Json.Nodes.JsonObject;

                if (root == null) return string.IsNullOrWhiteSpace(json) ? "{\"inputs\":[],\"approvals\":[],\"cooperations\":[],\"version\":\"converted\"}" : json!;

                System.Text.Json.Nodes.JsonNode? ParseArrayOrEmpty(string? raw)
                {
                    if (string.IsNullOrWhiteSpace(raw)) return System.Text.Json.Nodes.JsonNode.Parse("[]");
                    try { var n = System.Text.Json.Nodes.JsonNode.Parse(raw); return n is System.Text.Json.Nodes.JsonArray ? n : System.Text.Json.Nodes.JsonNode.Parse("[]"); }
                    catch { return System.Text.Json.Nodes.JsonNode.Parse("[]"); }
                }

                root["approvals"] = ParseArrayOrEmpty(approvalsJson);
                root["cooperations"] = ParseArrayOrEmpty(cooperationsJson);
                if (root.ContainsKey("Approvals")) root.Remove("Approvals");
                if (root.ContainsKey("Cooperations")) root.Remove("Cooperations");
                if (root["version"] == null) root["version"] = "converted";

                return root.ToJsonString(new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, WriteIndented = false });
            }
            catch { return string.IsNullOrWhiteSpace(json) ? "{\"inputs\":[],\"approvals\":[],\"cooperations\":[],\"version\":\"converted\"}" : json!; }
        }

        private (List<string> emails, List<string> diag) GetInitialRecipients(ComposePostDto dto, string normalizedDesc, IConfiguration cfg, string? docId = null, string? templateCode = null)
        {
            var emails = new List<string>();
            var diag = new List<string>();

            // 2026.06.18 Added: 메일 수신자 해석용 현재 작성자 UserId 추가 Contents 추후 메일 사용 시 기안자 예약값도 기존 수신자 해석 경로를 사용
            var currentUserId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;

            static bool LooksLikeEmail(string s) => !string.IsNullOrWhiteSpace(s) && s.Contains("@") && s.Contains(".");

            void AppendTokens(string raw, SqlConnection? conn)
            {
                if (string.IsNullOrWhiteSpace(raw)) return;
                var tokens = raw.Split(new[] { ';', ',', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Distinct(StringComparer.OrdinalIgnoreCase);
                foreach (var tk in tokens)
                {
                    var tokenValue = ResolveDrafterApproverValue(tk, currentUserId);

                    if (string.IsNullOrWhiteSpace(tokenValue))
                    {
                        diag.Add($"'{tk}' -> drafter-empty");
                        continue;
                    }

                    if (LooksLikeEmail(tokenValue)) { emails.Add(tokenValue); diag.Add($"'{tk}' -> email"); continue; }
                    if (conn == null) { diag.Add($"'{tk}' -> no-conn"); continue; }
                    string? email = null;
                    using (var cmd = conn.CreateCommand()) { cmd.CommandText = @"SELECT TOP 1 Email FROM dbo.AspNetUsers WHERE Id=@v OR UserName=@v OR NormalizedUserName=UPPER(@v) OR Email=@v OR NormalizedEmail=UPPER(@v);"; cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = tokenValue }); email = cmd.ExecuteScalar() as string; }
                    if (!string.IsNullOrWhiteSpace(email) && LooksLikeEmail(email)) { emails.Add(email!); diag.Add($"'{tk}' -> AspNetUsers '{email}'"); continue; }
                    using (var cmd = conn.CreateCommand()) { cmd.CommandText = @"SELECT TOP 1 COALESCE(u.Email, p.Email) FROM dbo.UserProfiles p LEFT JOIN dbo.AspNetUsers u ON u.Id = p.UserId WHERE p.DisplayName=@n OR p.Name=@n OR p.UserId=@n OR p.Email=@n;"; cmd.Parameters.Add(new SqlParameter("@n", SqlDbType.NVarChar, 256) { Value = tokenValue }); email = cmd.ExecuteScalar() as string; }
                    if (!string.IsNullOrWhiteSpace(email) && LooksLikeEmail(email)) { emails.Add(email!); diag.Add($"'{tk}' -> UserProfiles '{email}'"); }
                    else diag.Add($"'{tk}' not resolved");
                }
            }

            var cs = cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            using var conn = string.IsNullOrWhiteSpace(cs) ? null : new SqlConnection(cs);
            if (conn != null) conn.Open();

            if (dto?.Mail?.TO is { Count: > 0 }) foreach (var s in dto.Mail.TO) AppendTokens(s, conn);
            if (emails.Count == 0 && dto?.Approvals?.To is { Count: > 0 }) foreach (var s in dto.Approvals.To) AppendTokens(s, conn);
            if (emails.Count == 0 && dto?.Approvals?.Steps is { Count: > 0 }) foreach (var st in dto.Approvals.Steps) { if (!string.IsNullOrWhiteSpace(st?.Value)) AppendTokens(st!.Value!, conn); }

            if (emails.Count == 0)
            {
                try
                {
                    using var doc = JsonDocument.Parse(normalizedDesc);
                    if (doc.RootElement.TryGetProperty("approvals", out var arr) && arr.ValueKind == JsonValueKind.Array)
                        foreach (var el in arr.EnumerateArray())
                        {
                            var val = el.TryGetProperty("value", out var v) ? v.GetString() : el.TryGetProperty("ApproverValue", out var v2) ? v2.GetString() : null;
                            if (!string.IsNullOrWhiteSpace(val)) AppendTokens(val!, conn);
                        }
                }
                catch { }
            }

            if (emails.Count == 0 && !string.IsNullOrWhiteSpace(docId) && conn != null)
            {
                try
                {
                    using var q = conn.CreateCommand();
                    q.CommandText = @"SELECT ApproverValue FROM dbo.DocumentApprovals WHERE DocId=@id AND ApproverValue IS NOT NULL AND LTRIM(RTRIM(ApproverValue))<>'';";
                    q.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                    using var rd = q.ExecuteReader();
                    while (rd.Read()) AppendTokens(rd.GetString(0), conn);
                    diag.Add("fallback: DocumentApprovals used");
                }
                catch (Exception ex) { diag.Add("db fallback ex: " + ex.Message); }
            }

            if (emails.Count == 0 && !string.IsNullOrWhiteSpace(templateCode) && conn != null)
            {
                try
                {
                    using var cmd = conn.CreateCommand();
                    cmd.CommandText = @"SELECT TOP (1) v.DescriptorJson FROM dbo.DocTemplateVersion v JOIN dbo.DocTemplateMaster m ON m.Id=v.TemplateId WHERE m.DocCode=@code ORDER BY v.VersionNo DESC;";
                    cmd.Parameters.Add(new SqlParameter("@code", SqlDbType.NVarChar, 100) { Value = templateCode });
                    var raw = cmd.ExecuteScalar() as string;
                    if (!string.IsNullOrWhiteSpace(raw))
                    {
                        using var jd = JsonDocument.Parse(raw);
                        var root = jd.RootElement;
                        var path = root.TryGetProperty("approvals", out var _) ? "approvals" : (root.TryGetProperty("Approvals", out var __) ? "Approvals" : null);
                        if (path != null)
                        {
                            var arr = root.GetProperty(path);
                            foreach (var el in arr.EnumerateArray())
                            {
                                var val = el.TryGetProperty("value", out var v) ? v.GetString() : el.TryGetProperty("ApproverValue", out var v2) ? v2.GetString() : null;
                                if (!string.IsNullOrWhiteSpace(val)) AppendTokens(val!, conn);
                            }
                            diag.Add("fallback: Template.DescriptorJson approvals used");
                        }
                    }
                }
                catch (Exception ex) { diag.Add("template fallback ex: " + ex.Message); }
            }

            return (emails.Distinct(StringComparer.OrdinalIgnoreCase).ToList(), diag);
        }

        private async Task FillDocumentApprovalsFromEmailsAsync(string docId, IEnumerable<string>? approverEmails)
        {
            var list = approverEmails?.Where(e => !string.IsNullOrWhiteSpace(e)).Select(e => e.Trim()).ToList() ?? new List<string>();
            if (list.Count == 0) { _log.LogWarning("CreateDX FillDocumentApprovalsFromEmailsAsync docId={docId} email list empty.", docId); return; }

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            for (var i = 0; i < list.Count; i++)
            {
                var step = i + 1;
                var email = list[i];
                string? userId = null;
                await using (var userCmd = conn.CreateCommand())
                {
                    userCmd.CommandText = @"SELECT TOP (1) Id FROM dbo.AspNetUsers WHERE NormalizedEmail = UPPER(@E) OR NormalizedUserName = UPPER(@E) OR Id = @E;";
                    userCmd.Parameters.Add(new SqlParameter("@E", SqlDbType.NVarChar, 256) { Value = email });
                    userId = await userCmd.ExecuteScalarAsync() as string;
                }
                await using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"UPDATE a SET a.ApproverValue = @Email, a.UserId = COALESCE(a.UserId, @UserId) FROM dbo.DocumentApprovals AS a WHERE a.DocId = @DocId AND a.StepOrder = @Step AND (ISNULL(a.ApproverValue, N'') = N'' OR ISNULL(a.UserId, N'') = N'');";
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    cmd.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = step });
                    cmd.Parameters.Add(new SqlParameter("@Email", SqlDbType.NVarChar, 256) { Value = email });
                    cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = (object?)userId ?? DBNull.Value });
                    await cmd.ExecuteNonQueryAsync();
                }
            }
        }

        private async Task<(string docId, string outputPath)> SaveDocumentWithCollisionGuardAsync(string docId, string templateCode, string title, string status, string outputPath, Dictionary<string, string> inputs, string approvalsJson, string cooperationsJson, string descriptorJson, string userId, string userName, string compCd, string? departmentId)
        {
            static string WithNewSuffix(string baseDocId)
            {
                var core = string.IsNullOrWhiteSpace(baseDocId) ? "DOC_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") : baseDocId;
                var head = core.Length >= 4 ? core[..^4] : core;
                return head + RandomAlphaNumUpper(4);
            }
            static string EnsurePathFor(string path, string newDocId)
            {
                var dir = string.IsNullOrWhiteSpace(path) ? AppContext.BaseDirectory : (Path.GetDirectoryName(path) ?? AppContext.BaseDirectory);
                var ext = Path.GetExtension(path); if (string.IsNullOrWhiteSpace(ext)) ext = ".xlsx";
                return Path.Combine(dir, $"{newDocId}{ext}");
            }

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    await SaveDocumentAsync(docId, templateCode, title, status, outputPath, inputs, approvalsJson, cooperationsJson, descriptorJson, userId, userName, compCd, departmentId);
                    return (docId, outputPath);
                }
                catch (SqlException se) when (se.Number == 2627 || se.Number == 2601)
                {
                    _log.LogWarning(se, "CreateDX DocId unique violation attempt={attempt} docId={docId}", attempt + 1, docId);
                    var newDocId = WithNewSuffix(docId);
                    var newPath = EnsurePathFor(outputPath, newDocId);
                    if (System.IO.File.Exists(newPath)) { newDocId = WithNewSuffix(newDocId); newPath = EnsurePathFor(outputPath, newDocId); }
                    try { if (!string.IsNullOrWhiteSpace(outputPath) && System.IO.File.Exists(outputPath) && !string.Equals(outputPath, newPath, StringComparison.OrdinalIgnoreCase)) System.IO.File.Move(outputPath, newPath, overwrite: false); }
                    catch (Exception exMove) { _log.LogWarning(exMove, "CreateDX rename failed old={old} new={new}", outputPath, newPath); }
                    docId = newDocId; outputPath = newPath;
                }
            }
            throw new InvalidOperationException("DocId collision could not be resolved after retries");
        }

        private async Task<string> BuildSnapshotDisplayNameAsync(SqlConnection conn, SqlTransaction tx, string? userId, string? fallbackText, string? compCd)
        {
            static string ComposeDisplay(string? positionName, string? displayName, string fallback)
            {
                var pos = (positionName ?? string.Empty).Trim();
                var name = (displayName ?? string.Empty).Trim();
                var fb = (fallback ?? string.Empty).Trim();

                if (string.IsNullOrWhiteSpace(name))
                    name = fb;

                if (string.IsNullOrWhiteSpace(name))
                    return fb;

                if (string.IsNullOrWhiteSpace(pos))
                    return name;

                if (name.StartsWith(pos + " ", StringComparison.OrdinalIgnoreCase))
                    return name;

                return $"{pos} {name}".Trim();
            }

            var fallback = (fallbackText ?? string.Empty).Trim();

            var resolvedUserId = string.IsNullOrWhiteSpace(userId)
                ? await TryResolveUserIdAsync(conn, tx, fallback)
                : userId!.Trim();

            if (string.IsNullOrWhiteSpace(resolvedUserId))
                return fallback;

            await using var cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = @"
;WITH C AS
(
    SELECT
        u.Id,
        u.UserName,
        u.Email,
        up.CompCd,
        up.PositionId,
        COALESCE(NULLIF(LTRIM(RTRIM(up.DisplayName)), N''), NULLIF(LTRIM(RTRIM(u.UserName)), N''), NULLIF(LTRIM(RTRIM(u.Email)), N''), N'') AS DisplayName
    FROM dbo.AspNetUsers u
    LEFT JOIN dbo.UserProfiles up
        ON up.UserId = u.Id
    WHERE u.Id = @UserId
)
SELECT TOP 1
       COALESCE(NULLIF(LTRIM(RTRIM(pl.Name)), N''), NULLIF(LTRIM(RTRIM(pm.Name)), N''), N'') AS PositionName,
       COALESCE(NULLIF(LTRIM(RTRIM(C.DisplayName)), N''), NULLIF(LTRIM(RTRIM(@FallbackText)), N''), N'') AS DisplayName
FROM C
LEFT JOIN dbo.PositionMasters pm
       ON pm.Id = C.PositionId
      AND pm.CompCd = C.CompCd
LEFT JOIN dbo.PositionMasterLoc pl
       ON pl.PositionId = pm.Id
      AND pl.LangCode = N'ko'
ORDER BY
    CASE WHEN @CompCd <> '' AND C.CompCd = @CompCd THEN 0 ELSE 1 END,
    CASE WHEN COALESCE(NULLIF(LTRIM(RTRIM(pl.Name)), N''), NULLIF(LTRIM(RTRIM(pm.Name)), N''), N'') <> N'' THEN 0 ELSE 1 END,
    CASE WHEN COALESCE(NULLIF(LTRIM(RTRIM(C.DisplayName)), N''), N'') <> N'' THEN 0 ELSE 1 END;";

            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = resolvedUserId });
            cmd.Parameters.Add(new SqlParameter("@CompCd", SqlDbType.VarChar, 10) { Value = compCd ?? string.Empty });
            cmd.Parameters.Add(new SqlParameter("@FallbackText", SqlDbType.NVarChar, 256) { Value = fallback });

            await using var rd = await cmd.ExecuteReaderAsync();
            if (await rd.ReadAsync())
            {
                var positionName = rd["PositionName"] as string ?? string.Empty;
                var displayName = rd["DisplayName"] as string ?? string.Empty;
                var result = ComposeDisplay(positionName, displayName, fallback);
                if (!string.IsNullOrWhiteSpace(result))
                    return result;
            }

            return fallback;
        }

        private async Task SaveDocumentAsync(string docId, string templateCode, string title, string status, string outputPath, Dictionary<string, string> inputs, string approvalsJson, string cooperationsJson, string descriptorJson, string userId, string userName, string compCd, string? departmentId, int? templateVersionId = null)
        {
            if (string.IsNullOrWhiteSpace(compCd)) throw new InvalidOperationException("CompCd is required.");

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();
            await using var tx = await conn.BeginTransactionAsync();

            try
            {
                int? effectiveTemplateVersionId = templateVersionId;
                if (!effectiveTemplateVersionId.HasValue)
                {
                    await using var vCmd = conn.CreateCommand();
                    vCmd.Transaction = (SqlTransaction)tx;
                    vCmd.CommandText = @"SELECT TOP (1) v.Id FROM dbo.DocTemplateVersion AS v JOIN dbo.DocTemplateMaster AS m ON v.TemplateId = m.Id WHERE m.DocCode = @TemplateCode ORDER BY v.VersionNo DESC, v.Id DESC;";
                    vCmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 100) { Value = templateCode });
                    var obj = await vCmd.ExecuteScalarAsync();
                    if (obj != null && obj != DBNull.Value && int.TryParse(obj.ToString(), out var vid))
                        effectiveTemplateVersionId = vid;
                }

                descriptorJson = UpsertDescriptorApprovalAndCooperationArrays(descriptorJson, approvalsJson, cooperationsJson);

                // 저장 전체를 막지 않도록 snapshot 조회는 안전 fallback 처리
                var createdByDisplayText = (userName ?? string.Empty).Trim();
                try
                {
                    var snap = await BuildSnapshotDisplayNameAsync(
                        conn,
                        (SqlTransaction)tx,
                        userId,
                        userName,
                        compCd);

                    if (!string.IsNullOrWhiteSpace(snap))
                        createdByDisplayText = snap.Trim();
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "SaveDocument BuildSnapshotDisplayNameAsync failed for CreatedBy. docId={docId}, userId={userId}, compCd={compCd}", docId, userId, compCd);
                }

                await using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = (SqlTransaction)tx;
                    cmd.CommandText = @"
INSERT INTO dbo.Documents
(
    DocId,
    TemplateCode,
    TemplateVersionId,
    TemplateTitle,
    Status,
    OutputPath,
    CompCd,
    DepartmentId,
    CreatedBy,
    CreatedByName,
    CreatedAt,
    DescriptorJson
)
VALUES
(
    @DocId,
    @TemplateCode,
    @TemplateVersionId,
    @TemplateTitle,
    @Status,
    @OutputPath,
    @CompCd,
    @DepartmentId,
    @CreatedBy,
    @CreatedByName,
    SYSUTCDATETIME(),
    @DescriptorJson
);";
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    cmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 100) { Value = templateCode });
                    cmd.Parameters.Add(new SqlParameter("@TemplateVersionId", SqlDbType.Int) { Value = (object?)effectiveTemplateVersionId ?? DBNull.Value });
                    cmd.Parameters.Add(new SqlParameter("@TemplateTitle", SqlDbType.NVarChar, 400) { Value = (object?)title ?? DBNull.Value });
                    cmd.Parameters.Add(new SqlParameter("@Status", SqlDbType.NVarChar, 20) { Value = status });
                    cmd.Parameters.Add(new SqlParameter("@OutputPath", SqlDbType.NVarChar, 500) { Value = outputPath });
                    cmd.Parameters.Add(new SqlParameter("@CompCd", SqlDbType.VarChar, 10) { Value = compCd });
                    cmd.Parameters.Add(new SqlParameter("@DepartmentId", SqlDbType.VarChar, 12) { Value = string.IsNullOrWhiteSpace(departmentId) ? DBNull.Value : departmentId });
                    cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 450) { Value = userId ?? string.Empty });
                    cmd.Parameters.Add(new SqlParameter("@CreatedByName", SqlDbType.NVarChar, 200) { Value = string.IsNullOrWhiteSpace(createdByDisplayText) ? DBNull.Value : createdByDisplayText });
                    cmd.Parameters.Add(new SqlParameter("@DescriptorJson", SqlDbType.NVarChar, -1) { Value = (object?)descriptorJson ?? DBNull.Value });
                    await cmd.ExecuteNonQueryAsync();
                }

                if (inputs?.Count > 0)
                {
                    foreach (var kv in inputs)
                    {
                        if (string.IsNullOrWhiteSpace(kv.Key)) continue;

                        await using var icmd = conn.CreateCommand();
                        icmd.Transaction = (SqlTransaction)tx;
                        icmd.CommandText = @"INSERT INTO dbo.DocumentInputs (DocId, FieldKey, FieldValue, CreatedAt) VALUES (@DocId, @FieldKey, @FieldValue, SYSUTCDATETIME());";
                        icmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        icmd.Parameters.Add(new SqlParameter("@FieldKey", SqlDbType.NVarChar, 200) { Value = kv.Key });
                        icmd.Parameters.Add(new SqlParameter("@FieldValue", SqlDbType.NVarChar, -1) { Value = (object?)kv.Value ?? DBNull.Value });
                        await icmd.ExecuteNonQueryAsync();
                    }
                }

                try
                {
                    using var aj = JsonDocument.Parse(approvalsJson ?? "[]");
                    var toInsert = new List<(int step, string roleKey, string approverValue)>();
                    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var el in aj.RootElement.EnumerateArray())
                    {
                        if (el.ValueKind != JsonValueKind.Object) continue;

                        string roleKey = string.Empty;
                        if (el.TryGetProperty("roleKey", out var rk) && rk.ValueKind == JsonValueKind.String) roleKey = rk.GetString() ?? string.Empty;
                        else if (el.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? string.Empty;

                        int step = 0;
                        if (el.TryGetProperty("Slot", out var pSlot))
                        {
                            if (pSlot.ValueKind == JsonValueKind.Number) pSlot.TryGetInt32(out step);
                            else if (pSlot.ValueKind == JsonValueKind.String) int.TryParse(pSlot.GetString(), out step);
                        }
                        if (step <= 0 && el.TryGetProperty("order", out var pOrd))
                        {
                            if (pOrd.ValueKind == JsonValueKind.Number) pOrd.TryGetInt32(out step);
                            else if (pOrd.ValueKind == JsonValueKind.String) int.TryParse(pOrd.GetString(), out step);
                        }
                        if (step <= 0 && !string.IsNullOrWhiteSpace(roleKey))
                        {
                            var m = Regex.Match(roleKey.Trim(), @"(?i)^A(\d+)$");
                            if (m.Success) int.TryParse(m.Groups[1].Value, out step);
                        }
                        if (step <= 0) step = 1;

                        var rkNorm = string.IsNullOrWhiteSpace(roleKey) ? $"A{step}" : roleKey.Trim();

                        string? approverVal = null;
                        if (el.TryGetProperty("ApproverValue", out var pAV) && pAV.ValueKind == JsonValueKind.String) approverVal = pAV.GetString();
                        if (string.IsNullOrWhiteSpace(approverVal) && el.TryGetProperty("value", out var pV) && pV.ValueKind == JsonValueKind.String) approverVal = pV.GetString();
                        if (string.IsNullOrWhiteSpace(approverVal)) continue;

                        var key = $"{step}|{rkNorm}|{approverVal.Trim()}";
                        if (!seen.Add(key)) continue;

                        toInsert.Add((step, rkNorm, approverVal.Trim()));
                    }

                    foreach (var item in toInsert.OrderBy(x => x.step))
                    {
                        // 2026.06.18 Added: 문서별 결재자 기안자 예약값 치환 Contents 현재 작성자 UserId로 저장하여 WebPush 대상 조회가 가능하도록 함
                        var approverValue = ResolveDrafterApproverValue(item.approverValue, userId);
                        string? resolvedUserId = await TryResolveUserIdAsync(conn, (SqlTransaction)tx, approverValue);

                        string approverDisplayText = approverValue;
                        try
                        {
                            var snap = await BuildSnapshotDisplayNameAsync(
                                conn,
                                (SqlTransaction)tx,
                                resolvedUserId,
                                approverValue,
                                compCd);

                            if (!string.IsNullOrWhiteSpace(snap))
                                approverDisplayText = snap.Trim();
                        }
                        catch (Exception ex)
                        {
                            _log.LogWarning(ex, "SaveDocument approval snapshot failed docId={docId}, roleKey={roleKey}, approverValue={approverValue}", docId, item.roleKey, approverValue);
                        }

                        await using var acmd = conn.CreateCommand();
                        acmd.Transaction = (SqlTransaction)tx;
                        acmd.CommandText = @"
INSERT INTO dbo.DocumentApprovals
(
    DocId,
    StepOrder,
    RoleKey,
    ApproverValue,
    UserId,
    ActorName,
    ApproverDisplayText,
    Status,
    CreatedAt
)
VALUES
(
    @DocId,
    @StepOrder,
    @RoleKey,
    @ApproverValue,
    @UserId,
    @ActorName,
    @ApproverDisplayText,
    'Pending',
    SYSUTCDATETIME()
);";
                        acmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        acmd.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = item.step });
                        acmd.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 10) { Value = item.roleKey });
                        acmd.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 256) { Value = approverValue });
                        acmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = (object?)resolvedUserId ?? DBNull.Value });
                        acmd.Parameters.Add(new SqlParameter("@ActorName", SqlDbType.NVarChar, 200) { Value = string.IsNullOrWhiteSpace(approverDisplayText) ? DBNull.Value : approverDisplayText });
                        acmd.Parameters.Add(new SqlParameter("@ApproverDisplayText", SqlDbType.NVarChar, 200) { Value = string.IsNullOrWhiteSpace(approverDisplayText) ? DBNull.Value : approverDisplayText });
                        await acmd.ExecuteNonQueryAsync();
                    }
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "SaveDocument approvals insert failed docId={docId}", docId);
                }

                try
                {
                    using var cj = JsonDocument.Parse(cooperationsJson ?? "[]");
                    var toInsert = new List<(string lineType, string roleKey, string approverValue)>();
                    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    int seq = 1;

                    foreach (var el in cj.RootElement.EnumerateArray())
                    {
                        if (el.ValueKind != JsonValueKind.Object) continue;

                        string roleKey = string.Empty;
                        if (el.TryGetProperty("roleKey", out var rk) && rk.ValueKind == JsonValueKind.String) roleKey = rk.GetString() ?? string.Empty;
                        else if (el.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? string.Empty;

                        string lineType = "Cooperation";
                        if (el.TryGetProperty("lineType", out var lt1) && lt1.ValueKind == JsonValueKind.String) lineType = string.IsNullOrWhiteSpace(lt1.GetString()) ? "Cooperation" : lt1.GetString()!;
                        else if (el.TryGetProperty("LineType", out var lt2) && lt2.ValueKind == JsonValueKind.String) lineType = string.IsNullOrWhiteSpace(lt2.GetString()) ? "Cooperation" : lt2.GetString()!;

                        var roleKeyNorm = string.IsNullOrWhiteSpace(roleKey) ? $"C{seq}" : roleKey.Trim();
                        var approverVal = EnumerateCooperationValues(el).Select(v => (v ?? string.Empty).Trim()).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
                        if (string.IsNullOrWhiteSpace(approverVal))
                        {
                            seq++;
                            continue;
                        }

                        var key = $"{lineType}|{roleKeyNorm}|{approverVal}";
                        if (!seen.Add(key))
                        {
                            seq++;
                            continue;
                        }

                        toInsert.Add((lineType, roleKeyNorm, approverVal));
                        seq++;
                    }

                    foreach (var item in toInsert)
                    {
                        // 2026.06.18 Added: 문서별 협조자 기안자 예약값 치환 Contents 현재 작성자 UserId로 저장하여 WebPush 대상 조회가 가능하도록 함
                        var approverValue = ResolveDrafterApproverValue(item.approverValue, userId);
                        string? resolvedUserId = await TryResolveUserIdAsync(conn, (SqlTransaction)tx, approverValue);

                        string approverDisplayText = approverValue;
                        try
                        {
                            var snap = await BuildSnapshotDisplayNameAsync(
                                conn,
                                (SqlTransaction)tx,
                                resolvedUserId,
                                approverValue,
                                compCd);

                            if (!string.IsNullOrWhiteSpace(snap))
                                approverDisplayText = snap.Trim();
                        }
                        catch (Exception ex)
                        {
                            _log.LogWarning(ex, "SaveDocument cooperation snapshot failed docId={docId}, roleKey={roleKey}, approverValue={approverValue}", docId, item.roleKey, approverValue);
                        }

                        await using var ccmd = conn.CreateCommand();
                        ccmd.Transaction = (SqlTransaction)tx;
                        ccmd.CommandText = @"
INSERT INTO dbo.DocumentCooperations
(
    DocId,
    LineType,
    RoleKey,
    ApproverValue,
    UserId,
    ActorName,
    ApproverDisplayText,
    Status,
    CreatedAt
)
VALUES
(
    @DocId,
    @LineType,
    @RoleKey,
    @ApproverValue,
    @UserId,
    @ActorName,
    @ApproverDisplayText,
    'Pending',
    SYSUTCDATETIME()
);";
                        ccmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        ccmd.Parameters.Add(new SqlParameter("@LineType", SqlDbType.NVarChar, 20) { Value = item.lineType });
                        ccmd.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 100) { Value = item.roleKey });
                        ccmd.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 128) { Value = approverValue });
                        ccmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = (object?)resolvedUserId ?? DBNull.Value });
                        ccmd.Parameters.Add(new SqlParameter("@ActorName", SqlDbType.NVarChar, 200) { Value = string.IsNullOrWhiteSpace(approverDisplayText) ? DBNull.Value : approverDisplayText });
                        ccmd.Parameters.Add(new SqlParameter("@ApproverDisplayText", SqlDbType.NVarChar, 200) { Value = string.IsNullOrWhiteSpace(approverDisplayText) ? DBNull.Value : approverDisplayText });
                        await ccmd.ExecuteNonQueryAsync();
                    }
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "SaveDocument cooperations insert failed docId={docId}", docId);
                }

                await tx.CommitAsync();
            }
            catch
            {
                await tx.RollbackAsync();
                throw;
            }
        }

        private static async Task<string?> TryResolveUserIdAsync(SqlConnection conn, SqlTransaction? tx, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            await using var cmd = conn.CreateCommand();
            if (tx != null) cmd.Transaction = tx;
            cmd.CommandText = @"SELECT TOP 1 Id FROM dbo.AspNetUsers WHERE Id=@v OR UserName=@v OR NormalizedUserName=UPPER(@v) OR Email=@v OR NormalizedEmail=UPPER(@v);";
            cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            var id = (string?)await cmd.ExecuteScalarAsync();
            if (!string.IsNullOrWhiteSpace(id)) return id;
            await using var cmd2 = conn.CreateCommand();
            if (tx != null) cmd2.Transaction = tx;
            cmd2.CommandText = @"SELECT TOP 1 COALESCE(p.UserId,u.Id) FROM dbo.UserProfiles p LEFT JOIN dbo.AspNetUsers u ON u.Id=p.UserId WHERE p.UserId=@v OR p.Email=@v OR p.DisplayName=@v OR p.Name=@v;";
            cmd2.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            id = (string?)await cmd2.ExecuteScalarAsync();
            return string.IsNullOrWhiteSpace(id) ? null : id;
        }

        private async Task EnsureApprovalsAndSyncAsync(string docId, string templateCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();
            await using var tx = await conn.BeginTransactionAsync();

            try
            {
                string approvalsJson = "[]";
                string compCd = string.Empty;
                string drafterUserId = string.Empty;

                await using (var metaCmd = conn.CreateCommand())
                {
                    metaCmd.Transaction = (SqlTransaction)tx;
                    metaCmd.CommandText = @"
SELECT TOP 1 DescriptorJson, CompCd, CreatedBy
FROM dbo.Documents
WHERE DocId = @DocId;";
                    metaCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

                    await using var rd = await metaCmd.ExecuteReaderAsync();
                    if (await rd.ReadAsync())
                    {
                        approvalsJson = rd["DescriptorJson"] as string ?? "[]";
                        compCd = rd["CompCd"] as string ?? string.Empty;
                        drafterUserId = rd["CreatedBy"] as string ?? string.Empty;
                    }
                }

                approvalsJson = ExtractApprovalsJson(approvalsJson);

                var approvals = new List<(int StepOrder, string RoleKey, string ApproverValue)>();
                try
                {
                    using var aj = JsonDocument.Parse(string.IsNullOrWhiteSpace(approvalsJson) ? "[]" : approvalsJson);
                    if (aj.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        int seq = 1;
                        foreach (var a in aj.RootElement.EnumerateArray())
                        {
                            if (a.ValueKind != JsonValueKind.Object) continue;

                            string? approverValue = null;
                            if (a.TryGetProperty("ApproverValue", out var av1) && av1.ValueKind == JsonValueKind.String) approverValue = av1.GetString();
                            else if (a.TryGetProperty("approverValue", out var av2) && av2.ValueKind == JsonValueKind.String) approverValue = av2.GetString();
                            else if (a.TryGetProperty("value", out var av3) && av3.ValueKind == JsonValueKind.String) approverValue = av3.GetString();

                            approverValue = (approverValue ?? string.Empty).Trim();
                            if (string.IsNullOrWhiteSpace(approverValue)) continue;

                            string roleKey = string.Empty;
                            if (a.TryGetProperty("roleKey", out var rk1) && rk1.ValueKind == JsonValueKind.String) roleKey = rk1.GetString() ?? string.Empty;
                            else if (a.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? string.Empty;

                            int stepOrder = 0;
                            if (a.TryGetProperty("Slot", out var pSlot))
                            {
                                if (pSlot.ValueKind == JsonValueKind.Number) pSlot.TryGetInt32(out stepOrder);
                                else if (pSlot.ValueKind == JsonValueKind.String) int.TryParse(pSlot.GetString(), out stepOrder);
                            }
                            if (stepOrder <= 0 && a.TryGetProperty("order", out var pOrd))
                            {
                                if (pOrd.ValueKind == JsonValueKind.Number) pOrd.TryGetInt32(out stepOrder);
                                else if (pOrd.ValueKind == JsonValueKind.String) int.TryParse(pOrd.GetString(), out stepOrder);
                            }
                            if (stepOrder <= 0 && !string.IsNullOrWhiteSpace(roleKey))
                            {
                                var m = Regex.Match(roleKey.Trim(), @"(?i)^A(\d+)$");
                                if (m.Success) int.TryParse(m.Groups[1].Value, out stepOrder);
                            }
                            if (stepOrder <= 0) stepOrder = seq;

                            approvals.Add((
                                stepOrder,
                                string.IsNullOrWhiteSpace(roleKey) ? $"A{stepOrder}" : roleKey.Trim(),
                                approverValue));

                            seq++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "EnsureApprovalsAndSyncAsync parse failed docId={docId}", docId);
                }

                await using (var del = conn.CreateCommand())
                {
                    del.Transaction = (SqlTransaction)tx;
                    del.CommandText = @"DELETE FROM dbo.DocumentApprovals WHERE DocId = @DocId;";
                    del.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    await del.ExecuteNonQueryAsync();
                }

                foreach (var a in approvals
                    .GroupBy(x => new
                    {
                        StepOrder = x.StepOrder,
                        RoleKey = (x.RoleKey ?? string.Empty).Trim().ToUpperInvariant(),
                        ApproverValue = (x.ApproverValue ?? string.Empty).Trim().ToUpperInvariant()
                    })
                    .Select(g => g.First())
                    .OrderBy(x => x.StepOrder))
                {
                    // 2026.06.18 Added: 재동기화 시 결재자 기안자 예약값 치환 Contents 문서 작성자 UserId 기준으로 결재라인을 확정
                    var approverValue = ResolveDrafterApproverValue(a.ApproverValue, drafterUserId);
                    string? resolvedUserId = await TryResolveUserIdAsync(conn, (SqlTransaction)tx, approverValue);

                    string snapshotText = approverValue;
                    try
                    {
                        var snap = await BuildSnapshotDisplayNameAsync(
                            conn,
                            (SqlTransaction)tx,
                            resolvedUserId,
                            approverValue,
                            compCd);

                        if (!string.IsNullOrWhiteSpace(snap))
                            snapshotText = snap.Trim();
                    }
                    catch (Exception ex)
                    {
                        _log.LogWarning(ex, "EnsureApprovalsAndSyncAsync snapshot failed docId={docId} roleKey={roleKey} approverValue={approverValue}", docId, a.RoleKey, approverValue);
                    }

                    await using var ins = conn.CreateCommand();
                    ins.Transaction = (SqlTransaction)tx;
                    ins.CommandText = @"
INSERT INTO dbo.DocumentApprovals
(
    DocId,
    StepOrder,
    RoleKey,
    ApproverValue,
    UserId,
    ActorName,
    ApproverDisplayText,
    Status,
    CreatedAt
)
VALUES
(
    @DocId,
    @StepOrder,
    @RoleKey,
    @ApproverValue,
    @UserId,
    @ActorName,
    @ApproverDisplayText,
    N'Pending',
    SYSUTCDATETIME()
);";
                    ins.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    ins.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = a.StepOrder });
                    ins.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 100) { Value = a.RoleKey });
                    ins.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 128) { Value = approverValue });
                    ins.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = (object?)resolvedUserId ?? DBNull.Value });
                    ins.Parameters.Add(new SqlParameter("@ActorName", SqlDbType.NVarChar, 200) { Value = string.IsNullOrWhiteSpace(snapshotText) ? DBNull.Value : snapshotText });
                    ins.Parameters.Add(new SqlParameter("@ApproverDisplayText", SqlDbType.NVarChar, 200) { Value = string.IsNullOrWhiteSpace(snapshotText) ? DBNull.Value : snapshotText });
                    await ins.ExecuteNonQueryAsync();
                }

                await tx.CommitAsync();
            }
            catch (Exception ex)
            {
                try { await tx.RollbackAsync(); } catch { }
                _log.LogError(ex, "EnsureApprovalsAndSyncAsync failed docId={docId}", docId);
            }
        }

        private async Task EnsureCooperationsAndSyncAsync(string docId, string cooperationsJson)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();
            await using var tx = await conn.BeginTransactionAsync();

            try
            {
                string compCd = string.Empty;
                string drafterUserId = string.Empty;
                await using (var compCmd = conn.CreateCommand())
                {
                    compCmd.Transaction = (SqlTransaction)tx;
                    compCmd.CommandText = @"SELECT TOP 1 CompCd, CreatedBy FROM dbo.Documents WHERE DocId = @DocId;";
                    compCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    await using var compRd = await compCmd.ExecuteReaderAsync();
                    if (await compRd.ReadAsync())
                    {
                        compCd = (compRd["CompCd"] as string ?? string.Empty).Trim();
                        drafterUserId = (compRd["CreatedBy"] as string ?? string.Empty).Trim();
                    }
                }

                var cooperations = new List<(string LineType, string RoleKey, string ApproverValue)>();
                try
                {
                    using var cj = JsonDocument.Parse(string.IsNullOrWhiteSpace(cooperationsJson) ? "[]" : cooperationsJson);
                    if (cj.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        int seq = 1;
                        foreach (var c in cj.RootElement.EnumerateArray())
                        {
                            if (c.ValueKind != JsonValueKind.Object) continue;

                            var approverValue = EnumerateCooperationValues(c)
                                .Select(v => (v ?? string.Empty).Trim())
                                .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));

                            if (string.IsNullOrWhiteSpace(approverValue)) continue;

                            string roleKey = string.Empty;
                            if (c.TryGetProperty("roleKey", out var rk1) && rk1.ValueKind == JsonValueKind.String) roleKey = rk1.GetString() ?? string.Empty;
                            else if (c.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? string.Empty;

                            string lineType = NormalizeCooperationLineType(null);
                            if (c.TryGetProperty("lineType", out var lt1) && lt1.ValueKind == JsonValueKind.String) lineType = NormalizeCooperationLineType(lt1.GetString());
                            else if (c.TryGetProperty("LineType", out var lt2) && lt2.ValueKind == JsonValueKind.String) lineType = NormalizeCooperationLineType(lt2.GetString());

                            cooperations.Add((
                                lineType,
                                string.IsNullOrWhiteSpace(roleKey) ? $"C{seq}" : roleKey.Trim(),
                                approverValue));

                            seq++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "EnsureCooperationsAndSyncAsync parse failed docId={docId}", docId);
                }

                await using (var del = conn.CreateCommand())
                {
                    del.Transaction = (SqlTransaction)tx;
                    del.CommandText = @"DELETE FROM dbo.DocumentCooperations WHERE DocId = @DocId;";
                    del.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    await del.ExecuteNonQueryAsync();
                }

                foreach (var c in cooperations
                    .GroupBy(x => new
                    {
                        LineType = (x.LineType ?? string.Empty).Trim().ToUpperInvariant(),
                        RoleKey = (x.RoleKey ?? string.Empty).Trim().ToUpperInvariant(),
                        ApproverValue = (x.ApproverValue ?? string.Empty).Trim().ToUpperInvariant()
                    })
                    .Select(g => g.First()))
                {
                    // 2026.06.18 Added: 재동기화 시 협조자 기안자 예약값 치환 Contents 문서 작성자 UserId 기준으로 협조라인을 확정
                    var approverValue = ResolveDrafterApproverValue(c.ApproverValue, drafterUserId);
                    string? resolvedUserId = await TryResolveUserIdAsync(conn, (SqlTransaction)tx, approverValue);

                    string snapshotText = approverValue;
                    try
                    {
                        var snap = await BuildSnapshotDisplayNameAsync(
                            conn,
                            (SqlTransaction)tx,
                            resolvedUserId,
                            approverValue,
                            compCd);

                        if (!string.IsNullOrWhiteSpace(snap))
                            snapshotText = snap.Trim();
                    }
                    catch (Exception ex)
                    {
                        _log.LogWarning(ex, "EnsureCooperationsAndSyncAsync snapshot failed docId={docId} roleKey={roleKey} approverValue={approverValue}", docId, c.RoleKey, approverValue);
                    }

                    await using var ins = conn.CreateCommand();
                    ins.Transaction = (SqlTransaction)tx;
                    ins.CommandText = @"
INSERT INTO dbo.DocumentCooperations
(
    DocId,
    LineType,
    RoleKey,
    ApproverValue,
    UserId,
    ActorName,
    ApproverDisplayText,
    Status,
    CreatedAt
)
VALUES
(
    @DocId,
    @LineType,
    @RoleKey,
    @ApproverValue,
    @UserId,
    @ActorName,
    @ApproverDisplayText,
    N'Pending',
    SYSUTCDATETIME()
);";
                    ins.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    ins.Parameters.Add(new SqlParameter("@LineType", SqlDbType.NVarChar, 20) { Value = c.LineType });
                    ins.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 100) { Value = c.RoleKey });
                    ins.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 128) { Value = approverValue });
                    ins.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = (object?)resolvedUserId ?? DBNull.Value });
                    ins.Parameters.Add(new SqlParameter("@ActorName", SqlDbType.NVarChar, 200) { Value = string.IsNullOrWhiteSpace(snapshotText) ? DBNull.Value : snapshotText });
                    ins.Parameters.Add(new SqlParameter("@ApproverDisplayText", SqlDbType.NVarChar, 200) { Value = string.IsNullOrWhiteSpace(snapshotText) ? DBNull.Value : snapshotText });
                    await ins.ExecuteNonQueryAsync();
                }

                await tx.CommitAsync();
            }
            catch (Exception ex)
            {
                try { await tx.RollbackAsync(); } catch { }
                _log.LogError(ex, "EnsureCooperationsAndSyncAsync failed docId={docId}", docId);
            }
        }

        // 2026.06.12 Added: 최종 저장 문서에서는 작성 화면용 매핑셀 배경색을 제거하고 다시 Lock 처리한다.
        private static void ClearFinalOutputInputCellBackgrounds(string workbookPath, IEnumerable<string> mappedA1s)
        {
            if (string.IsNullOrWhiteSpace(workbookPath) || !System.IO.File.Exists(workbookPath))
                return;

            var cells = (mappedA1s ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (cells.Count == 0)
                return;

            TryRemoveWorksheetProtectionForCompose(workbookPath);

            using var wb = new XLWorkbook(workbookPath);
            var firstSheet = wb.Worksheets
                .FirstOrDefault(w => !string.Equals(w.Name, "EB_META", StringComparison.OrdinalIgnoreCase))
                ?? wb.Worksheets.FirstOrDefault();

            if (firstSheet == null)
                return;

            var touchedSheets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var rawA1 in cells)
            {
                try
                {
                    var (sheetNameFromA1, localA1) = SplitSheetAndA1ForCompose(rawA1);
                    if (string.IsNullOrWhiteSpace(localA1))
                        continue;

                    var ws = !string.IsNullOrWhiteSpace(sheetNameFromA1)
                        ? wb.Worksheets.FirstOrDefault(w => string.Equals(w.Name, sheetNameFromA1, StringComparison.OrdinalIgnoreCase)) ?? firstSheet
                        : firstSheet;

                    if (ws == null)
                        continue;

                    try { if (ws.IsProtected) ws.Unprotect(); } catch { }

                    var range = ws.Range(localA1);
                    var targetRange = range.FirstCell().IsMerged()
                        ? range.FirstCell().MergedRange()
                        : range;

                    targetRange.Style.Fill.PatternType = XLFillPatternValues.None;
                    targetRange.Style.Protection.Locked = true;
                    touchedSheets.Add(ws.Name);
                }
                catch
                {
                }
            }

            foreach (var sheetName in touchedSheets)
            {
                try
                {
                    var ws = wb.Worksheet(sheetName);
                    if (ws.IsProtected)
                    {
                        try { ws.Unprotect(); } catch { }
                    }

                    ws.Protect(allowedElements:
                        XLSheetProtectionElements.SelectLockedCells |
                        XLSheetProtectionElements.SelectUnlockedCells);
                }
                catch
                {
                }
            }

            wb.Save();
        }

        // 2026.04.06 Changed: 최종 저장 문서에도 시트 격자선 숨김을 적용하여 DetailDX에서 기본 엑셀 격자가 보이지 않게 함
        private async Task<string> GenerateExcelFromInputsAsync(long templateVersionId, string templateExcelFullPath, Dictionary<string, string> inputs, string? descriptorVersion)
        {
            if (string.IsNullOrWhiteSpace(templateExcelFullPath) || !System.IO.File.Exists(templateExcelFullPath))
                throw new FileNotFoundException("Template excel not found", templateExcelFullPath);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            var maps = new List<(string key, string a1, string type)>();

            await using (var conn = new SqlConnection(cs))
            {
                await conn.OpenAsync();
                await using var cmd = conn.CreateCommand();
                cmd.CommandText = @"SELECT [Key], [Type], A1, CellRow, CellColumn
                    FROM dbo.DocTemplateField
                    WHERE VersionId = @vid
                    ORDER BY Id;";
                cmd.Parameters.Add(new SqlParameter("@vid", SqlDbType.BigInt) { Value = templateVersionId });
                await using var rd = await cmd.ExecuteReaderAsync();
                while (await rd.ReadAsync())
                {
                    var key = rd["Key"] as string ?? string.Empty;
                    var typ = rd["Type"] as string ?? "Text";
                    string a1 = rd["A1"] as string ?? string.Empty;
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

            var outDir = Path.Combine(
                _env.ContentRootPath ?? AppContext.BaseDirectory, "App_Data", "Docs");
            Directory.CreateDirectory(outDir);
            var docId = GenerateDocumentId();
            var outPath = Path.Combine(outDir, $"{docId}.xlsx");
            var targetSheetName = string.Empty;

            System.IO.File.Copy(templateExcelFullPath, outPath, true);

            using (var spreadDoc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(outPath, true))
            {
                var wbPart = spreadDoc.WorkbookPart
                    ?? throw new InvalidOperationException("WorkbookPart not found");

                var sheet = wbPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>()
                    .FirstOrDefault()
                    ?? throw new InvalidOperationException("No sheet found");

                targetSheetName = sheet.Name?.Value ?? string.Empty;

                var wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)wbPart.GetPartById(sheet.Id!);
                var sheetData = wsPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()
                    ?? throw new InvalidOperationException("SheetData not found");

                var sstPart = wbPart.SharedStringTablePart
                    ?? wbPart.AddNewPart<DocumentFormat.OpenXml.Packaging.SharedStringTablePart>();
                sstPart.SharedStringTable ??= new DocumentFormat.OpenXml.Spreadsheet.SharedStringTable();

                var stylesPart = wbPart.WorkbookStylesPart
                    ?? wbPart.AddNewPart<DocumentFormat.OpenXml.Packaging.WorkbookStylesPart>();

                if (stylesPart.Stylesheet == null)
                {
                    stylesPart.Stylesheet = new DocumentFormat.OpenXml.Spreadsheet.Stylesheet(
                        new DocumentFormat.OpenXml.Spreadsheet.Fonts(
                            new DocumentFormat.OpenXml.Spreadsheet.Font()),
                        new DocumentFormat.OpenXml.Spreadsheet.Fills(
                            new DocumentFormat.OpenXml.Spreadsheet.Fill(
                                new DocumentFormat.OpenXml.Spreadsheet.PatternFill { PatternType = DocumentFormat.OpenXml.Spreadsheet.PatternValues.None }),
                            new DocumentFormat.OpenXml.Spreadsheet.Fill(
                                new DocumentFormat.OpenXml.Spreadsheet.PatternFill { PatternType = DocumentFormat.OpenXml.Spreadsheet.PatternValues.Gray125 })),
                        new DocumentFormat.OpenXml.Spreadsheet.Borders(
                            new DocumentFormat.OpenXml.Spreadsheet.Border()),
                        new DocumentFormat.OpenXml.Spreadsheet.CellStyleFormats(
                            new DocumentFormat.OpenXml.Spreadsheet.CellFormat()),
                        new DocumentFormat.OpenXml.Spreadsheet.CellFormats(
                            new DocumentFormat.OpenXml.Spreadsheet.CellFormat())
                    );
                }

                stylesPart.Stylesheet.NumberingFormats ??= new DocumentFormat.OpenXml.Spreadsheet.NumberingFormats();
                stylesPart.Stylesheet.CellFormats ??= new DocumentFormat.OpenXml.Spreadsheet.CellFormats(new DocumentFormat.OpenXml.Spreadsheet.CellFormat());

                static (int row, int col) ParseA1(string a1)
                {
                    var m = System.Text.RegularExpressions.Regex.Match(
                        a1.Trim().ToUpperInvariant(), @"^([A-Z]+)(\d+)$");
                    if (!m.Success) return (0, 0);
                    int col = 0;
                    foreach (char c in m.Groups[1].Value) col = col * 26 + (c - 'A' + 1);
                    return (int.Parse(m.Groups[2].Value), col);
                }

                static string ToA1(int row, int col)
                {
                    string letters = string.Empty;
                    int n = col;
                    while (n > 0)
                    {
                        int r = (n - 1) % 26;
                        letters = (char)('A' + r) + letters;
                        n = (n - 1) / 26;
                    }
                    return $"{letters}{row}";
                }

                int AddSharedString(string text)
                {
                    var items = sstPart.SharedStringTable.Elements<DocumentFormat.OpenXml.Spreadsheet.SharedStringItem>().ToList();
                    for (int i = 0; i < items.Count; i++)
                    {
                        if (items[i].InnerText == text) return i;
                    }

                    sstPart.SharedStringTable.AppendChild(
                        new DocumentFormat.OpenXml.Spreadsheet.SharedStringItem(
                            new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
                    sstPart.SharedStringTable.Save();
                    return items.Count;
                }

                DocumentFormat.OpenXml.Spreadsheet.Row GetOrCreateRow(int rowIndex)
                {
                    var row = sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>()
                        .FirstOrDefault(r => r.RowIndex?.Value == (uint)rowIndex);
                    if (row != null) return row;

                    row = new DocumentFormat.OpenXml.Spreadsheet.Row { RowIndex = (uint)rowIndex };
                    var refRow = sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>()
                        .FirstOrDefault(r => r.RowIndex?.Value > (uint)rowIndex);

                    if (refRow != null) sheetData.InsertBefore(row, refRow);
                    else sheetData.AppendChild(row);

                    return row;
                }

                DocumentFormat.OpenXml.Spreadsheet.Cell GetOrCreateCell(
                    DocumentFormat.OpenXml.Spreadsheet.Row row, string cellRef)
                {
                    var cell = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                        .FirstOrDefault(c => c.CellReference?.Value == cellRef);
                    if (cell != null) return cell;

                    cell = new DocumentFormat.OpenXml.Spreadsheet.Cell { CellReference = cellRef };
                    var (_, newCol) = ParseA1(cellRef);
                    var refCell = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>()
                        .FirstOrDefault(c =>
                        {
                            var (_, ec) = ParseA1(c.CellReference?.Value ?? "A1");
                            return ec > newCol;
                        });

                    if (refCell != null) row.InsertBefore(cell, refCell);
                    else row.AppendChild(cell);

                    return cell;
                }

                uint EnsureDateStyleIndex(uint baseStyleIndex, string formatCode)
                {
                    var stylesheet = stylesPart.Stylesheet!;
                    var numberingFormats = stylesheet.NumberingFormats!;
                    var cellFormats = stylesheet.CellFormats!;

                    uint numberFormatId = 0;

                    foreach (var nf in numberingFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>())
                    {
                        if (string.Equals(nf.FormatCode?.Value, formatCode, StringComparison.OrdinalIgnoreCase))
                        {
                            numberFormatId = nf.NumberFormatId?.Value ?? 0U;
                            break;
                        }
                    }

                    if (numberFormatId == 0)
                    {
                        var usedIds = numberingFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>()
                            .Select(x => x.NumberFormatId?.Value ?? 0U)
                            .Where(x => x >= 164U)
                            .ToList();

                        numberFormatId = usedIds.Count == 0 ? 164U : (usedIds.Max() + 1U);

                        numberingFormats.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat
                        {
                            NumberFormatId = numberFormatId,
                            FormatCode = formatCode
                        });
                        numberingFormats.Count = (uint)numberingFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.NumberingFormat>().Count();
                    }

                    var formats = cellFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().ToList();
                    var safeBaseIndex = (int)Math.Min(baseStyleIndex, (uint)Math.Max(0, formats.Count - 1));
                    var baseFormat = formats.Count > 0
                        ? (DocumentFormat.OpenXml.Spreadsheet.CellFormat)formats[safeBaseIndex].CloneNode(true)
                        : new DocumentFormat.OpenXml.Spreadsheet.CellFormat();

                    for (int i = 0; i < formats.Count; i++)
                    {
                        var cf = formats[i];
                        if ((cf.NumberFormatId?.Value ?? 0U) != numberFormatId) continue;
                        if ((cf.FontId?.Value ?? 0U) != (baseFormat.FontId?.Value ?? 0U)) continue;
                        if ((cf.FillId?.Value ?? 0U) != (baseFormat.FillId?.Value ?? 0U)) continue;
                        if ((cf.BorderId?.Value ?? 0U) != (baseFormat.BorderId?.Value ?? 0U)) continue;
                        if ((cf.FormatId?.Value ?? 0U) != (baseFormat.FormatId?.Value ?? 0U)) continue;
                        return (uint)i;
                    }

                    baseFormat.NumberFormatId = numberFormatId;
                    baseFormat.ApplyNumberFormat = true;

                    cellFormats.AppendChild(baseFormat);
                    cellFormats.Count = (uint)cellFormats.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().Count();
                    return cellFormats.Count!.Value - 1U;
                }

                void MarkWorkbookForFullRecalculation()
                {
                    wbPart.Workbook.CalculationProperties ??= new DocumentFormat.OpenXml.Spreadsheet.CalculationProperties();
                    wbPart.Workbook.CalculationProperties.CalculationMode = DocumentFormat.OpenXml.Spreadsheet.CalculateModeValues.Auto;
                    wbPart.Workbook.CalculationProperties.ForceFullCalculation = true;
                    wbPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;
                    wbPart.Workbook.CalculationProperties.CalculationCompleted = false;

                    if (wbPart.CalculationChainPart != null)
                    {
                        wbPart.DeletePart(wbPart.CalculationChainPart);
                    }
                }

                var dateFormatCode = GetExcelDateFormatByCulture(System.Globalization.CultureInfo.CurrentUICulture);
                if (string.IsNullOrWhiteSpace(dateFormatCode))
                {
                    dateFormatCode = "yyyy-mm-dd";
                }

                foreach (var m in maps)
                {
                    if (!inputs.TryGetValue(m.key, out var raw)) continue;

                    var (rowIdx, colIdx) = ParseA1(m.a1);
                    if (rowIdx <= 0 || colIdx <= 0) continue;

                    var cellRef = ToA1(rowIdx, colIdx);
                    var row = GetOrCreateRow(rowIdx);
                    var cell = GetOrCreateCell(row, cellRef);
                    var baseStyleIndex = cell.StyleIndex?.Value ?? 0U;

                    cell.CellFormula = null;

                    if (string.IsNullOrEmpty(raw))
                    {
                        cell.CellValue = null;
                        cell.DataType = null;
                        continue;
                    }

                    switch (m.type.ToLowerInvariant())
                    {
                        case "date":
                            if (DateTime.TryParse(raw, out var dt))
                            {
                                cell.DataType = null;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(
                                    dt.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture));
                                cell.StyleIndex = EnsureDateStyleIndex(baseStyleIndex, dateFormatCode);
                            }
                            else
                            {
                                int si = AddSharedString(raw);
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(si.ToString());
                            }
                            break;

                        case "num":
                            var normalized = raw.Replace(",", string.Empty);
                            if (decimal.TryParse(
                                normalized,
                                System.Globalization.NumberStyles.Any,
                                System.Globalization.CultureInfo.InvariantCulture,
                                out var dec))
                            {
                                cell.DataType = null;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(
                                    dec.ToString(System.Globalization.CultureInfo.InvariantCulture));
                            }
                            else
                            {
                                int si = AddSharedString(raw);
                                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
                                cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(si.ToString());
                            }
                            break;

                        default:
                            int ssIdx = AddSharedString(raw);
                            cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString;
                            cell.CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(ssIdx.ToString());
                            break;
                    }
                }

                MarkWorkbookForFullRecalculation();

                wsPart.Worksheet.Save();
                stylesPart.Stylesheet.Save();
                wbPart.Workbook.Save();
            }

            try
            {
                // 2026.06.12 Changed: 작성 화면 표시용 입력셀 배경색은 최종 저장 문서에서는 제거한다.
                // DetailDX는 조회/결재 화면이므로 매핑셀 안내색을 그대로 보여주지 않는다.
                ClearFinalOutputInputCellBackgrounds(outPath, maps.Select(x => x.a1));
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "GenerateExcelFromInputsAsync input background clear failed path={path}", outPath);
            }

            if (!string.IsNullOrWhiteSpace(targetSheetName))
            {
                try
                {
                    HideComposeDxSheetGridLines(outPath, targetSheetName);
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "GenerateExcelFromInputsAsync gridline hide failed path={path}", outPath);
                }
            }

            return outPath;

            static string A1FromRowCol(int row, int col)
            {
                if (row < 1 || col < 1) return string.Empty;
                string letters = string.Empty;
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
                var s = (t ?? string.Empty).Trim().ToLowerInvariant();
                if (s.StartsWith("date")) return "date";
                if (s.StartsWith("num") || s.Contains("number") ||
                    s.Contains("decimal") || s.Contains("integer")) return "num";
                return "text";
            }
        }

        private static string GetExcelDateFormatByCulture(CultureInfo culture)
        {
            var name = (culture?.Name ?? string.Empty).Trim().ToLowerInvariant();
            return name switch { "ko-kr" => "yyyy-mm-dd", "vi-vn" => "dd/mm/yyyy", "en-us" => "mm/dd/yyyy", "id-id" => "dd/mm/yyyy", "zh-cn" => "yyyy-mm-dd", _ => "yyyy-mm-dd" };
        }

        private static string GenerateDocumentId()
        {
            return $"DOC_{DateTime.Now:yyyyMMddHHmmssfff}{RandomAlphaNumUpper(4)}";
        }

        private static string RandomAlphaNumUpper(int len)
        {
            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            Span<byte> buffer = stackalloc byte[len];
            RandomNumberGenerator.Fill(buffer);
            char[] chars = new char[len];
            for (int i = 0; i < len; i++) chars[i] = alphabet[buffer[i] % alphabet.Length];
            return new string(chars);
        }

        private static bool HasCells(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return false;
            if (!json.Contains("\"cells\"", StringComparison.Ordinal)) return false;
            try { using var doc = JsonDocument.Parse(json); if (doc.RootElement.TryGetProperty("cells", out var cells) && cells.ValueKind == JsonValueKind.Array && cells.GetArrayLength() > 0) return cells[0].ValueKind == JsonValueKind.Array; }
            catch { }
            return false;
        }

        private static bool TryParseJsonFlexible(string? json, out JsonDocument doc)
        {
            doc = null!;
            if (string.IsNullOrWhiteSpace(json)) return false;
            try { doc = JsonDocument.Parse(json); return true; }
            catch { return false; }
        }

        private static (string BaseKey, int? Index) ParseKey(string key)
        {
            var s = (key ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(s)) return (string.Empty, null);
            var m = Regex.Match(s, @"^(.*?)(?:_(\d+))?$", RegexOptions.CultureInvariant);
            if (!m.Success) return (s, null);
            var baseKey = (m.Groups[1].Value ?? string.Empty).Trim();
            if (int.TryParse(m.Groups[2].Value, out var n)) return (baseKey, n);
            return (baseKey, null);
        }

        private string NormalizeTemplateExcelPath(string? raw)
        {
            var s = (raw ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(s)) return s;
            if (s.StartsWith("App_Data\\", StringComparison.OrdinalIgnoreCase) || s.StartsWith("App_Data/", StringComparison.OrdinalIgnoreCase)) return s.Replace('/', '\\');
            var idx = s.IndexOf("App_Data\\", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) idx = s.IndexOf("App_Data/", StringComparison.OrdinalIgnoreCase);
            return idx >= 0 ? s[idx..].Replace('/', '\\') : s;
        }

        [AllowAnonymous]
        [HttpGet("ComposeDXWarmup")]
        [ResponseCache(NoStore = true, Location = ResponseCacheLocation.None)]
        public async Task<IActionResult> ComposeDXWarmup()
        {
            if (!IsComposeDxWarmupLoopbackRequest())
                return NotFound();

            var sw = System.Diagnostics.Stopwatch.StartNew();
            var lastMs = 0.0;
            string? copiedPath = null;

            void Mark(string stage)
            {
                var now = sw.Elapsed.TotalMilliseconds;
                _log.LogInformation(
                    "COMPOSEDX-WARMUP {Stage} elapsedMs={ElapsedMs} deltaMs={DeltaMs}",
                    stage,
                    now,
                    now - lastMs);
                lastMs = now;
            }

            try
            {
                Mark("start");

                var sourcePath = EnsureComposeDxWarmupWorkbookPath();
                Mark("source-ready");

                var descriptorJson = BuildComposeDxWarmupDescriptorJson();
                var dxDocumentId = "compose_warmup_" + Guid.NewGuid().ToString("N");

                copiedPath = await CreateComposeDxWorkbookCopyAsync(
                    "COMPOSE_WARMUP",
                    sourcePath,
                    dxDocumentId,
                    descriptorJson);

                Mark("workbook-copy");

                TryDeleteComposeDxWarmupOutputFile(copiedPath);
                Mark("cleanup");

                sw.Stop();

                _log.LogInformation(
                    "COMPOSEDX-WARMUP completed elapsedMs={ElapsedMs} copiedPath={CopiedPath}",
                    sw.Elapsed.TotalMilliseconds,
                    copiedPath ?? string.Empty);

                return Json(new
                {
                    ok = true,
                    elapsedMs = sw.Elapsed.TotalMilliseconds
                });
            }
            catch (Exception ex)
            {
                sw.Stop();

                TryDeleteComposeDxWarmupOutputFile(copiedPath);

                _log.LogWarning(
                    ex,
                    "COMPOSEDX-WARMUP failed elapsedMs={ElapsedMs}",
                    sw.Elapsed.TotalMilliseconds);

                return StatusCode(500, new
                {
                    ok = false,
                    elapsedMs = sw.Elapsed.TotalMilliseconds
                });
            }
        }

        // 2026.06.15 Added: 서버 내부 warm-up 호출만 허용하여 외부에서 ComposeDX warm-up 액션을 직접 사용할 수 없게 한다.
        private bool IsComposeDxWarmupLoopbackRequest()
        {
            var remoteIp = HttpContext?.Connection?.RemoteIpAddress;
            if (remoteIp == null)
                return false;

            if (System.Net.IPAddress.IsLoopback(remoteIp))
                return true;

            if (remoteIp.IsIPv4MappedToIPv6 && System.Net.IPAddress.IsLoopback(remoteIp.MapToIPv4()))
                return true;

            return false;
        }

        // 2026.06.15 Added: 실제 템플릿을 변경하지 않고 workbook-copy 처리 경로만 예열하기 위한 최소 xlsx 파일을 생성한다.
        private string EnsureComposeDxWarmupWorkbookPath()
        {
            var warmupDir = Path.Combine(_env.ContentRootPath, "App_Data", "DxWarmup", "Compose");
            Directory.CreateDirectory(warmupDir);

            var warmupPath = Path.Combine(warmupDir, "compose_warmup_source.xlsx");
            var tempPath = warmupPath + ".tmp.xlsx";

            if (System.IO.File.Exists(tempPath))
                System.IO.File.Delete(tempPath);

            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Sheet1");

                ws.Cell("A1").Value = "Warmup Text";
                ws.Cell("B2").Value = 123;
                ws.Cell("C3").Value = DateTime.Today;

                ws.Range("D4:E4").Merge();
                ws.Cell("D4").Value = "Warmup Merge";

                ws.Range("A1:E8").Style.Protection.Locked = true;
                ws.Range("A1:E8").Style.Alignment.WrapText = false;

                try
                {
                    ws.PageSetup.PrintAreas.Add("A1:E8");
                }
                catch
                {
                }

                try
                {
                    ws.Protect(allowedElements:
                        XLSheetProtectionElements.SelectLockedCells |
                        XLSheetProtectionElements.SelectUnlockedCells);
                }
                catch
                {
                }

                wb.SaveAs(tempPath);
            }

            if (System.IO.File.Exists(warmupPath))
                System.IO.File.Delete(warmupPath);

            System.IO.File.Move(tempPath, warmupPath);

            return warmupPath;
        }

        // 2026.06.15 Added: Text Date Num 병합셀 경로를 함께 예열할 수 있도록 warm-up 전용 descriptor 를 생성한다.
        private static string BuildComposeDxWarmupDescriptorJson()
        {
            return """
            {
              "inputs": [
                { "key": "WarmupText", "type": "Text", "required": false, "sheet": "Sheet1", "a1": "A1" },
                { "key": "WarmupNumber", "type": "Num", "required": false, "sheet": "Sheet1", "a1": "B2" },
                { "key": "WarmupDate", "type": "Date", "required": false, "sheet": "Sheet1", "a1": "C3" },
                { "key": "WarmupMerge", "type": "Text", "required": false, "sheet": "Sheet1", "a1": "D4:E4" }
              ],
              "approvals": [],
              "cooperations": [],
              "version": "warmup"
            }
            """;
        }

        // 2026.06.15 Added: ComposeDX warm-up 으로 생성된 작성용 임시 복사본만 삭제한다.
        private void TryDeleteComposeDxWarmupOutputFile(string? path)
        {
            if (string.IsNullOrWhiteSpace(path))
                return;

            try
            {
                var composeRoot = Path.Combine(_env.ContentRootPath, "App_Data", "DocDxCompose");
                var fullPath = Path.GetFullPath(path);
                var fullRoot = Path.GetFullPath(composeRoot);

                if (!fullPath.StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase))
                    return;

                var fileName = Path.GetFileName(fullPath);
                if (!fileName.StartsWith("COMPOSE_WARMUP_", StringComparison.OrdinalIgnoreCase))
                    return;

                if (System.IO.File.Exists(fullPath))
                    System.IO.File.Delete(fullPath);

                var editableTempPath = fullPath + ".editable.tmp.xlsx";
                if (System.IO.File.Exists(editableTempPath))
                    System.IO.File.Delete(editableTempPath);
            }
            catch (Exception ex)
            {
                _log.LogDebug(ex, "COMPOSEDX-WARMUP cleanup skipped path={Path}", path);
            }
        }
        // =========================================================
        // DTOs
        // =========================================================
        public sealed class ComposePostDto
        {
            [JsonPropertyName("templateCode")] public string? TemplateCode { get; set; }
            [JsonPropertyName("inputs")] public Dictionary<string, string>? Inputs { get; set; }
            [JsonPropertyName("approvals")] public ApprovalsDto? Approvals { get; set; }
            [JsonPropertyName("cooperations")] public CooperationsDto? Cooperations { get; set; }
            [JsonPropertyName("mail")] public MailDto? Mail { get; set; }
            [JsonPropertyName("descriptorVersion")] public string? DescriptorVersion { get; set; }
            [JsonPropertyName("selectedRecipientUserIds")] public List<string>? SelectedRecipientUserIds { get; set; }
            [JsonPropertyName("attachments")] public List<ComposeAttachmentDto>? Attachments { get; set; }
        }
        public sealed class ComposeAttachmentDto
        {
            [JsonPropertyName("FileKey")] public string? FileKey { get; set; }
            [JsonPropertyName("OriginalName")] public string? OriginalName { get; set; }
            [JsonPropertyName("ContentType")] public string? ContentType { get; set; }
            [JsonPropertyName("ByteSize")] public long? ByteSize { get; set; }
        }
        public sealed class ApprovalsDto { public List<string>? To { get; set; } public List<ApprovalStepDto>? Steps { get; set; } }
        public sealed class ApprovalStepDto { public string? RoleKey { get; set; } public string? ApproverType { get; set; } public string? Value { get; set; } }
        public sealed class CooperationsDto { public List<string>? To { get; set; } public List<CooperationStepDto>? Steps { get; set; } }
        public sealed class CooperationStepDto { public string? RoleKey { get; set; } public string? ApproverType { get; set; } public string? Value { get; set; } public string? LineType { get; set; } }
        public sealed class MailDto { public List<string>? TO { get; set; } public List<string>? CC { get; set; } public List<string>? BCC { get; set; } public string? Subject { get; set; } public string? Body { get; set; } public bool Send { get; set; } = true; public string? From { get; set; } }
        public sealed class DxBatchEditRequest { [JsonPropertyName("spreadsheetState")] public SpreadsheetClientState? SpreadsheetState { get; set; } [JsonPropertyName("operations")] public List<DxBatchEditOperation>? Operations { get; set; } }
        public sealed class DxBatchEditOperation { [JsonPropertyName("a1")] public string? A1 { get; set; } [JsonPropertyName("value")] public string? Value { get; set; } [JsonPropertyName("clear")] public bool Clear { get; set; } [JsonPropertyName("wrap")] public bool? Wrap { get; set; } [JsonPropertyName("type")] public string? Type { get; set; } }
        private sealed class DescriptorDto
        {
            [JsonPropertyName("version")] public string? Version { get; set; }
            [JsonPropertyName("inputs")] public List<object>? Inputs { get; set; }
            [JsonPropertyName("approvals")] public List<object>? Approvals { get; set; }
            [JsonPropertyName("cooperations")] public List<object>? Cooperations { get; set; }
            [JsonPropertyName("flowGroups")] public List<FlowGroupDto>? FlowGroups { get; set; }
        }
        private sealed class InputFieldDto { public string Key { get; set; } = string.Empty; public string? A1 { get; set; } public string? Type { get; set; } }
        private sealed class FlowGroupDto { public string ID { get; set; } = string.Empty; public List<string> Keys { get; set; } = new(); }
        public sealed class DeleteDxTempRequest
        {
            [System.Text.Json.Serialization.JsonPropertyName("dxDocumentId")]
            public string? DxDocumentId { get; set; }
        }
    }
}
