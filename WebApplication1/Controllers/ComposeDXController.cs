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
            var meta = await _tpl.LoadMetaAsync(templateCode);

            var descriptorJson = string.IsNullOrWhiteSpace(meta.descriptorJson)
                ? "{}"
                : meta.descriptorJson;

            var previewJson = "{}";
            var excelAbsPath = meta.excelFilePath ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(excelAbsPath) && System.IO.File.Exists(excelAbsPath))
            {
                try
                {
                    var rebuilt = BuildPreviewJsonFromExcel(excelAbsPath);
                    if (!string.IsNullOrWhiteSpace(rebuilt) && HasCells(rebuilt))
                        previewJson = rebuilt;
                }
                catch { }
            }

            if (previewJson == "{}")
            {
                var metaPreview = string.IsNullOrWhiteSpace(meta.previewJson) ? "{}" : meta.previewJson;
                if (!string.IsNullOrWhiteSpace(metaPreview) && HasCells(metaPreview))
                    previewJson = metaPreview;
            }

            try
            {
                var orgNodes = await _helper.BuildOrgTreeNodesAsync("ko");
                ViewBag.OrgTreeNodes = orgNodes;
            }
            catch
            {
                ViewBag.OrgTreeNodes = Array.Empty<OrgTreeNode>();
            }

            descriptorJson = BuildDescriptorJsonWithFlowGroups(templateCode, descriptorJson);

            var dxDocumentId = $"compose_{Guid.NewGuid():N}";
            var composeDxExcelPath = excelAbsPath;
            try
            {
                if (!string.IsNullOrWhiteSpace(excelAbsPath) && System.IO.File.Exists(excelAbsPath))
                {
                    composeDxExcelPath = await BuildComposeDxEditableWorkbookAsync(templateCode, excelAbsPath, descriptorJson, dxDocumentId);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "CreateDX editable workbook build failed. templateCode={tc}", templateCode);
                composeDxExcelPath = excelAbsPath;
            }

            var visualExtentPath = !string.IsNullOrWhiteSpace(composeDxExcelPath) && System.IO.File.Exists(composeDxExcelPath)
                ? composeDxExcelPath
                : excelAbsPath;

            try
            {
                if (!string.IsNullOrWhiteSpace(visualExtentPath) && System.IO.File.Exists(visualExtentPath))
                {
                    var (lastRow1, lastCol1, heightPx, widthPx, mode) = ComputeVisualExtentPxForCompose(visualExtentPath);

                    ViewBag.TargetHeightPx = heightPx;
                    ViewBag.TargetWidthPx = widthPx;
                    ViewBag.LastRow = Math.Max(0, lastRow1 - 1);
                    ViewBag.LastCol = Math.Max(0, lastCol1 - 1);
                    ViewBag.LastA1 = ToA1ForCompose(lastRow1, lastCol1);
                    ViewBag.TargetA1 = ViewBag.LastA1;
                    ViewBag.TargetRow1 = lastRow1;
                    ViewBag.TargetCol1 = lastCol1;
                    ViewBag.DxVisualMode = mode;
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "CreateDX visual extent compute failed. templateCode={tc}", templateCode);
                ViewBag.TargetHeightPx = 900;
                ViewBag.TargetWidthPx = 900;
                ViewBag.LastRow = 0;
                ViewBag.LastCol = 0;
                ViewBag.LastA1 = "A1";
                ViewBag.TargetA1 = "A1";
                ViewBag.TargetRow1 = 1;
                ViewBag.TargetCol1 = 1;
                ViewBag.DxVisualMode = "Error";
            }

            ViewBag.DescriptorJson = descriptorJson;
            ViewBag.PreviewJson = previewJson;
            ViewBag.TemplateTitle = meta.templateTitle ?? string.Empty;
            ViewBag.TemplateCode = templateCode;
            ViewBag.ExcelPath = composeDxExcelPath;
            ViewBag.DxCallbackUrl = "/Doc/dx-callback";
            ViewBag.DxDocumentId = dxDocumentId;
            ViewBag.HideCellPicker = true;

            return View("~/Views/Doc/ComposeDX.cshtml");

            static bool HasCells(string json)
            {
                if (string.IsNullOrWhiteSpace(json)) return false;
                if (!json.Contains("\"cells\"", StringComparison.Ordinal)) return false;
                try
                {
                    using var doc = JsonDocument.Parse(json);
                    if (doc.RootElement.TryGetProperty("cells", out var cells)
                        && cells.ValueKind == JsonValueKind.Array
                        && cells.GetArrayLength() > 0)
                    {
                        return cells[0].ValueKind == JsonValueKind.Array;
                    }
                }
                catch { }
                return false;
            }
        }

        // 2026.04.06 Changed: ComposeDX 가시 범위 계산 후 마지막 셀에 행 열 1칸 보정값을 추가하여 Q40 인식 시 R41까지 포함되도록 임시 보정함
        private static (int lastRow1, int lastCol1, int heightPx, int widthPx, string mode) ComputeVisualExtentPxForCompose(string excelAbsPath)
        {
            using var wb = new XLWorkbook(excelAbsPath);
            var ws = wb.Worksheets.First();

            IXLRange? baseRange = null;

            try
            {
                var pas = ws.PageSetup.PrintAreas;
                if (pas != null && pas.Any())
                {
                    baseRange = pas.First();
                }
            }
            catch { }

            if (baseRange == null)
            {
                baseRange = ws.RangeUsed(XLCellsUsedOptions.All);
            }

            if (baseRange == null)
            {
                return (1, 1, 600, 800, "Empty");
            }

            var firstRow1 = baseRange.RangeAddress.FirstAddress.RowNumber;
            var firstCol1 = baseRange.RangeAddress.FirstAddress.ColumnNumber;
            var candLastRow = baseRange.RangeAddress.LastAddress.RowNumber;
            var candLastCol = baseRange.RangeAddress.LastAddress.ColumnNumber;

            var merged = ws.MergedRanges?.ToList() ?? new List<IXLRange>();
            if (merged.Count > 0)
            {
                foreach (var mr in merged)
                {
                    var lr = mr.RangeAddress.LastAddress.RowNumber;
                    var lc = mr.RangeAddress.LastAddress.ColumnNumber;
                    if (lr > candLastRow) candLastRow = lr;
                    if (lc > candLastCol) candLastCol = lc;
                }
            }

            bool RowHasInk(int r)
            {
                for (int i = 0; i < merged.Count; i++)
                {
                    var a = merged[i].RangeAddress;
                    if (a.FirstAddress.RowNumber <= r && r <= a.LastAddress.RowNumber) return true;
                }

                for (int c = firstCol1; c <= candLastCol; c++)
                {
                    var cell = ws.Cell(r, c);
                    if (!cell.IsEmpty()) return true;
                    if (cell.HasFormula) return true;
                    if (HasAnyBorderForCompose(cell)) return true;
                }
                return false;
            }

            bool ColHasInk(int c)
            {
                for (int i = 0; i < merged.Count; i++)
                {
                    var a = merged[i].RangeAddress;
                    if (a.FirstAddress.ColumnNumber <= c && c <= a.LastAddress.ColumnNumber) return true;
                }

                for (int r = firstRow1; r <= candLastRow; r++)
                {
                    var cell = ws.Cell(r, c);
                    if (!cell.IsEmpty()) return true;
                    if (cell.HasFormula) return true;
                    if (HasAnyBorderForCompose(cell)) return true;
                }
                return false;
            }

            var lastRow1 = candLastRow;
            for (int r = candLastRow; r >= firstRow1; r--)
            {
                if (RowHasInk(r))
                {
                    lastRow1 = r;
                    break;
                }
            }

            var lastCol1 = candLastCol;
            for (int c = candLastCol; c >= firstCol1; c--)
            {
                if (ColHasInk(c))
                {
                    lastCol1 = c;
                    break;
                }
            }

            // 2026.04.06 Changed: 마지막 셀 인식 후 행 열 1칸씩 추가 보정
            lastRow1 += 1;
            lastCol1 += 1;

            var openXmlColWidths = new Dictionary<int, double>();
            var openXmlRowHeights = new Dictionary<int, double>();
            var defaultColWidth = ws.ColumnWidth > 0 ? ws.ColumnWidth : 8.43;
            var defaultRowHeightPt = ws.RowHeight > 0 ? ws.RowHeight : 15.0;
            var hasOpenXmlMetrics = TryReadOpenXmlSheetMetricsForCompose(
                excelAbsPath,
                ws.Name,
                openXmlColWidths,
                openXmlRowHeights,
                ref defaultColWidth,
                ref defaultRowHeightPt);

            int SumRowHeightsPx(int startRow1, int endRow1)
            {
                double sum = 0;
                for (int r = startRow1; r <= endRow1; r++)
                {
                    double pt;
                    if (hasOpenXmlMetrics && openXmlRowHeights.TryGetValue(r, out var oxPt) && oxPt > 0)
                    {
                        pt = oxPt;
                    }
                    else
                    {
                        var row = ws.Row(r);
                        pt = row.Height;
                        if (pt <= 0) pt = defaultRowHeightPt;
                    }

                    sum += (pt * 96.0 / 72.0);
                }

                sum += 12;
                return (int)Math.Ceiling(sum);
            }

            int SumColWidthsPx(int startCol1, int endCol1)
            {
                double sum = 0;
                for (int c = startCol1; c <= endCol1; c++)
                {
                    double w;
                    if (hasOpenXmlMetrics && openXmlColWidths.TryGetValue(c, out var oxW) && oxW > 0)
                    {
                        w = oxW;
                    }
                    else
                    {
                        var col = ws.Column(c);
                        w = col.Width;
                        if (w <= 0) w = defaultColWidth;
                    }

                    sum += ExcelColumnWidthToPixelsForCompose(w);
                }

                sum += 12;
                return (int)Math.Ceiling(sum);
            }

            var heightPx = SumRowHeightsPx(firstRow1, lastRow1);
            var widthPx = SumColWidthsPx(firstCol1, lastCol1);

            var mode = (ws.PageSetup.PrintAreas != null && ws.PageSetup.PrintAreas.Any())
                ? "PrintArea+VisualTrim"
                : "RangeUsed(All)+VisualTrim";

            if (hasOpenXmlMetrics)
            {
                mode += "+OpenXmlMetrics";
            }

            return (lastRow1, lastCol1, heightPx, widthPx, mode);
        }

        private static bool TryReadOpenXmlSheetMetricsForCompose(
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

                if (sheet.Id == null) return false;

                var wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)wbPart.GetPartById(sheet.Id);
                var worksheet = wsPart.Worksheet;
                if (worksheet == null) return false;

                var sfp = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetFormatProperties>();
                if (sfp?.DefaultColumnWidth?.Value != null && sfp.DefaultColumnWidth.Value > 0)
                {
                    defaultColWidth = sfp.DefaultColumnWidth.Value;
                }

                if (sfp?.DefaultRowHeight?.Value != null && sfp.DefaultRowHeight.Value > 0)
                {
                    defaultRowHeightPt = sfp.DefaultRowHeight.Value;
                }

                foreach (var cols in worksheet.Elements<DocumentFormat.OpenXml.Spreadsheet.Columns>())
                {
                    foreach (var col in cols.Elements<DocumentFormat.OpenXml.Spreadsheet.Column>())
                    {
                        var min = (int)(col.Min?.Value ?? 0U);
                        var max = (int)(col.Max?.Value ?? 0U);
                        var width = col.Width?.Value;

                        if (min <= 0 || max <= 0 || width == null || width.Value <= 0) continue;

                        for (int i = min; i <= max; i++)
                        {
                            openXmlColWidths[i] = width.Value;
                        }
                    }
                }

                var sheetData = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>();
                if (sheetData != null)
                {
                    foreach (var row in sheetData.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>())
                    {
                        var rowIndex = (int)(row.RowIndex?.Value ?? 0U);
                        if (rowIndex <= 0) continue;

                        if (row.Height?.Value != null && row.Height.Value > 0)
                        {
                            openXmlRowHeights[rowIndex] = row.Height.Value;
                        }
                    }
                }

                return openXmlColWidths.Count > 0 || openXmlRowHeights.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        private static bool HasAnyBorderForCompose(IXLCell cell)
        {
            try
            {
                var b = cell.Style.Border;
                return b.LeftBorder != XLBorderStyleValues.None
                    || b.RightBorder != XLBorderStyleValues.None
                    || b.TopBorder != XLBorderStyleValues.None
                    || b.BottomBorder != XLBorderStyleValues.None;
            }
            catch
            {
                return false;
            }
        }

        //private static double ExcelColumnWidthToPixelsForCompose(double excelWidth)
        //{
        //    if (excelWidth <= 0) return 0;
        //    var px = Math.Truncate(((256.0 * excelWidth + Math.Truncate(128.0 / 7.0)) / 256.0) * 7.0);
        //    if (px < 1) px = 1;
        //    return px;
        //}
        private static double ExcelColumnWidthToPixelsForCompose(double excelWidth)
        {
            if (excelWidth <= 0) return 0;

            var px = Math.Round(excelWidth * 7.0 + 5.0, MidpointRounding.AwayFromZero);
            if (px < 1) px = 1;

            return px;
        }

        private static string ToA1ForCompose(int row1, int col1)
        {
            string ColLetter(int col)
            {
                var s = "";
                var n = col;
                while (n > 0)
                {
                    int mod = (n - 1) % 26;
                    s = ((char)('A' + mod)) + s;
                    n = (n - 1) / 26;
                }
                return s;
            }

            return ColLetter(col1) + row1.ToString(CultureInfo.InvariantCulture);
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
                var pattern = $"*_{request.DxDocumentId}.xlsx";
                var files = Directory.GetFiles(dir, pattern);

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
            return SpreadsheetRequestProcessor.GetResponse(HttpContext);
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

        private sealed class ComposeDxEditableField
        {
            public string Key { get; set; } = string.Empty;
            public string Type { get; set; } = "Text";
            public string A1 { get; set; } = string.Empty;
        }

        // 2026.04.06 Changed: DX 작성용 임시 워크북에서 기본 격자선을 숨겨 문서 서식 외 빈 셀 격자가 노출되지 않도록 보정함
        // 2026.04.06 Changed: ClosedXML에서 지원되지 않는 SheetView.ShowGridLines 대신 OpenXml로 시트 표시 격자선 숨김 플래그를 저장 후 직접 설정함
        private async Task<string> BuildComposeDxEditableWorkbookAsync(string templateCode, string sourceExcelFullPath, string descriptorJson, string dxDocumentId)
        {
            var outRoot = Path.Combine(_env.ContentRootPath, "App_Data", "DocDxCompose");
            Directory.CreateDirectory(outRoot);
            var outPath = Path.Combine(outRoot,
                $"{templateCode}_{DateTime.Now:yyyyMMddHHmmssfff}_{dxDocumentId}.xlsx");
            System.IO.File.Copy(sourceExcelFullPath, outPath, true);
            var editableFields = ParseEditableFields(descriptorJson);
            var tempPath = outPath + "_tmp.xlsx";
            var targetSheetName = string.Empty;

            try
            {
                using (var wb = new XLWorkbook(outPath))
                {
                    var ws = wb.Worksheets.FirstOrDefault();
                    if (ws != null)
                    {
                        targetSheetName = ws.Name;

                        if (ws.IsProtected) ws.Unprotect();

                        ws.Cells().Style.Protection.Locked = true;

                        foreach (var f in editableFields)
                        {
                            if (string.IsNullOrWhiteSpace(f.A1)) continue;

                            try
                            {
                                var cell = ws.Cell(f.A1);
                                var targetCell = cell.IsMerged()
                                    ? cell.MergedRange().FirstCell()
                                    : cell;

                                targetCell.Style.Protection.Locked = false;
                                targetCell.Style.Fill.BackgroundColor = XLColor.FromHtml("#EAF6FF");
                                targetCell.Style.Alignment.WrapText = false;

                                if (string.Equals(f.Type, "Date", StringComparison.OrdinalIgnoreCase))
                                {
                                    try
                                    {
                                        var fmt = GetExcelDateFormatByCulture(CultureInfo.CurrentUICulture);
                                        targetCell.Value = DateTime.Today;
                                        targetCell.Style.DateFormat.Format = fmt;
                                    }
                                    catch { }
                                }
                            }
                            catch { }
                        }

                        ws.Protect(allowedElements:
                            XLSheetProtectionElements.SelectLockedCells |
                            XLSheetProtectionElements.SelectUnlockedCells);

                        wb.SaveAs(tempPath);
                    }
                }

                if (System.IO.File.Exists(tempPath))
                {
                    // 2026.04.06 Changed: 문서 서식 외 빈 셀에 표시되는 기본 엑셀 격자선을 OpenXml로 숨김
                    HideComposeDxSheetGridLines(tempPath, targetSheetName);

                    System.IO.File.Delete(outPath);
                    System.IO.File.Move(tempPath, outPath);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "BuildComposeDx failed tc={tc}", templateCode);
                if (System.IO.File.Exists(tempPath))
                {
                    try { System.IO.File.Delete(tempPath); } catch { }
                }
            }

            await Task.CompletedTask;
            return outPath;
        }

        // 2026.04.06 Added: OpenXml 시트 표시 옵션에 ShowGridLines false 를 기록하여 DX 렌더 시 빈 영역 격자 노출을 줄임
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

        private static List<ComposeDxEditableField> ParseEditableFields(string descriptorJson)
        {
            var list = new List<ComposeDxEditableField>();
            if (string.IsNullOrWhiteSpace(descriptorJson)) return list;
            try
            {
                using var doc = JsonDocument.Parse(descriptorJson);
                var root = doc.RootElement;
                if (!root.TryGetProperty("inputs", out var inputs) || inputs.ValueKind != JsonValueKind.Array)
                    return list;
                foreach (var item in inputs.EnumerateArray())
                {
                    var key = item.TryGetProperty("key", out var k) && k.ValueKind == JsonValueKind.String
                        ? (k.GetString() ?? string.Empty).Trim() : string.Empty;
                    var type = item.TryGetProperty("type", out var t) && t.ValueKind == JsonValueKind.String
                        ? (t.GetString() ?? "Text").Trim() : "Text";
                    var a1 = item.TryGetProperty("a1", out var a) && a.ValueKind == JsonValueKind.String
                        ? (a.GetString() ?? string.Empty).Trim().ToUpperInvariant() : string.Empty;
                    if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(a1)) continue;
                    list.Add(new ComposeDxEditableField { Key = key, Type = type, A1 = a1 });
                }
            }
            catch { return new List<ComposeDxEditableField>(); }

            return list
                .GroupBy(x => $"{x.Key}||{x.A1}", StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();
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
                    User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty,
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
                foreach (var kv in bestByStep.OrderBy(k => k.Key)) rebuilt.Add(kv.Value);
                if (root.ContainsKey("approvals")) root["approvals"] = rebuilt;
                else if (root.ContainsKey("Approvals")) root["Approvals"] = rebuilt;

                return root.ToJsonString(new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, WriteIndented = false });
            }
            catch { return json; }
        }

        private long GetLatestVersionId(string templateCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cs)) return 0;
            try
            {
                using var conn = new SqlConnection(cs);
                conn.Open();
                using var cmd = new SqlCommand(@"SELECT TOP 1 COALESCE(CAST(v.Id AS BIGINT), CAST(v.VersionId AS BIGINT)) AS VersionId FROM DocTemplateMaster m JOIN DocTemplateVersion v ON v.TemplateId = m.Id WHERE m.DocCode = @code ORDER BY v.VersionNo DESC;", conn);
                cmd.Parameters.Add(new SqlParameter("@code", SqlDbType.NVarChar, 100) { Value = templateCode ?? string.Empty });
                var obj = cmd.ExecuteScalar();
                if (obj is long l) return l;
                if (obj is int i) return i;
                if (obj != null && long.TryParse(obj.ToString(), out var p)) return p;
            }
            catch { }
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
            static bool LooksLikeEmail(string s) => !string.IsNullOrWhiteSpace(s) && s.Contains("@") && s.Contains(".");

            void AppendTokens(string raw, SqlConnection? conn)
            {
                if (string.IsNullOrWhiteSpace(raw)) return;
                var tokens = raw.Split(new[] { ';', ',', ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).Distinct(StringComparer.OrdinalIgnoreCase);
                foreach (var tk in tokens)
                {
                    if (LooksLikeEmail(tk)) { emails.Add(tk); diag.Add($"'{tk}' -> email"); continue; }
                    if (conn == null) { diag.Add($"'{tk}' -> no-conn"); continue; }
                    string? email = null;
                    using (var cmd = conn.CreateCommand()) { cmd.CommandText = @"SELECT TOP 1 Email FROM dbo.AspNetUsers WHERE Id=@v OR UserName=@v OR NormalizedUserName=UPPER(@v) OR Email=@v OR NormalizedEmail=UPPER(@v);"; cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = tk }); email = cmd.ExecuteScalar() as string; }
                    if (!string.IsNullOrWhiteSpace(email) && LooksLikeEmail(email)) { emails.Add(email!); diag.Add($"'{tk}' -> AspNetUsers '{email}'"); continue; }
                    using (var cmd = conn.CreateCommand()) { cmd.CommandText = @"SELECT TOP 1 COALESCE(u.Email, p.Email) FROM dbo.UserProfiles p LEFT JOIN dbo.AspNetUsers u ON u.Id = p.UserId WHERE p.DisplayName=@n OR p.Name=@n OR p.UserId=@n OR p.Email=@n;"; cmd.Parameters.Add(new SqlParameter("@n", SqlDbType.NVarChar, 256) { Value = tk }); email = cmd.ExecuteScalar() as string; }
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

        private async Task<(string docId, string outputPath)> SaveDocumentWithCollisionGuardAsync(
            string docId, string templateCode, string title, string status, string outputPath,
            Dictionary<string, string> inputs, string approvalsJson, string cooperationsJson,
            string descriptorJson, string userId, string userName, string compCd, string? departmentId)
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

        // 2026.04.09 Changed: 문서 저장 시 CreatedByName과 ApproverDisplayText를 직급 이름 스냅샷으로 저장하도록 수정함

        private async Task<string> BuildSnapshotDisplayNameAsync(
            SqlConnection conn,
            SqlTransaction tx,
            string? userId,
            string? fallbackText,
            string? compCd)
        {
            var fallback = (fallbackText ?? string.Empty).Trim();

            var effectiveUserId = string.IsNullOrWhiteSpace(userId)
                ? await TryResolveUserIdAsync(conn, tx, fallback)
                : userId?.Trim();

            if (string.IsNullOrWhiteSpace(effectiveUserId))
                return fallback;

            await using var cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = @"
SELECT TOP 1
       COALESCE(NULLIF(LTRIM(RTRIM(pl.Name)), N''), NULLIF(LTRIM(RTRIM(pm.Name)), N''), N'') AS PositionName,
       COALESCE(NULLIF(LTRIM(RTRIM(up.DisplayName)), N''), NULLIF(LTRIM(RTRIM(@FallbackText)), N''), u.UserName, u.Email, N'') AS DisplayName
FROM dbo.UserProfiles up
LEFT JOIN dbo.PositionMasters pm
       ON pm.CompCd = up.CompCd
      AND pm.Id = up.PositionId
LEFT JOIN dbo.PositionMasterLoc pl
       ON pl.PositionId = pm.Id
      AND pl.LangCode = N'ko'
LEFT JOIN dbo.AspNetUsers u
       ON u.Id = up.UserId
WHERE up.UserId = @UserId
  AND (@CompCd = '' OR up.CompCd = @CompCd);";
            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = effectiveUserId });
            cmd.Parameters.Add(new SqlParameter("@CompCd", SqlDbType.VarChar, 10) { Value = compCd ?? string.Empty });
            cmd.Parameters.Add(new SqlParameter("@FallbackText", SqlDbType.NVarChar, 200) { Value = fallback });

            await using var rd = await cmd.ExecuteReaderAsync();
            if (!await rd.ReadAsync())
                return fallback;

            var positionName = rd["PositionName"] as string ?? string.Empty;
            var displayName = rd["DisplayName"] as string ?? string.Empty;

            positionName = positionName.Trim();
            displayName = displayName.Trim();

            if (string.IsNullOrWhiteSpace(displayName))
                return fallback;

            if (string.IsNullOrWhiteSpace(positionName))
                return displayName;

            if (displayName.StartsWith(positionName + " ", StringComparison.Ordinal))
                return displayName;

            return $"{positionName} {displayName}".Trim();
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
                    if (obj != null && obj != DBNull.Value && int.TryParse(obj.ToString(), out var vid)) effectiveTemplateVersionId = vid;
                }

                descriptorJson = UpsertDescriptorApprovalAndCooperationArrays(descriptorJson, approvalsJson, cooperationsJson);

                var createdByDisplayText = await BuildSnapshotDisplayNameAsync(
                    conn,
                    (SqlTransaction)tx,
                    userId,
                    userName,
                    compCd);

                await using (var cmd = conn.CreateCommand())
                {
                    cmd.Transaction = (SqlTransaction)tx;
                    cmd.CommandText = @"INSERT INTO dbo.Documents (DocId, TemplateCode, TemplateVersionId, TemplateTitle, Status, OutputPath, CompCd, DepartmentId, CreatedBy, CreatedByName, CreatedAt, DescriptorJson) VALUES (@DocId, @TemplateCode, @TemplateVersionId, @TemplateTitle, @Status, @OutputPath, @CompCd, @DepartmentId, @CreatedBy, @CreatedByName, SYSUTCDATETIME(), @DescriptorJson);";
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    cmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 100) { Value = templateCode });
                    cmd.Parameters.Add(new SqlParameter("@TemplateVersionId", SqlDbType.Int) { Value = (object?)effectiveTemplateVersionId ?? DBNull.Value });
                    cmd.Parameters.Add(new SqlParameter("@TemplateTitle", SqlDbType.NVarChar, 400) { Value = (object?)title ?? DBNull.Value });
                    cmd.Parameters.Add(new SqlParameter("@Status", SqlDbType.NVarChar, 20) { Value = status });
                    cmd.Parameters.Add(new SqlParameter("@OutputPath", SqlDbType.NVarChar, 500) { Value = outputPath });
                    cmd.Parameters.Add(new SqlParameter("@CompCd", SqlDbType.VarChar, 10) { Value = compCd });
                    cmd.Parameters.Add(new SqlParameter("@DepartmentId", SqlDbType.VarChar, 12) { Value = string.IsNullOrWhiteSpace(departmentId) ? DBNull.Value : departmentId });
                    cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 450) { Value = userId ?? string.Empty });
                    cmd.Parameters.Add(new SqlParameter("@CreatedByName", SqlDbType.NVarChar, 200) { Value = (object?)createdByDisplayText ?? DBNull.Value });
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
                        string? resolvedUserId = await TryResolveUserIdAsync(conn, (SqlTransaction)tx, item.approverValue);
                        var approverDisplayText = await BuildSnapshotDisplayNameAsync(
                            conn,
                            (SqlTransaction)tx,
                            resolvedUserId,
                            item.approverValue,
                            compCd);

                        await using var acmd = conn.CreateCommand();
                        acmd.Transaction = (SqlTransaction)tx;
                        acmd.CommandText = @"INSERT INTO dbo.DocumentApprovals (DocId, StepOrder, RoleKey, ApproverValue, UserId, ApproverDisplayText, Status, CreatedAt) VALUES (@DocId, @StepOrder, @RoleKey, @ApproverValue, @UserId, @ApproverDisplayText, 'Pending', SYSUTCDATETIME());";
                        acmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        acmd.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = item.step });
                        acmd.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 10) { Value = item.roleKey });
                        acmd.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 256) { Value = item.approverValue });
                        acmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = (object?)resolvedUserId ?? DBNull.Value });
                        acmd.Parameters.Add(new SqlParameter("@ApproverDisplayText", SqlDbType.NVarChar, 200) { Value = (object?)approverDisplayText ?? DBNull.Value });
                        await acmd.ExecuteNonQueryAsync();
                    }
                }
                catch
                {
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
                        string? resolvedUserId = await TryResolveUserIdAsync(conn, (SqlTransaction)tx, item.approverValue);
                        var approverDisplayText = await BuildSnapshotDisplayNameAsync(
                            conn,
                            (SqlTransaction)tx,
                            resolvedUserId,
                            item.approverValue,
                            compCd);

                        await using var ccmd = conn.CreateCommand();
                        ccmd.Transaction = (SqlTransaction)tx;
                        ccmd.CommandText = @"INSERT INTO dbo.DocumentCooperations (DocId, LineType, RoleKey, ApproverValue, UserId, ApproverDisplayText, Status, CreatedAt) VALUES (@DocId, @LineType, @RoleKey, @ApproverValue, @UserId, @ApproverDisplayText, 'Pending', SYSUTCDATETIME());";
                        ccmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        ccmd.Parameters.Add(new SqlParameter("@LineType", SqlDbType.NVarChar, 20) { Value = item.lineType });
                        ccmd.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 100) { Value = item.roleKey });
                        ccmd.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 128) { Value = item.approverValue });
                        ccmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = (object?)resolvedUserId ?? DBNull.Value });
                        ccmd.Parameters.Add(new SqlParameter("@ApproverDisplayText", SqlDbType.NVarChar, 200) { Value = (object?)approverDisplayText ?? DBNull.Value });
                        await ccmd.ExecuteNonQueryAsync();
                    }
                }
                catch
                {
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

        private async Task EnsureApprovalsAndSyncAsync(string docId, string docCode)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();
            await using var tx = await conn.BeginTransactionAsync();
            try
            {
                string? descriptorJson = null;
                await using (var cmd = conn.CreateCommand()) { cmd.Transaction = (SqlTransaction)tx; cmd.CommandText = @"SELECT TOP (1) DescriptorJson FROM dbo.Documents WHERE DocId = @DocId;"; cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId }); var obj = await cmd.ExecuteScalarAsync(); if (obj != null && obj != DBNull.Value) descriptorJson = Convert.ToString(obj); }
                if (string.IsNullOrWhiteSpace(descriptorJson)) { await tx.CommitAsync(); return; }

                var approvals = new List<(int StepOrder, string RoleKey, string? ApproverValue)>();
                try
                {
                    using var dj = JsonDocument.Parse(descriptorJson);
                    var root = dj.RootElement;
                    if (root.TryGetProperty("approvals", out var arr) && arr.ValueKind == JsonValueKind.Array)
                    {
                        int i = 0;
                        foreach (var a in arr.EnumerateArray())
                        {
                            string? approverValue = null;
                            if (a.TryGetProperty("value", out var v1) && v1.ValueKind == JsonValueKind.String) approverValue = v1.GetString();
                            else if (a.TryGetProperty("approverValue", out var v2) && v2.ValueKind == JsonValueKind.String) approverValue = v2.GetString();
                            else if (a.TryGetProperty("ApproverValue", out var v3) && v3.ValueKind == JsonValueKind.String) approverValue = v3.GetString();
                            if (string.IsNullOrWhiteSpace(approverValue)) { i++; continue; }

                            int stepOrder = i + 1;
                            if (a.TryGetProperty("order", out var ordEl)) { if (ordEl.ValueKind == JsonValueKind.Number && ordEl.TryGetInt32(out var on)) stepOrder = on; else if (ordEl.ValueKind == JsonValueKind.String && int.TryParse(ordEl.GetString(), out var os)) stepOrder = os; }
                            else if (a.TryGetProperty("slot", out var slotEl)) { if (slotEl.ValueKind == JsonValueKind.Number && slotEl.TryGetInt32(out var sn)) stepOrder = sn; else if (slotEl.ValueKind == JsonValueKind.String && int.TryParse(slotEl.GetString(), out var ss)) stepOrder = ss; }

                            string roleKey = string.Empty;
                            if (a.TryGetProperty("roleKey", out var rk1) && rk1.ValueKind == JsonValueKind.String) roleKey = rk1.GetString() ?? string.Empty;
                            else if (a.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? string.Empty;
                            if (string.IsNullOrWhiteSpace(roleKey)) roleKey = $"A{stepOrder}";
                            approvals.Add((stepOrder, roleKey, approverValue!.Trim())); i++;
                        }
                    }
                }
                catch (Exception ex) { _log.LogWarning(ex, "EnsureApprovalsAndSyncAsync parse failed docId={docId}", docId); }

                if (approvals.Count == 0) { await tx.CommitAsync(); return; }

                await using (var del = conn.CreateCommand()) { del.Transaction = (SqlTransaction)tx; del.CommandText = @"DELETE FROM dbo.DocumentApprovals WHERE DocId = @DocId;"; del.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId }); await del.ExecuteNonQueryAsync(); }

                var ordered = approvals.OrderBy(a => a.StepOrder).ThenBy(a => a.RoleKey, StringComparer.OrdinalIgnoreCase).ToList();
                for (int idx = 0; idx < ordered.Count; idx++)
                {
                    var a = ordered[idx]; var stepOrder = idx + 1; var roleKey = $"A{stepOrder}";
                    await using var ins = conn.CreateCommand();
                    ins.Transaction = (SqlTransaction)tx;
                    ins.CommandText = @"INSERT INTO dbo.DocumentApprovals (DocId, StepOrder, RoleKey, ApproverValue, Status, CreatedAt) VALUES (@DocId, @StepOrder, @RoleKey, @ApproverValue, N'Pending', SYSUTCDATETIME());";
                    ins.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    ins.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = stepOrder });
                    ins.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 100) { Value = roleKey });
                    ins.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 128) { Value = a.ApproverValue! });
                    await ins.ExecuteNonQueryAsync();
                }
                await tx.CommitAsync();
            }
            catch (Exception ex) { try { await tx.RollbackAsync(); } catch { } _log.LogError(ex, "EnsureApprovalsAndSyncAsync failed docId={docId}", docId); }
        }

        private async Task EnsureCooperationsAndSyncAsync(string docId, string cooperationsJson)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();
            await using var tx = await conn.BeginTransactionAsync();
            try
            {
                var cooperations = new List<(string LineType, string RoleKey, string ApproverValue)>();
                try
                {
                    using var cj = JsonDocument.Parse(string.IsNullOrWhiteSpace(cooperationsJson) ? "[]" : cooperationsJson);
                    if (cj.RootElement.ValueKind == JsonValueKind.Array)
                    {
                        int i = 1;
                        foreach (var c in cj.RootElement.EnumerateArray())
                        {
                            if (c.ValueKind != JsonValueKind.Object) continue;
                            var value = EnumerateCooperationValues(c).Select(v => (v ?? string.Empty).Trim()).FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
                            if (string.IsNullOrWhiteSpace(value)) continue;
                            string roleKey = string.Empty;
                            if (c.TryGetProperty("roleKey", out var rk1) && rk1.ValueKind == JsonValueKind.String) roleKey = rk1.GetString() ?? string.Empty;
                            else if (c.TryGetProperty("RoleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? string.Empty;
                            string lineType = NormalizeCooperationLineType(null);
                            if (c.TryGetProperty("lineType", out var lt1) && lt1.ValueKind == JsonValueKind.String) lineType = NormalizeCooperationLineType(lt1.GetString());
                            else if (c.TryGetProperty("LineType", out var lt2) && lt2.ValueKind == JsonValueKind.String) lineType = NormalizeCooperationLineType(lt2.GetString());
                            cooperations.Add((lineType, string.IsNullOrWhiteSpace(roleKey) ? $"C{i}" : roleKey.Trim(), value)); i++;
                        }
                    }
                }
                catch (Exception ex) { _log.LogWarning(ex, "EnsureCooperationsAndSyncAsync parse failed docId={docId}", docId); }

                await using (var del = conn.CreateCommand()) { del.Transaction = (SqlTransaction)tx; del.CommandText = @"DELETE FROM dbo.DocumentCooperations WHERE DocId = @DocId;"; del.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId }); await del.ExecuteNonQueryAsync(); }

                foreach (var c in cooperations.GroupBy(x => new { L = (x.LineType ?? "").Trim().ToUpperInvariant(), R = (x.RoleKey ?? "").Trim().ToUpperInvariant(), A = (x.ApproverValue ?? "").Trim().ToUpperInvariant() }).Select(g => g.First()))
                {
                    await using var ins = conn.CreateCommand();
                    ins.Transaction = (SqlTransaction)tx;
                    ins.CommandText = @"INSERT INTO dbo.DocumentCooperations (DocId, LineType, RoleKey, ApproverValue, Status, CreatedAt) VALUES (@DocId, @LineType, @RoleKey, @ApproverValue, N'Pending', SYSUTCDATETIME());";
                    ins.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    ins.Parameters.Add(new SqlParameter("@LineType", SqlDbType.NVarChar, 20) { Value = c.LineType });
                    ins.Parameters.Add(new SqlParameter("@RoleKey", SqlDbType.NVarChar, 100) { Value = c.RoleKey });
                    ins.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 128) { Value = c.ApproverValue });
                    await ins.ExecuteNonQueryAsync();
                }
                await tx.CommitAsync();
            }
            catch (Exception ex) { try { await tx.RollbackAsync(); } catch { } _log.LogError(ex, "EnsureCooperationsAndSyncAsync failed docId={docId}", docId); }
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
