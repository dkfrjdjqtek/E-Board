// 2026.02.27 Changed: DxSpreadsheetRequest 라우트가 중복 등록되어 AmbiguousMatchException 이 발생하므로 절대경로 중복 라우트를 제거하고 단일 엔드포인트로 통합

using ClosedXML.Excel;
using DevExpress.AspNetCore.Spreadsheet;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("DxSheet")]
    public class DxSheetController : Controller
    {
        // ✅ /DxSheet/Test
        [HttpGet("Test")]
        public IActionResult Test(string? openName = null)
        {
            // ✅ 이 플래그가 켜진 View에서는 _Layout 이 DX Spreadsheet 리소스를 head에 로드합니다.
            ViewData["UseDxSpreadsheet"] = true;

            // (유지) 기본 파일명
            openName = string.IsNullOrWhiteSpace(openName) ? "sample.xlsx" : Path.GetFileName(openName);

            var excelAbsPath = Path.Combine(Directory.GetCurrentDirectory(), "App_Data", openName);

            ViewBag.OpenName = openName;
            ViewBag.OpenRel = Path.Combine("App_Data", openName);

            if (!System.IO.File.Exists(excelAbsPath))
            {
                ViewBag.DebugExists = false;
                ViewBag.DebugErr = "Excel file not found";
                ViewBag.TargetHeightPx = 600;
                ViewBag.TargetWidthPx = 800;
                ViewBag.LastRow = 0;
                ViewBag.LastCol = 0;
                ViewBag.LastA1 = "A1";
                return View("~/Views/DxSheet/Test.cshtml");
            }

            try
            {
                var (lastRow1, lastCol1, heightPx, widthPx, mode) = ComputeVisualExtentPx(excelAbsPath);

                ViewBag.DebugExists = true;
                ViewBag.DebugErr = "";
                ViewBag.DebugUsedMode = mode;

                ViewBag.LastRow = Math.Max(0, lastRow1 - 1);  // r0
                ViewBag.LastCol = Math.Max(0, lastCol1 - 1);  // c0
                ViewBag.LastA1 = ToA1(lastRow1, lastCol1);

                ViewBag.TargetHeightPx = heightPx;
                ViewBag.TargetWidthPx = widthPx;
                ViewBag.TargetA1 = ViewBag.LastA1;
                ViewBag.TargetRow1 = lastRow1;
                ViewBag.TargetCol1 = lastCol1;
            }
            catch (Exception ex)
            {
                ViewBag.DebugExists = true;
                ViewBag.DebugErr = ex.Message;
                ViewBag.DebugUsedMode = "Error";

                ViewBag.TargetHeightPx = 600;
                ViewBag.TargetWidthPx = 800;
                ViewBag.LastRow = 0;
                ViewBag.LastCol = 0;
                ViewBag.LastA1 = "A1";
            }

            return View("~/Views/DxSheet/Test.cshtml");
        }

        // ✅ DX Spreadsheet 콜백 (중복 라우트 금지: 여기만 남겨야 합니다)
        // View에서 var dxCbUrl = "/DxSheet/DxSpreadsheetRequest"; 로 호출되는 엔드포인트
        [AcceptVerbs("GET", "POST")]
        [Route("DxSpreadsheetRequest")]
        public IActionResult DxSpreadsheetRequest()
        {
            return SpreadsheetRequestProcessor.GetResponse(HttpContext);
        }

        // (옵션) 기존 DxRequest를 쓰는 코드가 남아있다면 유지
        [AcceptVerbs("GET", "POST")]
        [Route("DxRequest")]
        public IActionResult DxRequest()
        {
            return SpreadsheetRequestProcessor.GetResponse(HttpContext);
        }

        private static (int lastRow1, int lastCol1, int heightPx, int widthPx, string mode) ComputeVisualExtentPx(string excelAbsPath)
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

                for (int c = 1; c <= candLastCol; c++)
                {
                    var cell = ws.Cell(r, c);
                    if (!cell.IsEmpty()) return true;
                    if (cell.HasFormula) return true;
                    if (HasAnyBorder(cell)) return true;
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

                for (int r = 1; r <= candLastRow; r++)
                {
                    var cell = ws.Cell(r, c);
                    if (!cell.IsEmpty()) return true;
                    if (cell.HasFormula) return true;
                    if (HasAnyBorder(cell)) return true;
                }
                return false;
            }

            var lastRow1 = candLastRow;
            for (int r = candLastRow; r >= 1; r--)
            {
                if (RowHasInk(r)) { lastRow1 = r; break; }
            }

            var lastCol1 = candLastCol;
            for (int c = candLastCol; c >= 1; c--)
            {
                if (ColHasInk(c)) { lastCol1 = c; break; }
            }

            int SumRowHeightsPx(int r1)
            {
                double sum = 0;
                for (int r = 1; r <= r1; r++)
                {
                    var row = ws.Row(r);
                    var pt = row.Height;
                    if (pt <= 0) pt = ws.RowHeight > 0 ? ws.RowHeight : 15.0;
                    sum += (pt * 96.0 / 72.0);
                }
                sum += 12;
                return (int)Math.Ceiling(sum);
            }

            int SumColWidthsPx(int c1)
            {
                double sum = 0;
                for (int c = 1; c <= c1; c++)
                {
                    var col = ws.Column(c);
                    var w = col.Width;
                    if (w <= 0) w = ws.ColumnWidth > 0 ? ws.ColumnWidth : 8.43;
                    sum += ExcelColumnWidthToPixels(w);
                }
                sum += 12;
                return (int)Math.Ceiling(sum);
            }

            var heightPx = SumRowHeightsPx(lastRow1);
            var widthPx = SumColWidthsPx(lastCol1);

            var mode = (ws.PageSetup.PrintAreas != null && ws.PageSetup.PrintAreas.Any())
                ? "PrintArea+VisualTrim"
                : "RangeUsed(All)+VisualTrim";

            return (lastRow1, lastCol1, heightPx, widthPx, mode);
        }

        private static bool HasAnyBorder(IXLCell cell)
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

        private static double ExcelColumnWidthToPixels(double excelWidth)
        {
            if (excelWidth <= 0) return 0;
            var px = Math.Truncate(((256.0 * excelWidth + Math.Truncate(128.0 / 7.0)) / 256.0) * 7.0);
            if (px < 1) px = 1;
            return px;
        }

        private static string ToA1(int row1, int col1)
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
    }
}