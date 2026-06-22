using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Data.SqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using WebApplication1.Controllers;

namespace WebApplication1.Services
{
    // 2026.06.15 Added: 기존 템플릿 버전 xlsx 파일을 현재 저장 규칙과 동일하게 재준비하고 표시 메트릭을 계산한다.
    public sealed class TemplateVersionPrepareService
    {
        private readonly IWebHostEnvironment _env;

        public TemplateVersionPrepareService(IWebHostEnvironment env)
        {
            _env = env;
        }

        public TemplateVersionPrepareResult PrepareExistingVersionFile(SqlConnection cn, long versionId, string excelAbsPath)
        {
            if (cn == null)
                throw new ArgumentNullException(nameof(cn));

            if (versionId <= 0)
                throw new InvalidOperationException("Template version id is invalid.");

            if (string.IsNullOrWhiteSpace(excelAbsPath) || !File.Exists(excelAbsPath))
                throw new FileNotFoundException("Template Excel file not found.", excelAbsPath);

            var fields = LoadTemplateFieldCellsForVersion(cn, versionId);
            var prepare = PrepareTemplateExcelForVersionByDbFields(excelAbsPath, fields);

            prepare.ExcelFileSize = new FileInfo(excelAbsPath).Length;
            prepare.TemplateFileHash = ComputeSha256Hex(excelAbsPath);

            try
            {
                prepare.PreviewJson = DocControllerHelper.BuildPreviewJsonFromExcel(excelAbsPath);
            }
            catch
            {
                prepare.PreviewJson = "{}";
            }

            return prepare;
        }

        public string CreateBackupFile(string absolutePath)
        {
            if (string.IsNullOrWhiteSpace(absolutePath) || !File.Exists(absolutePath))
                throw new FileNotFoundException("Template Excel file not found.", absolutePath);

            var contentRoot = Path.GetFullPath(_env.ContentRootPath);
            var fullPath = Path.GetFullPath(absolutePath);
            var relative = Path.GetRelativePath(contentRoot, fullPath);

            var backupRoot = Path.Combine(
                _env.ContentRootPath,
                "App_Data",
                "DocTemplates_Backup",
                "MetadataBackfill",
                DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture)
            );

            var backupPath = Path.Combine(backupRoot, relative);
            var backupDir = Path.GetDirectoryName(backupPath);

            if (!string.IsNullOrWhiteSpace(backupDir))
                Directory.CreateDirectory(backupDir);

            File.Copy(fullPath, backupPath, overwrite: false);
            return backupPath;
        }

        public static void RestoreBackupFile(string? backupPath, string? targetPath)
        {
            if (string.IsNullOrWhiteSpace(backupPath) || string.IsNullOrWhiteSpace(targetPath))
                return;

            try
            {
                if (File.Exists(backupPath))
                    File.Copy(backupPath, targetPath, overwrite: true);
            }
            catch
            {
            }
        }

        private static List<TemplateFieldCellInfo> LoadTemplateFieldCellsForVersion(SqlConnection cn, long versionId)
        {
            var list = new List<TemplateFieldCellInfo>();

            using var cmd = cn.CreateCommand();
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
                var type = NormalizeTemplateFieldType(ReadString(rd, "Type"));
                var sheet = ReadString(rd, "Sheet");
                var a1 = FirstNotEmpty(ReadString(rd, "A1"), ReadString(rd, "CellA1")) ?? string.Empty;

                if (string.IsNullOrWhiteSpace(a1))
                {
                    var row = FirstPositive(ReadInt(rd, "CellRow"), ReadInt(rd, "Row"));
                    var col = FirstPositive(ReadInt(rd, "CellColumn"), ReadInt(rd, "Column"));

                    if (row > 0 && col > 0)
                        a1 = ToA1FromRowColForTemplateVersion(row, col);
                }

                if (string.IsNullOrWhiteSpace(a1))
                    continue;

                var split = SplitSheetAndA1ForTemplateVersion(a1);

                if (string.IsNullOrWhiteSpace(sheet))
                    sheet = split.SheetName;

                if (string.IsNullOrWhiteSpace(split.LocalA1))
                    continue;

                list.Add(new TemplateFieldCellInfo
                {
                    Key = key ?? string.Empty,
                    Type = type,
                    Sheet = sheet ?? string.Empty,
                    A1 = split.LocalA1
                });
            }

            return list;
        }

        private static TemplateVersionPrepareResult PrepareTemplateExcelForVersionByDbFields(string excelAbsPath, IReadOnlyList<TemplateFieldCellInfo> fields)
        {
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
                        ws.Unprotect();
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

            return new TemplateVersionPrepareResult
            {
                ProtectionRuleCode = "ResetLockByMappingV1",
                VisualMetricRuleCode = "OpenXmlRangePxV1",
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
                if (used != null)
                    ranges.Add(used);
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
                        if (pa != null)
                            ranges.Add(pa);
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
                        cell.Style.Fill.BackgroundColor = XLColor.NoColor;
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

                    if (!intersects)
                        continue;

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
                if (wbPart?.Workbook == null)
                    return false;

                var sheets = wbPart.Workbook.Sheets?.Elements<DocumentFormat.OpenXml.Spreadsheet.Sheet>()?.ToList();
                if (sheets == null || sheets.Count == 0)
                    return false;

                var sheet = sheets.FirstOrDefault(x =>
                    string.Equals(x.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
                    ?? sheets.First();

                var relId = sheet.Id?.Value;
                if (string.IsNullOrWhiteSpace(relId))
                    return false;

                var wsPart = (DocumentFormat.OpenXml.Packaging.WorksheetPart)wbPart.GetPartById(relId);
                var worksheet = wsPart.Worksheet;
                if (worksheet == null)
                    return false;

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

                        if (hidden || min <= 0 || max <= 0 || width == null || width.Value <= 0)
                            continue;

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

                        if (rowIndex <= 0)
                            continue;

                        if (row.Hidden?.Value == true)
                            continue;

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
            if (excelWidth <= 0)
                return 0;

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

        private static string ToA1FromRowColForTemplateVersion(int row, int col)
        {
            if (row < 1 || col < 1)
                return string.Empty;

            return ToColumnLettersForTemplateVersion(col) + row.ToString(CultureInfo.InvariantCulture);
        }

        private static (string SheetName, string LocalA1) SplitSheetAndA1ForTemplateVersion(string? a1)
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

            var split = SplitSheetAndA1ForTemplateVersion(a1);
            var localA1 = (split.LocalA1 ?? string.Empty).Trim().ToUpperInvariant();

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

            if (lastRow < firstRow)
                (firstRow, lastRow) = (lastRow, firstRow);

            if (lastCol < firstCol)
                (firstCol, lastCol) = (lastCol, firstCol);

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

        private static int ColLettersToIndex(string letters)
        {
            if (string.IsNullOrWhiteSpace(letters))
                return 0;

            int n = 0;

            foreach (var ch in letters.Trim().ToUpperInvariant())
            {
                if (ch < 'A' || ch > 'Z')
                    break;

                n = n * 26 + (ch - 'A' + 1);
            }

            return n;
        }

        private static int FirstPositive(params int?[] values)
        {
            foreach (var value in values)
            {
                if (value.HasValue && value.Value > 0)
                    return value.Value;
            }

            return 0;
        }

        private static string NormalizeTemplateFieldType(string? value)
        {
            var t = (value ?? string.Empty).Trim().ToLowerInvariant();

            if (t.StartsWith("date", StringComparison.OrdinalIgnoreCase))
                return "Date";

            if (t.StartsWith("num", StringComparison.OrdinalIgnoreCase) ||
                t.Contains("number", StringComparison.OrdinalIgnoreCase) ||
                t.Contains("decimal", StringComparison.OrdinalIgnoreCase) ||
                t.Contains("integer", StringComparison.OrdinalIgnoreCase))
                return "Num";

            return "Text";
        }

        private static string ComputeSha256Hex(string filePath)
        {
            using var sha = SHA256.Create();
            using var stream = File.OpenRead(filePath);
            return Convert.ToHexString(sha.ComputeHash(stream));
        }

        private static string? FirstNotEmpty(params string?[] values)
        {
            return values.FirstOrDefault(x => !string.IsNullOrWhiteSpace(x));
        }

        private static string? ReadString(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            return rd.IsDBNull(ordinal) ? null : rd.GetString(ordinal);
        }

        private static int? ReadInt(SqlDataReader rd, string name)
        {
            var ordinal = rd.GetOrdinal(name);
            return rd.IsDBNull(ordinal) ? null : rd.GetInt32(ordinal);
        }

        private sealed class TemplateFieldCellInfo
        {
            public string Key { get; set; } = string.Empty;
            public string Type { get; set; } = "Text";
            public string Sheet { get; set; } = string.Empty;
            public string A1 { get; set; } = string.Empty;
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
    }

    public sealed class TemplateVersionPrepareResult
    {
        public string ProtectionRuleCode { get; set; } = "ResetLockByMappingV1";
        public string VisualMetricRuleCode { get; set; } = "OpenXmlRangePxV1";
        public string VisualSource { get; set; } = string.Empty;
        public string VisualRangeA1 { get; set; } = string.Empty;
        public int VisualWidthPx { get; set; }
        public int VisualHeightPx { get; set; }
        public string PreviewJson { get; set; } = "{}";
        public string TemplateFileHash { get; set; } = string.Empty;
        public long ExcelFileSize { get; set; }
    }
}