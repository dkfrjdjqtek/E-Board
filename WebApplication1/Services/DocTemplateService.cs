using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace WebApplication1.Services
{
    public class DocTemplateService : IDocTemplateService
    {
        private readonly IConfiguration _cfg;
        private readonly ILogger<DocTemplateService> _log;
        private readonly IWebHostEnvironment _env;

        public DocTemplateService(IConfiguration cfg, ILogger<DocTemplateService> log, IWebHostEnvironment env)
        {
            _cfg = cfg;
            _log = log;
            _env = env;
        }

        public async Task<(string descriptorJson, string previewJson, string templateTitle, long versionId, string? excelFilePath)>
            LoadMetaAsync(string templateCode)
        {
            if (string.IsNullOrWhiteSpace(templateCode)) return ("{}", "{}", "", 0L, null);
            var cs = _cfg.GetConnectionString("DefaultConnection");
            if (string.IsNullOrWhiteSpace(cs)) return ("{}", "{}", "", 0L, null);

            string templateTitle = "";
            long versionId = 0;
            string? excelPath = null;
            string descriptorJson = "{}";
            string previewJson = "{}";

            using (var conn = new SqlConnection(cs))
            {
                await conn.OpenAsync();

                // 1. 기본 정보
                using (var cmd1 = conn.CreateCommand())
                {
                    cmd1.CommandText = @"
SELECT TOP 1 m.DocName, v.Id AS VersionId, v.DescriptorJson, v.PreviewJson, v.ExcelFilePath
FROM DocTemplateMaster m
JOIN DocTemplateVersion v ON v.TemplateId = m.Id
WHERE m.DocCode = @code
ORDER BY v.VersionNo DESC, v.Id DESC;";
                    cmd1.Parameters.Add(new SqlParameter("@code", SqlDbType.NVarChar, 100) { Value = templateCode });
                    using var rd1 = await cmd1.ExecuteReaderAsync();
                    if (await rd1.ReadAsync())
                    {
                        templateTitle = rd1["DocName"] as string ?? "";
                        versionId = Convert.ToInt64(rd1["VersionId"]);
                        descriptorJson = rd1["DescriptorJson"] as string ?? "{}";
                        previewJson = rd1["PreviewJson"] as string ?? "{}";
                        excelPath = rd1["ExcelFilePath"] as string;
                    }
                    else return ("{}", "{}", "", 0L, null);
                }

                // 2. 입력 필드
                var inputs = new List<Dictionary<string, object>>();
                using (var cmd2 = conn.CreateCommand())
                {
                    cmd2.CommandText = @"SELECT Id,[Key],[Type],A1,CellRow,CellColumn FROM DocTemplateField WHERE VersionId=@vid ORDER BY Id;";
                    cmd2.Parameters.Add(new SqlParameter("@vid", SqlDbType.BigInt) { Value = versionId });
                    using var rd2 = await cmd2.ExecuteReaderAsync();
                    while (await rd2.ReadAsync())
                    {
                        string key = rd2["Key"] as string ?? "";
                        string typ = MapFieldType(rd2["Type"] as string);
                        string a1 = rd2["A1"] as string ?? "";
                        if (string.IsNullOrWhiteSpace(a1))
                        {
                            int r = rd2["CellRow"] is DBNull ? 0 : Convert.ToInt32(rd2["CellRow"]);
                            int c = rd2["CellColumn"] is DBNull ? 0 : Convert.ToInt32(rd2["CellColumn"]);
                            if (r > 0 && c > 0) a1 = A1FromRowCol(r, c);
                        }
                        inputs.Add(new Dictionary<string, object> { ["key"] = key, ["type"] = typ, ["required"] = false, ["a1"] = a1 });
                    }
                }

                // 3. 결재/협조 슬롯
                var approvals = new List<Dictionary<string, object>>();
                var cooperations = new List<Dictionary<string, object>>();
                using (var cmd3 = conn.CreateCommand())
                {
                    cmd3.CommandText = @"SELECT * FROM dbo.DocTemplateApproval WHERE VersionId=@vid ORDER BY Slot;";
                    cmd3.Parameters.Add(new SqlParameter("@vid", SqlDbType.BigInt) { Value = versionId });
                    using var rd3 = await cmd3.ExecuteReaderAsync();
                    int ordSlot = GetOrdinalOrThrow(rd3, "Slot");
                    int ordType = GetOrdinalOrMinusOne(rd3, "ApproverType", "Type");
                    int ordValue = GetOrdinalOrMinusOne(rd3, "ApproverValue", "Value");
                    int ordLineType = GetOrdinalOrMinusOne(rd3, "LineType");
                    int ordA1 = GetOrdinalOrMinusOne(rd3, "CellA1", "A1");
                    while (await rd3.ReadAsync())
                    {
                        int slot = rd3.GetInt32(ordSlot);
                        string typ = (ordType >= 0 && !rd3.IsDBNull(ordType)) ? rd3.GetString(ordType) : "Person";
                        string val = (ordValue >= 0 && !rd3.IsDBNull(ordValue)) ? rd3.GetString(ordValue) ?? "" : "";
                        string cellA1 = (ordA1 >= 0 && !rd3.IsDBNull(ordA1)) ? rd3.GetString(ordA1) ?? "" : "";
                        string lineType = (ordLineType >= 0 && !rd3.IsDBNull(ordLineType)) ? rd3.GetString(ordLineType) ?? "" : "";
                        bool isCoop = string.Equals(lineType, "Cooperation", StringComparison.OrdinalIgnoreCase);
                        if (isCoop)
                            cooperations.Add(new Dictionary<string, object> { ["roleKey"] = $"C{slot}", ["approverType"] = MapApproverType(typ), ["lineType"] = "Cooperation", ["value"] = val, ["cellA1"] = cellA1 });
                        else
                            approvals.Add(new Dictionary<string, object> { ["roleKey"] = $"A{slot}", ["approverType"] = MapApproverType(typ), ["required"] = (object)false, ["value"] = val, ["cellA1"] = cellA1 });
                    }
                }

                // 4. DescriptorJson Cooperations 보완 (Cell.A1 파싱)
                if (!string.IsNullOrWhiteSpace(descriptorJson) && descriptorJson != "{}")
                {
                    try
                    {
                        using var descDoc = JsonDocument.Parse(descriptorJson);
                        JsonElement coopsEl;
                        bool hasCoops =
                            descDoc.RootElement.TryGetProperty("Cooperations", out coopsEl) ||
                            descDoc.RootElement.TryGetProperty("cooperations", out coopsEl);

                        if (hasCoops && coopsEl.ValueKind == JsonValueKind.Array)
                        {
                            cooperations.Clear();
                            int idx = 1;
                            foreach (var c in coopsEl.EnumerateArray())
                            {
                                string typ = "Person";
                                if (c.TryGetProperty("ApproverType", out var at) && at.ValueKind == JsonValueKind.String) typ = at.GetString() ?? "Person";

                                string val = "";
                                if (c.TryGetProperty("ApproverValue", out var av) && av.ValueKind == JsonValueKind.String) val = av.GetString() ?? "";
                                else if (c.TryGetProperty("value", out var vv) && vv.ValueKind == JsonValueKind.String) val = vv.GetString() ?? "";

                                // A1 탐색: Cell.A1 → cell.A1 → A1 → a1 → cellA1 → CellA1
                                string cellA1 = "";
                                if (c.TryGetProperty("Cell", out var ce) && ce.ValueKind == JsonValueKind.Object)
                                {
                                    if (ce.TryGetProperty("A1", out var e1) && e1.ValueKind == JsonValueKind.String) cellA1 = e1.GetString() ?? "";
                                    if (string.IsNullOrWhiteSpace(cellA1) && ce.TryGetProperty("a1", out var e2) && e2.ValueKind == JsonValueKind.String) cellA1 = e2.GetString() ?? "";
                                }
                                if (string.IsNullOrWhiteSpace(cellA1) && c.TryGetProperty("cell", out var ce2) && ce2.ValueKind == JsonValueKind.Object)
                                {
                                    if (ce2.TryGetProperty("A1", out var e3) && e3.ValueKind == JsonValueKind.String) cellA1 = e3.GetString() ?? "";
                                    if (string.IsNullOrWhiteSpace(cellA1) && ce2.TryGetProperty("a1", out var e4) && e4.ValueKind == JsonValueKind.String) cellA1 = e4.GetString() ?? "";
                                }
                                if (string.IsNullOrWhiteSpace(cellA1) && c.TryGetProperty("A1", out var dA) && dA.ValueKind == JsonValueKind.String) cellA1 = dA.GetString() ?? "";
                                if (string.IsNullOrWhiteSpace(cellA1) && c.TryGetProperty("a1", out var dAl) && dAl.ValueKind == JsonValueKind.String) cellA1 = dAl.GetString() ?? "";
                                if (string.IsNullOrWhiteSpace(cellA1) && c.TryGetProperty("cellA1", out var ca1) && ca1.ValueKind == JsonValueKind.String) cellA1 = ca1.GetString() ?? "";
                                if (string.IsNullOrWhiteSpace(cellA1) && c.TryGetProperty("CellA1", out var ca2) && ca2.ValueKind == JsonValueKind.String) cellA1 = ca2.GetString() ?? "";

                                string roleKey = "";
                                if (c.TryGetProperty("RoleKey", out var rk) && rk.ValueKind == JsonValueKind.String) roleKey = rk.GetString() ?? "";
                                else if (c.TryGetProperty("roleKey", out var rk2) && rk2.ValueKind == JsonValueKind.String) roleKey = rk2.GetString() ?? "";
                                if (string.IsNullOrWhiteSpace(roleKey)) roleKey = $"C{idx}";

                                cooperations.Add(new Dictionary<string, object> { ["roleKey"] = roleKey, ["approverType"] = MapApproverType(typ), ["lineType"] = "Cooperation", ["value"] = val, ["cellA1"] = cellA1 });
                                idx++;
                            }
                        }
                    }
                    catch { }
                }

                // 5. 최종 직렬화
                descriptorJson = JsonSerializer.Serialize(new Dictionary<string, object>
                {
                    ["inputs"] = inputs,
                    ["approvals"] = approvals,
                    ["cooperations"] = cooperations,
                    ["version"] = "db"
                });
            }

            // 6. 엑셀 경로 정규화
            if (!string.IsNullOrWhiteSpace(excelPath))
            {
                excelPath = NormalizeTemplateExcelPath(excelPath);
                var norm = excelPath.Replace('/', Path.DirectorySeparatorChar).Replace('\\', Path.DirectorySeparatorChar);
                if (!Path.IsPathRooted(norm)) norm = Path.Combine(_env.ContentRootPath, norm);
                excelPath = norm;
            }

            // 7. 프리뷰 빌드
            if (!IsPreviewJsonUsable(previewJson) && !string.IsNullOrWhiteSpace(excelPath) && File.Exists(excelPath))
            {
                try { previewJson = BuildPreviewJsonFromExcel(excelPath); }
                catch (Exception ex) { _log.LogWarning(ex, "PreviewJson build failed. excelPath={Path}", excelPath); previewJson = "{}"; }
            }
            if (string.IsNullOrWhiteSpace(previewJson)) previewJson = "{}";

            return (descriptorJson, previewJson, templateTitle, versionId, excelPath);
        }

        private static string NormalizeTemplateExcelPath(string? raw)
        {
            var s = (raw ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return s;
            if (s.StartsWith("App_Data\\", StringComparison.OrdinalIgnoreCase) || s.StartsWith("App_Data/", StringComparison.OrdinalIgnoreCase))
                return s.Replace('/', '\\');
            var idx = s.IndexOf("App_Data\\", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) idx = s.IndexOf("App_Data/", StringComparison.OrdinalIgnoreCase);
            return idx >= 0 ? s.Substring(idx).Replace('/', '\\') : s;
        }

        private static bool IsPreviewJsonUsable(string? json)
        {
            var s = (json ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s) || s == "{}" || s == "null" || s == "[]") return false;
            try
            {
                using var doc = JsonDocument.Parse(s);
                return doc.RootElement.ValueKind == JsonValueKind.Object &&
                       doc.RootElement.TryGetProperty("cells", out var cells) &&
                       cells.ValueKind == JsonValueKind.Array;
            }
            catch { return false; }
        }

        private static int GetOrdinalOrThrow(System.Data.Common.DbDataReader r, string name)
        {
            for (int i = 0; i < r.FieldCount; i++)
                if (string.Equals(r.GetName(i), name, StringComparison.OrdinalIgnoreCase)) return i;
            throw new InvalidOperationException($"Column '{name}' not found.");
        }

        private static int GetOrdinalOrMinusOne(System.Data.Common.DbDataReader r, params string[] candidates)
        {
            for (int i = 0; i < r.FieldCount; i++) { var n = r.GetName(i); foreach (var c in candidates) if (string.Equals(n, c, StringComparison.OrdinalIgnoreCase)) return i; }
            return -1;
        }

        private static string MapFieldType(string? t)
        {
            var s = (t ?? "").Trim().ToLowerInvariant();
            if (s.StartsWith("date")) return "Date";
            if (s.StartsWith("num") || s.Contains("number") || s.Contains("decimal") || s.Contains("integer")) return "Num";
            return "Text";
        }

        private static string MapApproverType(string? t) { t = (t ?? "").Trim(); return (t == "Person" || t == "Role" || t == "Rule") ? t : "Person"; }

        private static string A1FromRowCol(int row, int col)
        {
            if (row < 1 || col < 1) return "";
            string letters = ""; int n = col;
            while (n > 0) { int r = (n - 1) % 26; letters = (char)('A' + r) + letters; n = (n - 1) / 26; }
            return $"{letters}{row}";
        }

        private static string BuildPreviewJsonFromExcel(string excelPath, int maxRows = 50, int maxCols = 26)
        {
            using var wb = new XLWorkbook(excelPath);
            var ws = wb.Worksheets.First();

            var cells = new List<List<string>>(maxRows);
            for (int r = 1; r <= maxRows; r++) { var row = new List<string>(maxCols); for (int c = 1; c <= maxCols; c++) row.Add(ws.Cell(r, c).GetString()); cells.Add(row); }

            var merges = new List<int[]>();
            foreach (var mr in ws.MergedRanges)
            {
                var a = mr.RangeAddress;
                int r1 = a.FirstAddress.RowNumber, c1 = a.FirstAddress.ColumnNumber;
                int r2 = Math.Min(a.LastAddress.RowNumber, maxRows), c2 = Math.Min(a.LastAddress.ColumnNumber, maxCols);
                if (r1 <= maxRows && c1 <= maxCols && r1 <= r2 && c1 <= c2) merges.Add(new[] { r1, c1, r2, c2 });
            }

            var colW = new List<double>(maxCols); for (int c = 1; c <= maxCols; c++) colW.Add(ws.Column(c).Width);
            var rowH = new List<double>(maxRows); for (int r = 1; r <= maxRows; r++) rowH.Add(ws.Row(r).Height);

            var styles = new Dictionary<string, object>();
            for (int r = 1; r <= maxRows; r++)
                for (int c = 1; c <= maxCols; c++)
                {
                    var cell = ws.Cell(r, c); var st = cell.Style;
                    string? bg = null;
                    try { if (st.Fill.BackgroundColor.ColorType == XLColorType.Color) { var cc = st.Fill.BackgroundColor.Color; bg = $"#{cc.R:X2}{cc.G:X2}{cc.B:X2}"; } } catch { }
                    styles[$"{r},{c}"] = new Dictionary<string, object>
                    {
                        ["font"] = new Dictionary<string, object> { ["name"] = st.Font.FontName ?? "", ["size"] = st.Font.FontSize, ["bold"] = st.Font.Bold },
                        ["align"] = new Dictionary<string, object> { ["h"] = st.Alignment.Horizontal.ToString(), ["v"] = st.Alignment.Vertical.ToString(), ["wrap"] = st.Alignment.WrapText },
                        ["border"] = new Dictionary<string, object> { ["l"] = st.Border.LeftBorder.ToString(), ["r"] = st.Border.RightBorder.ToString(), ["t"] = st.Border.TopBorder.ToString(), ["b"] = st.Border.BottomBorder.ToString() },
                        ["fill"] = new Dictionary<string, object?> { ["bg"] = bg }
                    };
                }

            return JsonSerializer.Serialize(new Dictionary<string, object>
            {
                ["sheet"] = ws.Name,
                ["rows"] = maxRows,
                ["cols"] = maxCols,
                ["cells"] = cells,
                ["merges"] = merges,
                ["colW"] = colW,
                ["rowH"] = rowH,
                ["styles"] = styles
            });
        }
    }
}
