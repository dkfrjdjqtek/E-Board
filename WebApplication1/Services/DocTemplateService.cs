using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace WebApplication1.Services
{
    public class DocTemplateService : IDocTemplateService
    {
        private readonly IConfiguration _cfg;
        private readonly ILogger<DocTemplateService> _log;

        public DocTemplateService(IConfiguration cfg, ILogger<DocTemplateService> log)
        {
            _cfg = cfg;
            _log = log;
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

                // 1) 최신 버전 메타
                using (var cmd1 = conn.CreateCommand())
                {
                    cmd1.CommandText = @"
SELECT TOP 1 m.DocName, v.Id AS VersionId, v.DescriptorJson, v.PreviewJson, v.ExcelFilePath
FROM DocTemplateMaster m
JOIN DocTemplateVersion v ON v.TemplateId = m.Id
WHERE m.DocCode = @code
ORDER BY v.VersionNo DESC;";
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
                    else
                    {
                        return ("{}", "{}", "", 0L, null);
                    }
                }

                // 2) fields
                var inputs = new List<object>();
                using (var cmd2 = conn.CreateCommand())
                {
                    cmd2.CommandText = @"
SELECT Id, [Key], [Type], A1, CellRow, CellColumn
FROM DocTemplateField
WHERE VersionId = @vid
ORDER BY Id;";
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

                        inputs.Add(new { key, type = typ, required = false, a1 });
                    }
                }

                // 3) approvals  ← 위에 제시한 코드
                var approvals = new List<object>();

                using (var cmd3 = conn.CreateCommand())
                {
                    cmd3.CommandText = @"
        SELECT *
        FROM dbo.DocTemplateApproval
        WHERE VersionId = @vid
        ORDER BY Slot;";
                    cmd3.Parameters.Add(new SqlParameter("@vid", SqlDbType.BigInt) { Value = versionId });

                    using var rd3 = await cmd3.ExecuteReaderAsync();

                    // 동적으로 컬럼 ordinal 찾기
                    int ordSlot = GetOrdinalOrThrow(rd3, "Slot");
                    int ordType = GetOrdinalOrMinusOne(rd3, "ApproverType", "Type");
                    int ordValue = GetOrdinalOrMinusOne(rd3, "ApproverValue", "Value");

                    while (await rd3.ReadAsync())
                    {
                        int slot = rd3.GetInt32(ordSlot);

                        string typ = "Person";
                        if (ordType >= 0 && !rd3.IsDBNull(ordType))
                            typ = rd3.GetString(ordType);

                        string val = "";
                        if (ordValue >= 0 && !rd3.IsDBNull(ordValue))
                            val = rd3.GetString(ordValue);

                        approvals.Add(new
                        {
                            roleKey = $"A{slot}",
                            approverType = MapApproverType(typ),
                            required = false,
                            value = val
                        });
                    }
                }

                // 4) descriptorJson 조합
                descriptorJson = JsonSerializer.Serialize(new
                {
                    inputs,
                    approvals,
                    version = "db"
                });
            }

            // 5) previewJson
            if (string.IsNullOrWhiteSpace(previewJson) && !string.IsNullOrWhiteSpace(excelPath))
            {
                try { previewJson = BuildPreviewJsonFromExcel(excelPath); }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "PreviewJson build failed. excelPath={Path}", excelPath);
                    previewJson = "{}";
                }
            }
            if (string.IsNullOrWhiteSpace(previewJson)) previewJson = "{}";

            return (descriptorJson, previewJson, templateTitle, versionId, excelPath);
        }

        // ===== Helpers =====
        static int GetOrdinalOrThrow(System.Data.Common.DbDataReader r, string name)
        {
            for (int i = 0; i < r.FieldCount; i++)
                if (string.Equals(r.GetName(i), name, StringComparison.OrdinalIgnoreCase)) return i;
            throw new InvalidOperationException($"Column '{name}' not found.");
        }
        static int GetOrdinalOrMinusOne(System.Data.Common.DbDataReader r, params string[] candidates)
        {
            for (int i = 0; i < r.FieldCount; i++)
            {
                var n = r.GetName(i);
                foreach (var c in candidates)
                    if (string.Equals(n, c, StringComparison.OrdinalIgnoreCase)) return i;
            }
            return -1; // 없으면 -1
        }
        private static string MapFieldType(string? t)
        {
            var s = (t ?? "").Trim().ToLowerInvariant();
            if (s.StartsWith("date")) return "Date";
            if (s.StartsWith("num") || s.Contains("number") || s.Contains("decimal") || s.Contains("integer")) return "Num";
            return "Text";
        }

        private static string MapApproverType(string? t)
        {
            t = (t ?? "").Trim();
            return (t == "Person" || t == "Role" || t == "Rule") ? t : "Person";
        }

        private static string A1FromRowCol(int row, int col)
        {
            if (row < 1 || col < 1) return "";
            string letters = "";
            int n = col;
            while (n > 0)
            {
                int r = (n - 1) % 26;
                letters = (char)('A' + r) + letters;
                n = (n - 1) / 26;
            }
            return $"{letters}{row}";
        }

        /// <summary>엑셀에서 간단 미리보기 JSON 생성(값/병합/열폭/간단 스타일)</summary>
        private static string BuildPreviewJsonFromExcel(string excelPath, int maxRows = 50, int maxCols = 26)
        {
            using var wb = new XLWorkbook(excelPath);
            var ws0 = wb.Worksheets.First();

            // cells
            var cells = new List<List<string>>(maxRows);
            for (int r = 1; r <= maxRows; r++)
            {
                var row = new List<string>(maxCols);
                for (int c = 1; c <= maxCols; c++) row.Add(ws0.Cell(r, c).GetString());
                cells.Add(row);
            }

            // merges
            var merges = new List<int[]>();
            foreach (var mr in ws0.MergedRanges)
            {
                var a = mr.RangeAddress;
                int r1 = a.FirstAddress.RowNumber, c1 = a.FirstAddress.ColumnNumber;
                int r2 = a.LastAddress.RowNumber, c2 = a.LastAddress.ColumnNumber;
                if (r1 > maxRows || c1 > maxCols) continue;
                r2 = Math.Min(r2, maxRows); c2 = Math.Min(c2, maxCols);
                if (r1 <= r2 && c1 <= c2) merges.Add(new[] { r1, c1, r2, c2 });
            }

            // col widths
            var colW = new List<double>(maxCols);
            for (int c = 1; c <= maxCols; c++) colW.Add(ws0.Column(c).Width);

            // very light styles
            var styles = new Dictionary<string, object>();
            for (int r = 1; r <= maxRows; r++)
            {
                for (int c = 1; c <= maxCols; c++)
                {
                    var cell = ws0.Cell(r, c);
                    var st = cell.Style;
                    string? bg = null;
                    try
                    {
                        if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Color)
                        {
                            var cc = cell.Style.Fill.BackgroundColor.Color;
                            bg = $"#{cc.R:X2}{cc.G:X2}{cc.B:X2}";
                        }
                    }
                    catch { }

                    styles[$"{r},{c}"] = new
                    {
                        font = new { name = st.Font.FontName, size = st.Font.FontSize, bold = st.Font.Bold },
                        align = new
                        {
                            h = st.Alignment.Horizontal.ToString(),
                            v = st.Alignment.Vertical.ToString(),
                            wrap = st.Alignment.WrapText
                        },
                        border = new
                        {
                            l = st.Border.LeftBorder.ToString(),
                            r = st.Border.RightBorder.ToString(),
                            t = st.Border.TopBorder.ToString(),
                            b = st.Border.BottomBorder.ToString(),
                        },
                        fill = new { bg }
                    };
                }
            }

            return JsonSerializer.Serialize(new
            {
                sheet = ws0.Name,
                rows = maxRows,
                cols = maxCols,
                cells,
                merges,
                colW,
                styles
            });
        }
    }
}
