// 2026.01.23 Changed Doc 작성 기능을 ComposeController로 분리
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Logging;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class ComposeController : DocControllerHelper
    {
        public ComposeController(
            IConfiguration cfg,
            IWebHostEnvironment env,
            ILoggerFactory loggerFactory,
            IStringLocalizer<SharedResource> S,
            IWebPushNotifier webPushNotifier
        ) : base(cfg, env, loggerFactory, S, webPushNotifier)
        {
        }

        // ========= 입력값 → 엑셀 =========
        protected async Task<string> GenerateExcelFromInputsAsync(
            long versionId,
            string templateExcelPath,
            Dictionary<string, string> inputs,
            string? descriptorVersion)
        {
            if (string.IsNullOrWhiteSpace(templateExcelPath) || !System.IO.File.Exists(templateExcelPath))
                throw new FileNotFoundException("Template excel not found", templateExcelPath);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var maps = new List<(string key, string a1, string type)>();

            await using (var conn = new Microsoft.Data.SqlClient.SqlConnection(cs))
            {
                await conn.OpenAsync();
                await using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
SELECT [Key], [Type], A1, CellRow, CellColumn
FROM dbo.DocTemplateField
WHERE VersionId = @vid
ORDER BY Id;";
                    cmd.Parameters.Add(new Microsoft.Data.SqlClient.SqlParameter("@vid", System.Data.SqlDbType.BigInt) { Value = versionId });

                    await using var rd = await cmd.ExecuteReaderAsync();
                    while (await rd.ReadAsync())
                    {
                        var key = rd["Key"] as string ?? "";
                        var typ = rd["Type"] as string ?? "Text";
                        string a1 = rd["A1"] as string ?? "";
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
            }

            using var wb = new XLWorkbook(templateExcelPath);
            var ws = wb.Worksheets?.FirstOrDefault()
                     ?? throw new InvalidOperationException("Template has no worksheet");

            foreach (var m in maps)
            {
                if (!inputs.TryGetValue(m.key, out var raw)) continue;
                var cell = ws.Cell(m.a1);

                if (string.IsNullOrEmpty(raw))
                {
                    cell.Clear(XLClearOptions.Contents);
                    continue;
                }

                if (raw.Contains('\n') || raw.Contains('\r')) cell.Style.Alignment.WrapText = true;

                switch (m.type.ToLowerInvariant())
                {
                    case "date":
                        if (DateTime.TryParse(raw, out var dt))
                        {
                            cell.Value = dt;
                            cell.Style.DateFormat.Format = "yyyy-MM-dd";
                        }
                        else cell.Value = raw;
                        break;
                    case "num":
                        if (decimal.TryParse(raw, out var dec)) cell.Value = dec;
                        else cell.Value = raw;
                        break;
                    default:
                        cell.Value = raw;
                        break;
                }
            }

            var outDir = Path.Combine(_env.ContentRootPath ?? AppContext.BaseDirectory, "App_Data", "Docs");
            Directory.CreateDirectory(outDir);

            var docId = GenerateDocumentId();
            var outPath = Path.Combine(outDir, $"{docId}.xlsx");

            wb.SaveAs(outPath);
            return outPath;

            static string A1FromRowCol(int row, int col)
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

            static string MapFieldType(string? t)
            {
                var s = (t ?? "").Trim().ToLowerInvariant();
                if (s.StartsWith("date")) return "date";
                if (s.StartsWith("num") || s.Contains("number") || s.Contains("decimal") || s.Contains("integer")) return "num";
                return "text";
            }
        }

        private static string GenerateDocumentId()
        {
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmssfff");
            string suffix = RandomAlphaNumUpper(4);
            return $"DOC_{timestamp}{suffix}";
        }

        private static string RandomAlphaNumUpper(int len)
        {
            const string alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
            Span<byte> buffer = stackalloc byte[len];
            RandomNumberGenerator.Fill(buffer);
            char[] chars = new char[len];
            for (int i = 0; i < len; i++)
                chars[i] = alphabet[buffer[i] % alphabet.Length];
            return new string(chars);
        }

        // ========= (기존 DocController의 Compose 관련 액션은 이 컨트롤러로 이동) =========
        // - New/Select
        // - Create(GET/POST)
        // - Csrf
        // - Upload (방법 B 정책 포함)
        // - 템플릿 로딩/descriptor/preview 준비 등
        // 위 액션들은 기존 시그니처 그대로 이 클래스로 옮기면 됩니다.
    }
}
