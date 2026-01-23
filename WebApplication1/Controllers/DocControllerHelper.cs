// 2026.01.23 Changed DocController 공통 멤버와 헬퍼를 DocControllerHelper 베이스로 분리
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Security.Cryptography;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Logging;

namespace WebApplication1.Controllers
{
    /// <summary>
    /// Doc 관련 컨트롤러 공통 베이스.
    /// - 필드/DI
    /// - DTO
    /// - 공통 헬퍼(프리뷰, 시간대, 수신자/유틸 등)
    /// </summary>
    [Authorize]
    public abstract class DocControllerHelper : Controller
    {
        protected readonly IConfiguration _cfg;
        protected readonly IWebHostEnvironment _env;
        protected readonly ILogger _log;
        protected readonly IStringLocalizer<SharedResource> _S;
        protected readonly IWebPushNotifier _webPushNotifier;

        protected DocControllerHelper(
            IConfiguration cfg,
            IWebHostEnvironment env,
            ILoggerFactory loggerFactory,
            IStringLocalizer<SharedResource> S,
            IWebPushNotifier webPushNotifier
        )
        {
            _cfg = cfg;
            _env = env;
            _S = S;
            _webPushNotifier = webPushNotifier;
            _log = loggerFactory.CreateLogger(GetType());
        }

        // ========= DTOs =========
        public sealed class ComposePostDto
        {
            [JsonPropertyName("templateCode")]
            public string? TemplateCode { get; set; }

            [JsonPropertyName("inputs")]
            public Dictionary<string, string>? Inputs { get; set; }

            [JsonPropertyName("approvals")]
            public ApprovalsDto? Approvals { get; set; }

            [JsonPropertyName("mail")]
            public MailDto? Mail { get; set; }

            [JsonPropertyName("descriptorVersion")]
            public string? DescriptorVersion { get; set; }

            // 2025.12.23 Added: Compose 화면 조직 멀티 콤보박스 선택 사용자(공유 대상)
            [JsonPropertyName("selectedRecipientUserIds")]
            public List<string>? SelectedRecipientUserIds { get; set; }

            // 2025.12.23 Added: Compose 첨부파일(임시 업로드 결과) 목록
            [JsonPropertyName("attachments")]
            public List<ComposeAttachmentDto>? Attachments { get; set; }
        }

        public sealed class ComposeAttachmentDto
        {
            [JsonPropertyName("fileKey")]
            public string? FileKey { get; set; }        // /DocFile/Upload 응답의 fileKey

            [JsonPropertyName("originalName")]
            public string? OriginalName { get; set; }   // 원본 파일명(표시용)

            [JsonPropertyName("contentType")]
            public string? ContentType { get; set; }

            [JsonPropertyName("byteSize")]
            public long? ByteSize { get; set; }
        }

        public sealed class ApprovalsDto
        {
            public List<string>? To { get; set; }
            public List<ApprovalStepDto>? Steps { get; set; }
        }

        public sealed class ApprovalStepDto
        {
            public string? RoleKey { get; set; }
            public string? ApproverType { get; set; }
            public string? Value { get; set; }
        }

        public sealed class MailDto
        {
            public List<string>? TO { get; set; }
            public List<string>? CC { get; set; }
            public List<string>? BCC { get; set; }
            public string? Subject { get; set; }
            public string? Body { get; set; }
            public bool Send { get; set; } = true;
            public string? From { get; set; }
        }

        // Detail 공유자 변경용 DTO
        public sealed class UpdateSharesDto
        {
            public string? DocId { get; set; }
            public List<string>? SelectedRecipientUserIds { get; set; }
        }

        protected sealed class DescriptorDto
        {
            public string? Version { get; set; }
            public List<InputFieldDto>? Inputs { get; set; }
            public Dictionary<string, object>? Styles { get; set; }
            public List<object>? Approvals { get; set; }
            public List<FlowGroupDto>? FlowGroups { get; set; }
        }

        protected sealed class InputFieldDto
        {
            public string Key { get; set; } = "";
            public string? A1 { get; set; }
            public string? Type { get; set; }
        }

        protected sealed class FlowGroupDto
        {
            public string ID { get; set; } = "";
            public List<string> Keys { get; set; } = new();
        }

        // ========= Preview 생성 =========
        protected static string BuildPreviewJsonFromExcel(string excelPath, int maxRows = 50, int maxCols = 26)
        {
            using var wb = new XLWorkbook(excelPath);
            var ws0 = wb.Worksheets.First();

            double defaultRowPt = ws0.RowHeight; if (defaultRowPt <= 0) defaultRowPt = 15.0;
            double defaultColChar = ws0.ColumnWidth; if (defaultColChar <= 0) defaultColChar = 8.43;

            var cells = new List<List<string>>(maxRows);
            for (int r = 1; r <= maxRows; r++)
            {
                var row = new List<string>(maxCols);
                for (int c = 1; c <= maxCols; c++)
                {
                    var cell = ws0.Cell(r, c);
                    row.Add(cell.GetFormattedString());
                }
                cells.Add(row);
            }

            var merges = new List<int[]>();
            foreach (var mr in ws0.MergedRanges)
            {
                var a = mr.RangeAddress;
                int r1 = a.FirstAddress.RowNumber, c1 = a.FirstAddress.ColumnNumber;
                int r2 = a.LastAddress.RowNumber, c2 = a.LastAddress.ColumnNumber;
                if (r1 > maxRows || c1 > maxCols) continue;
                r2 = Math.Min(r2, maxRows);
                c2 = Math.Min(c2, maxCols);
                if (r1 <= r2 && c1 <= c2) merges.Add(new[] { r1, c1, r2, c2 });
            }

            var colW = new List<double>(maxCols);
            for (int c = 1; c <= maxCols; c++)
            {
                var w = ws0.Column(c).Width;
                if (w <= 0) w = defaultColChar;
                colW.Add(w);
            }

            var rowH = new List<double>(maxRows);
            for (int r = 1; r <= maxRows; r++)
            {
                var h = ws0.Row(r).Height;
                if (h <= 0) h = defaultRowPt;
                rowH.Add(h);
            }

            var styles = new Dictionary<string, object>();
            for (int r = 1; r <= maxRows; r++)
            {
                for (int c = 1; c <= maxCols; c++)
                {
                    var cell = ws0.Cell(r, c);
                    var st = cell.Style;
                    string? bgHex = ToHexIfRgb(cell);
                    styles[$"{r},{c}"] = new
                    {
                        font = new
                        {
                            name = st.Font.FontName,
                            size = st.Font.FontSize,
                            bold = st.Font.Bold,
                            italic = st.Font.Italic,
                            underline = st.Font.Underline != XLFontUnderlineValues.None
                        },
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
                            b = st.Border.BottomBorder.ToString()
                        },
                        fill = new { bg = bgHex }
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
                rowH,
                styles
            });
        }

        protected static string? ToHexIfRgb(IXLCell cell)
        {
            try
            {
                var bg = cell?.Style?.Fill?.BackgroundColor;
                if (bg != null && bg.ColorType == XLColorType.Color)
                {
                    var c = bg.Color;
                    return $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                }
            }
            catch { }
            return null;
        }

        protected static bool TryParseJsonFlexible(string? json, out JsonDocument doc)
        {
            doc = null!;
            if (string.IsNullOrWhiteSpace(json)) return false;

            try
            {
                var first = JsonDocument.Parse(json);
                if (first.RootElement.ValueKind == JsonValueKind.String)
                {
                    var inner = first.RootElement.GetString();
                    first.Dispose();
                    if (string.IsNullOrWhiteSpace(inner)) return false;
                    doc = JsonDocument.Parse(inner!);
                    return true;
                }
                doc = first;
                return true;
            }
            catch
            {
                return false;
            }
        }

        // ========= 시간대/유틸 =========
        protected string ResolveTimeZoneIdForCurrentUser()
        {
            var tzFromClaim = User?.Claims?.FirstOrDefault(c => c.Type == "TimeZoneId")?.Value;
            if (!string.IsNullOrWhiteSpace(tzFromClaim)) return tzFromClaim;
            return "Korea Standard Time";
        }

        protected string ToLocalStringFromUtc(DateTime utc)
        {
            if (utc.Kind != DateTimeKind.Utc)
                utc = DateTime.SpecifyKind(utc, DateTimeKind.Utc);

            var tzId = ResolveTimeZoneIdForCurrentUser();
            TimeZoneInfo tzi;
            try { tzi = TimeZoneInfo.FindSystemTimeZoneById(tzId); }
            catch { tzi = TimeZoneInfo.FindSystemTimeZoneById("Korea Standard Time"); }

            var local = TimeZoneInfo.ConvertTimeFromUtc(utc, tzi);
            return local.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        }

        protected static string FallbackNameFromEmail(string? email)
        {
            if (string.IsNullOrWhiteSpace(email)) return string.Empty;
            var at = email.IndexOf('@');
            return at > 0 ? email[..at] : email;
        }

        protected string ComposeAddress(string? email, string? displayName)
        {
            if (string.IsNullOrWhiteSpace(email)) return string.Empty;
            var name = (displayName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(name)) return email.Trim();
            if (email.Contains('<') && email.Contains('>')) return email.Trim();
            return $"{name} <{email.Trim()}>";
        }

        // ========= Claims → 회사/부서 =========
        protected (string compCd, string? departmentId) GetUserCompDept()
        {
            var comp = User.FindFirstValue("compCd");
            var dept = User.FindFirstValue("departmentId");

            if (!string.IsNullOrWhiteSpace(comp))
                return (comp, string.IsNullOrWhiteSpace(dept) ? null : dept);

            try
            {
                var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
                if (!string.IsNullOrEmpty(userId))
                {
                    var cs = _cfg.GetConnectionString("DefaultConnection");
                    using var conn = new SqlConnection(cs);
                    conn.Open();
                    using var cmd = new SqlCommand(
                        @"SELECT TOP 1 CompCd, DepartmentId FROM dbo.UserProfiles WHERE UserId=@uid", conn);
                    cmd.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 450) { Value = userId });
                    using var rd = cmd.ExecuteReader();
                    if (rd.Read())
                    {
                        var c = rd.IsDBNull(0) ? null : rd.GetString(0);
                        var d = rd.IsDBNull(1) ? null : rd.GetInt32(1).ToString();
                        return (c ?? "", d);
                    }
                }
            }
            catch { }
            return ("", string.IsNullOrWhiteSpace(dept) ? null : dept);
        }

        // ========= 첨부 임시경로 =========
        protected string ResolveTempUploadPath(string fileKey)
        {
            if (string.IsNullOrWhiteSpace(fileKey))
                return string.Empty;

            var safeKey = Path.GetFileName(fileKey.Trim());
            return Path.Combine(_env.ContentRootPath, "App_Data", "Uploads", "Temp", safeKey);
        }

        // ========= (기존 DocController의 기타 공통 헬퍼는 여기로 이동) =========
        // - GetInitialRecipients(...)
        // - ResolveCompanyTimeZone(), BuildSubmissionMail(...)
        // - EnsureApprovalsAndSyncAsync, FillDocumentApprovalsFromEmailsAsync 등
        // 위 메서드들은 시그니처/내부 로직 변경 없이 "그대로" 이 파일(베이스)로 옮기면 됩니다.
    }
}
