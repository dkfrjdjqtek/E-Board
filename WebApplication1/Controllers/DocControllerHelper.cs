using ClosedXML.Excel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Text.Json;
using System.Threading.Tasks;
using WebApplication1.Models;
using WebApplication1.Services;

namespace WebApplication1.Controllers
{
    public sealed class DocControllerHelper
    {
        private readonly IConfiguration _cfg;
        private readonly IWebHostEnvironment _env;

        // 실행 시점에 현재 User를 얻기 위한 accessor
        private readonly Func<ClaimsPrincipal> _userAccessor;

        // ★ 기존 코드가 _user 를 참조해도 컴파일되도록 "프로퍼티"로 복원
        // (컨트롤러 생성 시점 User가 비어있어도, 실제 실행 시점에 평가됨)
        private ClaimsPrincipal _user => _userAccessor();

        // 기존 시그니처(호환 유지)
        public DocControllerHelper(IConfiguration cfg, IWebHostEnvironment env, ClaimsPrincipal user)
            : this(cfg, env, () => user)
        {
        }

        // 새 시그니처
        public DocControllerHelper(IConfiguration cfg, IWebHostEnvironment env, Func<ClaimsPrincipal> userAccessor)
        {
            _cfg = cfg;
            _env = env;
            _userAccessor = userAccessor ?? throw new ArgumentNullException(nameof(userAccessor));
        }


        private ClaimsPrincipal User => _userAccessor();

        // ============================================================
        // ORG TREE (회사/부서/사용자)
        // ============================================================
        public async Task<List<OrgTreeNode>> BuildOrgTreeNodesAsync(string langCode)
        {
            var orgNodes = new List<OrgTreeNode>();
            var compMap = new Dictionary<string, OrgTreeNode>(StringComparer.OrdinalIgnoreCase);
            var deptMap = new Dictionary<string, OrgTreeNode>(StringComparer.OrdinalIgnoreCase);

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
SELECT
    a.CompCd,
    ISNULL(g.Name, a.CompCd) AS CompName,
    a.DepartmentId,
    ISNULL(f.Name, CAST(a.DepartmentId AS nvarchar(50))) AS DeptName,
    a.UserId,
    ISNULL(d.Name, '') AS PositionName,
    a.DisplayName
FROM dbo.UserProfiles a
LEFT JOIN dbo.PositionMasters c      ON a.CompCd = c.CompCd AND a.PositionId = c.Id
LEFT JOIN dbo.PositionMasterLoc d    ON c.Id = d.PositionId AND d.LangCode = @LangCode
LEFT JOIN dbo.DepartmentMasters e    ON a.CompCd = e.CompCd AND a.DepartmentId = e.Id
LEFT JOIN dbo.DepartmentMasterLoc f  ON e.Id = f.DepartmentId AND f.LangCode = @LangCode
LEFT JOIN dbo.CompMasters g          ON a.CompCd = g.CompCd
ORDER BY a.CompCd, e.SortOrder, c.RankLevel, a.DisplayName;";
                cmd.Parameters.Add(new SqlParameter("@LangCode", SqlDbType.NVarChar, 8) { Value = langCode });

                await using var rd = await cmd.ExecuteReaderAsync();
                while (await rd.ReadAsync())
                {
                    var compCd = rd["CompCd"] as string ?? "";
                    if (string.IsNullOrWhiteSpace(compCd)) continue;

                    var compName = rd["CompName"] as string ?? compCd;

                    var deptIdObj = rd["DepartmentId"];
                    if (deptIdObj == DBNull.Value) continue;

                    var deptId = Convert.ToString(deptIdObj) ?? "";
                    if (string.IsNullOrWhiteSpace(deptId)) continue;

                    var deptName = rd["DeptName"] as string ?? deptId;

                    var userId = rd["UserId"] as string ?? "";
                    if (string.IsNullOrWhiteSpace(userId)) continue;

                    var positionName = rd["PositionName"] as string ?? "";
                    var displayName = rd["DisplayName"] as string ?? "";

                    if (!compMap.TryGetValue(compCd, out var compNode))
                    {
                        compNode = new OrgTreeNode
                        {
                            NodeId = compCd,
                            Name = compName,
                            NodeType = "Branch",
                            ParentId = null,
                            Children = new List<OrgTreeNode>()
                        };
                        compMap[compCd] = compNode;
                        orgNodes.Add(compNode);
                    }
                    else
                    {
                        compNode.Children ??= new List<OrgTreeNode>();
                    }

                    var deptKey = $"{compCd}:{deptId}";
                    if (!deptMap.TryGetValue(deptKey, out var deptNode))
                    {
                        deptNode = new OrgTreeNode
                        {
                            NodeId = deptKey,
                            Name = deptName,
                            NodeType = "Dept",
                            ParentId = compNode.NodeId,
                            Children = new List<OrgTreeNode>()
                        };
                        deptMap[deptKey] = deptNode;
                        compNode.Children.Add(deptNode);
                    }
                    else
                    {
                        deptNode.Children ??= new List<OrgTreeNode>();
                    }

                    var caption = string.IsNullOrWhiteSpace(positionName)
                        ? displayName
                        : $"{positionName} {displayName}";

                    deptNode.Children.Add(new OrgTreeNode
                    {
                        NodeId = userId,
                        Name = caption,
                        NodeType = "User",
                        ParentId = deptNode.NodeId,
                        Children = new List<OrgTreeNode>()
                    });
                }
            }
            catch
            {
                // 실패 시 빈 리스트 반환 (기존 동작 유지)
            }

            return orgNodes;
        }

        // ============================================================
        // PATH UTILS
        // ============================================================
        public string ToContentRootAbsolute(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;

            var normalized = path.Replace('/', Path.DirectorySeparatorChar)
                                 .Replace('\\', Path.DirectorySeparatorChar);

            return Path.IsPathRooted(normalized)
                ? normalized
                : Path.Combine(_env.ContentRootPath, normalized);
        }

        public string ToContentRootRelative(string? fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath)) return string.Empty;

            var root = _env.ContentRootPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

            var normalized = fullPath.Replace('/', Path.DirectorySeparatorChar)
                                     .Replace('\\', Path.DirectorySeparatorChar);

            if (Path.IsPathRooted(normalized) &&
                normalized.StartsWith(root, StringComparison.OrdinalIgnoreCase))
            {
                return Path.GetRelativePath(root, normalized);
            }

            return normalized;
        }

        public static string NormalizeTemplateExcelPath(string? raw)
        {
            var s = (raw ?? "").Trim();
            if (string.IsNullOrWhiteSpace(s)) return s;

            if (s.StartsWith("App_Data\\", StringComparison.OrdinalIgnoreCase) ||
                s.StartsWith("App_Data/", StringComparison.OrdinalIgnoreCase))
                return s.Replace('/', '\\');

            var idx = s.IndexOf("App_Data\\", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) idx = s.IndexOf("App_Data/", StringComparison.OrdinalIgnoreCase);

            return idx >= 0 ? s.Substring(idx).Replace('/', '\\') : s;
        }

        public static (string BaseKey, int? Index) ParseKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return ("", null);

            var i = key.LastIndexOf('_');
            return (i > 0 && int.TryParse(key[(i + 1)..], out var n))
                ? (key[..i], n)
                : (key, null);
        }

        // ============================================================
        // PREVIEW BUILDER (스타일 포함)
        // ============================================================
        public static string BuildPreviewJsonFromExcel(string excelPath, int maxRows = 500, int maxCols = 100, Microsoft.Extensions.Logging.ILogger? log = null)
        {
            log?.LogInformation("BuildPreviewJsonFromExcel called. path={Path} exists={Exists}", excelPath, System.IO.File.Exists(excelPath));
            if (string.IsNullOrWhiteSpace(excelPath) || !System.IO.File.Exists(excelPath))
            {
                log?.LogWarning("BuildPreviewJsonFromExcel early return. path={Path}", excelPath);
                return "{}";
            }
            try
            {
                using var wb = new XLWorkbook(excelPath);
                var ws = wb.Worksheets.First();
                var cells = new List<List<string>>(maxRows);
                var merges = new List<int[]>();
                var colW = new List<double>(maxCols);
                var rowH = new List<double>(maxRows);
                var styles = new Dictionary<string, object>(maxRows * maxCols);
                double defaultRowPt = ws.RowHeight <= 0 ? 15 : ws.RowHeight;
                double defaultColChar = ws.ColumnWidth <= 0 ? 8.43 : ws.ColumnWidth;

                int actualMaxC = 1;
                int actualMaxR = 1;
                //foreach (var row in ws.RowsUsed(XLCellsUsedOptions.AllContents))
                    foreach (var row in ws.RowsUsed(XLCellsUsedOptions.All))
                    {
                    var rowNum = row.RowNumber();
                    if (rowNum > actualMaxR) actualMaxR = rowNum;
                    //var last = row.LastCellUsed(XLCellsUsedOptions.AllContents);
                    var last = row.LastCellUsed(XLCellsUsedOptions.All);
                    if (last != null && last.Address.ColumnNumber > actualMaxC)
                        actualMaxC = last.Address.ColumnNumber;
                }
                foreach (var mr in ws.MergedRanges)
                {
                    var lastR = mr.RangeAddress.LastAddress.RowNumber;
                    var lastC = mr.RangeAddress.LastAddress.ColumnNumber;
                    if (lastR > actualMaxR) actualMaxR = lastR;
                    if (lastC > actualMaxC) actualMaxC = lastC; // 
                    // ★ lastC로 actualMaxC 확장 제거
                }
                actualMaxR = Math.Min(actualMaxR, maxRows);
                actualMaxC = Math.Min(actualMaxC , maxCols);

                for (int r = 1; r <= actualMaxR; r++)
                {
                    var row = new List<string>(actualMaxC);
                    for (int c = 1; c <= actualMaxC; c++)
                        row.Add(ws.Cell(r, c).GetFormattedString());
                    cells.Add(row);
                }

                foreach (var mr in ws.MergedRanges)
                {
                    var a = mr.RangeAddress;
                    int r1 = a.FirstAddress.RowNumber, c1 = a.FirstAddress.ColumnNumber;
                    int r2 = Math.Min(a.LastAddress.RowNumber, actualMaxR);
                    int c2 = Math.Min(a.LastAddress.ColumnNumber, actualMaxC);
                    if (r1 > r2 || c1 > c2) continue;
                    merges.Add(new[] { r1, c1, r2, c2 });
                }

                for (int c = 1; c <= actualMaxC; c++)
                    colW.Add(ws.Column(c).Width <= 0 ? defaultColChar : ws.Column(c).Width);

                for (int r = 1; r <= actualMaxR; r++)
                    rowH.Add(ws.Row(r).Height <= 0 ? defaultRowPt : ws.Row(r).Height);

                for (int r = 1; r <= actualMaxR; r++)
                {
                    for (int c = 1; c <= actualMaxC; c++)
                    {
                        var cell = ws.Cell(r, c);
                        var st = cell.Style;
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
                            fill = new { bg = ToHexIfRgb(cell) }
                        };
                    }
                }
                var colPx = new List<double>(maxCols);
                for (int c = 1; c <= actualMaxC; c++)
                {
                    var w = ws.Column(c).Width <= 0 ? defaultColChar : ws.Column(c).Width;
                    colPx.Add(Math.Round(w * 7 + 5));
                }

                return JsonSerializer.Serialize(new
                {
                    sheet = ws.Name,
                    rows = actualMaxR,
                    cols = actualMaxC,
                    cells,
                    merges,
                    colW,
                    colPx, // ★ 추가
                    rowH,
                    styles
                });
            }
            catch (Exception ex)
            {
                log?.LogWarning(ex, "BuildPreviewJsonFromExcel failed. path={Path}", excelPath);
                return "{}";
            }
        }

        private static string? ToHexIfRgb(IXLCell cell)
        {
            try
            {
                var bg = cell.Style.Fill.BackgroundColor;
                if (bg.ColorType == XLColorType.Color)
                {
                    var c = bg.Color;
                    return $"#{c.R:X2}{c.G:X2}{c.B:X2}";
                }
            }
            catch { }
            return null;
        }

        public static bool TryParseJsonFlexible(string? json, out JsonDocument doc)
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

        // ============================================================
        // TIMEZONE
        // ============================================================
        public TimeZoneInfo ResolveCompanyTimeZone()
        {
            // 우선순위: 사용자 클레임 TimeZoneId -> CompMasters.TimeZoneId -> 기본 후보
            var tzFromClaim = _user?.Claims?.FirstOrDefault(c => c.Type == "TimeZoneId")?.Value;

            string? tzFromDb = null;
            try
            {
                // (수정) _user가 null일 수 있으므로 null-조건 연산자 적용
                var compCd = _user?.FindFirstValue("compCd") ?? "";
                var cs = _cfg.GetConnectionString("DefaultConnection");
                if (!string.IsNullOrWhiteSpace(cs) && !string.IsNullOrWhiteSpace(compCd))
                {
                    using var conn = new SqlConnection(cs);
                    conn.Open();
                    using var cmd = new SqlCommand(@"SELECT TOP 1 TimeZoneId FROM dbo.CompMasters WHERE CompCd=@c", conn);
                    cmd.Parameters.Add(new SqlParameter("@c", SqlDbType.VarChar, 10) { Value = compCd });
                    tzFromDb = cmd.ExecuteScalar() as string;
                }
            }
            catch { }

            var candidates = new List<string>();
            if (!string.IsNullOrWhiteSpace(tzFromClaim)) candidates.Add(tzFromClaim!.Trim());
            if (!string.IsNullOrWhiteSpace(tzFromDb)) candidates.Add(tzFromDb!.Trim());
            candidates.Add("Asia/Seoul");
            candidates.Add("Korea Standard Time");

            foreach (var id in candidates.Where(x => !string.IsNullOrWhiteSpace(x)))
            {
                if (TryResolveTimeZone(id!, out var tz))
                    return tz;
            }

            return TimeZoneInfo.Utc;
        }

        private static bool TryResolveTimeZone(string id, out TimeZoneInfo tz)
        {
            tz = TimeZoneInfo.Utc;
            try
            {
                tz = TimeZoneInfo.FindSystemTimeZoneById(id);
                return true;
            }
            catch
            {
                // Windows/Linux 매핑 최소치
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                {
                    ["Asia/Seoul"] = "Korea Standard Time",
                    ["Korea Standard Time"] = "Asia/Seoul",
                    ["Asia/Ho_Chi_Minh"] = "SE Asia Standard Time",
                    ["Asia/Jakarta"] = "SE Asia Standard Time"
                };
                if (map.TryGetValue(id, out var mapped))
                {
                    try
                    {
                        tz = TimeZoneInfo.FindSystemTimeZoneById(mapped);
                        return true;
                    }
                    catch { }
                }
                return false;
            }
        }

        public string ToLocalStringFromUtc(DateTime utc)
        {
            if (utc.Kind != DateTimeKind.Utc)
                utc = DateTime.SpecifyKind(utc, DateTimeKind.Utc);

            var tz = ResolveCompanyTimeZone();
            var local = TimeZoneInfo.ConvertTimeFromUtc(utc, tz);
            return local.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        }

        // ============================================================
        // USER / PROFILE HELPERS
        // ============================================================
        public string GetCurrentUserEmail()
        {
            try
            {
                var uid = _user.FindFirstValue(ClaimTypes.NameIdentifier);
                if (string.IsNullOrWhiteSpace(uid)) return "";

                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = new SqlCommand("SELECT TOP 1 Email FROM dbo.AspNetUsers WHERE Id=@id", conn);
                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 450) { Value = uid });

                return cmd.ExecuteScalar() as string ?? "";
            }
            catch { return ""; }
        }

        public string GetCurrentUserDisplayNameStrict()
        {
            try
            {
                var uid = _user.FindFirstValue(ClaimTypes.NameIdentifier);
                if (string.IsNullOrWhiteSpace(uid)) return "";

                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = new SqlCommand(
                    @"SELECT TOP 1 NULLIF(LTRIM(RTRIM(DisplayName)), N'') FROM dbo.UserProfiles WHERE UserId=@u", conn);
                cmd.Parameters.Add(new SqlParameter("@u", SqlDbType.NVarChar, 450) { Value = uid });

                var name = cmd.ExecuteScalar() as string;
                return string.IsNullOrWhiteSpace(name) ? "" : name!;
            }
            catch { return ""; }
        }

        public string GetDisplayNameByEmailStrict(string? email)
        {
            if (string.IsNullOrWhiteSpace(email)) return "";

            try
            {
                var em = email.Trim();

                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = new SqlCommand(@"
SELECT TOP 1 NULLIF(LTRIM(RTRIM(p.DisplayName)), N'')
FROM dbo.AspNetUsers u
JOIN dbo.UserProfiles p ON u.Id = p.UserId
WHERE LTRIM(RTRIM(u.Email)) = @e OR u.NormalizedEmail = UPPER(@e);", conn);

                cmd.Parameters.Add(new SqlParameter("@e", SqlDbType.NVarChar, 256) { Value = em });

                var name = cmd.ExecuteScalar() as string;
                return string.IsNullOrWhiteSpace(name) ? "" : name!;
            }
            catch { return ""; }
        }

        public static string FallbackNameFromEmail(string? email)
        {
            if (string.IsNullOrWhiteSpace(email)) return "";
            var at = email.IndexOf('@');
            return at > 0 ? email[..at] : email;
        }

        public string ComposeAddress(string? email, string? displayName)
        {
            if (string.IsNullOrWhiteSpace(email)) return "";
            var em = email.Trim();
            var name = (displayName ?? "").Trim();
            if (string.IsNullOrWhiteSpace(name)) return em;
            if (em.Contains('<') && em.Contains('>')) return em;
            return $"{name} <{em}>";
        }

        public (string compCd, string? departmentId) GetUserCompDept()
        {
            var comp = _user.FindFirstValue("compCd");
            var dept = _user.FindFirstValue("departmentId");

            if (!string.IsNullOrWhiteSpace(comp))
                return (comp, string.IsNullOrWhiteSpace(dept) ? null : dept);

            try
            {
                var uid = _user.FindFirstValue(ClaimTypes.NameIdentifier);
                if (string.IsNullOrWhiteSpace(uid)) return ("", null);

                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = new SqlCommand(
                    @"SELECT TOP 1 CompCd, DepartmentId FROM dbo.UserProfiles WHERE UserId=@u", conn);
                cmd.Parameters.Add(new SqlParameter("@u", SqlDbType.NVarChar, 450) { Value = uid });

                using var r = cmd.ExecuteReader();
                if (r.Read())
                {
                    var c = r.IsDBNull(0) ? "" : (r.GetString(0) ?? "");
                    var d = r.IsDBNull(1) ? null : r.GetInt32(1).ToString(CultureInfo.InvariantCulture);
                    return (c, d);
                }
            }
            catch { }

            return ("", null);
        }

        public static string CalculateViewerRole(
            string? createdBy,
            string viewerId,
            List<(int StepOrder, string? UserId, string Status)> approvals)
        {
            if (string.IsNullOrWhiteSpace(viewerId))
                return "Unknown";

            var isCreator = !string.IsNullOrWhiteSpace(createdBy) &&
                            string.Equals(createdBy, viewerId, StringComparison.OrdinalIgnoreCase);

            var isApprover = approvals.Exists(a =>
                !string.IsNullOrWhiteSpace(a.UserId) &&
                string.Equals(a.UserId, viewerId, StringComparison.OrdinalIgnoreCase));

            if (isCreator && isApprover) return "Creator+Approver";
            if (isCreator) return "Creator";
            if (isApprover) return "Approver";
            return "Viewer";
        }

        // Push 알림 결재용
        public static async Task SendApprovalPendingBadgeAsync(IWebPushNotifier notifier,IStringLocalizer? S, SqlConnection conn, IEnumerable<string> targetUserIds, string url = "/", string tag = "badge-approval-pending")
        {
            if (notifier == null) throw new ArgumentNullException(nameof(notifier));
            if (conn == null) throw new ArgumentNullException(nameof(conn));
            if (targetUserIds == null) return;

            if (conn.State != ConnectionState.Open)
                await conn.OpenAsync();

            //  사용자가 원하는 PUSH_* 우선, 기존 LOGIN_* 키는 fallback
            var titleText = (S?["PUSH_SummaryTitle"] ?? "PUSH_SummaryTitle").ToString();
            var bodyTpl = (S?["PUSH_ApprovalPending"] ?? "PUSH_ApprovalPending").ToString(); // 예: "결재 대기 {0}건"

            foreach (var uid in targetUserIds
                         .Where(x => !string.IsNullOrWhiteSpace(x))
                         .Select(x => x.Trim())
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var n = await GetApprovalPendingCountAsync(conn, uid);

                await notifier.SendToUserIdAsync(
                    userId: uid,
                    title: titleText,
                    body: string.Format(CultureInfo.CurrentCulture, bodyTpl, n),
                    url: url,
                    tag: tag
                );
            }
        }

        //  BoardBadges(sqlApprovalPending)와 동일 기준으로 집계 (PendingHold 포함, Pending% 포함)
        private static async Task<int> GetApprovalPendingCountAsync(SqlConnection conn, string targetUserId)
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentApprovals a
JOIN dbo.Documents d ON a.DocId = d.DocId
WHERE a.UserId = @UserId
  AND ISNULL(a.Status, N'') LIKE N'Pending%'
  AND (
        ISNULL(d.Status, N'') = N'PendingA'     + CAST(a.StepOrder AS nvarchar(10))
     OR ISNULL(d.Status, N'') = N'PendingHoldA' + CAST(a.StepOrder AS nvarchar(10))
  )
  AND ISNULL(d.Status, N'') NOT LIKE N'Recalled%';";
            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
            return Convert.ToInt32(await cmd.ExecuteScalarAsync());
        }

        //  conn을 외부에서 재사용하는 구조이므로, 여기서 무조건 OpenAsync 하지 않도록 수정
        public static async Task<string> GetA1ApproverUserIdByDocAsync(SqlConnection conn, string xDocId)
        {
            if (conn == null) throw new ArgumentNullException(nameof(conn));
            if (conn.State != ConnectionState.Open)
                await conn.OpenAsync();

            var sql = @"
SELECT TOP (1) a.UserId
FROM dbo.DocumentApprovals a
WHERE a.DocId = @DocId
  AND a.StepOrder = 1
ORDER BY a.Id;";

            await using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = xDocId });

            var obj = await cmd.ExecuteScalarAsync();
            return (obj == null || obj == DBNull.Value) ? "" : (obj.ToString() ?? "").Trim();
        }

        // Push 알림 공유용
        public static async Task SendSharedUnreadBadgeAsync(IWebPushNotifier notifier, IStringLocalizer? S, SqlConnection conn, IEnumerable<string> targetUserIds, string url = "/", string tag = "badge-shared")
        {
            if (notifier == null) throw new ArgumentNullException(nameof(notifier));
            if (conn == null) throw new ArgumentNullException(nameof(conn));

            if (conn.State != ConnectionState.Open)
                await conn.OpenAsync();

            var ids = (targetUserIds ?? Array.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (ids.Count == 0) return;

            //  사용자가 원하는 PUSH_* 우선, 기존 LOGIN_* 키는 fallback
            var titleText = (S?["PUSH_SummaryTitle"] ?? "PUSH_SummaryTitle").ToString();

            //  사용자가 변경한 PUSH__SharedUnread 우선 지원 (double underscore)
            var bodyTpl = (S?["PUSH_SharedUnread"] ?? "PUSH_SharedUnread").ToString();

            foreach (var uid in ids)
            {
                var n = await GetSharedUnreadCountAsync(conn, uid);

                await notifier.SendToUserIdAsync(
                    userId: uid,
                    title: titleText,
                    body: string.Format(CultureInfo.CurrentCulture, bodyTpl, n),
                    url: url,
                    tag: tag
                );
            }
        }

        //  BoardBadges(sqlSharedUnread)와 동일 기준(IsRead=0)으로 집계
        private static async Task<int> GetSharedUnreadCountAsync(SqlConnection conn, string targetUserId)
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentShares s
WHERE s.UserId = @UserId
  AND ISNULL(s.IsRevoked, 0) = 0
  AND (s.ExpireAt IS NULL OR s.ExpireAt > SYSUTCDATETIME())
  AND ISNULL(s.IsRead, 0) = 0;";
            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
            return Convert.ToInt32(await cmd.ExecuteScalarAsync());
        }

        public static async Task SendCooperationPendingBadgeAsync(IWebPushNotifier notifier, IStringLocalizer? S, SqlConnection conn, IEnumerable<string> targetUserIds, string docId, string url = "/", string tag = "badge-cooperation-pending")
        {
            if (notifier == null) throw new ArgumentNullException(nameof(notifier));
            if (conn == null) throw new ArgumentNullException(nameof(conn));
            if (targetUserIds == null) return;

            if (conn.State != ConnectionState.Open)
                await conn.OpenAsync();

            var titleText = (S?["PUSH_SummaryTitle"] ?? "PUSH_SummaryTitle").ToString();
            var bodyTpl = (S?["PUSH_CooperationPending"] ?? "PUSH_CooperationPending").ToString();

            foreach (var uid in targetUserIds
                         .Where(x => !string.IsNullOrWhiteSpace(x))
                         .Select(x => x.Trim())
                         .Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var n = await GetCooperationPendingCountAsync(conn, uid);

                await notifier.SendToUserIdAsync(
                    userId: uid,
                    title: titleText,
                    body: string.Format(CultureInfo.CurrentCulture, bodyTpl, n),
                    url: url.Contains("{docId}") ? url.Replace("{docId}", Uri.EscapeDataString(docId)) : url,
                    tag: tag
                );
            }
        }

        private static async Task<int> GetCooperationPendingCountAsync(SqlConnection conn, string targetUserId)
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentCooperations c
JOIN dbo.Documents d ON d.DocId = c.DocId
WHERE c.UserId = @UserId
  AND ISNULL(c.Status, N'Pending') = N'Pending'
  AND ISNULL(d.Status, N'') NOT LIKE N'Recalled%';";
            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
            return Convert.ToInt32(await cmd.ExecuteScalarAsync());
        }
    }
}