// 2026.01.26 Changed: DocControllerHelper 공용 함수 누락 보완 및 컴파일 가능하도록 의존성 주입과 접근 제어 수정
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using WebApplication1.Models;

namespace WebApplication1.Controllers
{
    public class DocControllerHelper
    {
        private readonly IConfiguration _cfg;
        private readonly IWebHostEnvironment _env;
        private readonly ClaimsPrincipal _user;

        public DocControllerHelper(IConfiguration cfg, IWebHostEnvironment env, ClaimsPrincipal user)
        {
            _cfg = cfg;
            _env = env;
            _user = user;
        }

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
    ISNULL(g.Name, a.CompCd)        AS CompName,
    a.DepartmentId,
    ISNULL(f.Name, CAST(a.DepartmentId AS nvarchar(50))) AS DeptName,
    a.UserId,
    ISNULL(d.Name, '')              AS PositionName,
    a.DisplayName
FROM UserProfiles a
LEFT JOIN AspNetUsers b          ON a.UserId = b.Id
LEFT JOIN PositionMasters c      ON a.CompCd = c.CompCd AND a.PositionId = c.Id
LEFT JOIN PositionMasterLoc d    ON c.Id = d.PositionId AND d.LangCode = @LangCode
LEFT JOIN DepartmentMasters e    ON a.CompCd = e.CompCd AND a.DepartmentId = e.Id
LEFT JOIN DepartmentMasterLoc f  ON e.Id = f.DepartmentId AND f.LangCode = @LangCode
LEFT JOIN CompMasters g          ON a.CompCd = g.CompCd
ORDER BY a.CompCd, e.SortOrder, c.RankLevel, a.DisplayName;";
                cmd.Parameters.Add(new SqlParameter("@LangCode", SqlDbType.NVarChar, 8) { Value = langCode });

                await using var rd = await cmd.ExecuteReaderAsync();
                while (await rd.ReadAsync())
                {
                    var compCd = rd["CompCd"] as string ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(compCd))
                        continue;

                    var compName = rd["CompName"] as string ?? compCd;

                    var deptIdObj = rd["DepartmentId"];
                    if (deptIdObj == null || deptIdObj == DBNull.Value)
                        continue;

                    var deptId = Convert.ToString(deptIdObj) ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(deptId))
                        continue;

                    var deptName = rd["DeptName"] as string ?? deptId;

                    var userId = rd["UserId"] as string ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(userId))
                        continue;

                    var positionName = rd["PositionName"] as string ?? string.Empty;
                    var displayName = rd["DisplayName"] as string ?? string.Empty;

                    // 회사 노드
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

                    // 부서 노드 (유니크 NodeId: compCd:deptId)
                    var deptKey = compCd + ":" + deptId;
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

                    // 사용자 노드
                    var caption = string.IsNullOrWhiteSpace(positionName)
                        ? displayName
                        : positionName + " " + displayName;

                    var userNode = new OrgTreeNode
                    {
                        NodeId = userId,
                        Name = caption,
                        NodeType = "User",
                        ParentId = deptNode.NodeId,
                        Children = new List<OrgTreeNode>()
                    };

                    deptNode.Children.Add(userNode);
                }
            }
            catch
            {
                // 실패 시 빈 리스트 반환 (기존 동작 유지)
            }

            return orgNodes;
        }

        public string ToContentRootAbsolute(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;

            var normalized = path
                .Replace('/', Path.DirectorySeparatorChar)
                .Replace('\\', Path.DirectorySeparatorChar);

            if (Path.IsPathRooted(normalized))
                return normalized;

            return Path.Combine(_env.ContentRootPath, normalized);
        }

        public string ToContentRootRelative(string? fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath)) return string.Empty;

            var root = _env.ContentRootPath.TrimEnd(
                Path.DirectorySeparatorChar,
                Path.AltDirectorySeparatorChar);

            var normalized = fullPath
                .Replace('/', Path.DirectorySeparatorChar)
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
            var s = (raw ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(s)) return s;

            if (s.StartsWith("App_Data\\", StringComparison.OrdinalIgnoreCase) || s.StartsWith("App_Data/", StringComparison.OrdinalIgnoreCase))
                return s.Replace('/', '\\');

            var idx = s.IndexOf("App_Data\\", StringComparison.OrdinalIgnoreCase);
            if (idx < 0) idx = s.IndexOf("App_Data/", StringComparison.OrdinalIgnoreCase);

            if (idx >= 0)
                return s.Substring(idx).Replace('/', '\\');

            return s;
        }

        public static (string BaseKey, int? Index) ParseKey(string key)
        {
            if (string.IsNullOrWhiteSpace(key)) return ("", null);
            var i = key.LastIndexOf('_');
            if (i > 0 && int.TryParse(key[(i + 1)..], out var n))
                return (key[..i], n);
            return (key, null);
        }

        public static string BuildPreviewJsonFromExcel(string excelPath, int maxRows = 50, int maxCols = 26)
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

        private static string? ToHexIfRgb(IXLCell cell)
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

        public TimeZoneInfo ResolveCompanyTimeZone()
        {
            var compCd = _user.FindFirstValue("compCd") ?? "";

            string? tzIdFromDb = null;
            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection");
                if (!string.IsNullOrWhiteSpace(cs) && !string.IsNullOrWhiteSpace(compCd))
                {
                    using var conn = new SqlConnection(cs);
                    conn.Open();
                    using var cmd = new SqlCommand(
                        @"SELECT TOP 1 TimeZoneId FROM dbo.CompMasters WHERE CompCd = @comp", conn);
                    cmd.Parameters.Add(new SqlParameter("@comp", SqlDbType.VarChar, 10) { Value = compCd });
                    tzIdFromDb = cmd.ExecuteScalar() as string;
                }
            }
            catch { }

            var candidates = new List<string>();
            if (!string.IsNullOrWhiteSpace(tzIdFromDb)) candidates.Add(tzIdFromDb.Trim());
            candidates.Add("Asia/Seoul");
            candidates.Add("Korea Standard Time");

            foreach (var id in candidates)
            {
                try { return TimeZoneInfo.FindSystemTimeZoneById(id); }
                catch
                {
                    var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["Asia/Seoul"] = "Korea Standard Time",
                        ["Korea Standard Time"] = "Asia/Seoul",
                        ["Asia/Ho_Chi_Minh"] = "SE Asia Standard Time",
                        ["Asia/Jakarta"] = "SE Asia Standard Time",
                    };
                    if (map.TryGetValue(id, out var mapped))
                    {
                        try { return TimeZoneInfo.FindSystemTimeZoneById(mapped); } catch { }
                    }
                }
            }

            try { return TimeZoneInfo.FindSystemTimeZoneById("Korea Standard Time"); } catch { }
            try { return TimeZoneInfo.FindSystemTimeZoneById("Asia/Seoul"); } catch { }
            return TimeZoneInfo.Utc;
        }

        public string ResolveTimeZoneIdForCurrentUser()
        {
            var tzFromClaim = _user?.Claims?.FirstOrDefault(c => c.Type == "TimeZoneId")?.Value;
            if (!string.IsNullOrWhiteSpace(tzFromClaim)) return tzFromClaim;
            return "Korea Standard Time";
        }

        public string ToLocalStringFromUtc(DateTime utc)
        {
            if (utc.Kind != DateTimeKind.Utc)
                utc = DateTime.SpecifyKind(utc, DateTimeKind.Utc);

            var tzId = ResolveTimeZoneIdForCurrentUser();
            TimeZoneInfo tzi;
            try
            {
                tzi = TimeZoneInfo.FindSystemTimeZoneById(tzId);
            }
            catch
            {
                tzi = TimeZoneInfo.FindSystemTimeZoneById("Korea Standard Time");
            }

            var local = TimeZoneInfo.ConvertTimeFromUtc(utc, tzi);
            return local.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        }

        public string GetCurrentUserEmail()
        {
            try
            {
                var uid = _user?.FindFirstValue(ClaimTypes.NameIdentifier);
                if (string.IsNullOrWhiteSpace(uid)) return string.Empty;

                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                conn.Open();
                using var cmd = new SqlCommand("SELECT TOP 1 Email FROM dbo.AspNetUsers WHERE Id=@id", conn);
                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 450) { Value = uid });
                return cmd.ExecuteScalar() as string ?? string.Empty;
            }
            catch { return string.Empty; }
        }

        public string GetCurrentUserDisplayNameStrict()
        {
            try
            {
                var uid = _user?.FindFirstValue(ClaimTypes.NameIdentifier);
                if (string.IsNullOrWhiteSpace(uid)) return string.Empty;

                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = new SqlCommand(@"
SELECT TOP 1 LTRIM(RTRIM(p.DisplayName))
FROM dbo.UserProfiles p
WHERE p.UserId = @uid;", conn);
                cmd.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 450) { Value = uid });

                var name = cmd.ExecuteScalar() as string;
                return string.IsNullOrWhiteSpace(name) ? string.Empty : name!;
            }
            catch { return string.Empty; }
        }

        public string GetDisplayNameByEmailStrict(string? email)
        {
            if (string.IsNullOrWhiteSpace(email)) return string.Empty;

            try
            {
                var em = email.Trim();
                var cs = _cfg.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = new SqlCommand(@"
SELECT TOP 1 LTRIM(RTRIM(p.DisplayName))
FROM dbo.AspNetUsers u
JOIN dbo.UserProfiles p ON p.UserId = u.Id
WHERE LTRIM(RTRIM(u.Email)) = @em OR u.NormalizedEmail = UPPER(@em);", conn);
                cmd.Parameters.Add(new SqlParameter("@em", SqlDbType.NVarChar, 256) { Value = em });

                var name = cmd.ExecuteScalar() as string;
                return string.IsNullOrWhiteSpace(name) ? string.Empty : name!;
            }
            catch { return string.Empty; }
        }

        public static string FallbackNameFromEmail(string? email)
        {
            if (string.IsNullOrWhiteSpace(email)) return string.Empty;
            var at = email.IndexOf('@');
            return at > 0 ? email[..at] : email;
        }

        public string ComposeAddress(string? email, string? displayName)
        {
            if (string.IsNullOrWhiteSpace(email)) return string.Empty;
            var name = (displayName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(name)) return email.Trim();
            if (email.Contains('<') && email.Contains('>')) return email.Trim();
            return $"{name} <{email.Trim()}>";
        }

        public (string compCd, string? departmentId) GetUserCompDept()
        {
            var comp = _user.FindFirstValue("compCd");
            var dept = _user.FindFirstValue("departmentId");

            if (!string.IsNullOrWhiteSpace(comp))
                return (comp, string.IsNullOrWhiteSpace(dept) ? null : dept);

            try
            {
                var userId = _user.FindFirstValue(ClaimTypes.NameIdentifier);
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
    }
}
