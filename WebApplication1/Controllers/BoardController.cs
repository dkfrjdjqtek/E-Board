using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Logging;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class BoardController : Controller
    {
        private readonly IConfiguration _cfg;
        private readonly ILogger<BoardController> _log;
        private readonly IStringLocalizer<SharedResource> _S;

        public BoardController(
            IConfiguration cfg,
            ILogger<BoardController> log,
            IStringLocalizer<SharedResource> S)
        {
            _cfg = cfg;
            _log = log;
            _S = S;
        }

        [HttpGet("Board")]
        public IActionResult Board()
        {
            return View("~/Views/Doc/Board.cshtml");
        }

        // 2026.01.19 Changed: shared 탭의 titleFilter를 문서상태가 아닌 열람 미열람 기준으로 동작하도록 BoardData 쿼리 분기 추가
        [HttpGet("BoardData")]
        public async Task<IActionResult> BoardData(
            string tab = "created",
            int page = 1,
            int pageSize = 20,
            string titleFilter = "all",
            string sort = "created_desc",
            string? q = null,
            string approvalView = "all")
        {
            if (page < 1) page = 1;
            if (pageSize < 1 || pageSize > 100) pageSize = 20;

            var userId = User?.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;

            var langCode = (CultureInfo.CurrentUICulture?.TwoLetterISOLanguageName ?? "ko").ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(langCode)) langCode = "ko";

            string orderBy = sort switch
            {
                "created_asc" => "ORDER BY d.CreatedAt ASC",
                "title_asc" => "ORDER BY d.TemplateTitle ASC",
                "title_desc" => "ORDER BY d.TemplateTitle DESC",
                _ => "ORDER BY d.CreatedAt DESC"
            };

            string whereSearch = string.IsNullOrWhiteSpace(q) ? "" : " AND d.TemplateTitle LIKE @Q ";

            // ✅ titleFilter: shared 탭은 "열람/미열람" 기준으로 별도 처리해야 함
            // - 기존 값(approved/rejected/...)은 created/approval에서만 사용
            // - shared에서는 viewed/unviewed(또는 read/unread) 값을 사용하도록 분기
            string whereTitleFilter = "";
            if (string.Equals(tab, "shared", StringComparison.OrdinalIgnoreCase))
            {
                // shared에서는 상태필터를 쓰지 않음(문서 상태가 아니라 열람 여부가 목적)
                whereTitleFilter = "";
            }
            else
            {
                whereTitleFilter = titleFilter?.ToLowerInvariant() switch
                {
                    "approved" => " AND ISNULL(d.Status, N'') LIKE N'Approved%' ",
                    "rejected" => " AND ISNULL(d.Status, N'') LIKE N'Rejected%' ",
                    "pending" => " AND ISNULL(d.Status, N'') LIKE N'Pending%' ",
                    "onhold" => " AND ISNULL(d.Status, N'') LIKE N'OnHold%' ",
                    "recalled" => " AND ISNULL(d.Status, N'') LIKE N'Recalled%' ",
                    _ => ""
                };
            }

            const string outerApplyResultActor = @"
OUTER APPLY (
    SELECT
        CASE
            WHEN ISNULL(d.Status, N'') LIKE N'Recalled%' THEN N'Recalled'
            WHEN ISNULL(d.Status, N'') LIKE N'PendingA%'  THEN N'Pending'
            WHEN ISNULL(d.Status, N'') LIKE N'Rejected%'  THEN N'Rejected'
            WHEN ISNULL(d.Status, N'') LIKE N'OnHold%'    THEN N'OnHold'
            WHEN ISNULL(d.Status, N'') LIKE N'Approved%'  THEN N'Approved'
            ELSE N''
        END AS ResultVerbKey,
        CASE
            WHEN ISNULL(d.Status, N'') LIKE N'Recalled%' THEN COALESCE(upC.DisplayName, N'')
            WHEN ISNULL(d.Status, N'') LIKE N'PendingA%' THEN (
                SELECT TOP (1) COALESCE(upN.DisplayName, N'')
                FROM DocumentApprovals daN
                LEFT JOIN UserProfiles upN ON upN.UserId = daN.UserId
                WHERE daN.DocId = d.DocId
                  AND daN.StepOrder = TRY_CONVERT(int, SUBSTRING(ISNULL(d.Status, N''), 9, 10))
                ORDER BY daN.StepOrder ASC
            )
            WHEN ISNULL(d.Status, N'') LIKE N'Rejected%' THEN (
                SELECT TOP (1) COALESCE(upR.DisplayName, N'')
                FROM DocumentApprovals daR
                LEFT JOIN UserProfiles upR ON upR.UserId = daR.UserId
                WHERE daR.DocId = d.DocId
                  AND (ISNULL(daR.Action, N'') = N'Rejected' OR ISNULL(daR.Status, N'') LIKE N'Rejected%')
                ORDER BY ISNULL(daR.ActedAt, '1900-01-01') DESC, ISNULL(daR.StepOrder, 0) DESC
            )
            WHEN ISNULL(d.Status, N'') LIKE N'OnHold%' THEN (
                SELECT TOP (1) COALESCE(upH.DisplayName, N'')
                FROM DocumentApprovals daH
                LEFT JOIN UserProfiles upH ON upH.UserId = daH.UserId
                WHERE daH.DocId = d.DocId
                  AND (ISNULL(daH.Action, N'') = N'OnHold' OR ISNULL(daH.Status, N'') LIKE N'OnHold%')
                ORDER BY ISNULL(daH.ActedAt, '1900-01-01') DESC, ISNULL(daH.StepOrder, 0) DESC
            )
            WHEN ISNULL(d.Status, N'') LIKE N'Approved%' THEN (
                SELECT TOP (1) COALESCE(upA.DisplayName, N'')
                FROM DocumentApprovals daA
                LEFT JOIN UserProfiles upA ON upA.UserId = daA.UserId
                WHERE daA.DocId = d.DocId
                  AND (ISNULL(daA.Action, N'') = N'Approved' OR ISNULL(daA.Status, N'') = N'Approved')
                ORDER BY ISNULL(daA.ActedAt, '1900-01-01') DESC, ISNULL(daA.StepOrder, 0) DESC
            )
            ELSE N''
        END AS ResultActorName,
        CASE
            WHEN ISNULL(d.Status, N'') LIKE N'Recalled%' THEN COALESCE(pmCLoc.Name, pmC.Name, N'')
            WHEN ISNULL(d.Status, N'') LIKE N'PendingA%' THEN (
                SELECT TOP (1) COALESCE(pmNLoc.Name, pmN.Name, N'')
                FROM DocumentApprovals daN
                LEFT JOIN UserProfiles upN ON upN.UserId = daN.UserId
                LEFT JOIN PositionMasters pmN ON pmN.CompCd = upN.CompCd AND pmN.Id = upN.PositionId
                LEFT JOIN PositionMasterLoc pmNLoc ON pmNLoc.PositionId = pmN.Id AND pmNLoc.LangCode = @LangCode
                WHERE daN.DocId = d.DocId
                  AND daN.StepOrder = TRY_CONVERT(int, SUBSTRING(ISNULL(d.Status, N''), 9, 10))
                ORDER BY daN.StepOrder ASC
            )
            WHEN ISNULL(d.Status, N'') LIKE N'Rejected%' THEN (
                SELECT TOP (1) COALESCE(pmRLoc.Name, pmR.Name, N'')
                FROM DocumentApprovals daR
                LEFT JOIN UserProfiles upR ON upR.UserId = daR.UserId
                LEFT JOIN PositionMasters pmR ON pmR.CompCd = upR.CompCd AND pmR.Id = upR.PositionId
                LEFT JOIN PositionMasterLoc pmRLoc ON pmRLoc.PositionId = pmR.Id AND pmRLoc.LangCode = @LangCode
                WHERE daR.DocId = d.DocId
                  AND (ISNULL(daR.Action, N'') = N'Rejected' OR ISNULL(daR.Status, N'') LIKE N'Rejected%')
                ORDER BY ISNULL(daR.ActedAt, '1900-01-01') DESC, ISNULL(daR.StepOrder, 0) DESC
            )
            WHEN ISNULL(d.Status, N'') LIKE N'OnHold%' THEN (
                SELECT TOP (1) COALESCE(pmHLoc.Name, pmH.Name, N'')
                FROM DocumentApprovals daH
                LEFT JOIN UserProfiles upH ON upH.UserId = daH.UserId
                LEFT JOIN PositionMasters pmH ON pmH.CompCd = upH.CompCd AND pmH.Id = upH.PositionId
                LEFT JOIN PositionMasterLoc pmHLoc ON pmHLoc.PositionId = pmH.Id AND pmHLoc.LangCode = @LangCode
                WHERE daH.DocId = d.DocId
                  AND (ISNULL(daH.Action, N'') = N'OnHold' OR ISNULL(daH.Status, N'') LIKE N'OnHold%')
                ORDER BY ISNULL(daH.ActedAt, '1900-01-01') DESC, ISNULL(daH.StepOrder, 0) DESC
            )
            WHEN ISNULL(d.Status, N'') LIKE N'Approved%' THEN (
                SELECT TOP (1) COALESCE(pmALoc.Name, pmA.Name, N'')
                FROM DocumentApprovals daA
                LEFT JOIN UserProfiles upA ON upA.UserId = daA.UserId
                LEFT JOIN PositionMasters pmA ON pmA.CompCd = upA.CompCd AND pmA.Id = upA.PositionId
                LEFT JOIN PositionMasterLoc pmALoc ON pmALoc.PositionId = pmA.Id AND pmALoc.LangCode = @LangCode
                WHERE daA.DocId = d.DocId
                  AND (ISNULL(daA.Action, N'') = N'Approved' OR ISNULL(daA.Status, N'') = N'Approved')
                ORDER BY ISNULL(daA.ActedAt, '1900-01-01') DESC, ISNULL(daA.StepOrder, 0) DESC
            )
            ELSE N''
        END AS ResultActorPosition
    FROM UserProfiles upC
    LEFT JOIN PositionMasters pmC ON pmC.CompCd = upC.CompCd AND pmC.Id = upC.PositionId
    LEFT JOIN PositionMasterLoc pmCLoc ON pmCLoc.PositionId = pmC.Id AND pmCLoc.LangCode = @LangCode
    WHERE upC.UserId = d.CreatedBy
) rs";

            const string outerApplyIsReadAny = @"
OUTER APPLY (
    SELECT TOP (1) 1 AS HasLog
    FROM DocumentViewLogs v2
    WHERE v2.DocId = d.DocId
      AND v2.ViewerId = @UserId
) vAny";

            const string outerApplyIsReadShared = @"
OUTER APPLY (
    SELECT TOP (1) 1 AS HasLog
    FROM DocumentViewLogs v2
    WHERE v2.DocId = d.DocId
      AND v2.ViewerId = @UserId
      AND ISNULL(v2.ViewerRole, N'') = N'Shared'
) vShared";

            string sqlCount;
            string sqlList;
            var offset = (page - 1) * pageSize;

            if (tab == "created")
            {
                sqlCount = @"
SELECT COUNT(1)
FROM Documents d
WHERE d.CreatedBy = @UserId" + whereSearch + whereTitleFilter + ";";

                sqlList = @"
SELECT d.DocId,
       d.TemplateTitle,
       d.CreatedAt,
       ISNULL(d.Status, N'') AS Status,
       up.DisplayName AS AuthorName,
       (SELECT COUNT(1)
          FROM DocumentComments c
         WHERE c.DocId = d.DocId
           AND c.IsDeleted = 0) AS CommentCount,
       CAST(0 AS bit) AS HasAttachment,
       (SELECT COUNT(1) FROM DocumentApprovals da WHERE da.DocId = d.DocId) AS TotalSteps,
       (SELECT COUNT(1) FROM DocumentApprovals da WHERE da.DocId = d.DocId
           AND (ISNULL(da.Action, N'') = N'Approved' OR ISNULL(da.Status, N'') = N'Approved')) AS CompletedSteps,
       ISNULL(rs.ResultVerbKey, N'') AS ResultVerbKey,
       ISNULL(rs.ResultActorName, N'') AS ResultActorName,
       ISNULL(rs.ResultActorPosition, N'') AS ResultActorPosition,
       CAST(CASE WHEN vAny.HasLog IS NULL THEN 0 ELSE 1 END AS bit) AS IsRead
FROM Documents d
LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
" + outerApplyResultActor + @"
" + outerApplyIsReadAny + @"
WHERE d.CreatedBy = @UserId" + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }
            else if (tab == "approval")
            {
                string whereApprovalView = (approvalView ?? string.Empty).ToLowerInvariant() switch
                {
                    "pending" => " AND ISNULL(d.Status, N'') LIKE N'Pending%' ",
                    "approved" => " AND (ISNULL(a.Action, N'') = N'Approved' OR ISNULL(a.Status, N'') = N'Approved') ",
                    _ => ""
                };

                const string whereRecalledFilter = @"
  AND ISNULL(d.Status, N'') NOT LIKE N'Recalled%'
  AND ISNULL(a.Status, N'') NOT LIKE N'Recalled%'";

                sqlCount = @"
SELECT COUNT(1)
FROM DocumentApprovals a
JOIN Documents d ON a.DocId = d.DocId
WHERE a.UserId = @UserId" + whereRecalledFilter + whereApprovalView + whereSearch + whereTitleFilter + ";";

                sqlList = @"
SELECT d.DocId,
       d.TemplateTitle,
       d.CreatedAt,
       CASE 
           WHEN ISNULL(a.Action, N'') <> N'' THEN a.Action 
           ELSE ISNULL(d.Status, N'') 
       END AS Status,
       up.DisplayName AS AuthorName,
       (SELECT COUNT(1)
          FROM DocumentComments c
         WHERE c.DocId = d.DocId
           AND c.IsDeleted = 0) AS CommentCount,
       CAST(0 AS bit) AS HasAttachment,
       (SELECT COUNT(1) FROM DocumentApprovals da WHERE da.DocId = d.DocId) AS TotalSteps,
       (SELECT COUNT(1) FROM DocumentApprovals da WHERE da.DocId = d.DocId
           AND (ISNULL(da.Action, N'') = N'Approved' OR ISNULL(da.Status, N'') = N'Approved')) AS CompletedSteps,
       ISNULL(rs.ResultVerbKey, N'') AS ResultVerbKey,
       ISNULL(rs.ResultActorName, N'') AS ResultActorName,
       ISNULL(rs.ResultActorPosition, N'') AS ResultActorPosition,
       CAST(CASE WHEN vAny.HasLog IS NULL THEN 0 ELSE 1 END AS bit) AS IsRead
FROM DocumentApprovals a
JOIN Documents d ON a.DocId = d.DocId
LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
" + outerApplyResultActor + @"
" + outerApplyIsReadAny + @"
WHERE a.UserId = @UserId" + whereRecalledFilter + whereApprovalView + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }
            else
            {
                const string whereShareActive = @"
 AND ISNULL(s.IsRevoked, 0) = 0
 AND (s.ExpireAt IS NULL OR s.ExpireAt > SYSUTCDATETIME())";

                // ✅ shared 열람 미열람 필터(Count용: EXISTS/NOT EXISTS)
                string whereSharedReadCount = "";
                var tf = (titleFilter ?? string.Empty).Trim().ToLowerInvariant();
                if (tf == "viewed" || tf == "read")
                {
                    whereSharedReadCount = @"
 AND EXISTS (
     SELECT 1
     FROM DocumentViewLogs v
     WHERE v.DocId = s.DocId
       AND v.ViewerId = @UserId
       AND ISNULL(v.ViewerRole, N'') = N'Shared'
 )";
                }
                else if (tf == "unviewed" || tf == "unread")
                {
                    whereSharedReadCount = @"
 AND NOT EXISTS (
     SELECT 1
     FROM DocumentViewLogs v
     WHERE v.DocId = s.DocId
       AND v.ViewerId = @UserId
       AND ISNULL(v.ViewerRole, N'') = N'Shared'
 )";
                }

                // ✅ shared 열람 미열람 필터(List용: OUTER APPLY vShared 결과 사용)
                string whereSharedReadList = "";
                if (tf == "viewed" || tf == "read")
                    whereSharedReadList = " AND vShared.HasLog IS NOT NULL ";
                else if (tf == "unviewed" || tf == "unread")
                    whereSharedReadList = " AND vShared.HasLog IS NULL ";

                sqlCount = @"
SELECT COUNT(1)
FROM DocumentShares s
JOIN Documents d ON s.DocId = d.DocId
WHERE s.UserId = @UserId" + whereShareActive + whereSharedReadCount + whereSearch + whereTitleFilter + ";";

                sqlList = @"
SELECT d.DocId,
       d.TemplateTitle,
       d.CreatedAt,
       ISNULL(d.Status, N'') AS Status,
       up.DisplayName AS AuthorName,
       (SELECT COUNT(1)
          FROM DocumentComments c
         WHERE c.DocId = d.DocId
           AND c.IsDeleted = 0) AS CommentCount,
       CAST(0 AS bit) AS HasAttachment,
       (SELECT COUNT(1) FROM DocumentApprovals da WHERE da.DocId = d.DocId) AS TotalSteps,
       (SELECT COUNT(1) FROM DocumentApprovals da WHERE da.DocId = d.DocId
           AND (ISNULL(da.Action, N'') = N'Approved' OR ISNULL(da.Status, N'') = N'Approved')) AS CompletedSteps,
       ISNULL(rs.ResultVerbKey, N'') AS ResultVerbKey,
       ISNULL(rs.ResultActorName, N'') AS ResultActorName,
       ISNULL(rs.ResultActorPosition, N'') AS ResultActorPosition,
       CAST(CASE WHEN vShared.HasLog IS NULL THEN 0 ELSE 1 END AS bit) AS IsRead
FROM DocumentShares s
JOIN Documents d ON s.DocId = d.DocId
LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
" + outerApplyResultActor + @"
" + outerApplyIsReadShared + @"
WHERE s.UserId = @UserId" + whereShareActive + whereSharedReadList + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }

            string TemplateKey(string verbKey)
            {
                switch ((verbKey ?? string.Empty).Trim())
                {
                    case "Recalled": return "DOC_RS_Recalled";
                    case "Pending": return "DOC_RS_Pending";
                    case "Rejected": return "DOC_RS_Rejected";
                    case "OnHold": return "DOC_RS_OnHold";
                    case "Approved": return "DOC_RS_ApprovedDone";
                    default: return "DOC_RS_Fallback";
                }
            }

            string BuildResultSummarySentence(string verbKey, string actorName, string actorPos)
            {
                actorName = (actorName ?? string.Empty).Trim();
                actorPos = (actorPos ?? string.Empty).Trim();

                if (string.IsNullOrWhiteSpace(actorName) && string.IsNullOrWhiteSpace(actorPos))
                    return string.Empty;

                var key = TemplateKey(verbKey);
                var fmt = _S[key].Value;

                if (string.IsNullOrWhiteSpace(fmt) || string.Equals(fmt, key, StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrWhiteSpace(actorName) && !string.IsNullOrWhiteSpace(actorPos))
                        return actorName + " " + actorPos;
                    return !string.IsNullOrWhiteSpace(actorName) ? actorName : actorPos;
                }

                return string.Format(fmt, actorName, actorPos).Trim();
            }

            var connStr = _cfg.GetConnectionString("DefaultConnection");
            using var conn = new SqlConnection(connStr);
            await conn.OpenAsync();

            using var cmdCount = new SqlCommand(sqlCount, conn);
            cmdCount.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
            if (!string.IsNullOrWhiteSpace(q))
                cmdCount.Parameters.Add(new SqlParameter("@Q", SqlDbType.NVarChar, 200) { Value = $"%{q}%" });

            var total = Convert.ToInt32(await cmdCount.ExecuteScalarAsync());

            using var cmdList = new SqlCommand(sqlList, conn);
            cmdList.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
            cmdList.Parameters.Add(new SqlParameter("@LangCode", SqlDbType.NVarChar, 10) { Value = langCode });
            if (!string.IsNullOrWhiteSpace(q))
                cmdList.Parameters.Add(new SqlParameter("@Q", SqlDbType.NVarChar, 200) { Value = $"%{q}%" });
            cmdList.Parameters.Add(new SqlParameter("@Offset", SqlDbType.Int) { Value = offset });
            cmdList.Parameters.Add(new SqlParameter("@PageSize", SqlDbType.Int) { Value = pageSize });

            var items = new List<object>();
            using (var rdr = await cmdList.ExecuteReaderAsync())
            {
                int ordIsRead = -1;
                try { ordIsRead = rdr.GetOrdinal("IsRead"); } catch { ordIsRead = -1; }

                while (await rdr.ReadAsync())
                {
                    var createdAtLocal = (rdr["CreatedAt"] is DateTime utc)
                        ? ToLocalStringFromUtc(utc)
                        : string.Empty;

                    var rawStatus = rdr["Status"]?.ToString() ?? string.Empty;

                    var totalSteps = rdr["TotalSteps"] is int ts ? ts : Convert.ToInt32(rdr["TotalSteps"]);
                    var completedSteps = rdr["CompletedSteps"] is int cs ? cs : Convert.ToInt32(rdr["CompletedSteps"]);
                    var commentCount = rdr["CommentCount"] is int cc ? cc : Convert.ToInt32(rdr["CommentCount"]);

                    var verbKey = rdr["ResultVerbKey"]?.ToString() ?? string.Empty;
                    var actorName = rdr["ResultActorName"]?.ToString() ?? string.Empty;
                    var actorPos = rdr["ResultActorPosition"]?.ToString() ?? string.Empty;

                    var resultSummary = BuildResultSummarySentence(verbKey, actorName, actorPos);

                    bool? isRead = null;
                    if (ordIsRead >= 0 && rdr[ordIsRead] != DBNull.Value)
                        isRead = Convert.ToInt32(rdr[ordIsRead]) == 1;

                    items.Add(new
                    {
                        docId = rdr["DocId"]?.ToString() ?? string.Empty,
                        templateTitle = rdr["TemplateTitle"]?.ToString() ?? string.Empty,
                        authorName = rdr["AuthorName"]?.ToString() ?? string.Empty,
                        createdAt = createdAtLocal,
                        status = rawStatus,
                        statusCode = rawStatus,
                        totalApprovers = totalSteps,
                        completedApprovers = completedSteps,
                        commentCount = commentCount,
                        hasAttachment = Convert.ToInt32(rdr["HasAttachment"]) == 1,
                        resultSummary = resultSummary,
                        isRead = isRead
                    });
                }
            }

            return Json(new { total, page, pageSize, items });
        }


        [HttpGet("BoardBadges")]
        public async Task<IActionResult> BoardBadges()
        {
            var userId = User?.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
            var connStr = _cfg.GetConnectionString("DefaultConnection");
            using var conn = new SqlConnection(connStr);
            await conn.OpenAsync();

            var sqlApprovalPending = @"
SELECT COUNT(1)
FROM DocumentApprovals a
JOIN Documents d ON a.DocId = d.DocId
WHERE a.UserId = @UserId
  AND ISNULL(a.Status, N'') = N'Pending'
  AND ISNULL(d.Status, N'') = N'PendingA' + CAST(a.StepOrder AS nvarchar(10))
  AND ISNULL(d.Status, N'') NOT LIKE N'Recalled%';";

            // created: 내가 만든 문서 중 열람로그가 한 번도 없으면 미열람으로 계산
            var sqlCreatedUnread = @"
SELECT COUNT(1)
FROM Documents d
WHERE d.CreatedBy = @UserId
  AND NOT EXISTS (
      SELECT 1
      FROM DocumentViewLogs v
      WHERE v.DocId = d.DocId
        AND v.ViewerId = @UserId
  );";

            // shared: 공유 문서함 미열람은 ViewerRole=Shared 열람로그만 기준으로 계산
            var sqlSharedUnread = @"
SELECT COUNT(1)
FROM DocumentShares s
WHERE s.UserId = @UserId
  AND ISNULL(s.IsRevoked, 0) = 0
  AND (s.ExpireAt IS NULL OR s.ExpireAt > SYSUTCDATETIME())
  AND NOT EXISTS (
      SELECT 1
      FROM DocumentViewLogs v
      WHERE v.DocId = s.DocId
        AND v.ViewerId = @UserId
        AND ISNULL(v.ViewerRole, N'') = N'Shared'
  );";

            int approvalPending = 0;
            int sharedUnread = 0;
            int createdUnread = 0;

            using (var cmd = new SqlCommand(sqlApprovalPending, conn))
            {
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
                approvalPending = Convert.ToInt32(await cmd.ExecuteScalarAsync());
            }

            using (var cmd = new SqlCommand(sqlSharedUnread, conn))
            {
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
                sharedUnread = Convert.ToInt32(await cmd.ExecuteScalarAsync());
            }

            using (var cmd = new SqlCommand(sqlCreatedUnread, conn))
            {
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
                createdUnread = Convert.ToInt32(await cmd.ExecuteScalarAsync());
            }

            return Json(new { created = createdUnread, approvalPending, sharedUnread });
        }

        private static string ToLocalStringFromUtc(DateTime utc)
        {
            try
            {
                if (utc.Kind == DateTimeKind.Unspecified)
                    utc = DateTime.SpecifyKind(utc, DateTimeKind.Utc);

                if (utc.Kind == DateTimeKind.Local)
                    return utc.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

                TimeZoneInfo tz;
                try { tz = TimeZoneInfo.FindSystemTimeZoneById("Korea Standard Time"); }
                catch { tz = TimeZoneInfo.FindSystemTimeZoneById("Asia/Seoul"); }

                var local = TimeZoneInfo.ConvertTimeFromUtc(DateTime.SpecifyKind(utc, DateTimeKind.Utc), tz);
                return local.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
            }
            catch
            {
                return utc.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
            }
        }
    }
}