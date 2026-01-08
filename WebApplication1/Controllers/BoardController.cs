// 2026.01.08 Added: Board 화면/데이터/배지 엔드포인트를 DocController에서 분리(라우트는 /Doc/* 그대로 유지)
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
using Microsoft.Extensions.Logging;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class BoardController : Controller
    {
        private readonly IConfiguration _cfg;
        private readonly ILogger<BoardController> _log;

        public BoardController(IConfiguration cfg, ILogger<BoardController> log)
        {
            _cfg = cfg;
            _log = log;
        }

        [HttpGet("Board")]
        public IActionResult Board()
        {
            // 기존 Views/Doc/Board.cshtml 을 그대로 사용
            return View("~/Views/Doc/Board.cshtml");
        }

        [HttpGet("BoardData")]
        // 2025.12.23 Changed: 공유 문서함 조회에 DocumentShares IsRevoked 조건과 만료 조건을 추가하여 철회 및 만료 공유를 제외
        // 2025.12.09 Changed: 결재 문서함 승인 탭에서 선행 단계 보류 반려 문서가 후행 승인자에게 표시되지 않도록 필터 보강
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

            string orderBy = sort switch
            {
                "created_asc" => "ORDER BY d.CreatedAt ASC",
                "title_asc" => "ORDER BY d.TemplateTitle ASC",
                "title_desc" => "ORDER BY d.TemplateTitle DESC",
                _ => "ORDER BY d.CreatedAt DESC"
            };

            string whereSearch = string.IsNullOrWhiteSpace(q) ? "" : " AND d.TemplateTitle LIKE @Q ";

            // 2025.12.08 Changed: 보류/반려 문서가 OnHoldA1 / RejectedA1 처럼 저장된 경우도 필터에 포함되도록 LIKE 로 변경
            string whereTitleFilter = titleFilter?.ToLowerInvariant() switch
            {
                "approved" => " AND ISNULL(d.Status, N'') LIKE N'Approved%' ",
                "rejected" => " AND ISNULL(d.Status, N'') LIKE N'Rejected%' ",
                "pending" => " AND ISNULL(d.Status, N'') LIKE N'Pending%' ",
                "onhold" => " AND ISNULL(d.Status, N'') LIKE N'OnHold%' ",
                "recalled" => " AND ISNULL(d.Status, N'') LIKE N'Recalled%' ",
                _ => ""
            };

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
       (SELECT COUNT(1)
          FROM DocumentApprovals da
         WHERE da.DocId = d.DocId) AS TotalSteps,
       (SELECT COUNT(1)
          FROM DocumentApprovals da
         WHERE da.DocId = d.DocId
           AND (ISNULL(da.Action, N'') = N'Approved'
             OR ISNULL(da.Status, N'') = N'Approved')) AS CompletedSteps
FROM Documents d
LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
WHERE d.CreatedBy = @UserId" + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }
            else if (tab == "approval")
            {
                // 2025.12.08 기존 코드 유지
                string whereApprovalView = (approvalView ?? string.Empty).ToLowerInvariant() switch
                {
                    "pending" =>
                        " AND ISNULL(a.Status, N'') = N'Pending' " +
                        " AND ISNULL(d.Status, N'') = N'PendingA' + CAST(a.StepOrder AS nvarchar(10)) ",
                    "approved" =>
                        " AND (ISNULL(a.Action, N'') = N'Approved' OR ISNULL(a.Status, N'') = N'Approved') ",
                    _ => @"
  AND (
        (
          ISNULL(a.Status, N'') = N'Pending'
          AND ISNULL(d.Status, N'') = N'PendingA' + CAST(a.StepOrder AS nvarchar(10))
        )
        OR
        (
          ISNULL(a.Action, N'') <> N''
          OR (
               ISNULL(a.Status, N'') <> N''
           AND ISNULL(a.Status, N'') <> N'Pending'
             )
        )
      )"
                };

                const string whereRecalledFilter = @"
  AND ISNULL(d.Status, N'') <> N'Recalled'
  AND ISNULL(a.Status, N'') <> N'Recalled'";

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
       (SELECT COUNT(1)
          FROM DocumentApprovals da
         WHERE da.DocId = d.DocId) AS TotalSteps,
       (SELECT COUNT(1)
          FROM DocumentApprovals da
         WHERE da.DocId = d.DocId
           AND (ISNULL(da.Action, N'') = N'Approved'
             OR ISNULL(da.Status, N'') = N'Approved')) AS CompletedSteps
FROM DocumentApprovals a
JOIN Documents d ON a.DocId = d.DocId
LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
WHERE a.UserId = @UserId" + whereRecalledFilter + whereApprovalView + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }
            else // shared
            {
                // 2025.12.23 Changed: 공유 문서함은 철회(IsRevoked=0) + 만료(ExpireAt) 조건을 반드시 포함
                const string whereShareActive = @"
 AND ISNULL(s.IsRevoked, 0) = 0
 AND (s.ExpireAt IS NULL OR s.ExpireAt > SYSUTCDATETIME())";

                sqlCount = @"
SELECT COUNT(1)
FROM DocumentShares s
JOIN Documents d ON s.DocId = d.DocId
WHERE s.UserId = @UserId" + whereShareActive + whereSearch + whereTitleFilter + ";";

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
       (SELECT COUNT(1)
          FROM DocumentApprovals da
         WHERE da.DocId = d.DocId) AS TotalSteps,
       (SELECT COUNT(1)
          FROM DocumentApprovals da
         WHERE da.DocId = d.DocId
           AND (ISNULL(da.Action, N'') = N'Approved'
             OR ISNULL(da.Status, N'') = N'Approved')) AS CompletedSteps
FROM DocumentShares s
JOIN Documents d ON s.DocId = d.DocId
LEFT JOIN UserProfiles up ON d.CreatedBy = up.UserId
WHERE s.UserId = @UserId" + whereShareActive + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
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
            if (!string.IsNullOrWhiteSpace(q))
                cmdList.Parameters.Add(new SqlParameter("@Q", SqlDbType.NVarChar, 200) { Value = $"%{q}%" });
            cmdList.Parameters.Add(new SqlParameter("@Offset", SqlDbType.Int) { Value = offset });
            cmdList.Parameters.Add(new SqlParameter("@PageSize", SqlDbType.Int) { Value = pageSize });

            var items = new List<object>();
            using (var rdr = await cmdList.ExecuteReaderAsync())
            {
                while (await rdr.ReadAsync())
                {
                    var createdAtLocal = (rdr["CreatedAt"] is DateTime utc)
                        ? ToLocalStringFromUtc(utc)
                        : string.Empty;

                    var rawStatus = rdr["Status"]?.ToString() ?? string.Empty;

                    var totalSteps = rdr["TotalSteps"] is int ts ? ts : Convert.ToInt32(rdr["TotalSteps"]);
                    var completedSteps = rdr["CompletedSteps"] is int cs ? cs : Convert.ToInt32(rdr["CompletedSteps"]);
                    var commentCount = rdr["CommentCount"] is int cc ? cc : Convert.ToInt32(rdr["CommentCount"]);

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
                        hasAttachment = Convert.ToInt32(rdr["HasAttachment"]) == 1
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

            // 2025.12.08 Changed: 회수된 문서는 결재 대기 배지에서 제외되도록 Documents 와 조인
            // 2025.12.09 Changed: a.Status=Pending 인 행 중에서 실제 현재 차수(PendingA{StepOrder})에 해당하는 것만 집계
            var sqlApprovalPending = @"
SELECT COUNT(1)
FROM DocumentApprovals a
JOIN Documents d ON a.DocId = d.DocId
WHERE a.UserId = @UserId
  AND ISNULL(a.Status, N'') = N'Pending'
  AND ISNULL(d.Status, N'') = N'PendingA' + CAST(a.StepOrder AS nvarchar(10))
  AND ISNULL(d.Status, N'') <> N'Recalled';";

            var sqlSharedUnread = @"
SELECT COUNT(1)
FROM DocumentShares s
WHERE s.UserId = @UserId
AND NOT EXISTS (
    SELECT 1 
    FROM DocumentAuditLogs l
    WHERE l.DocId = s.DocId AND l.ActorId = @UserId AND l.ActionCode = N'READ'
);";

            int approvalPending = 0;
            int sharedUnread = 0;

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

            // 기존 동작 유지(created=0)
            return Json(new { created = 0, approvalPending, sharedUnread });
        }

        private static string ToLocalStringFromUtc(DateTime utc)
        {
            try
            {
                if (utc.Kind == DateTimeKind.Unspecified)
                    utc = DateTime.SpecifyKind(utc, DateTimeKind.Utc);

                if (utc.Kind == DateTimeKind.Local)
                    return utc.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);

                // Windows: "Korea Standard Time", Linux: "Asia/Seoul"
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
