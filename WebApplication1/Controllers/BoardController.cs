using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Logging;
using NuGet.Packaging.Signing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using WebApplication1.Services;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class BoardController : Controller
    {
        private readonly IConfiguration _cfg;
        private readonly ILogger<BoardController> _log;
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IWebHostEnvironment _env;
        private readonly IWebPushNotifier _webPushNotifier;

        public sealed class BulkApproveDto
        {
            public List<string>? DocIds { get; set; }
        }

        private sealed class BulkApproveItemResult
        {
            public string DocId { get; set; } = "";
            public string Result { get; set; } = "";
            public string? Status { get; set; }
            public string? Reason { get; set; }
        }

        public BoardController(
            IConfiguration cfg,
            ILogger<BoardController> log,
            IStringLocalizer<SharedResource> S,
            IWebHostEnvironment env,
            IWebPushNotifier webPushNotifier)
        {
            _cfg = cfg;
            _log = log;
            _S = S;
            _env = env;
            _webPushNotifier = webPushNotifier;
        }

        [HttpGet("Board")]
        public IActionResult Board()
        {
            return View("~/Views/Doc/Board.cshtml");
        }

        [HttpPost("BulkApprove")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> BulkApprove([FromBody] BulkApproveDto dto)
        {
            var DocIds = dto?.DocIds?
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList() ?? new List<string>();

            if (DocIds.Count == 0)
                return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" } });

            if (DocIds.Count > 50)
                return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" } });

            var approverId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            if (string.IsNullOrWhiteSpace(approverId))
                return Forbid();

            var approverName =
                (User?.FindFirstValue("name") ?? User?.FindFirstValue(ClaimTypes.Name) ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(approverName))
                approverName = approverId;

            string? approverDisplayText = null;
            string? signatureRelativePath = null;

            try
            {
                var csProfile = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using (var connProfile = new SqlConnection(csProfile))
                {
                    await connProfile.OpenAsync();

                    await using var up = connProfile.CreateCommand();
                    up.CommandText = @"
SELECT TOP (1)
       u.DisplayName,
       COALESCE(pl.Name, pm.Name, N'') AS PositionName,
       COALESCE(dl.Name, dm.Name, N'') AS DepartmentName,
       u.SignatureRelativePath AS SignaturePath
FROM dbo.UserProfiles AS u
LEFT JOIN dbo.DepartmentMasters     AS dm ON dm.CompCd = u.CompCd AND dm.Id       = u.DepartmentId
LEFT JOIN dbo.DepartmentMasterLoc   AS dl ON dl.DepartmentId = dm.Id AND dl.LangCode = N'ko'
LEFT JOIN dbo.PositionMasters       AS pm ON pm.CompCd = u.CompCd AND pm.Id       = u.PositionId
LEFT JOIN dbo.PositionMasterLoc     AS pl ON pl.PositionId   = pm.Id AND pl.LangCode = N'ko'
WHERE u.UserId = @uid;";
                    up.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });

                    await using var r = await up.ExecuteReaderAsync();
                    if (await r.ReadAsync())
                    {
                        var disp = r["DisplayName"] as string ?? string.Empty;
                        var pos = r["PositionName"] as string ?? string.Empty;
                        var dept = r["DepartmentName"] as string ?? string.Empty;
                        var sig = r["SignaturePath"] as string ?? string.Empty;

                        approverDisplayText = string.Join(" ",
                            new[] { dept, pos, disp }.Where(s => !string.IsNullOrWhiteSpace(s)));

                        signatureRelativePath = string.IsNullOrWhiteSpace(sig) ? null : sig;
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "BulkApprove: 프로필 스냅샷 조회 실패 uid={uid}", approverId);
            }

            var clientIp = HttpContext?.Connection?.RemoteIpAddress?.ToString();
            var userAgent = Request?.Headers["User-Agent"].ToString();
            if (!string.IsNullOrWhiteSpace(userAgent) && userAgent.Length > 1024)
                userAgent = userAgent.Substring(0, 1024);

            var results = new List<BulkApproveItemResult>(DocIds.Count);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            static async Task InsertViewLogIfMissingAsync(SqlConnection conn2, string docId2, string viewerId2, string role2, string? ip2, string? ua2)
            {
                await using var cmd = conn2.CreateCommand();
                cmd.CommandText = @"
IF NOT EXISTS (
    SELECT 1
    FROM dbo.DocumentViewLogs v
    WHERE v.DocId = @DocId
      AND v.ViewerId = TRY_CONVERT(uniqueidentifier, @ViewerId)
)
BEGIN
    INSERT INTO dbo.DocumentViewLogs (DocId, ViewerId, ViewerRole, ViewedAt, ClientIp, UserAgent)
    VALUES (@DocId, TRY_CONVERT(uniqueidentifier, @ViewerId), @ViewerRole, SYSUTCDATETIME(), @ClientIp, @UserAgent);
END";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 100) { Value = docId2 });
                cmd.Parameters.Add(new SqlParameter("@ViewerId", SqlDbType.NVarChar, 64) { Value = viewerId2 });
                cmd.Parameters.Add(new SqlParameter("@ViewerRole", SqlDbType.NVarChar, 40) { Value = role2 });
                cmd.Parameters.Add(new SqlParameter("@ClientIp", SqlDbType.NVarChar, 128) { Value = (object?)ip2 ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@UserAgent", SqlDbType.NVarChar, 1024) { Value = (object?)ua2 ?? DBNull.Value });

                await cmd.ExecuteNonQueryAsync();
            }

            var nextApproverIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            async Task<int> GetApprovalPendingCountAsync(string targetUserId)
            {
                await using var connCnt = new SqlConnection(cs);
                await connCnt.OpenAsync();

                var sqlCnt = @"
SELECT COUNT(1)
FROM DocumentApprovals a
JOIN Documents d ON a.DocId = d.DocId
WHERE a.UserId = @UserId
  AND ISNULL(a.Status, N'') = N'Pending'
  AND ISNULL(d.Status, N'') = N'PendingA' + CAST(a.StepOrder AS nvarchar(10))
  AND ISNULL(d.Status, N'') NOT LIKE N'Recalled%';";

                await using var cmdCnt = new SqlCommand(sqlCnt, connCnt);
                cmdCnt.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                return Convert.ToInt32(await cmdCnt.ExecuteScalarAsync());
            }

            foreach (var docId in DocIds)
            {
                var item = new BulkApproveItemResult { DocId = docId };

                try
                {
                    string outXlsx = string.Empty;
                    await using (var cmdOut = conn.CreateCommand())
                    {
                        cmdOut.CommandText = @"
SELECT TOP (1) OutputPath
FROM dbo.Documents
WHERE DocId = @DocId;";
                        cmdOut.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

                        var obj = await cmdOut.ExecuteScalarAsync();
                        outXlsx = (obj as string) ?? string.Empty;
                    }

                    if (!string.IsNullOrWhiteSpace(outXlsx) && !System.IO.Path.IsPathRooted(outXlsx))
                    {
                        outXlsx = System.IO.Path.Combine(
                            _env.ContentRootPath,
                            outXlsx.TrimStart('\\', '/')
                        );
                    }

                    if (string.IsNullOrWhiteSpace(outXlsx) || !System.IO.File.Exists(outXlsx))
                    {
                        item.Result = "skipped";
                        item.Reason = "not_found";
                        results.Add(item);
                        continue;
                    }

                    int? currentStep = null;
                    await using (var findStep = conn.CreateCommand())
                    {
                        findStep.CommandText = @"
SELECT TOP (1) StepOrder
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND ISNULL(Status,N'Pending') LIKE N'Pending%'
  AND (
        UserId = @UserId
        OR ApproverValue = @UserId
      )
ORDER BY StepOrder;";
                        findStep.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        findStep.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = approverId });

                        var stepObj = await findStep.ExecuteScalarAsync();
                        if (stepObj != null && stepObj != DBNull.Value)
                            currentStep = Convert.ToInt32(stepObj);
                    }

                    if (currentStep == null)
                    {
                        item.Result = "skipped";
                        item.Reason = "not_my_turn";
                        results.Add(item);
                        continue;
                    }

                    var step = currentStep.Value;

                    await using (var u = conn.CreateCommand())
                    {
                        u.CommandText = @"
UPDATE dbo.DocumentApprovals
SET Action              = N'approve',
    ActedAt             = SYSUTCDATETIME(),
    ActorName           = @actor,
    Status              = N'Approved',
    UserId              = COALESCE(UserId, @uid),
    ApproverDisplayText = COALESCE(@displayText, ApproverDisplayText),
    SignaturePath       = COALESCE(@sigPath, SignaturePath)
WHERE DocId = @id AND StepOrder = @step;";
                        u.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = approverName });
                        u.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });
                        u.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                        u.Parameters.Add(new SqlParameter("@step", SqlDbType.Int) { Value = step });
                        u.Parameters.Add(new SqlParameter("@displayText", SqlDbType.NVarChar, 400) { Value = (object?)approverDisplayText ?? DBNull.Value });
                        u.Parameters.Add(new SqlParameter("@sigPath", SqlDbType.NVarChar, 400) { Value = (object?)signatureRelativePath ?? DBNull.Value });

                        var affected = await u.ExecuteNonQueryAsync();
                        if (affected <= 0)
                        {
                            item.Result = "failed";
                            item.Reason = "bad_request";
                            results.Add(item);
                            continue;
                        }
                    }

                    var next = step + 1;

                    var hasNext = false;
                    await using (var chk = conn.CreateCommand())
                    {
                        chk.CommandText = @"
SELECT TOP (1) 1
FROM dbo.DocumentApprovals
WHERE DocId = @DocId AND StepOrder = @NextStep;";
                        chk.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        chk.Parameters.Add(new SqlParameter("@NextStep", SqlDbType.Int) { Value = next });

                        var o = await chk.ExecuteScalarAsync();
                        hasNext = (o != null && o != DBNull.Value);
                    }

                    var newStatus = hasNext ? $"PendingA{next}" : "Approved";

                    await using (var u2 = conn.CreateCommand())
                    {
                        u2.CommandText = @"UPDATE dbo.Documents SET Status = @st WHERE DocId = @id;";
                        u2.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = newStatus });
                        u2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                        await u2.ExecuteNonQueryAsync();
                    }

                    await InsertViewLogIfMissingAsync(conn, docId, approverId, "Approval", clientIp, userAgent);

                    if (hasNext)
                    {
                        await using var cmdNext = conn.CreateCommand();
                        cmdNext.CommandText = @"
SELECT DISTINCT a.UserId
FROM dbo.DocumentApprovals a
WHERE a.DocId = @DocId
  AND a.StepOrder = @NextStep
  AND ISNULL(a.UserId, N'') <> N'';";
                        cmdNext.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                        cmdNext.Parameters.Add(new SqlParameter("@NextStep", SqlDbType.Int) { Value = next });

                        await using var rr = await cmdNext.ExecuteReaderAsync();
                        while (await rr.ReadAsync())
                        {
                            var uid = (rr[0] as string ?? string.Empty).Trim();
                            if (!string.IsNullOrWhiteSpace(uid))
                                nextApproverIds.Add(uid);
                        }
                    }

                    item.Result = "approved";
                    item.Status = newStatus;
                    results.Add(item);
                }
                catch (Exception ex)
                {
                    _log.LogError(ex, "BulkApprove: failed docId={docId}", docId);
                    item.Result = "failed";
                    item.Reason = "exception";
                    results.Add(item);
                }
            }

            try
            {
                foreach (var uid in nextApproverIds)
                {
                    var n = await GetApprovalPendingCountAsync(uid);

                    await _webPushNotifier.SendToUserIdAsync(
                        userId: uid,
                        title: "E-BOARD",
                        body: $"{n}개의 문서가 결재 대기 중입니다.",
                        url: "/",
                        tag: "badge-approval-pending"
                    );
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "BulkApprove: WebPush notify(next approver) failed");
            }

            var approved = results.Count(x => x.Result == "approved");
            var skipped = results.Count(x => x.Result == "skipped");
            var failed = results.Count(x => x.Result == "failed");

            return Json(new
            {
                ok = true,
                totals = new { approved, skipped, failed },
                results
            });
        }
                
        // 2026.04.10 Changed: 상신 결재 협조 공유 문서함의 진행 완료 판정을 문서 전체 기준(회수 반려 전체완료 보류 포함)으로 통일하고 협조 탭의 회수 문서 및 협조 반려를 정상 반영하도록 수정
        [HttpGet("BoardData")]
        public async Task<IActionResult> BoardData(string tab = "created", int page = 1, int pageSize = 20, string titleFilter = "all", string sort = "created_desc", string? q = null, string approvalView = "all", string approvalSub = "ongoing", string createdSub = "ongoing", string cooperationSub = "ongoing", string sharedSub = "ongoing")
        {
            if (page < 1) page = 1;
            if (pageSize < 1 || pageSize > 100) pageSize = 20;

            var userId = User?.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;

            var langCode = (CultureInfo.CurrentUICulture?.TwoLetterISOLanguageName ?? "ko").ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(langCode)) langCode = "ko";

            var tabKey = (tab ?? "created").Trim().ToLowerInvariant();

            string orderBy = sort switch
            {
                "created_asc" => "ORDER BY d.CreatedAt ASC",
                "title_asc" => "ORDER BY d.TemplateTitle ASC",
                "title_desc" => "ORDER BY d.TemplateTitle DESC",
                _ => "ORDER BY d.CreatedAt DESC"
            };

            string whereSearch = string.IsNullOrWhiteSpace(q) ? "" : " AND d.TemplateTitle LIKE @Q ";

            string whereTitleFilter = "";
            if (!string.Equals(tabKey, "shared", StringComparison.OrdinalIgnoreCase))
            {
                whereTitleFilter = titleFilter?.ToLowerInvariant() switch
                {
                    "approved" => " AND docStage.IsFullyApproved = 1 ",
                    "rejected" => " AND (docStage.IsApprovalRejected = 1 OR docStage.IsCoopRejected = 1) ",
                    "pending" => " AND docStage.IsOngoing = 1 ",
                    "onhold" => " AND docStage.IsOnHold = 1 ",
                    "recalled" => " AND docStage.IsRecalled = 1 ",
                    _ => ""
                };
            }

            const string outerApplyStepAgg = @"
OUTER APPLY(
    SELECT
        COUNT(1) AS TotalSteps,
        SUM(CASE WHEN ISNULL(da.Action,N'') = N'approve' OR ISNULL(da.Status,N'') = N'Approved' THEN 1 ELSE 0 END) AS DoneSteps,
        MAX(CASE WHEN ISNULL(da.Action,N'') = N'reject' OR ISNULL(da.Status,N'') LIKE N'Rejected%' THEN 1 ELSE 0 END) AS HasRejected
    FROM dbo.DocumentApprovals da
    WHERE da.DocId = d.DocId
) stepAgg";

            const string outerApplyCoopAgg = @"
OUTER APPLY(
    SELECT
        COUNT(1) AS CoopTotal,
        SUM(CASE
                WHEN ISNULL(cc2.Status,N'') = N'Cooperated'
                  OR ISNULL(cc2.Action,N'') = N'Cooperate'
                THEN 1 ELSE 0
            END) AS CoopDoneCount,
        SUM(CASE
                WHEN ISNULL(cc2.Status,N'') LIKE N'Rejected%'
                  OR ISNULL(cc2.Action,N'') IN (N'Reject', N'Rejected')
                THEN 1 ELSE 0
            END) AS CoopRejectedCount,
        (
            SELECT STRING_AGG(CONVERT(nvarchar(20), TRY_CONVERT(int, SUBSTRING(cc4.RoleKey, 2, 10))), ',')
            FROM dbo.DocumentCooperations cc4
            WHERE cc4.DocId = d.DocId
              AND ISNULL(cc4.Status, N'') <> N'Recalled'
              AND ISNULL(cc4.Action, N'') <> N'Recalled'
              AND (
                    ISNULL(cc4.Status, N'') = N'Cooperated'
                 OR ISNULL(cc4.Action, N'') = N'Cooperate'
              )
        ) AS CoopDoneKeys,
        (
            SELECT STRING_AGG(CONVERT(nvarchar(20), TRY_CONVERT(int, SUBSTRING(cc5.RoleKey, 2, 10))), ',')
            FROM dbo.DocumentCooperations cc5
            WHERE cc5.DocId = d.DocId
              AND ISNULL(cc5.Status, N'') <> N'Recalled'
              AND ISNULL(cc5.Action, N'') <> N'Recalled'
              AND (
                    ISNULL(cc5.Status, N'') LIKE N'Rejected%'
                 OR ISNULL(cc5.Action, N'') IN (N'Reject', N'Rejected')
              )
        ) AS CoopRejectedKeys,
        (
            SELECT TOP(1) COALESCE(upCo.DisplayName, N'')
            FROM dbo.DocumentCooperations cc3
            LEFT JOIN dbo.UserProfiles upCo ON upCo.UserId = cc3.UserId
            WHERE cc3.DocId = d.DocId
              AND ISNULL(cc3.Status, N'') NOT IN (N'Cooperated', N'Recalled', N'Rejected')
              AND ISNULL(cc3.Action, N'') NOT IN (N'Cooperate', N'Recalled', N'Reject', N'Rejected')
            ORDER BY cc3.RoleKey ASC
        ) AS CoopPendingName,
        (
            SELECT TOP(1) COALESCE(pmCoLoc.Name, pmCo.Name, N'')
            FROM dbo.DocumentCooperations cc3
            LEFT JOIN dbo.UserProfiles upCo ON upCo.UserId = cc3.UserId
            LEFT JOIN dbo.PositionMasters pmCo ON pmCo.CompCd = upCo.CompCd AND pmCo.Id = upCo.PositionId
            LEFT JOIN dbo.PositionMasterLoc pmCoLoc ON pmCoLoc.PositionId = pmCo.Id AND pmCoLoc.LangCode = @LangCode
            WHERE cc3.DocId = d.DocId
              AND ISNULL(cc3.Status, N'') NOT IN (N'Cooperated', N'Recalled', N'Rejected')
              AND ISNULL(cc3.Action, N'') NOT IN (N'Cooperate', N'Recalled', N'Reject', N'Rejected')
            ORDER BY cc3.RoleKey ASC
        ) AS CoopPendingPosition
    FROM dbo.DocumentCooperations cc2
    WHERE cc2.DocId = d.DocId
      AND ISNULL(cc2.Status, N'') <> N'Recalled'
      AND ISNULL(cc2.Action, N'') <> N'Recalled'
) coopAgg";

            const string outerApplyCoopAggEmpty = @"
OUTER APPLY(
    SELECT
        CAST(0 AS int)             AS CoopTotal,
        CAST(0 AS int)             AS CoopDoneCount,
        CAST(0 AS int)             AS CoopRejectedCount,
        CAST(N'' AS nvarchar(max)) AS CoopDoneKeys,
        CAST(N'' AS nvarchar(max)) AS CoopRejectedKeys,
        CAST(N'' AS nvarchar(200)) AS CoopPendingName,
        CAST(N'' AS nvarchar(200)) AS CoopPendingPosition
) coopAgg";

            const string outerApplyDocStage = @"
OUTER APPLY(
    SELECT
        CASE WHEN ISNULL(d.Status, N'') LIKE N'Recalled%' OR ISNULL(d.Status, N'') LIKE N'Recall%' THEN 1 ELSE 0 END AS IsRecalled,
        CASE WHEN ISNULL(d.Status, N'') LIKE N'PendingHold%' OR ISNULL(d.Status, N'') LIKE N'OnHold%' THEN 1 ELSE 0 END AS IsOnHold,
        CASE WHEN ISNULL(stepAgg.HasRejected, 0) > 0 OR ISNULL(d.Status, N'') LIKE N'Rejected%' THEN 1 ELSE 0 END AS IsApprovalRejected,
        CASE WHEN ISNULL(coopAgg.CoopRejectedCount, 0) > 0 THEN 1 ELSE 0 END AS IsCoopRejected,
        CASE
            WHEN ISNULL(stepAgg.TotalSteps, 0) = 0 THEN 1
            WHEN ISNULL(stepAgg.DoneSteps, 0) >= ISNULL(stepAgg.TotalSteps, 0) THEN 1
            ELSE 0
        END AS IsApprovalDone,
        CASE
            WHEN ISNULL(coopAgg.CoopTotal, 0) = 0 THEN 1
            WHEN ISNULL(coopAgg.CoopDoneCount, 0) >= ISNULL(coopAgg.CoopTotal, 0) THEN 1
            ELSE 0
        END AS IsCoopDone,
        CASE
            WHEN ISNULL(d.Status, N'') LIKE N'Recalled%' OR ISNULL(d.Status, N'') LIKE N'Recall%' THEN 0
            WHEN ISNULL(stepAgg.HasRejected, 0) > 0 OR ISNULL(d.Status, N'') LIKE N'Rejected%' THEN 0
            WHEN ISNULL(coopAgg.CoopRejectedCount, 0) > 0 THEN 0
            WHEN
                (
                    CASE
                        WHEN ISNULL(stepAgg.TotalSteps, 0) = 0 THEN 1
                        WHEN ISNULL(stepAgg.DoneSteps, 0) >= ISNULL(stepAgg.TotalSteps, 0) THEN 1
                        ELSE 0
                    END
                ) = 1
             AND
                (
                    CASE
                        WHEN ISNULL(coopAgg.CoopTotal, 0) = 0 THEN 1
                        WHEN ISNULL(coopAgg.CoopDoneCount, 0) >= ISNULL(coopAgg.CoopTotal, 0) THEN 1
                        ELSE 0
                    END
                ) = 1
            THEN 1
            ELSE 0
        END AS IsFullyApproved,
        CASE
            WHEN ISNULL(d.Status, N'') LIKE N'Recalled%' OR ISNULL(d.Status, N'') LIKE N'Recall%' THEN 1
            WHEN ISNULL(stepAgg.HasRejected, 0) > 0 OR ISNULL(d.Status, N'') LIKE N'Rejected%' THEN 1
            WHEN ISNULL(coopAgg.CoopRejectedCount, 0) > 0 THEN 1
            WHEN
                (
                    CASE
                        WHEN ISNULL(stepAgg.TotalSteps, 0) = 0 THEN 1
                        WHEN ISNULL(stepAgg.DoneSteps, 0) >= ISNULL(stepAgg.TotalSteps, 0) THEN 1
                        ELSE 0
                    END
                ) = 1
             AND
                (
                    CASE
                        WHEN ISNULL(coopAgg.CoopTotal, 0) = 0 THEN 1
                        WHEN ISNULL(coopAgg.CoopDoneCount, 0) >= ISNULL(coopAgg.CoopTotal, 0) THEN 1
                        ELSE 0
                    END
                ) = 1
            THEN 1
            ELSE 0
        END AS IsCompleted,
        CASE
            WHEN ISNULL(d.Status, N'') LIKE N'Recalled%' OR ISNULL(d.Status, N'') LIKE N'Recall%' THEN 0
            WHEN ISNULL(stepAgg.HasRejected, 0) > 0 OR ISNULL(d.Status, N'') LIKE N'Rejected%' THEN 0
            WHEN ISNULL(coopAgg.CoopRejectedCount, 0) > 0 THEN 0
            WHEN
                (
                    CASE
                        WHEN ISNULL(stepAgg.TotalSteps, 0) = 0 THEN 1
                        WHEN ISNULL(stepAgg.DoneSteps, 0) >= ISNULL(stepAgg.TotalSteps, 0) THEN 1
                        ELSE 0
                    END
                ) = 1
             AND
                (
                    CASE
                        WHEN ISNULL(coopAgg.CoopTotal, 0) = 0 THEN 1
                        WHEN ISNULL(coopAgg.CoopDoneCount, 0) >= ISNULL(coopAgg.CoopTotal, 0) THEN 1
                        ELSE 0
                    END
                ) = 1
            THEN 0
            ELSE 1
        END AS IsOngoing
) docStage";

            const string outerApplyResultActor = @"
OUTER APPLY(
    SELECT
        CASE
            WHEN docStage.IsRecalled = 1 THEN N'Recalled'
            WHEN docStage.IsOnHold = 1 THEN N'OnHold'
            WHEN docStage.IsApprovalRejected = 1 OR docStage.IsCoopRejected = 1 THEN N'Rejected'
            WHEN docStage.IsFullyApproved = 1 THEN N'Approved'
            ELSE N'Pending'
        END AS ResultVerbKey,
        COALESCE(
            CASE
                WHEN docStage.IsOnHold = 1 OR docStage.IsOngoing = 1 THEN
                (
                    SELECT TOP(1) COALESCE(upP.DisplayName, N'')
                    FROM dbo.DocumentApprovals pa
                    LEFT JOIN dbo.UserProfiles upP ON upP.UserId = pa.UserId
                    WHERE pa.DocId = d.DocId
                      AND ISNULL(pa.Status, N'Pending') LIKE N'Pending%'
                    ORDER BY pa.StepOrder ASC
                )
            END,
            CASE
                WHEN docStage.IsOnHold = 1 OR docStage.IsOngoing = 1 THEN
                (
                    SELECT TOP(1) COALESCE(upCP.DisplayName, N'')
                    FROM dbo.DocumentCooperations pc
                    LEFT JOIN dbo.UserProfiles upCP ON upCP.UserId = pc.UserId
                    WHERE pc.DocId = d.DocId
                      AND ISNULL(pc.Status, N'Pending') = N'Pending'
                    ORDER BY pc.RoleKey ASC
                )
            END,
            (
                SELECT TOP(1) COALESCE(upA.DisplayName, da.ActorName, N'')
                FROM dbo.DocumentApprovals da
                LEFT JOIN dbo.UserProfiles upA ON upA.UserId = da.UserId
                WHERE da.DocId = d.DocId
                  AND (
                        ISNULL(da.Action, N'') IN (N'approve', N'reject', N'hold', N'Recalled')
                     OR ISNULL(da.Status, N'') IN (N'Approved', N'OnHold', N'Recalled')
                     OR ISNULL(da.Status, N'') LIKE N'Rejected%'
                     OR ISNULL(da.Status, N'') LIKE N'PendingHold%'
                  )
                ORDER BY ISNULL(da.ActedAt, d.CreatedAt) DESC, da.StepOrder DESC
            ),
            (
                SELECT TOP(1) COALESCE(upC.DisplayName, dc.ActorName, N'')
                FROM dbo.DocumentCooperations dc
                LEFT JOIN dbo.UserProfiles upC ON upC.UserId = dc.UserId
                WHERE dc.DocId = d.DocId
                  AND (
                        ISNULL(dc.Action, N'') IN (N'Cooperate', N'Reject', N'Rejected')
                     OR ISNULL(dc.Status, N'') IN (N'Cooperated', N'Rejected')
                     OR ISNULL(dc.Status, N'') LIKE N'Rejected%'
                  )
                ORDER BY ISNULL(dc.ActedAt, d.CreatedAt) DESC, dc.RoleKey DESC
            ),
            N''
        ) AS ResultActorName,
        COALESCE(
            CASE
                WHEN docStage.IsOnHold = 1 OR docStage.IsOngoing = 1 THEN
                (
                    SELECT TOP(1) COALESCE(pmLoc.Name, pm.Name, N'')
                    FROM dbo.DocumentApprovals pa
                    LEFT JOIN dbo.UserProfiles upP ON upP.UserId = pa.UserId
                    LEFT JOIN dbo.PositionMasters pm ON pm.CompCd = upP.CompCd AND pm.Id = upP.PositionId
                    LEFT JOIN dbo.PositionMasterLoc pmLoc ON pmLoc.PositionId = pm.Id AND pmLoc.LangCode = @LangCode
                    WHERE pa.DocId = d.DocId
                      AND ISNULL(pa.Status, N'Pending') LIKE N'Pending%'
                    ORDER BY pa.StepOrder ASC
                )
            END,
            CASE
                WHEN docStage.IsOnHold = 1 OR docStage.IsOngoing = 1 THEN
                (
                    SELECT TOP(1) COALESCE(pmLoc.Name, pm.Name, N'')
                    FROM dbo.DocumentCooperations pc
                    LEFT JOIN dbo.UserProfiles upCP ON upCP.UserId = pc.UserId
                    LEFT JOIN dbo.PositionMasters pm ON pm.CompCd = upCP.CompCd AND pm.Id = upCP.PositionId
                    LEFT JOIN dbo.PositionMasterLoc pmLoc ON pmLoc.PositionId = pm.Id AND pmLoc.LangCode = @LangCode
                    WHERE pc.DocId = d.DocId
                      AND ISNULL(pc.Status, N'Pending') = N'Pending'
                    ORDER BY pc.RoleKey ASC
                )
            END,
            (
                SELECT TOP(1) COALESCE(pmLoc.Name, pm.Name, N'')
                FROM dbo.DocumentApprovals da
                LEFT JOIN dbo.UserProfiles upA ON upA.UserId = da.UserId
                LEFT JOIN dbo.PositionMasters pm ON pm.CompCd = upA.CompCd AND pm.Id = upA.PositionId
                LEFT JOIN dbo.PositionMasterLoc pmLoc ON pmLoc.PositionId = pm.Id AND pmLoc.LangCode = @LangCode
                WHERE da.DocId = d.DocId
                  AND (
                        ISNULL(da.Action, N'') IN (N'approve', N'reject', N'hold', N'Recalled')
                     OR ISNULL(da.Status, N'') IN (N'Approved', N'OnHold', N'Recalled')
                     OR ISNULL(da.Status, N'') LIKE N'Rejected%'
                     OR ISNULL(da.Status, N'') LIKE N'PendingHold%'
                  )
                ORDER BY ISNULL(da.ActedAt, d.CreatedAt) DESC, da.StepOrder DESC
            ),
            (
                SELECT TOP(1) COALESCE(pmLoc.Name, pm.Name, N'')
                FROM dbo.DocumentCooperations dc
                LEFT JOIN dbo.UserProfiles upC ON upC.UserId = dc.UserId
                LEFT JOIN dbo.PositionMasters pm ON pm.CompCd = upC.CompCd AND pm.Id = upC.PositionId
                LEFT JOIN dbo.PositionMasterLoc pmLoc ON pmLoc.PositionId = pm.Id AND pmLoc.LangCode = @LangCode
                WHERE dc.DocId = d.DocId
                  AND (
                        ISNULL(dc.Action, N'') IN (N'Cooperate', N'Reject', N'Rejected')
                     OR ISNULL(dc.Status, N'') IN (N'Cooperated', N'Rejected')
                     OR ISNULL(dc.Status, N'') LIKE N'Rejected%'
                  )
                ORDER BY ISNULL(dc.ActedAt, d.CreatedAt) DESC, dc.RoleKey DESC
            ),
            N''
        ) AS ResultActorPosition
) rs";

            const string outerApplyIsReadAny = @"
OUTER APPLY(
    SELECT TOP(1) 1 AS HasLog
    FROM dbo.DocumentViewLogs v
    WHERE v.DocId = d.DocId
      AND v.ViewerId = @UserId
) vAny";

            string whereCreatedSub = "";
            {
                var sub = (createdSub ?? string.Empty).Trim().ToLowerInvariant();
                if (sub == "completed") whereCreatedSub = " AND docStage.IsCompleted = 1 ";
                else if (sub == "ongoing") whereCreatedSub = " AND docStage.IsOngoing = 1 ";
            }

            string whereApprovalSub = "";
            {
                var sub = (approvalSub ?? string.Empty).Trim().ToLowerInvariant();
                if (sub == "completed") whereApprovalSub = " AND docStage.IsCompleted = 1 ";
                else if (sub == "ongoing") whereApprovalSub = " AND docStage.IsOngoing = 1 ";
            }

            string whereCooperationSub = "";
            {
                var sub = (cooperationSub ?? string.Empty).Trim().ToLowerInvariant();
                if (sub == "completed") whereCooperationSub = " AND docStage.IsCompleted = 1 ";
                else if (sub == "ongoing") whereCooperationSub = " AND docStage.IsOngoing = 1 ";
            }

            string whereSharedSub = "";
            {
                var sub = (sharedSub ?? string.Empty).Trim().ToLowerInvariant();
                if (sub == "completed") whereSharedSub = " AND docStage.IsCompleted = 1 ";
                else if (sub == "ongoing") whereSharedSub = " AND docStage.IsOngoing = 1 ";
            }

            const string coopSelectCols = @"
    ISNULL(coopAgg.CoopTotal, 0)              AS CoopTotalSteps,
    ISNULL(coopAgg.CoopDoneKeys, N'')         AS CoopDoneKeys,
    ISNULL(coopAgg.CoopRejectedKeys, N'')     AS CoopRejectedKeys,
    ISNULL(coopAgg.CoopPendingName, N'')      AS CoopPendingName,
    ISNULL(coopAgg.CoopPendingPosition, N'')  AS CoopPendingPosition,";

            string sqlCount;
            string sqlList;
            var offset = (page - 1) * pageSize;

            if (tabKey == "created")
            {
                sqlCount = @"
SELECT COUNT(1)
FROM dbo.Documents d
" + outerApplyStepAgg + @"
" + outerApplyCoopAggEmpty + @"
" + outerApplyDocStage + @"
WHERE d.CreatedBy = @UserId" + whereCreatedSub + whereSearch + whereTitleFilter + ";";

                sqlList = @"
SELECT
    d.DocId,
    d.TemplateTitle,
    d.CreatedAt,
    ISNULL(d.Status, N'') AS Status,
    up.DisplayName AS AuthorName,
    (SELECT COUNT(1) FROM dbo.DocumentComments c WHERE c.DocId = d.DocId AND c.IsDeleted = 0) AS CommentCount,
    CAST(0 AS bit) AS HasAttachment,
    ISNULL(stepAgg.TotalSteps, 0) AS TotalSteps,
    ISNULL(stepAgg.DoneSteps, 0) AS CompletedSteps,
    ISNULL(rs.ResultVerbKey, N'') AS ResultVerbKey,
    ISNULL(rs.ResultActorName, N'') AS ResultActorName,
    ISNULL(rs.ResultActorPosition, N'') AS ResultActorPosition,
" + coopSelectCols + @"
    CAST(CASE WHEN vAny.HasLog IS NULL THEN 0 ELSE 1 END AS bit) AS IsRead
FROM dbo.Documents d
LEFT JOIN dbo.UserProfiles up ON d.CreatedBy = up.UserId
" + outerApplyStepAgg + @"
" + outerApplyCoopAggEmpty + @"
" + outerApplyDocStage + @"
" + outerApplyResultActor + @"
" + outerApplyIsReadAny + @"
WHERE d.CreatedBy = @UserId" + whereCreatedSub + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }
            else if (tabKey == "approval")
            {
                sqlCount = @"
SELECT COUNT(DISTINCT d.DocId)
FROM dbo.DocumentApprovals a
JOIN dbo.Documents d ON a.DocId = d.DocId
" + outerApplyStepAgg + @"
" + outerApplyCoopAgg + @"
" + outerApplyDocStage + @"
WHERE a.UserId = @UserId" + whereApprovalSub + whereSearch + whereTitleFilter + ";";

                sqlList = @"
SELECT DISTINCT
    d.DocId,
    d.TemplateTitle,
    d.CreatedAt,
    ISNULL(d.Status, N'') AS Status,
    CASE
        WHEN docStage.IsOngoing = 1 THEN N'Pending'
        WHEN docStage.IsCompleted = 1 AND docStage.IsRecalled = 1 THEN N'Recalled'
        WHEN docStage.IsCompleted = 1 AND (docStage.IsApprovalRejected = 1 OR docStage.IsCoopRejected = 1) THEN N'Rejected'
        WHEN docStage.IsFullyApproved = 1 THEN N'Approved'
        ELSE ISNULL(d.Status, N'')
    END AS StatusCode,
    up.DisplayName AS AuthorName,
    (SELECT COUNT(1) FROM dbo.DocumentComments c WHERE c.DocId = d.DocId AND c.IsDeleted = 0) AS CommentCount,
    CAST(0 AS bit) AS HasAttachment,
    ISNULL(stepAgg.TotalSteps, 0) AS TotalSteps,
    ISNULL(stepAgg.DoneSteps, 0) AS CompletedSteps,
    ISNULL(rs.ResultVerbKey, N'') AS ResultVerbKey,
    ISNULL(rs.ResultActorName, N'') AS ResultActorName,
    ISNULL(rs.ResultActorPosition, N'') AS ResultActorPosition,
" + coopSelectCols + @"
    CAST(CASE WHEN vAny.HasLog IS NULL THEN 0 ELSE 1 END AS bit) AS IsRead
FROM dbo.DocumentApprovals a
JOIN dbo.Documents d ON a.DocId = d.DocId
LEFT JOIN dbo.UserProfiles up ON d.CreatedBy = up.UserId
" + outerApplyStepAgg + @"
" + outerApplyCoopAgg + @"
" + outerApplyDocStage + @"
" + outerApplyResultActor + @"
" + outerApplyIsReadAny + @"
WHERE a.UserId = @UserId" + whereApprovalSub + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }
            else if (tabKey == "cooperation")
            {
                const string cooperationBase = @"
FROM (
    SELECT
        c.DocId,
        MAX(c.Id) AS LastId
    FROM dbo.DocumentCooperations c
    WHERE c.UserId = @UserId
    GROUP BY c.DocId
) cc
JOIN dbo.DocumentCooperations c ON c.Id = cc.LastId
JOIN dbo.Documents d ON c.DocId = d.DocId";

                sqlCount = @"
SELECT COUNT(1)
" + cooperationBase + @"
" + outerApplyStepAgg + @"
" + outerApplyCoopAgg + @"
" + outerApplyDocStage + @"
WHERE 1 = 1" + whereCooperationSub + whereSearch + whereTitleFilter + ";";

                sqlList = @"
SELECT
    d.DocId,
    d.TemplateTitle,
    d.CreatedAt,
    ISNULL(d.Status, N'') AS Status,
    up.DisplayName AS AuthorName,
    (SELECT COUNT(1) FROM dbo.DocumentComments cmt WHERE cmt.DocId = d.DocId AND cmt.IsDeleted = 0) AS CommentCount,
    CAST(0 AS bit) AS HasAttachment,
    ISNULL(stepAgg.TotalSteps, 0) AS TotalSteps,
    ISNULL(stepAgg.DoneSteps, 0) AS CompletedSteps,
    ISNULL(rs.ResultVerbKey, N'') AS ResultVerbKey,
    ISNULL(rs.ResultActorName, N'') AS ResultActorName,
    ISNULL(rs.ResultActorPosition, N'') AS ResultActorPosition,
" + coopSelectCols + @"
    CAST(CASE WHEN vAny.HasLog IS NULL THEN 0 ELSE 1 END AS bit) AS IsRead
" + cooperationBase + @"
LEFT JOIN dbo.UserProfiles up ON d.CreatedBy = up.UserId
" + outerApplyStepAgg + @"
" + outerApplyCoopAgg + @"
" + outerApplyDocStage + @"
" + outerApplyResultActor + @"
" + outerApplyIsReadAny + @"
WHERE 1 = 1" + whereCooperationSub + whereSearch + whereTitleFilter + $@"
{orderBy}
OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY;";
            }
            else
            {
                const string whereShareActive = @"
 AND ISNULL(s.IsRevoked, 0) = 0
 AND (s.ExpireAt IS NULL OR s.ExpireAt > SYSUTCDATETIME())";

                string whereSharedReadCount = "";
                string whereSharedReadList = "";

                var tf = (titleFilter ?? string.Empty).Trim().ToLowerInvariant();
                if (tf == "viewed" || tf == "read")
                {
                    whereSharedReadCount = " AND vAny.HasLog IS NOT NULL ";
                    whereSharedReadList = " AND vAny.HasLog IS NOT NULL ";
                }
                else if (tf == "unviewed" || tf == "unread")
                {
                    whereSharedReadCount = " AND vAny.HasLog IS NULL ";
                    whereSharedReadList = " AND vAny.HasLog IS NULL ";
                }

                sqlCount = @"
SELECT COUNT(1)
FROM dbo.DocumentShares s
JOIN dbo.Documents d ON s.DocId = d.DocId
" + outerApplyStepAgg + @"
" + outerApplyCoopAgg + @"
" + outerApplyDocStage + @"
" + outerApplyIsReadAny + @"
WHERE s.UserId = @UserId" + whereShareActive + whereSharedSub + whereSharedReadCount + whereSearch + whereTitleFilter + ";";

                sqlList = @"
SELECT
    d.DocId,
    d.TemplateTitle,
    d.CreatedAt,
    ISNULL(d.Status, N'') AS Status,
    up.DisplayName AS AuthorName,
    (SELECT COUNT(1) FROM dbo.DocumentComments c WHERE c.DocId = d.DocId AND c.IsDeleted = 0) AS CommentCount,
    CAST(0 AS bit) AS HasAttachment,
    ISNULL(stepAgg.TotalSteps, 0) AS TotalSteps,
    ISNULL(stepAgg.DoneSteps, 0) AS CompletedSteps,
    ISNULL(rs.ResultVerbKey, N'') AS ResultVerbKey,
    ISNULL(rs.ResultActorName, N'') AS ResultActorName,
    ISNULL(rs.ResultActorPosition, N'') AS ResultActorPosition,
" + coopSelectCols + @"
    CAST(CASE WHEN vAny.HasLog IS NULL THEN 0 ELSE 1 END AS bit) AS IsRead
FROM dbo.DocumentShares s
JOIN dbo.Documents d ON s.DocId = d.DocId
LEFT JOIN dbo.UserProfiles up ON d.CreatedBy = up.UserId
" + outerApplyStepAgg + @"
" + outerApplyCoopAgg + @"
" + outerApplyDocStage + @"
" + outerApplyResultActor + @"
" + outerApplyIsReadAny + @"
WHERE s.UserId = @UserId" + whereShareActive + whereSharedSub + whereSharedReadList + whereSearch + whereTitleFilter + $@"
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

                var vk = (verbKey ?? string.Empty).Trim();
                if (vk == "Pending" || vk == "OnHold")
                {
                    if (!string.IsNullOrWhiteSpace(actorName) && !string.IsNullOrWhiteSpace(actorPos))
                        return actorName + " " + actorPos;
                    return !string.IsNullOrWhiteSpace(actorName) ? actorName : actorPos;
                }

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
            cmdCount.Parameters.Add(new SqlParameter("@LangCode", SqlDbType.NVarChar, 10) { Value = langCode });
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

                int ordStatusCode = -1;
                try { ordStatusCode = rdr.GetOrdinal("StatusCode"); } catch { ordStatusCode = -1; }

                while (await rdr.ReadAsync())
                {
                    var createdAtLocal = (rdr["CreatedAt"] is DateTime utc)
                        ? ToLocalStringFromUtc(utc)
                        : string.Empty;

                    var rawStatus = rdr["Status"]?.ToString() ?? string.Empty;

                    var rawStatusCode = rawStatus;
                    if (ordStatusCode >= 0 && rdr[ordStatusCode] != DBNull.Value)
                        rawStatusCode = rdr[ordStatusCode]?.ToString() ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(rawStatusCode))
                        rawStatusCode = rawStatus;

                    var totalSteps = rdr["TotalSteps"] is int ts ? ts : Convert.ToInt32(rdr["TotalSteps"]);
                    var completedSteps = rdr["CompletedSteps"] is int cs2 ? cs2 : Convert.ToInt32(rdr["CompletedSteps"]);
                    var commentCount = rdr["CommentCount"] is int cc ? cc : Convert.ToInt32(rdr["CommentCount"]);

                    var verbKey = rdr["ResultVerbKey"]?.ToString() ?? string.Empty;
                    var actorName = rdr["ResultActorName"]?.ToString() ?? string.Empty;
                    var actorPos = rdr["ResultActorPosition"]?.ToString() ?? string.Empty;

                    var resultSummary = BuildResultSummarySentence(verbKey, actorName, actorPos);

                    var isRead = false;
                    if (ordIsRead >= 0 && rdr[ordIsRead] != DBNull.Value)
                    {
                        var v = rdr[ordIsRead];
                        if (v is bool b) isRead = b;
                        else isRead = Convert.ToInt32(v) == 1;
                    }

                    var coopTotal = rdr["CoopTotalSteps"] is int ct ? ct : Convert.ToInt32(rdr["CoopTotalSteps"]);
                    var coopPendName = rdr["CoopPendingName"]?.ToString() ?? string.Empty;
                    var coopPendPos = rdr["CoopPendingPosition"]?.ToString() ?? string.Empty;
                    var coopDoneKeys = rdr["CoopDoneKeys"]?.ToString() ?? string.Empty;
                    var coopRejectedKeys = rdr["CoopRejectedKeys"]?.ToString() ?? string.Empty;

                    items.Add(new
                    {
                        docId = rdr["DocId"]?.ToString() ?? string.Empty,
                        templateTitle = rdr["TemplateTitle"]?.ToString() ?? string.Empty,
                        authorName = rdr["AuthorName"]?.ToString() ?? string.Empty,
                        createdAt = createdAtLocal,
                        status = rawStatus,
                        statusCode = rawStatusCode,
                        totalApprovers = totalSteps,
                        completedApprovers = completedSteps,
                        commentCount = commentCount,
                        hasAttachment = Convert.ToInt32(rdr["HasAttachment"]) == 1,
                        resultSummary = resultSummary,
                        isRead = isRead,
                        coopTotalSteps = coopTotal,
                        coopDoneKeys = coopDoneKeys,
                        coopRejectedKeys = coopRejectedKeys,
                        coopPendingName = coopPendName,
                        coopPendingPosition = coopPendPos
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
  AND ISNULL(a.Status, N'') LIKE N'Pending%'
  AND (
        ISNULL(d.Status, N'') = N'PendingA'     + CAST(a.StepOrder AS nvarchar(10))
     OR ISNULL(d.Status, N'') = N'PendingHoldA' + CAST(a.StepOrder AS nvarchar(10))
  )
  AND ISNULL(d.Status, N'') NOT LIKE N'Recalled%';";

            var sqlCooperationPending = @"
SELECT COUNT(1)
FROM (
    SELECT DISTINCT c.DocId
    FROM DocumentCooperations c
    JOIN Documents d ON c.DocId = d.DocId
    WHERE c.UserId = @UserId
      AND ISNULL(c.Status, N'') NOT LIKE N'Recalled%'
      AND NOT (
            ISNULL(c.Status, N'') LIKE N'Approved%'
         OR ISNULL(c.Status, N'') LIKE N'Rejected%'
         OR ISNULL(c.Status, N'') LIKE N'Recalled%'
         OR ISNULL(c.Status, N'') LIKE N'Recall%'
         OR ISNULL(c.Status, N'') LIKE N'Cooperated%'
      )
) x;";

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

            var sqlSharedUnread = @"
SELECT COUNT(1)
FROM DocumentShares s
WHERE s.UserId = @UserId
  AND ISNULL(s.IsRevoked, 0) = 0
  AND (s.ExpireAt IS NULL OR s.ExpireAt > SYSUTCDATETIME())
  AND ISNULL(s.IsRead, 0) = 0;";

            int approvalPending = 0;
            int cooperationPending = 0;
            int sharedUnread = 0;
            int createdUnread = 0;

            using (var cmd = new SqlCommand(sqlApprovalPending, conn))
            {
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
                approvalPending = Convert.ToInt32(await cmd.ExecuteScalarAsync());
            }

            using (var cmd = new SqlCommand(sqlCooperationPending, conn))
            {
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });
                cooperationPending = Convert.ToInt32(await cmd.ExecuteScalarAsync());
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

            return Json(new { created = createdUnread, approvalPending, cooperationPending, sharedUnread });
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