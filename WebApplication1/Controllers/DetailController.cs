// 2026.01.23 Changed Doc 상세 기능을 DetailController로 분리
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
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
    public class DetailController : DocControllerHelper
    {
        public DetailController(
            IConfiguration cfg,
            IWebHostEnvironment env,
            ILoggerFactory loggerFactory,
            IStringLocalizer<SharedResource> S,
            IWebPushNotifier webPushNotifier
        ) : base(cfg, env, loggerFactory, S, webPushNotifier)
        {
        }

        // =========================
        // STEP 3. UpdateShares 공유 변경 시 추가된 공유자에게 알림
        // =========================
        [HttpPost("UpdateShares")]
        [ValidateAntiForgeryToken]
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> UpdateShares([FromBody] UpdateSharesDto dto)
        {
            if (dto is null || string.IsNullOrWhiteSpace(dto.DocId))
                return BadRequest(new { ok = false, messages = new[] { "DOC_Err_SaveFailed" }, stage = "arg", detail = "docId null/empty" });

            var docId = dto.DocId!.Trim();
            var actorId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";

            var selected = (dto.SelectedRecipientUserIds ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(x => !string.Equals(x, actorId, StringComparison.OrdinalIgnoreCase))
                .ToList();

            var newlyAdded = new List<string>();

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();
                await using var tx = await conn.BeginTransactionAsync();

                try
                {
                    var currentActive = new List<string>();
                    await using (var sel = conn.CreateCommand())
                    {
                        sel.Transaction = (SqlTransaction)tx;
                        sel.CommandText = @"
SELECT UserId
FROM dbo.DocumentShares
WHERE DocId = @DocId AND ISNULL(IsRevoked,0) = 0;";
                        sel.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });

                        await using var r = await sel.ExecuteReaderAsync();
                        while (await r.ReadAsync())
                        {
                            var uid = (r["UserId"]?.ToString() ?? "").Trim();
                            if (!string.IsNullOrWhiteSpace(uid))
                                currentActive.Add(uid);
                        }
                    }

                    newlyAdded = selected
                        .Where(uid => !currentActive.Any(x => string.Equals(x, uid, StringComparison.OrdinalIgnoreCase)))
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList();

                    var toRevoke = currentActive
                        .Where(uid => !selected.Any(x => string.Equals(x, uid, StringComparison.OrdinalIgnoreCase)))
                        .ToList();

                    foreach (var targetUserId in toRevoke)
                    {
                        await using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = (SqlTransaction)tx;
                            cmd.CommandText = @"
UPDATE dbo.DocumentShares
   SET IsRevoked = 1,
       ExpireAt = SYSUTCDATETIME()
 WHERE DocId = @DocId AND UserId = @UserId;";
                            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                            await cmd.ExecuteNonQueryAsync();
                        }

                        await using (var logCmd = conn.CreateCommand())
                        {
                            logCmd.Transaction = (SqlTransaction)tx;
                            logCmd.CommandText = @"
INSERT INTO dbo.DocumentShareLogs (DocId, ActorId, ChangeCode, TargetUserId, BeforeJson, AfterJson, ChangedAt)
VALUES (@DocId, @ActorId, @ChangeCode, @TargetUserId, NULL, @AfterJson, SYSUTCDATETIME());";
                            logCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                            logCmd.Parameters.Add(new SqlParameter("@ActorId", SqlDbType.NVarChar, 64) { Value = actorId });
                            logCmd.Parameters.Add(new SqlParameter("@ChangeCode", SqlDbType.NVarChar, 50) { Value = "ShareRevoked" });
                            logCmd.Parameters.Add(new SqlParameter("@TargetUserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                            logCmd.Parameters.Add(new SqlParameter("@AfterJson", SqlDbType.NVarChar, -1) { Value = "{\"revoked\":true}" });
                            await logCmd.ExecuteNonQueryAsync();
                        }
                    }

                    foreach (var targetUserId in selected)
                    {
                        await using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = (SqlTransaction)tx;
                            cmd.CommandText = @"
IF EXISTS (SELECT 1 FROM dbo.DocumentShares WHERE DocId=@DocId AND UserId=@UserId)
BEGIN
    UPDATE dbo.DocumentShares
       SET IsRevoked = 0,
           ExpireAt = NULL
     WHERE DocId=@DocId AND UserId=@UserId;
END
ELSE
BEGIN
    INSERT INTO dbo.DocumentShares (DocId, UserId, AccessRole, ExpireAt, IsRevoked, CreatedBy, CreatedAt)
    VALUES (@DocId, @UserId, 'Commenter', NULL, 0, @CreatedBy, SYSUTCDATETIME());
END";
                            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                            cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 64) { Value = actorId });
                            await cmd.ExecuteNonQueryAsync();
                        }

                        await using (var logCmd = conn.CreateCommand())
                        {
                            logCmd.Transaction = (SqlTransaction)tx;
                            logCmd.CommandText = @"
INSERT INTO dbo.DocumentShareLogs (DocId, ActorId, ChangeCode, TargetUserId, BeforeJson, AfterJson, ChangedAt)
VALUES (@DocId, @ActorId, @ChangeCode, @TargetUserId, NULL, @AfterJson, SYSUTCDATETIME());";
                            logCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                            logCmd.Parameters.Add(new SqlParameter("@ActorId", SqlDbType.NVarChar, 64) { Value = actorId });
                            logCmd.Parameters.Add(new SqlParameter("@ChangeCode", SqlDbType.NVarChar, 50) { Value = "ShareAdded" });
                            logCmd.Parameters.Add(new SqlParameter("@TargetUserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                            logCmd.Parameters.Add(new SqlParameter("@AfterJson", SqlDbType.NVarChar, -1) { Value = "{\"accessRole\":\"Commenter\"}" });
                            await logCmd.ExecuteNonQueryAsync();
                        }
                    }

                    await ((SqlTransaction)tx).CommitAsync();
                }
                catch
                {
                    await ((SqlTransaction)tx).RollbackAsync();
                    throw;
                }

                // 추가된 공유자에게 푸시
                try
                {
                    if (newlyAdded.Count > 0)
                    {
                        foreach (var uid in newlyAdded
                                     .Where(x => !string.IsNullOrWhiteSpace(x))
                                     .Select(x => x.Trim())
                                     .Distinct(StringComparer.OrdinalIgnoreCase))
                        {
                            await _webPushNotifier.SendToUserIdAsync(
                                userId: uid,
                                title: "PUSH_SummaryTitle",
                                body: "PUSH_ApprovalShare",
                                url: "/",
                                tag: "badge-approval-Share"
                            );
                        }
                    }
                }
                catch (Exception exPush)
                {
                    _log.LogWarning(exPush, "UpdateShares push notify failed docId={docId}", docId);
                }

                return Json(new { ok = true, docId = docId, selectedCount = selected.Count, addedNotifyCount = newlyAdded.Count });
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "UpdateShares failed docId={docId}", docId);
                return BadRequest(new { ok = false, messages = new[] { "DOC_Err_SaveFailed" }, stage = "db", detail = ex.Message });
            }
        }

        [HttpGet("Download/{fileKey}")]
        public async Task<IActionResult> Download([FromRoute] string fileKey)
        {
            if (string.IsNullOrWhiteSpace(fileKey))
                return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                await using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
SELECT TOP (1) DocId, FileKey, OriginalName, StoragePath, ContentType, ByteSize
FROM dbo.DocumentFiles
WHERE FileKey = @FileKey
ORDER BY UploadedAt DESC, FileKey DESC;";
                cmd.Parameters.Add(new SqlParameter("@FileKey", SqlDbType.NVarChar, 200) { Value = fileKey });

                await using var r = await cmd.ExecuteReaderAsync(CommandBehavior.SingleRow);
                if (!await r.ReadAsync())
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                var storagePath = (r["StoragePath"] as string) ?? string.Empty;
                var originalName = (r["OriginalName"] as string) ?? fileKey;
                var contentType = (r["ContentType"] as string) ?? "application/octet-stream";

                var rel = storagePath.Replace('\\', '/').Trim();
                if (string.IsNullOrWhiteSpace(rel))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                if (rel.StartsWith("/") || rel.Contains(":") || rel.Contains(".."))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                var full = Path.GetFullPath(Path.Combine(_env.ContentRootPath, rel));
                var root = Path.GetFullPath(_env.ContentRootPath);

                if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                if (!System.IO.File.Exists(full))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                var stream = new FileStream(full, FileMode.Open, FileAccess.Read, FileShare.Read);
                return File(stream, contentType, originalName, enableRangeProcessing: true);
            }
            catch
            {
                return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });
            }
        }

        // 2026.01.19 Changed shared 탭으로 Detail 진입 시 ViewerRole을 Shared로 기록
        private async Task LogDocumentViewAsync(SqlConnection conn, string docId)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            var clientIp = HttpContext.Connection.RemoteIpAddress?.ToString() ?? "";
            var userAgent = Request.Headers["User-Agent"].ToString() ?? "";

            var tab = (Request?.Query["tab"].ToString() ?? string.Empty).Trim();
            var forceShared = string.Equals(tab, "shared", StringComparison.OrdinalIgnoreCase);

            string viewerRole = "Other";

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT 
    CreatedBy,
    CASE WHEN EXISTS (
        SELECT 1 
        FROM dbo.DocumentApprovals da 
        WHERE da.DocId = d.DocId AND da.UserId = @UserId
    ) THEN 1 ELSE 0 END AS IsApprover
FROM dbo.Documents d
WHERE d.DocId = @DocId;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = (object)userId ?? DBNull.Value });

                using var r = await cmd.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    var createdBy = r["CreatedBy"] as string;
                    var isApprover = (int)r["IsApprover"] == 1;

                    var isCreator = !string.IsNullOrWhiteSpace(createdBy) &&
                                    string.Equals(createdBy, userId, StringComparison.OrdinalIgnoreCase);

                    if (forceShared)
                        viewerRole = "Shared";
                    else if (isCreator && isApprover)
                        viewerRole = "Creator+Approver";
                    else if (isCreator)
                        viewerRole = "Creator";
                    else if (isApprover)
                        viewerRole = "Approver";
                    else
                        viewerRole = "SharedOrOther";
                }
            }

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
INSERT INTO dbo.DocumentViewLogs
    (DocId, ViewerId, ViewerRole, ViewedAt, ClientIp, UserAgent)
VALUES
    (@DocId, @ViewerId, @ViewerRole, SYSUTCDATETIME(), @ClientIp, @UserAgent);";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                cmd.Parameters.Add(new SqlParameter("@ViewerId", SqlDbType.NVarChar, 450) { Value = userId ?? (object)DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@ViewerRole", SqlDbType.NVarChar, 50) { Value = (object)viewerRole ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@ClientIp", SqlDbType.NVarChar, 64) { Value = (object)clientIp ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@UserAgent", SqlDbType.NVarChar, 512) { Value = (object)userAgent ?? DBNull.Value });

                await cmd.ExecuteNonQueryAsync();
            }
        }

        [HttpGet("DetailsData")]
        [Produces("application/json")]
        public async Task<IActionResult> DetailsData(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return BadRequest(new { message = "DOC_Err_BadRequest" });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT d.DocId,
       d.TemplateCode,
       d.TemplateTitle,
       d.Status,
       d.DescriptorJson,
       d.OutputPath,
       d.CreatedAt
FROM dbo.Documents d
WHERE d.DocId = @id;";
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = id });

            string? descriptor = null, output = null, title = null, code = null, status = null;
            DateTime? createdAt = null;

            await using (var rd = await cmd.ExecuteReaderAsync())
            {
                if (await rd.ReadAsync())
                {
                    code = rd["TemplateCode"] as string ?? "";
                    title = rd["TemplateTitle"] as string ?? "";
                    status = rd["Status"] as string ?? "";
                    descriptor = rd["DescriptorJson"] as string ?? "{}";
                    output = rd["OutputPath"] as string ?? "";
                    if (!rd.IsDBNull(rd.GetOrdinal("CreatedAt")))
                        createdAt = (DateTime)rd["CreatedAt"];
                }
                else
                {
                    return NotFound(new { message = "DOC_Err_DocumentNotFound" });
                }
            }

            var preview = string.IsNullOrWhiteSpace(output) || !System.IO.File.Exists(output)
                ? "{}"
                : BuildPreviewJsonFromExcel(output);

            var approvals = new List<object>();
            await using (var ac = conn.CreateCommand())
            {
                ac.CommandText = @"
SELECT  a.StepOrder,
        a.RoleKey,
        a.ApproverValue,
        a.UserId,
        a.Status,
        a.Action,
        a.ActedAt,
        a.ActorName,
        up.DisplayName
FROM dbo.DocumentApprovals a
LEFT JOIN dbo.UserProfiles up
       ON up.UserId = a.UserId
WHERE a.DocId = @DocId
ORDER BY a.StepOrder;";
                ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var r = await ac.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    DateTime? acted = r.IsDBNull(6) ? (DateTime?)null : r.GetDateTime(6);
                    string? when = acted.HasValue
                        ? ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc))
                        : null;

                    var approverValue = r["ApproverValue"]?.ToString();
                    var userId = r["UserId"]?.ToString();
                    var actorName = r["ActorName"]?.ToString();
                    var displayName = r["DisplayName"] as string ?? string.Empty;

                    var approverDisplayText =
                        !string.IsNullOrWhiteSpace(displayName) ? displayName :
                        !string.IsNullOrWhiteSpace(actorName) ? actorName :
                        approverValue ?? string.Empty;

                    approvals.Add(new
                    {
                        step = r.GetInt32(0),
                        roleKey = r["RoleKey"]?.ToString(),
                        approverValue,
                        userId,
                        status = r["Status"]?.ToString(),
                        action = r["Action"]?.ToString(),
                        actedAtText = when,
                        actorName,
                        ApproverDisplayText = approverDisplayText
                    });
                }
            }

            return Json(new
            {
                ok = true,
                docId = id,
                templateCode = code,
                templateTitle = title,
                status,
                createdAt = createdAt.HasValue
                    ? ToLocalStringFromUtc(DateTime.SpecifyKind(createdAt.Value, DateTimeKind.Utc))
                    : null,
                descriptorJson = descriptor,
                previewJson = preview,
                approvals
            });
        }

        // ========= (기존 DocController의 Detail/Approve/Board 관련 액션은 이 컨트롤러로 이동) =========
        // - Detail(id)
        // - ApproveOrHold(...)
        // - BoardData/BoardBadges(...)
        // - Comments/Files 등 Detail 화면 주변 API
        // 위 액션들은 기존 시그니처 그대로 이 클래스로 옮기면 됩니다.
    }
}
