// 2026.03.17 Added: DetailDX 구성
// 2026.03.18 Changed: DxTemp 임시파일 생성 제거 — ReadOnly 조회이므로 원본 파일 직접 사용
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Text.Json;
using System.Threading.Tasks;
using WebApplication1.Models;
using WebApplication1.Services;
using static WebApplication1.Controllers.DocControllerHelper;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class DetailDXController : Controller
    {
        private readonly IStringLocalizer<SharedResource> _S;
        private readonly IConfiguration _cfg;
        private readonly IWebHostEnvironment _env;
        private readonly ILogger<DetailDXController> _log;
        private readonly DocControllerHelper _helper;

        public DetailDXController(
            IStringLocalizer<SharedResource> S,
            IConfiguration cfg,
            IWebHostEnvironment env,
            ILogger<DetailDXController> log)
        {
            _S = S;
            _cfg = cfg;
            _env = env;
            _log = log;
            _helper = new DocControllerHelper(_cfg, _env, () => User);
        }

        [HttpGet("DetailDX/{id?}")]
        public async Task<IActionResult> DetailDX(string? id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                var qid = HttpContext?.Request?.Query["id"].ToString();
                if (!string.IsNullOrWhiteSpace(qid))
                    id = qid;
            }

            if (string.IsNullOrWhiteSpace(id))
                return Redirect("/Doc/Board");

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            string? templateCode = null;
            string? templateTitle = null;
            string? status = null;
            string? descriptorJson = null;
            string? outputPath = null;
            string? compCd = null;
            string? creatorUserId = null;
            string? creatorNameRaw = null;
            DateTime? createdAtUtc = null;

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT TOP 1
       d.DocId,
       d.TemplateCode,
       d.TemplateTitle,
       d.Status,
       d.DescriptorJson,
       d.OutputPath,
       d.CompCd,
       d.CreatedBy,
       d.CreatedByName,
       d.CreatedAt,
       COALESCE(p.DisplayName, d.CreatedByName, u.UserName, u.Email) AS CreatorDisplayName
FROM dbo.Documents d
LEFT JOIN dbo.AspNetUsers  u ON u.Id     = d.CreatedBy
LEFT JOIN dbo.UserProfiles p ON p.UserId = d.CreatedBy
WHERE d.DocId = @DocId;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var rd = await cmd.ExecuteReaderAsync();
                if (!await rd.ReadAsync())
                {
                    ViewBag.Message = _S["DOC_Err_DocumentNotFound"];
                    ViewBag.DocId = id;
                    ViewBag.DocumentId = id;
                    return View("~/Views/Doc/DetailDX.cshtml");
                }

                templateCode = rd["TemplateCode"] as string ?? string.Empty;
                templateTitle = rd["TemplateTitle"] as string ?? string.Empty;
                status = rd["Status"] as string ?? string.Empty;
                descriptorJson = rd["DescriptorJson"] as string ?? "{}";
                outputPath = rd["OutputPath"] as string ?? string.Empty;
                compCd = rd["CompCd"] as string ?? string.Empty;
                creatorUserId = rd["CreatedBy"] as string ?? string.Empty;
                creatorNameRaw = rd["CreatorDisplayName"] as string ?? string.Empty;
                createdAtUtc = rd["CreatedAt"] is DateTime cdt ? cdt : (DateTime?)null;
            }

            string? creatorDisplayName = creatorNameRaw;
            if (!string.IsNullOrWhiteSpace(creatorUserId))
            {
                await using var pc = conn.CreateCommand();
                pc.CommandText = @"
SELECT TOP 1 DisplayName
FROM dbo.UserProfiles
WHERE UserId = @UserId;";
                pc.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = creatorUserId! });

                var objName = await pc.ExecuteScalarAsync();
                if (objName != null && objName != DBNull.Value)
                    creatorDisplayName = Convert.ToString(objName) ?? creatorDisplayName;
            }

            string createdAtText = string.Empty;
            if (createdAtUtc.HasValue)
                createdAtText = _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(createdAtUtc.Value, DateTimeKind.Utc));

            string recalledAtText = string.Empty;
            try
            {
                await using var rcCmd = conn.CreateCommand();
                rcCmd.CommandText = @"
SELECT TOP 1 a.ActedAt
FROM dbo.DocumentApprovals a
WHERE a.DocId = @DocId
  AND (
        UPPER(ISNULL(a.Action,'')) = 'RECALLED'
     OR UPPER(ISNULL(a.Status,'')) = 'RECALLED'
  )
  AND a.ActedAt IS NOT NULL
ORDER BY a.ActedAt DESC;";
                rcCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                var obj = await rcCmd.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value)
                {
                    var dt = (obj is DateTime dtx) ? dtx : Convert.ToDateTime(obj);
                    recalledAtText = _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(dt, DateTimeKind.Utc));
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: recalled time load failed for {DocId}", id);
            }

            int? templateVersionId = null;
            if (!string.IsNullOrWhiteSpace(templateCode))
            {
                await using var tvCmd = conn.CreateCommand();
                tvCmd.CommandText = @"
SELECT TOP 1 v.Id
FROM dbo.DocTemplateVersion v
INNER JOIN dbo.DocTemplateMaster m ON v.TemplateId = m.Id
WHERE m.DocCode = @TemplateCode
ORDER BY v.VersionNo DESC, v.Id DESC;";
                tvCmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 50) { Value = templateCode! });

                var obj = await tvCmd.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value)
                    templateVersionId = Convert.ToInt32(obj);
            }

            // ★ 원본 파일 직접 사용 (ReadOnly 조회이므로 임시 파일 불필요)
            var dxOpenAbsPath = ResolveContentRootRelativePath(outputPath) ?? string.Empty;

            var previewJson = "{}";
            try
            {
                if (!string.IsNullOrWhiteSpace(dxOpenAbsPath) && System.IO.File.Exists(dxOpenAbsPath))
                    previewJson = BuildPreviewJsonFromExcel(dxOpenAbsPath, maxRows: 500, maxCols: 100);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: preview json build failed for {DocId}", id);
                previewJson = "{}";
            }

            var openRel = MakeRelativeToContentRoot(dxOpenAbsPath);

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
        COALESCE(a.ApproverDisplayText, up.DisplayName, a.ActorName, a.ApproverValue) AS ApproverDisplayText,
        a.SignaturePath
FROM dbo.DocumentApprovals a
LEFT JOIN dbo.UserProfiles up
       ON up.UserId = a.UserId
WHERE a.DocId = @DocId
ORDER BY a.StepOrder;";
                ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var r = await ac.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    DateTime? acted = r.IsDBNull(6) ? (DateTime?)null : r.GetDateTime(6);
                    approvals.Add(new
                    {
                        step = r.GetInt32(0),
                        roleKey = r["RoleKey"]?.ToString(),
                        approverValue = r["ApproverValue"]?.ToString(),
                        userId = r["UserId"]?.ToString(),
                        status = r["Status"]?.ToString(),
                        action = r["Action"]?.ToString(),
                        actedAtText = acted.HasValue
                            ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc))
                            : string.Empty,
                        actorName = r["ActorName"]?.ToString(),
                        approverDisplayText = r["ApproverDisplayText"]?.ToString() ?? string.Empty,
                        signaturePath = r["SignaturePath"]?.ToString()
                    });
                }
            }

            var approvalCells = new List<object>();
            if (templateVersionId.HasValue)
            {
                await using var ac2 = conn.CreateCommand();
                ac2.CommandText = @"
SELECT Slot,
       ISNULL(CellA1, A1) AS A1
FROM dbo.DocTemplateApproval
WHERE VersionId = @VersionId
ORDER BY Slot;";
                ac2.Parameters.Add(new SqlParameter("@VersionId", SqlDbType.Int) { Value = templateVersionId.Value });

                await using var rc = await ac2.ExecuteReaderAsync();
                while (await rc.ReadAsync())
                {
                    approvalCells.Add(new
                    {
                        Step = rc.GetInt32(0),
                        A1 = rc["A1"] as string ?? string.Empty
                    });
                }
            }

            var docFiles = new List<object>();
            try
            {
                await using var fc = conn.CreateCommand();
                fc.CommandText = @"
SELECT FileKey,
       OriginalName,
       ContentType,
       ByteSize,
       UploadedAt
FROM dbo.DocumentFiles
WHERE DocId = @DocId
ORDER BY UploadedAt, FileId ASC;";
                fc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var fr = await fc.ExecuteReaderAsync();
                while (await fr.ReadAsync())
                {
                    var uploadedAtText = string.Empty;
                    if (fr["UploadedAt"] != DBNull.Value)
                    {
                        var upUtc = Convert.ToDateTime(fr["UploadedAt"]);
                        uploadedAtText = _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(upUtc, DateTimeKind.Utc));
                    }

                    docFiles.Add(new
                    {
                        FileKey = fr["FileKey"] as string ?? string.Empty,
                        OriginalName = fr["OriginalName"] as string ?? string.Empty,
                        ContentType = fr["ContentType"] as string ?? string.Empty,
                        ByteSize = fr["ByteSize"] != DBNull.Value ? Convert.ToInt64(fr["ByteSize"]) : 0L,
                        uploadedAtText
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: load DocumentFiles failed for {DocId}", id);
            }

            var sharedRecipientUserIds = new List<string>();
            try
            {
                await using var sc = conn.CreateCommand();
                sc.CommandText = @"
SELECT s.UserId
FROM dbo.DocumentShares s
WHERE s.DocId = @DocId
  AND ISNULL(s.IsRevoked, 0) = 0
  AND (s.ExpireAt IS NULL OR s.ExpireAt > GETUTCDATE())
ORDER BY s.CreatedAt ASC, s.Id ASC;";
                sc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var sr = await sc.ExecuteReaderAsync();
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                while (await sr.ReadAsync())
                {
                    var uid = sr.IsDBNull(0) ? null : sr.GetString(0);
                    if (string.IsNullOrWhiteSpace(uid)) continue;
                    if (seen.Add(uid))
                        sharedRecipientUserIds.Add(uid);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: load DocumentShares failed for {DocId}", id);
            }

            try
            {
                ViewBag.OrgTreeNodes = await _helper.BuildOrgTreeNodesAsync("ko");
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: load OrgTreeNodes failed for {DocId}", id);
                ViewBag.OrgTreeNodes = Array.Empty<OrgTreeNode>();
            }

            try
            {
                await LogDocumentViewAsync(conn, id!);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: LogDocumentViewAsync failed for {DocId}", id);
            }

            var viewLogs = new List<object>();
            try
            {
                await using var vc = conn.CreateCommand();
                vc.CommandText = @"
SELECT TOP 200
       v.ViewedAt,
       v.ViewerId,
       v.ViewerRole,
       v.ClientIp,
       v.UserAgent,
       COALESCE(p.DisplayName, u.UserName, u.Email) AS UserName
FROM dbo.DocumentViewLogs v
LEFT JOIN dbo.UserProfiles p ON p.UserId = v.ViewerId
LEFT JOIN dbo.AspNetUsers  u ON u.Id     = v.ViewerId
WHERE v.DocId = @DocId
ORDER BY v.ViewedAt DESC;";
                vc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var vr = await vc.ExecuteReaderAsync();
                while (await vr.ReadAsync())
                {
                    DateTime? viewedUtc = vr.IsDBNull(0) ? (DateTime?)null : vr.GetDateTime(0);
                    var viewedAtText = viewedUtc.HasValue
                        ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(viewedUtc.Value, DateTimeKind.Utc))
                        : string.Empty;

                    viewLogs.Add(new
                    {
                        viewedAt = viewedUtc,
                        viewedAtText,
                        userId = vr.IsDBNull(1) ? null : Convert.ToString(vr.GetValue(1)),
                        viewerRole = vr.IsDBNull(2) ? null : Convert.ToString(vr.GetValue(2)),
                        clientIp = vr.IsDBNull(3) ? null : Convert.ToString(vr.GetValue(3)),
                        userAgent = vr.IsDBNull(4) ? null : Convert.ToString(vr.GetValue(4)),
                        userName = vr.IsDBNull(5) ? null : Convert.ToString(vr.GetValue(5))
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: load view logs failed for {DocId}", id);
            }

            var caps = await GetApprovalCapabilitiesAsync(id!);

            ViewData["Title"] = _S["DOC_Title_Detail"];
            ViewData["DisableDxAll"] = true;

            ViewBag.DocId = id;
            ViewBag.DocumentId = id;
            ViewBag.TemplateCode = templateCode ?? string.Empty;
            ViewBag.TemplateTitle = templateTitle ?? string.Empty;
            ViewBag.Status = status ?? string.Empty;
            ViewBag.CompCd = compCd ?? string.Empty;
            ViewBag.CreatorName = creatorDisplayName ?? string.Empty;

            ViewBag.DescriptorJson = descriptorJson ?? "{}";
            ViewBag.InputsJson = "{}";
            ViewBag.PreviewJson = previewJson;
            ViewBag.ApprovalsJson = JsonSerializer.Serialize(approvals);
            ViewBag.ViewLogsJson = JsonSerializer.Serialize(viewLogs);
            ViewBag.ApprovalCellsJson = JsonSerializer.Serialize(approvalCells);
            ViewBag.DocumentFilesJson = JsonSerializer.Serialize(docFiles);

            ViewBag.SelectedRecipientUserIds = sharedRecipientUserIds;
            ViewBag.SelectedRecipientUserIdsJson = JsonSerializer.Serialize(sharedRecipientUserIds);

            ViewBag.CreatedAtText = createdAtText;
            ViewBag.RecalledAtText = recalledAtText;

            ViewBag.ExcelPath = dxOpenAbsPath ?? string.Empty;
            ViewBag.OpenRel = openRel ?? string.Empty;
            ViewBag.DxCallbackUrl = "/Doc/dx-callback";
            ViewBag.DxDocumentId = "detail_" + (id ?? Guid.NewGuid().ToString("N"));

            ViewBag.CanRecall = caps.canRecall;
            ViewBag.CanApprove = caps.canApprove;
            ViewBag.CanHold = caps.canHold;
            ViewBag.CanReject = caps.canReject;
            ViewBag.CanBackToNew = caps.canRecall;
            ViewBag.CanPrint = true;
            ViewBag.CanExportPdf = true;

            return View("~/Views/Doc/DetailDX.cshtml");
        }

        private string ResolveContentRootRelativePath(string? path)
        {
            var raw = (path ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            if (System.IO.Path.IsPathRooted(raw))
                return raw;

            raw = raw.TrimStart('\\', '/')
                     .Replace('/', System.IO.Path.DirectorySeparatorChar)
                     .Replace('\\', System.IO.Path.DirectorySeparatorChar);

            return System.IO.Path.Combine(_env.ContentRootPath, raw);
        }

        private string MakeRelativeToContentRoot(string? absPath)
        {
            var raw = (absPath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            try
            {
                if (System.IO.Path.IsPathRooted(raw))
                    return System.IO.Path.GetRelativePath(_env.ContentRootPath, raw);
            }
            catch { }

            return raw;
        }

        private async Task LogDocumentViewAsync(SqlConnection conn, string docId)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
            var clientIp = HttpContext.Connection.RemoteIpAddress?.ToString() ?? string.Empty;
            var userAgent = Request.Headers["User-Agent"].ToString() ?? string.Empty;
            var tab = (Request?.Query["tab"].ToString() ?? string.Empty).Trim();
            var isSharedTab = string.Equals(tab, "shared", StringComparison.OrdinalIgnoreCase);

            var isSharedRecipient = false;
            try
            {
                await using var sc = conn.CreateCommand();
                sc.CommandText = @"
SELECT TOP 1 1
FROM dbo.DocumentShares s WITH (NOLOCK)
WHERE s.DocId = @DocId
  AND s.UserId = @UserId
  AND ISNULL(s.IsRevoked, 0) = 0
  AND (s.ExpireAt IS NULL OR s.ExpireAt > SYSUTCDATETIME());";
                sc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                sc.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });
                var o = await sc.ExecuteScalarAsync();
                isSharedRecipient = (o != null && o != DBNull.Value);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: shared recipient check failed. DocId={DocId}, UserId={UserId}", docId, userId);
            }

            string viewerRole = "Other";
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT
    CreatedBy,
    CASE WHEN EXISTS (
        SELECT 1 FROM dbo.DocumentApprovals da
        WHERE da.DocId = d.DocId AND da.UserId = @UserId
    ) THEN 1 ELSE 0 END AS IsApprover
FROM dbo.Documents d
WHERE d.DocId = @DocId;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });

                await using var r = await cmd.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    var createdBy = r["CreatedBy"] as string;
                    var isApprover = Convert.ToInt32(r["IsApprover"]) == 1;
                    var isCreator = !string.IsNullOrWhiteSpace(createdBy) &&
                                    string.Equals(createdBy, userId, StringComparison.OrdinalIgnoreCase);

                    if (isSharedTab)
                        viewerRole = "Shared";
                    else if (isCreator && isApprover)
                        viewerRole = "Creator+Approver";
                    else if (isCreator)
                        viewerRole = "Creator";
                    else if (isApprover)
                        viewerRole = "Approver";
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
                cmd.Parameters.Add(new SqlParameter("@ViewerId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });
                cmd.Parameters.Add(new SqlParameter("@ViewerRole", SqlDbType.NVarChar, 50) { Value = (object)viewerRole ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@ClientIp", SqlDbType.NVarChar, 64) { Value = (object)clientIp ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@UserAgent", SqlDbType.NVarChar, 512) { Value = (object)userAgent ?? DBNull.Value });
                await cmd.ExecuteNonQueryAsync();
            }

            if (isSharedTab && isSharedRecipient)
            {
                try
                {
                    await using var uc = conn.CreateCommand();
                    uc.CommandText = @"
UPDATE dbo.DocumentShares
SET IsRead = 1
WHERE DocId = @DocId
  AND UserId = @UserId
  AND ISNULL(IsRevoked, 0) = 0
  AND (ExpireAt IS NULL OR ExpireAt > SYSUTCDATETIME())
  AND ISNULL(IsRead, 0) = 0;";
                    uc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    uc.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });
                    await uc.ExecuteNonQueryAsync();
                }
                catch (Exception ex)
                {
                    _log.LogWarning(ex, "DetailDX: shared read update failed. DocId={DocId}, UserId={UserId}", docId, userId);
                }
            }
        }

        private async Task<(bool canRecall, bool canApprove, bool canHold, bool canReject)> GetApprovalCapabilitiesAsync(string docId)
        {
            var userId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(docId) || string.IsNullOrWhiteSpace(userId))
                return (false, false, false, false);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            string? status = null;
            string? createdBy = null;

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT TOP 1 Status, CreatedBy
FROM dbo.Documents
WHERE DocId = @id;";
                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });

                await using var r = await cmd.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    status = r["Status"] as string ?? string.Empty;
                    createdBy = r["CreatedBy"] as string ?? string.Empty;
                }
            }

            if (status == null)
                return (false, false, false, false);

            var st = status ?? string.Empty;
            var isCreator = string.Equals(createdBy ?? string.Empty, userId, StringComparison.OrdinalIgnoreCase);
            var isRecalled = st.Equals("Recalled", StringComparison.OrdinalIgnoreCase);
            var isFinal = st.StartsWith("Approved", StringComparison.OrdinalIgnoreCase)
                       || st.StartsWith("Rejected", StringComparison.OrdinalIgnoreCase);

            var canRecall = isCreator && st.StartsWith("Pending", StringComparison.OrdinalIgnoreCase);

            int currentStep = 0;
            await using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = @"
SELECT MIN(StepOrder)
FROM dbo.DocumentApprovals
WHERE DocId = @id
  AND ISNULL(Status, N'Pending') LIKE N'Pending%';";
                cmd2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });

                var obj = await cmd2.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value)
                    currentStep = Convert.ToInt32(obj);
            }

            if (currentStep == 0 || isFinal || isRecalled)
                return (canRecall, false, false, false);

            string? stepUserId = null;
            string stepStatus = string.Empty;

            await using (var cmd3 = conn.CreateCommand())
            {
                cmd3.CommandText = @"
SELECT TOP 1 UserId, ISNULL(Status, N'Pending') AS Status
FROM dbo.DocumentApprovals
WHERE DocId = @id
  AND StepOrder = @step;";
                cmd3.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                cmd3.Parameters.Add(new SqlParameter("@step", SqlDbType.Int) { Value = currentStep });

                await using var r = await cmd3.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    stepUserId = r["UserId"] as string ?? string.Empty;
                    stepStatus = r["Status"] as string ?? string.Empty;
                }
            }

            if (string.IsNullOrWhiteSpace(stepUserId))
                return (canRecall, false, false, false);

            var isMyStep = string.Equals(stepUserId, userId, StringComparison.OrdinalIgnoreCase);
            var isPending = stepStatus.StartsWith("Pending", StringComparison.OrdinalIgnoreCase);
            var canDo = isMyStep && isPending && !isRecalled && !isFinal;

            return (canRecall, canDo, canDo, canDo);
        }
    }
}
