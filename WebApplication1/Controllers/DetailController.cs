// 2026.01.26 Changed: DetailController 분리 후 누락/깨짐(헬퍼 메서드, WebPush, GetApprovalCapabilitiesAsync 시그니처, 프리뷰 빌더 접근 제한)으로 인한 컴파일 오류를 해결하고 _webPushNotifier 기반으로 통일함
using ClosedXML.Excel;
using Microsoft.AspNetCore.Antiforgery;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Localization;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Claims;
using System.Text.Json;
using System.Threading.Tasks;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Services;

namespace WebApplication1.Controllers
{
    [Authorize]
    [Route("Doc")]
    public class DetailController : Controller
    {
        private readonly IStringLocalizer _S;
        private readonly IConfiguration _cfg;
        private readonly IAntiforgery _antiforgery;
        private readonly IWebHostEnvironment _env;
        private readonly ApplicationDbContext _db;
        private readonly IDocTemplateService _tpl;
        private readonly ILogger _log;
        private readonly IEmailSender _emailSender;
        private readonly SmtpOptions _smtpOpt;
        private static readonly object _attachSeqLock = new object();
        private readonly IWebPushNotifier _webPushNotifier;

        public DetailController(
            IStringLocalizer S,
            IConfiguration cfg,
            IAntiforgery antiforgery,
            IWebHostEnvironment env,
            ApplicationDbContext db,
            IDocTemplateService tpl,
            IEmailSender emailSender,
            IOptions<SmtpOptions> smtpOptions,
            ILogger<DetailController> log,
            IWebPushNotifier webPushNotifier
        )
        {
            _S = S;
            _cfg = cfg;
            _antiforgery = antiforgery;
            _env = env;
            _db = db;
            _tpl = tpl;
            _emailSender = emailSender;
            _smtpOpt = smtpOptions?.Value ?? new SmtpOptions();
            _log = log;
            _webPushNotifier = webPushNotifier;
        }

        // =========================
        // STEP 2. ApproveOrHold: 승인 시 다음 결재자 알림, 보류/반려 시 작성자 알림
        // =========================
        [HttpPost("ApproveOrHold")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> ApproveOrHold([FromBody] ApproveDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto.docId) || string.IsNullOrWhiteSpace(dto.action))
                return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" } });

            var actionLower = dto.action!.ToLowerInvariant();

            // ===== 1) 회수 전용 처리 =====
            if (actionLower == "recall")
            {
                var csRecall = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using (var conn = new SqlConnection(csRecall))
                {
                    await conn.OpenAsync();

                    await using (var cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = @"
UPDATE dbo.Documents
SET Status = N'Recalled'
WHERE DocId = @DocId
  AND ISNULL(Status,N'') LIKE N'Pending%';";
                        cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        await cmd.ExecuteNonQueryAsync();
                    }

                    await using (var ac = conn.CreateCommand())
                    {
                        ac.CommandText = @"
UPDATE dbo.DocumentApprovals
SET Action    = N'Recalled',
    ActedAt   = SYSUTCDATETIME(),
    ActorName = @actor
WHERE DocId = @DocId
  AND ISNULL(Status,N'Pending') LIKE N'Pending%';";
                        ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        ac.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200)
                        {
                            Value = GetCurrentUserDisplayNameStrict() ?? string.Empty
                        });
                        await ac.ExecuteNonQueryAsync();
                    }
                }

                return Json(new { ok = true, docId = dto.docId, status = "Recalled" });
            }

            // ===== 2) 승인/보류/반려 공통 처리 =====
            string outXlsx = string.Empty;

            try
            {
                var csOut = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using (var connOut = new SqlConnection(csOut))
                {
                    await connOut.OpenAsync();

                    await using var cmdOut = connOut.CreateCommand();
                    cmdOut.CommandText = @"
SELECT TOP (1) OutputPath
FROM dbo.Documents
WHERE DocId = @DocId;";
                    cmdOut.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });

                    var obj = await cmdOut.ExecuteScalarAsync();
                    outXlsx = (obj as string) ?? string.Empty;
                }

                if (!string.IsNullOrWhiteSpace(outXlsx) && !Path.IsPathRooted(outXlsx))
                {
                    outXlsx = Path.Combine(_env.ContentRootPath, outXlsx.TrimStart('\\', '/'));
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ApproveOrHold: output path resolve failed docId={docId}", dto.docId);
                outXlsx = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(outXlsx) || !System.IO.File.Exists(outXlsx))
                return NotFound(new { messages = new[] { "DOC_Err_DocumentNotFound" } });

            var approverId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            var approverName = GetCurrentUserDisplayNameStrict();

            string? approverDisplayText = null;
            string? signatureRelativePath = null;

            // --- 프로필 스냅샷 (서명/표시명) ---
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
                _log.LogError(ex, "ApproveOrHold: 프로필 스냅샷 조회 실패 docId={docId}", dto.docId);
            }

            string newStatus = "Updated";

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                // --- 2-1) 현재 사용자 Pending 단계 StepOrder 찾기 ---
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
                    findStep.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    findStep.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = approverId });

                    var stepObj = await findStep.ExecuteScalarAsync();
                    if (stepObj != null && stepObj != DBNull.Value)
                        currentStep = Convert.ToInt32(stepObj);
                }

                if (currentStep == null)
                    return Forbid();

                var step = currentStep.Value;

                // --- 2-2) DocumentApprovals 업데이트 ---
                await using (var u = conn.CreateCommand())
                {
                    u.CommandText = @"
UPDATE dbo.DocumentApprovals
SET Action              = @act,
    ActedAt             = SYSUTCDATETIME(),
    ActorName           = @actor,
    Status              = CASE 
                             WHEN @act = N'approve' THEN N'Approved'
                             WHEN @act = N'hold'    THEN N'OnHold'
                             WHEN @act = N'reject'  THEN N'Rejected'
                             ELSE ISNULL(Status, N'Updated')
                          END,
    UserId              = COALESCE(UserId, @uid),
    ApproverDisplayText = COALESCE(@displayText, ApproverDisplayText),
    SignaturePath       = COALESCE(@sigPath, SignaturePath)
WHERE DocId = @id AND StepOrder = @step;";
                    u.Parameters.Add(new SqlParameter("@act", SqlDbType.NVarChar, 20) { Value = actionLower });
                    u.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = approverName ?? string.Empty });
                    u.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });
                    u.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    u.Parameters.Add(new SqlParameter("@step", SqlDbType.Int) { Value = step });
                    u.Parameters.Add(new SqlParameter("@displayText", SqlDbType.NVarChar, 400) { Value = (object?)approverDisplayText ?? DBNull.Value });
                    u.Parameters.Add(new SqlParameter("@sigPath", SqlDbType.NVarChar, 400) { Value = (object?)signatureRelativePath ?? DBNull.Value });

                    await u.ExecuteNonQueryAsync();
                }

                // --- 2-3) Documents.Status 업데이트 ---
                newStatus = actionLower switch
                {
                    "approve" => $"ApprovedA{step}",
                    "hold" => $"OnHoldA{step}",
                    "reject" => $"RejectedA{step}",
                    _ => "Updated"
                };

                await using (var u2 = conn.CreateCommand())
                {
                    u2.CommandText = @"UPDATE dbo.Documents SET Status = @st WHERE DocId = @id;";
                    u2.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = newStatus });
                    u2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    await u2.ExecuteNonQueryAsync();
                }

                // --- 2-4) 승인 시 다음 단계 Pending 전환 + 다음 결재자 웹푸시 ---
                if (actionLower == "approve")
                {
                    var next = step + 1;

                    // 다음 단계 Pending 승인자(UserId) 조회
                    var nextApproverUserIds = await GetNextApproverUserIdsFromDbAsync(conn, dto.docId, next);

                    if (nextApproverUserIds.Count == 0)
                    {
                        await using var upFinal = conn.CreateCommand();
                        upFinal.CommandText = @"UPDATE dbo.Documents SET Status = N'Approved' WHERE DocId = @id;";
                        upFinal.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        await upFinal.ExecuteNonQueryAsync();
                        newStatus = "Approved";
                    }
                    else
                    {
                        await using (var up2 = conn.CreateCommand())
                        {
                            up2.CommandText = @"UPDATE dbo.Documents SET Status = @st WHERE DocId = @id;";
                            up2.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = $"PendingA{next}" });
                            up2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            await up2.ExecuteNonQueryAsync();
                        }
                        newStatus = $"PendingA{next}";

                        // 다음 결재자 알림(_webPushNotifier)
                        try
                        {
                            var titleText = (_S?["PUSH_Approval_Next_Title"] ?? "E-BOARD").ToString();
                            var bodyText = (_S?["PUSH_Approval_Next_Body"] ?? "결재 대기 문서가 있습니다.").ToString();

                            foreach (var uid in nextApproverUserIds)
                            {
                                await _webPushNotifier.SendToUserIdAsync(
                                    userId: uid,
                                    title: titleText,
                                    body: bodyText,
                                    url: "/Doc/Detail?id=" + Uri.EscapeDataString(dto.docId),
                                    tag: "badge-approval-pending"
                                );
                            }
                        }
                        catch (Exception exPush)
                        {
                            _log.LogWarning(exPush, "ApproveOrHold: push notify(next approver) failed docId={docId}", dto.docId);
                        }
                    }
                }
                else if (actionLower == "hold" || actionLower == "reject")
                {
                    // 보류/반려 시 작성자 알림(_webPushNotifier)
                    try
                    {
                        var authorId = await GetDocumentAuthorUserIdAsync(conn, dto.docId);
                        if (!string.IsNullOrWhiteSpace(authorId)
                            && !string.Equals(authorId, approverId, StringComparison.OrdinalIgnoreCase))
                        {
                            var titleText = (_S?["PUSH_Approval_Author_Title"] ?? "E-BOARD").ToString();
                            var bodyText = (actionLower == "hold")
                                ? (_S?["PUSH_ApprovalHold"] ?? "문서가 보류 처리되었습니다.").ToString()
                                : (_S?["PUSH_ApprovalReject"] ?? "문서가 반려 처리되었습니다.").ToString();

                            var tag = (actionLower == "hold") ? "approval-author-hold" : "approval-author-reject";

                            await _webPushNotifier.SendToUserIdAsync(
                                userId: authorId.Trim(),
                                title: titleText,
                                body: bodyText,
                                url: "/Doc/Detail?id=" + Uri.EscapeDataString(dto.docId),
                                tag: tag
                            );
                        }
                    }
                    catch (Exception exPush)
                    {
                        _log.LogWarning(exPush, "ApproveOrHold: push notify(author) failed docId={docId}", dto.docId);
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ApproveOrHold flow failed docId={docId}", dto.docId);
            }

            var previewJson = BuildPreviewJsonFromExcel(outXlsx);
            return Json(new { ok = true, docId = dto.docId, status = newStatus, previewJson });
        }

        private async Task<string?> GetDocumentAuthorUserIdAsync(SqlConnection conn, string docId)
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT TOP (1) NULLIF(LTRIM(RTRIM(CreatedBy)), N'') AS CreatedBy
FROM dbo.Documents
WHERE DocId = @DocId;";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

            var obj = await cmd.ExecuteScalarAsync();
            var s = (obj == null || obj == DBNull.Value) ? null : (obj.ToString() ?? "").Trim();
            return string.IsNullOrWhiteSpace(s) ? null : s;
        }

        private async Task<List<string>> GetNextApproverUserIdsFromDbAsync(SqlConnection conn, string docId, int nextStep)
        {
            var list = new List<string>();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT DISTINCT NULLIF(LTRIM(RTRIM(UserId)), N'') AS UserId
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND StepOrder = @Step
  AND ISNULL(Status, N'Pending') LIKE N'Pending%';";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = nextStep });

            await using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                var uid = r["UserId"] as string;
                if (!string.IsNullOrWhiteSpace(uid))
                    list.Add(uid.Trim());
            }

            return list
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ------------------------------------------------------------
        // UpdateShares 공유 알림: _webPushNotifier 로만 발송 (Lib.Net.Http.WebPush 직접 사용 제거)
        // ------------------------------------------------------------
        private async Task NotifySharesByWebPushAsync(string docId, List<string> newlyAdded)
        {
            if (string.IsNullOrWhiteSpace(docId)) return;
            if (newlyAdded == null || newlyAdded.Count == 0) return;

            var targets = newlyAdded
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (targets.Count == 0) return;

            var titleText = (_S?["PUSH_Share_Title"] ?? "E-BOARD").ToString();
            var bodyText = (_S?["PUSH_Share_Body"] ?? "공유 문서가 도착했습니다.").ToString();

            foreach (var uid in targets)
            {
                await _webPushNotifier.SendToUserIdAsync(
                    userId: uid,
                    title: titleText,
                    body: bodyText,
                    url: "/Doc/Detail?id=" + Uri.EscapeDataString(docId),
                    tag: "badge-shared"
                );
            }
        }

        [HttpGet("Detail/{id?}")]
        public async Task<IActionResult> Detail(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                var qid = HttpContext?.Request?.Query["id"].ToString();
                if (!string.IsNullOrWhiteSpace(qid))
                    id = qid;
            }

            if (string.IsNullOrWhiteSpace(id))
                return RedirectToAction("Board");

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            // ---------------- 1) 기본 문서 정보 ----------------
            string? templateCode = null;
            string? templateTitle = null;
            string? status = null;
            string? descriptorJson = null;
            string? outputPath = null;
            string? compCd = null;

            string? creatorUserId = null;
            string? creatorNameRaw = null;

            DateTime? createdAtUtc = null;
            DateTime? updatedAtUtc = null;

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
       d.UpdatedAt,
       COALESCE(p.DisplayName, d.CreatedByName, u.UserName, u.Email) AS CreatorDisplayName
FROM dbo.Documents d
LEFT JOIN dbo.AspNetUsers  u ON u.Id     = d.CreatedBy
LEFT JOIN dbo.UserProfiles p ON p.UserId = d.CreatedBy
WHERE d.DocId = @DocId;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var rd = await cmd.ExecuteReaderAsync();
                if (!await rd.ReadAsync())
                {
                    ViewBag.Message = _S["DOC_Err_DocumentNotFound"];
                    ViewBag.DocId = id;
                    ViewBag.DocumentId = id;
                    return View("Detail");
                }

                templateCode = rd["TemplateCode"] as string ?? "";
                templateTitle = rd["TemplateTitle"] as string ?? "";
                status = rd["Status"] as string ?? "";
                descriptorJson = rd["DescriptorJson"] as string ?? "{}";
                outputPath = rd["OutputPath"] as string ?? "";
                compCd = rd["CompCd"] as string ?? "";

                creatorUserId = rd["CreatedBy"] as string ?? "";
                creatorNameRaw = rd["CreatorDisplayName"] as string ?? "";

                createdAtUtc = rd["CreatedAt"] is DateTime cdt ? cdt : (DateTime?)null;
                updatedAtUtc = rd["UpdatedAt"] is DateTime udt ? udt : (DateTime?)null;
            }

            // ---------------- 1-0) 작성 시각(표시용) ----------------
            DateTime? createdAtForUi = null;
            string createdAtText = string.Empty;

            if (createdAtUtc.HasValue)
            {
                createdAtForUi = DateTime.SpecifyKind(createdAtUtc.Value, DateTimeKind.Utc);
                createdAtText = ToLocalStringFromUtc(createdAtForUi.Value);
            }

            // ---------------- 2) 프리뷰 JSON ----------------
            var inputsJson = "{}";
            string previewJson = "{}";

            try
            {
                var resolved = ResolveDocOutputPath(outputPath);

                if (!string.IsNullOrWhiteSpace(resolved) && System.IO.File.Exists(resolved))
                {
                    previewJson = BuildPreviewJsonFromExcel(resolved);
                }
                else
                {
                    _log.LogWarning(
                        "Detail: output excel not found. DocId={DocId}, OutputPath(raw)={Raw}, OutputPath(resolved)={Resolved}, ContentRoot={Root}",
                        id, outputPath ?? "", resolved ?? "", _env.ContentRootPath
                    );
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Detail: BuildPreviewJsonFromExcel failed for {DocId}. OutputPath={OutputPath}", id, outputPath ?? string.Empty);
                previewJson = "{}";
            }

            // ---------------- 3) 승인 정보 ----------------
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
                ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var r = await ac.ExecuteReaderAsync();
                while (await r.ReadAsync())
                {
                    DateTime? acted = r.IsDBNull(6) ? (DateTime?)null : r.GetDateTime(6);
                    string? when = acted.HasValue
                        ? ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc))
                        : null;

                    approvals.Add(new
                    {
                        step = r.GetInt32(0),
                        roleKey = r["RoleKey"]?.ToString(),
                        approverValue = r["ApproverValue"]?.ToString(),
                        userId = r["UserId"]?.ToString(),
                        status = r["Status"]?.ToString(),
                        action = r["Action"]?.ToString(),
                        actedAtText = when,
                        actorName = r["ActorName"]?.ToString(),
                        approverDisplayText = r["ApproverDisplayText"]?.ToString() ?? "",
                        signaturePath = r["SignaturePath"]?.ToString()
                    });
                }
            }

            // ---------------- 3-2) 문서 첨부 파일 목록(DocumentFiles) ----------------
            var docFiles = new List<object>();
            try
            {
                await using var fc = conn.CreateCommand();
                fc.CommandText = @"
SELECT FileKey,
       OriginalName,
       StoragePath,
       ContentType,
       ByteSize,
       Sha256,
       UploadedBy,
       UploadedAt
FROM dbo.DocumentFiles
WHERE DocId = @DocId
ORDER BY UploadedAt, FileId ASC;";
                fc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var fr = await fc.ExecuteReaderAsync();
                while (await fr.ReadAsync())
                {
                    var fileKey = fr["FileKey"] as string ?? string.Empty;
                    var originalName = fr["OriginalName"] as string ?? string.Empty;
                    var contentType = fr["ContentType"] as string ?? string.Empty;
                    long byteSize = (fr["ByteSize"] != DBNull.Value) ? Convert.ToInt64(fr["ByteSize"]) : 0;

                    DateTime? upUtc = fr["UploadedAt"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(fr["UploadedAt"]);
                    string? upText = null;
                    if (upUtc.HasValue)
                    {
                        var up = DateTime.SpecifyKind(upUtc.Value, DateTimeKind.Utc);
                        upText = ToLocalStringFromUtc(up);
                    }

                    docFiles.Add(new
                    {
                        fileKey,
                        originalName,
                        contentType,
                        byteSize,
                        uploadedAtText = upText
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: load DocumentFiles failed for {DocId}", id);
            }

            // ---------------- 3-3) 공유 대상자 목록(DocumentShares) ----------------
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
                sc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

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
                _log.LogWarning(ex, "Detail: load DocumentShares failed for {DocId}", id);
            }

            // ---------------- 3-4) 공유 후보(OrgTreeNodes) 로드 ----------------
            try
            {
                var langCode = "ko";
                var orgNodes = await BuildOrgTreeNodesAsync(langCode);
                ViewBag.OrgTreeNodes = orgNodes;
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: load OrgTreeNodes failed for {DocId}, CompCd={CompCd}", id, compCd ?? string.Empty);
                ViewBag.OrgTreeNodes = Array.Empty<OrgTreeNode>();
            }

            // ---------------- 4) 조회 로그 기록 ----------------
            try
            {
                await LogDocumentViewAsync(conn, id);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: LogDocumentViewAsync failed for {DocId}", id);
            }

            // ---------------- 5) 조회 로그 조회 ----------------
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
                vc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                await using var vr = await vc.ExecuteReaderAsync();

                string? GetStringSafe(int ordinal)
                    => vr.IsDBNull(ordinal) ? null : Convert.ToString(vr.GetValue(ordinal));

                while (await vr.ReadAsync())
                {
                    DateTime? viewedUtc = vr.IsDBNull(0) ? (DateTime?)null : vr.GetDateTime(0);
                    string? when = viewedUtc.HasValue
                        ? ToLocalStringFromUtc(DateTime.SpecifyKind(viewedUtc.Value, DateTimeKind.Utc))
                        : null;

                    viewLogs.Add(new
                    {
                        viewedAtText = when,
                        userId = GetStringSafe(1),
                        viewerRole = GetStringSafe(2),
                        clientIp = GetStringSafe(3),
                        userAgent = GetStringSafe(4),
                        userName = GetStringSafe(5)
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: load view logs failed for {DocId}", id);
            }

            // ---------------- 7) ViewBag 세팅 ----------------
            ViewBag.DocId = id;
            ViewBag.DocumentId = id;
            ViewBag.TemplateCode = templateCode ?? "";
            ViewBag.TemplateTitle = templateTitle ?? "";
            ViewBag.Status = status ?? "";
            ViewBag.CompCd = compCd ?? "";

            ViewBag.CreatorName = creatorNameRaw ?? string.Empty;

            ViewBag.DescriptorJson = descriptorJson ?? "{}";
            ViewBag.InputsJson = inputsJson;
            ViewBag.PreviewJson = previewJson;
            ViewBag.ApprovalsJson = JsonSerializer.Serialize(approvals);
            ViewBag.ViewLogsJson = JsonSerializer.Serialize(viewLogs);

            ViewBag.DocumentFilesJson = JsonSerializer.Serialize(docFiles);

            ViewBag.SelectedRecipientUserIds = sharedRecipientUserIds;
            ViewBag.SelectedRecipientUserIdsJson = JsonSerializer.Serialize(sharedRecipientUserIds);

            ViewBag.CreatedAt = createdAtForUi;
            ViewBag.CreatedAtText = createdAtText;

            var caps = await GetApprovalCapabilitiesAsync(id);
            ViewBag.CanRecall = caps.canRecall;
            ViewBag.CanApprove = caps.canApprove;
            ViewBag.CanHold = caps.canHold;
            ViewBag.CanReject = caps.canReject;
            ViewBag.CanBackToNew = caps.canRecall;

            return View("Detail");
        }

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
                    }

                    await ((SqlTransaction)tx).CommitAsync();
                }
                catch
                {
                    await ((SqlTransaction)tx).RollbackAsync();
                    throw;
                }

                // 공유 변경 성공 후 "추가된 공유자"에게 웹푸시
                try
                {
                    if (newlyAdded.Count > 0)
                        await NotifySharesByWebPushAsync(docId, newlyAdded);
                }
                catch (Exception exPush)
                {
                    _log.LogWarning(exPush, "UpdateShares: push notify(shared) failed docId={docId}", docId);
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
                    var isApprover = Convert.ToInt32(r["IsApprover"]) == 1;

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

        // =========================
        // 권한(Recall/Approve/Hold/Reject) 계산
        // =========================
        private async Task<(bool canRecall, bool canApprove, bool canHold, bool canReject)> GetApprovalCapabilitiesAsync(string docId)
        {
            var userId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            if (string.IsNullOrWhiteSpace(docId) || string.IsNullOrWhiteSpace(userId))
                return (false, false, false, false);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
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
            var isFinal =
                st.StartsWith("Approved", StringComparison.OrdinalIgnoreCase) ||
                st.StartsWith("Rejected", StringComparison.OrdinalIgnoreCase);

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

        // =========================
        // 누락된 헬퍼들(분리로 인해 DetailController에 필요)
        // =========================
        private string? GetCurrentUserDisplayNameStrict()
        {
            try
            {
                var uid = User?.FindFirstValue(ClaimTypes.NameIdentifier);
                if (string.IsNullOrWhiteSpace(uid))
                    return User?.Identity?.Name;

                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                using var conn = new SqlConnection(cs);
                conn.Open();

                using var cmd = conn.CreateCommand();
                cmd.CommandText = @"
SELECT TOP 1 NULLIF(LTRIM(RTRIM(DisplayName)), N'')
FROM dbo.UserProfiles
WHERE UserId = @uid;";
                cmd.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = uid });

                var obj = cmd.ExecuteScalar();
                var name = (obj == null || obj == DBNull.Value) ? null : Convert.ToString(obj);
                if (!string.IsNullOrWhiteSpace(name)) return name;

                return User?.Identity?.Name;
            }
            catch
            {
                return User?.Identity?.Name;
            }
        }

        private TimeZoneInfo ResolveCompanyTimeZone()
        {
            // 설정 키는 프로젝트에 맞춰 유지/조정 가능
            // 우선순위: Company:TimeZone -> TimeZone -> 기본(Asia/Seoul)
            var tzId = _cfg["Company:TimeZone"] ?? _cfg["TimeZone"] ?? "Asia/Seoul";
            try
            {
                return TimeZoneInfo.FindSystemTimeZoneById(tzId);
            }
            catch
            {
                // Windows 서버 fallback
                try { return TimeZoneInfo.FindSystemTimeZoneById("Korea Standard Time"); }
                catch { return TimeZoneInfo.Local; }
            }
        }

        private string ToLocalStringFromUtc(DateTime utc)
        {
            var u = utc.Kind == DateTimeKind.Utc ? utc : DateTime.SpecifyKind(utc, DateTimeKind.Utc);
            var tz = ResolveCompanyTimeZone();
            var local = TimeZoneInfo.ConvertTimeFromUtc(u, tz);
            return local.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture);
        }

        private string? ResolveDocOutputPath(string? raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;

            var p = raw.Trim();

            if (!Path.IsPathRooted(p))
            {
                p = p.TrimStart('\\', '/');
                p = Path.Combine(_env.ContentRootPath, p);
            }

            p = p.Replace('/', Path.DirectorySeparatorChar)
                 .Replace('\\', Path.DirectorySeparatorChar);

            return p;
        }

        // 최소 프리뷰(JSON): 값 중심. (스타일/병합이 필요한 경우 DocController의 정식 빌더를 Helper로 이동해 공유하는 것이 최선)
        private string BuildPreviewJsonFromExcel(string fullPath)
        {
            try
            {
                using var wb = new XLWorkbook(fullPath);
                var ws = wb.Worksheets.FirstOrDefault();
                if (ws == null) return "{}";

                var used = ws.RangeUsed();
                if (used == null) return "{}";

                var firstRow = used.RangeAddress.FirstAddress.RowNumber;
                var firstCol = used.RangeAddress.FirstAddress.ColumnNumber;
                var lastRow = used.RangeAddress.LastAddress.RowNumber;
                var lastCol = used.RangeAddress.LastAddress.ColumnNumber;

                var cells = new Dictionary<string, object>(StringComparer.OrdinalIgnoreCase);

                for (int r = firstRow; r <= lastRow; r++)
                {
                    for (int c = firstCol; c <= lastCol; c++)
                    {
                        var cell = ws.Cell(r, c);
                        var v = cell.GetValue<string>();
                        if (string.IsNullOrEmpty(v)) continue;

                        var a1 = XLHelper.GetColumnLetterFromNumber(c) + r.ToString(CultureInfo.InvariantCulture);
                        cells[a1] = new { v };
                    }
                }

                var obj = new
                {
                    sheet = ws.Name,
                    range = new
                    {
                        r1 = firstRow,
                        c1 = firstCol,
                        r2 = lastRow,
                        c2 = lastCol
                    },
                    cells = cells
                };

                return JsonSerializer.Serialize(obj);
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "BuildPreviewJsonFromExcel(min) failed path={path}", fullPath);
                return "{}";
            }
        }

        // (프로젝트에 이미 존재하는 공용 OrgTree 로더가 있다면 그걸 호출하도록 교체)
        private async Task<OrgTreeNode[]> BuildOrgTreeNodesAsync(string langCode)
        {
            // NOTE: 실제 구현은 당신이 가진 DocController 공용 로더를 옮기는 것이 정답.
            // 여기서는 컴파일용 최소 형태로 빈 배열 반환.
            await Task.CompletedTask;
            return Array.Empty<OrgTreeNode>();
        }

        public sealed class ApproveDto
        {
            public string? docId { get; set; }
            public string? action { get; set; }
            public int slot { get; set; } = 1;
            public string? comment { get; set; }
        }

        public sealed class UpdateSharesDto
        {
            public string? DocId { get; set; }
            public List<string>? SelectedRecipientUserIds { get; set; }
        }
    }
}
