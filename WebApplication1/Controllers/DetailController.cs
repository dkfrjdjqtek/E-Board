using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Lib.Net.Http.WebPush;
using Microsoft.AspNetCore.Antiforgery;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Identity.UI.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
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
using System.Security.Cryptography;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using WebApplication1.Controllers;
using WebApplication1.Data;
using WebApplication1.Models;
using WebApplication1.Services;
using static WebApplication1.Controllers.DocController;

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

        // 2026.01.22 Changed: ApproveOrHold에서 WebPushClient 직접 사용(ResolveVapid WebPushClient VapidDetails PushSubscription WebPushException PushPayload SendWebPushToUserIdsAsync) 흐름을 제거하고 기존 _webPushNotifier 방식으로 다음 결재자 및 작성자 알림 발송; 기존 주석 블록은 삭제하지 않고 유지
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

                    // 1-1) 문서 상태를 Recalled 로 변경 (현재 Pending 단계에서만 허용)
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

                    // 1-2) 승인선은 Status 를 그대로 유지하고, 회수 이력만 남김
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

                    // (선택) 회수 알림 정책이 필요하면 여기서 작성자/현재 결재자에게 발송 가능
                    // - 현재 요청 범위는 "승인 다음결재자 / 보류반려 작성자" 이므로 회수는 미적용
                }

                // 클라이언트에는 문서 최종 상태만 전달
                return Json(new { ok = true, docId = dto.docId, status = "Recalled" });
            }

            // ===== 2) 승인 보류 반려 공통 처리 =====
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

                // DB에 상대 경로(App_Data\Docs\... 형태)로 저장되어 있는 경우 ContentRoot 기준으로 보정
                if (!string.IsNullOrWhiteSpace(outXlsx) && !System.IO.Path.IsPathRooted(outXlsx))
                {
                    outXlsx = System.IO.Path.Combine(
                        _env.ContentRootPath,
                        outXlsx.TrimStart('\\', '/')
                    );
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

            // --- 현재 로그인 사용자의 프로필 스냅샷 (결재 캡션 서명용) ---
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

            // (추가) 알림 수신자 계산용 변수
            int? actedStep = null;
            int? nextStepForNotify = null;

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                // --- 2-1) 현재 로그인 사용자가 담당인 Pending 단계 StepOrder 계산 ---
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

                // 현재 사용자가 결재할 수 있는 단계가 없으면 차단
                if (currentStep == null)
                    return Forbid();

                var step = currentStep.Value;
                actedStep = step;

                // --- 2-2) 현재 단계 DocumentApprovals 업데이트 ---
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

                // --- 2-3) Documents.Status 1차 업데이트 (현재 단계 기준) ---
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

                // --- 2-4) 승인일 때만 다음 단계 처리 및 메일 발송 ---
                if (actionLower == "approve")
                {
                    var next = step + 1;
                    nextStepForNotify = next;

                    var toList = await GetNextApproverEmailsFromDbAsync(conn, dto.docId, next);

                    if (toList.Count == 0)
                    {
                        await using var up2a = conn.CreateCommand();
                        up2a.CommandText = @"UPDATE dbo.Documents SET Status = N'Approved' WHERE DocId = @id;";
                        up2a.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        await up2a.ExecuteNonQueryAsync();
                        newStatus = "Approved";

                        // (선택) 최종 승인 알림 정책이 필요하면 여기서 작성자 등에게 발송 가능
                    }
                    else
                    {
                        // 1) 문서 기본 정보 조회 (작성자 이름 문서 제목 작성 시각)
                        string? mailAuthor = null;
                        string? mailDocTitle = null;
                        DateTime? createdAtLocalForMail = null;

                        await using (var q = conn.CreateCommand())
                        {
                            q.CommandText = @"
SELECT TOP 1
       ISNULL(up.DisplayName,
              ISNULL(d.CreatedByName, d.CreatedBy)) AS AuthorName,
       ISNULL(d.TemplateTitle, N'')                 AS TemplateTitle,
       d.CreatedAt                                  AS CreatedAtUtc
FROM dbo.Documents      AS d
LEFT JOIN dbo.UserProfiles AS up
       ON up.UserId = d.CreatedBy
WHERE d.DocId = @id;";
                            q.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });

                            await using var r = await q.ExecuteReaderAsync();
                            if (await r.ReadAsync())
                            {
                                mailAuthor = r["AuthorName"] as string ?? string.Empty;
                                mailDocTitle = r["TemplateTitle"] as string ?? string.Empty;

                                if (r["CreatedAtUtc"] != DBNull.Value)
                                {
                                    var createdAtUtc = (DateTime)r["CreatedAtUtc"];
                                    if (createdAtUtc.Kind != DateTimeKind.Utc)
                                        createdAtUtc = DateTime.SpecifyKind(createdAtUtc, DateTimeKind.Utc);

                                    var tz = ResolveCompanyTimeZone();
                                    createdAtLocalForMail = TimeZoneInfo.ConvertTimeFromUtc(createdAtUtc, tz);
                                }
                            }
                        }

                        if (string.IsNullOrWhiteSpace(mailAuthor))
                            mailAuthor = approverName ?? string.Empty;

                        if (string.IsNullOrWhiteSpace(mailDocTitle))
                            mailDocTitle = dto.docId;

                        // 2) 다음 단계 상태를 PendingA{next} 로 변경
                        await using (var up2 = conn.CreateCommand())
                        {
                            up2.CommandText = @"UPDATE dbo.Documents SET Status = @st WHERE DocId = @id;";
                            up2.Parameters.Add(new SqlParameter("@st", SqlDbType.NVarChar, 20) { Value = $"PendingA{next}" });
                            up2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                            await up2.ExecuteNonQueryAsync();
                        }
                        newStatus = $"PendingA{next}";

                        // 3) 중복 제거된 수신자 목록
                        //var distinctTo = toList
                        //    .Where(e => !string.IsNullOrWhiteSpace(e))
                        //    .Distinct(StringComparer.OrdinalIgnoreCase)
                        //    .ToList();

                        // 4) 상신 메일 포맷 그대로 사용하여 개별 발송
                        // 사장님 지시로 메일 발송 금지
                        //foreach (var to in distinctTo)
                        //{
                        //    try
                        //    {
                        //        var recipientDisplay = GetDisplayNameByEmailStrict(to);
                        //        if (string.IsNullOrWhiteSpace(recipientDisplay))
                        //            recipientDisplay = FallbackNameFromEmail(to);
                        //
                        //        var mail = BuildSubmissionMail(
                        //            mailAuthor!,
                        //            mailDocTitle!,
                        //            dto.docId,
                        //            recipientDisplay,
                        //            createdAtLocalForMail
                        //        );
                        //
                        //        await _emailSender.SendEmailAsync(to, mail.subject, mail.bodyHtml);
                        //
                        //        _log.LogInformation(
                        //            "NextApproverMail SEND OK DocId {DocId} NextStep {NextStep} To {To}",
                        //            dto.docId, next, to);
                        //    }
                        //    catch (Exception exMail)
                        //    {
                        //        _log.LogError(
                        //            exMail,
                        //            "NextApproverMail SEND ERROR DocId {DocId} NextStep {NextStep} To {To}",
                        //            dto.docId, next, to);
                        //    }
                        //}

                        // ===== (변경) 승인 성공 시 다음 결재자에게 웹푸시 알림: 기존 _webPushNotifier 방식 사용 =====
                        // - 다국어: title/body 는 "값"을 발송(리소스 키는 SW에서 해석 로직이 없으면 그대로 노출됨)
                        // - 기존 결재 알림 갱신(교체)로 보내려면 tag 를 badge-approval-pending 으로 고정
                        try
                        {
                            var nextApproverUserIds = await GetNextApproverUserIdsFromDbAsync(conn, dto.docId, next);
                            if (nextApproverUserIds != null && nextApproverUserIds.Count > 0)
                            {
                                var titleText = (_S?["PUSH_Approval_Next_Title"] ?? "E-BOARD").ToString();
                                var bodyText = (_S?["PUSH_Approval_Next_Body"] ?? "결재 대기 문서가 있습니다.").ToString();

                                foreach (var uid in nextApproverUserIds
                                             .Where(x => !string.IsNullOrWhiteSpace(x))
                                             .Select(x => x.Trim())
                                             .Distinct(StringComparer.OrdinalIgnoreCase))
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
                        }
                        catch (Exception exPush)
                        {
                            _log.LogWarning(exPush, "ApproveOrHold: push notify(next approver) failed docId={docId}", dto.docId);
                        }
                    }
                }
                else if (actionLower == "hold" || actionLower == "reject")
                {
                    // ===== (변경) 보류/반려 시 작성자에게 웹푸시 알림: 기존 _webPushNotifier 방식 사용 =====
                    // - 다국어: title/body 는 값으로 발송
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

            var previewJson = BuildPreviewJsonFromExcel(outXlsx!);
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
        // UpdateShares 내부 푸시: 기존 SendWebPushToUserIdsAsync/PushPayload 제거 없이 "호출부"만 _webPushNotifier로 교체
        // (기존 주석 블록 유지 목적: 아래는 UpdateShares의 해당 try 블록에 넣어 교체)
        // ------------------------------------------------------------
        private async Task NotifySharesByWebPushAsync(SqlConnection conn, string docId, List<string> newlyAdded)
        {
            if (conn == null) throw new ArgumentNullException(nameof(conn));
            if (string.IsNullOrWhiteSpace(docId)) return;
            if (newlyAdded == null || newlyAdded.Count == 0) return;

            var targets = newlyAdded
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (targets.Count == 0) return;

            // WebPush VAPID 설정 (키 이름은 프로젝트 설정에 맞게 유지/조정)
            var subject = _cfg["WebPush:Subject"] ?? "mailto:admin@example.com";
            var publicKey = _cfg["WebPush:PublicKey"];
            var privateKey = _cfg["WebPush:PrivateKey"];
            if (string.IsNullOrWhiteSpace(publicKey) || string.IsNullOrWhiteSpace(privateKey)) return;

            var vapid = new Lib.Net.Http.WebPush.Authentication.VapidDetails(subject, publicKey, privateKey);
            var client = new Lib.Net.Http.WebPush.WebPushClient();

            // 대상 사용자들의 구독 엔드포인트 조회
            // 테이블/컬럼명은 당신 프로젝트(예: dbo.WebPushSubscriptions) 기준으로 유지
            var paramNames = new List<string>();
            using var cmd = conn.CreateCommand();
            for (int i = 0; i < targets.Count; i++)
            {
                var pName = "@u" + i.ToString(CultureInfo.InvariantCulture);
                paramNames.Add(pName);
                cmd.Parameters.AddWithValue(pName, targets[i]);
            }

            cmd.CommandText = $@"
SELECT Id, UserId, Endpoint, P256dh, Auth
FROM dbo.WebPushSubscriptions WITH (NOLOCK)
WHERE IsActive = 1
  AND UserId IN ({string.Join(",", paramNames)})
";

            var subs = new List<(long Id, string UserId, string Endpoint, string P256dh, string Auth)>();
            using (var rd = await cmd.ExecuteReaderAsync())
            {
                while (await rd.ReadAsync())
                {
                    subs.Add((
                        rd.GetInt64(0),
                        rd.GetString(1),
                        rd.GetString(2),
                        rd.GetString(3),
                        rd.GetString(4)
                    ));
                }
            }

            if (subs.Count == 0) return;

            // payload는 최소 형태로 구성 (프론트 SW에서 tag/docId 처리 가능)
            var payloadObj = new
            {
                title = "E-BOARD",
                body = "공유 문서가 도착했습니다.",
                tag = "badge-shared",
                docId = docId
            };
            var payload = System.Text.Json.JsonSerializer.Serialize(payloadObj);

            foreach (var s in subs)
            {
                var pushSub = new Lib.Net.Http.WebPush.PushSubscription(
                    s.Endpoint,
                    s.P256dh,
                    s.Auth
                );

                try
                {
                    await client.SendNotificationAsync(pushSub, payload, vapid);
                }
                catch (Lib.Net.Http.WebPush.WebPushException)
                {
                    // 실패 시 비활성화(선택): 필요하면 여기서 IsActive=0 업데이트
                    // 원문 정책에 맞게 처리
                }
                catch
                {
                    // 원문 정책에 맞게 처리
                }
            }
        }

        private static async Task<string?> TryResolveUserIdAsync(SqlConnection conn, SqlTransaction? tx, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;

            await using var cmd = conn.CreateCommand();
            if (tx != null) cmd.Transaction = tx;
            cmd.CommandText = @"
SELECT TOP 1 Id
FROM dbo.AspNetUsers
WHERE Id=@v OR UserName=@v OR NormalizedUserName=UPPER(@v) OR Email=@v OR NormalizedEmail=UPPER(@v);";
            cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            var id = (string?)await cmd.ExecuteScalarAsync();
            if (!string.IsNullOrWhiteSpace(id)) return id;

            await using var cmd2 = conn.CreateCommand();
            if (tx != null) cmd2.Transaction = tx;
            cmd2.CommandText = @"
SELECT TOP 1 COALESCE(p.UserId, u.Id)
FROM dbo.UserProfiles p
LEFT JOIN dbo.AspNetUsers u ON u.Id = p.UserId
WHERE p.UserId=@v OR p.Email=@v OR p.DisplayName=@v OR p.Name=@v;";
            cmd2.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            id = (string?)await cmd2.ExecuteScalarAsync();
            return string.IsNullOrWhiteSpace(id) ? null : id;
        }

        private async Task<IList<object>> LoadApprovalsForDetailAsync(string docId)
        {
            var list = new List<object>();

            var cs = _cfg.GetConnectionString("DefaultConnection");
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            const string sql = @"
SELECT  a.DocId,
        a.StepOrder,
        a.RoleKey,
        a.ApproverValue,
        a.UserId,
        a.Status,
        a.Action,
        a.ActedAt,
        a.ActorName,
        a.ApproverDisplayText,
        up.DisplayName AS ApproverName
FROM    dbo.DocumentApprovals a
LEFT JOIN dbo.UserProfiles up
       ON up.UserId = a.UserId
WHERE   a.DocId = @DocId
ORDER BY a.StepOrder;";

            await using var cmd = new SqlCommand(sql, conn);
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

            await using var rdr = await cmd.ExecuteReaderAsync();
            while (await rdr.ReadAsync())
            {
                var stepOrder = rdr["StepOrder"] is int so ? so : Convert.ToInt32(rdr["StepOrder"]);
                var roleKey = rdr["RoleKey"]?.ToString() ?? string.Empty;
                var approverValue = rdr["ApproverValue"]?.ToString() ?? string.Empty;   // 이메일 등
                var approverName = rdr["ApproverName"]?.ToString() ?? string.Empty;    // UserProfiles.DisplayName
                var storedDisplay = rdr["ApproverDisplayText"]?.ToString() ?? string.Empty;
                var status = rdr["Status"]?.ToString() ?? string.Empty;
                var action = rdr["Action"]?.ToString() ?? string.Empty;
                var actedAtObj = rdr["ActedAt"];
                DateTime? actedAt = actedAtObj is DateTime dt ? dt : null;

                // 담당자 텍스트: 이름 > 기존 ApproverDisplayText > 이메일
                var displayText =
                    !string.IsNullOrWhiteSpace(approverName) ? approverName :
                    !string.IsNullOrWhiteSpace(storedDisplay) ? storedDisplay :
                    approverValue;

                list.Add(new
                {
                    step = stepOrder,
                    roleKey = roleKey,
                    approverValue = approverValue,
                    approverName = approverName,
                    approverDisplayText = displayText,
                    status = status,
                    action = action,
                    actedAt = actedAt,
                    actedAtText = actedAt.HasValue
                                           ? ToLocalStringFromUtc(actedAt.Value)
                                           : string.Empty,
                    actorName = rdr["ActorName"]?.ToString() ?? string.Empty
                });
            }

            return list;
        }

        [HttpGet("Detail/{id?}")]
        // 2025.12.02 Changed: Detail 에서 템플릿 버전별 결재 셀 A1 정보를 조회해 ViewBag.ApprovalCellsJson 으로 전달
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

            // 작성자 DisplayName 우선 해석
            string? creatorDisplayName = creatorNameRaw;

            if (!string.IsNullOrWhiteSpace(creatorUserId))
            {
                await using var pc = conn.CreateCommand();
                pc.CommandText = @"
SELECT TOP 1 DisplayName
FROM dbo.UserProfiles
WHERE UserId = @UserId;";
                pc.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = creatorUserId });

                var objName = await pc.ExecuteScalarAsync();
                if (objName != null && objName != DBNull.Value)
                {
                    creatorDisplayName = Convert.ToString(objName) ?? creatorDisplayName;
                }
            }

            // ---------------- 1-0) 작성/회수 시각(표시용) ----------------
            DateTime? createdAtForUi = null;
            string createdAtText = string.Empty;

            if (createdAtUtc.HasValue)
            {
                createdAtForUi = DateTime.SpecifyKind(createdAtUtc.Value, DateTimeKind.Utc);
                createdAtText = ToLocalStringFromUtc(createdAtForUi.Value);
            }

            DateTime? recalledAtForUi = null;
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
                rcCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });

                var obj = await rcCmd.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value)
                {
                    var dt = (obj is DateTime dtx) ? dtx : Convert.ToDateTime(obj);
                    recalledAtForUi = DateTime.SpecifyKind(dt, DateTimeKind.Utc);
                    recalledAtText = ToLocalStringFromUtc(recalledAtForUi.Value);
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "Detail: load recalled ActedAt failed for {DocId}", id);
            }

            // ---------------- 1-1) TemplateCode -> 템플릿 버전 Id ----------------
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
                tvCmd.Parameters.Add(new SqlParameter("@TemplateCode", SqlDbType.NVarChar, 50) { Value = templateCode });

                var obj = await tvCmd.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value)
                    templateVersionId = Convert.ToInt32(obj);
            }

            // ---------------- 2) 프리뷰 JSON ----------------
            var inputsJson = "{}";
            string previewJson = "{}";

            try
            {
                var raw = outputPath ?? string.Empty;
                var resolved = raw;

                if (!string.IsNullOrWhiteSpace(resolved))
                {
                    if (!System.IO.Path.IsPathRooted(resolved))
                    {
                        resolved = resolved.TrimStart('\\', '/');
                        resolved = System.IO.Path.Combine(_env.ContentRootPath, resolved);
                    }

                    resolved = resolved.Replace('/', System.IO.Path.DirectorySeparatorChar)
                                       .Replace('\\', System.IO.Path.DirectorySeparatorChar);
                }

                var exists = !string.IsNullOrWhiteSpace(resolved) && System.IO.File.Exists(resolved);
                if (!exists)
                {
                    _log.LogWarning(
                        "Detail: output excel not found. DocId={DocId}, OutputPath(raw)={Raw}, OutputPath(resolved)={Resolved}, CWD={Cwd}, ContentRoot={Root}",
                        id, raw, resolved, System.IO.Directory.GetCurrentDirectory(), _env.ContentRootPath
                    );
                }
                else
                {
                    previewJson = BuildPreviewJsonFromExcel(resolved);
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "Detail: BuildPreviewJsonFromExcel failed for {DocId}. OutputPath={OutputPath}", id, outputPath ?? string.Empty);
                previewJson = "{}";
            }

            // ---------------- 3) 승인 정보 DisplayName 조인 ----------------
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

                    var approverDisplayText = r["ApproverDisplayText"] as string ?? string.Empty;

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
                        approverDisplayText,
                        signaturePath = r["SignaturePath"]?.ToString()
                    });
                }
            }

            // ---------------- 3-1) 템플릿 결재 셀 A1 매핑 ----------------
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
                    long byteSize = 0;
                    if (fr["ByteSize"] != DBNull.Value) byteSize = Convert.ToInt64(fr["ByteSize"]);

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

            // ---------------- 3-4) (수정) 공유 대상자 콤보 후보 데이터(OrgTreeNodes) 로드 ----------------
            // 공통 헬퍼(BuildOrgTreeNodesAsync)로 통일 (Create(GET)과 동일 로직 재사용)
            try
            {
                var langCode = "ko";

                // (기존, 빌드 에러) var orgNodes = await BuildOrgTreeNodesAsync(conn, langCode);
                // (수정) BuildOrgTreeNodesAsync의 기존 시그니처에 맞춰 호출
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
                    DateTime? viewedLocal = viewedUtc?.ToLocalTime();
                    string? when = viewedLocal?.ToString("yyyy-MM-dd HH:mm");

                    viewLogs.Add(new
                    {
                        viewedAt = viewedLocal,
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

            ViewBag.CreatorName = creatorDisplayName ?? string.Empty;

            ViewBag.DescriptorJson = descriptorJson ?? "{}";
            ViewBag.InputsJson = inputsJson;
            ViewBag.PreviewJson = previewJson;
            ViewBag.ApprovalsJson = JsonSerializer.Serialize(approvals);
            ViewBag.ViewLogsJson = JsonSerializer.Serialize(viewLogs);
            ViewBag.ApprovalCellsJson = JsonSerializer.Serialize(approvalCells);

            ViewBag.DocumentFilesJson = JsonSerializer.Serialize(docFiles);

            ViewBag.SelectedRecipientUserIds = sharedRecipientUserIds;
            ViewBag.SelectedRecipientUserIdsJson = JsonSerializer.Serialize(sharedRecipientUserIds);

            ViewBag.CreatedAt = createdAtForUi;
            ViewBag.CreatedAtText = createdAtText;
            ViewBag.RecalledAt = recalledAtForUi;
            ViewBag.RecalledAtText = recalledAtText;

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

            // 선택값 정리(본인 제외, 중복 제거)
            var selected = (dto.SelectedRecipientUserIds ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(x => !string.Equals(x, actorId, StringComparison.OrdinalIgnoreCase))
                .ToList();

            // (추가) 알림 대상 계산용: 현재 활성 공유자 vs 선택된 공유자 차집합(추가된 사람만 알림)
            var newlyAdded = new List<string>();

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();
                await using var tx = await conn.BeginTransactionAsync();

                try
                {
                    // 1) 현재 활성 공유자(미회수) 목록 조회
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

                    // 2) 제외된 사람은 revoke 처리 + 로그
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

                    // 3) 선택된 사람은 upsert(기존 Create 로직 그대로) + 로그
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

                // ===== (추가) 공유 변경 성공 후 "추가된 공유자"에게 웹푸시 알림 =====
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
                                //title: "PUSH_Share_Title",
                                //body: "PUSH_Share_Body",
                                //url: "/Doc/Detail?id=" + Uri.EscapeDataString(docId),
                                //tag: "badge-shared"
                                title: "PUSH_SummaryTitle",          // 또는 PUSH_App_Title
                                body: "PUSH_ApprovalShare",        // 작성: "결재 대기 {0}건" 이지만 {0} 없이면 문구만
                                url: "/",                                  // 배지 갱신 성격이면 상세 이동 X
                                tag: "badge-approval-Share"
                            );
                        }
                    }
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

                // StoragePath는 반드시 "상대경로"여야 함 (절대경로 금지 정책)
                // 경로 탐색 방지: 루트/드라이브/상위(..) 차단 + ContentRootPath 결합
                var rel = storagePath.Replace('\\', '/').Trim();
                if (string.IsNullOrWhiteSpace(rel))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                if (rel.StartsWith("/") || rel.Contains(":") || rel.Contains(".."))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                var full = Path.GetFullPath(Path.Combine(_env.ContentRootPath, rel));
                var root = Path.GetFullPath(_env.ContentRootPath);

                // ContentRoot 하위만 허용
                if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                if (!System.IO.File.Exists(full))
                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

                // 스트림 반환
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
            var userName = User.Identity?.Name ?? "";
            var clientIp = HttpContext.Connection.RemoteIpAddress?.ToString() ?? "";
            var userAgent = Request.Headers["User-Agent"].ToString() ?? "";

            // ✅ 핵심: UI 탭이 shared면 ViewerRole을 Shared로 강제(결재자/작성자 여부보다 우선)
            var tab = (Request?.Query["tab"].ToString() ?? string.Empty).Trim();
            var forceShared = string.Equals(tab, "shared", StringComparison.OrdinalIgnoreCase);

            // ViewerRole 계산
            string viewerRole = "Other";

            // 기안자 / 결재자 여부 간단 판별
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

                    // ✅ 우선순위 1: shared 탭에서 열었다면 Shared로 기록
                    // (sharedUnread/IsRead 계산이 ViewerRole='Shared' 로그에 의존하므로 값은 정확히 'Shared'여야 함)
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
            // 필요하면 여기서 DocumentAuditLogs 에도 "View" 액션으로 1줄 남겨도 됨
        }

        // ========== DetailsData (상세 데이터 API) ==========
        [HttpGet("DetailsData")]
        [Produces("application/json")]
        public async Task<IActionResult> DetailsData(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return BadRequest(new { message = "DOC_Err_BadRequest" });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            // ----- 1) 문서 기본 정보 -----
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

            // ----- 2) 결재 정보 DisplayName 조인 -----
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

                    // UserProfiles DisplayName 이 있으면 그 값을 표시용 이름으로 사용
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

            // ----- 3) 결과 반환 -----
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
        private string ResolveTempUploadPath(string fileKey)
        {
            if (string.IsNullOrWhiteSpace(fileKey))
                return string.Empty;

            // path traversal 방지: 파일명만 사용
            var safeKey = Path.GetFileName(fileKey.Trim());

            return Path.Combine(_env.ContentRootPath, "App_Data", "Uploads", "Temp", safeKey);
        }


        [HttpGet("Comments")]
        [Produces("application/json")]
        public async Task<IActionResult> GetComments(
            [FromQuery] string? docId,
            [FromQuery] string? id,
            [FromQuery] string? documentId)
        {
            docId = !string.IsNullOrWhiteSpace(docId)
                ? docId
                : !string.IsNullOrWhiteSpace(id)
                    ? id
                    : documentId;

            if (string.IsNullOrWhiteSpace(docId))
                return Json(new { items = Array.Empty<object>() });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var items = new List<object>();

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            var childMap = new Dictionary<long, int>();
            {
                using var childCmd = conn.CreateCommand();
                childCmd.CommandText = @"
SELECT ParentCommentId, COUNT(*) AS ChildCount
FROM dbo.DocumentComments
WHERE DocId = @DocId AND IsDeleted = 0 AND ParentCommentId IS NOT NULL
GROUP BY ParentCommentId;";
                childCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

                using var childRd = await childCmd.ExecuteReaderAsync();
                while (await childRd.ReadAsync())
                {
                    var pid = Convert.ToInt64(childRd["ParentCommentId"]);
                    var cnt = Convert.ToInt32(childRd["ChildCount"]);
                    childMap[pid] = cnt;
                }
            }

            var cmd = conn.CreateCommand();

            cmd.CommandText = @"
SELECT  c.CommentId,
        c.DocId,
        c.ParentCommentId,
        c.ThreadRootId,
        c.Depth,
        c.Body,
        c.HasAttachment,
        c.IsDeleted,
        c.CreatedBy,
        c.CreatedAt,
        c.UpdatedAt,
        COALESCE(p.DisplayName, u.UserName, u.Email, c.CreatedBy) AS AuthorName
FROM dbo.DocumentComments c
LEFT JOIN dbo.AspNetUsers  u
       ON  u.Id       = c.CreatedBy
       OR  u.UserName = c.CreatedBy
       OR  u.Email    = c.CreatedBy
LEFT JOIN dbo.UserProfiles p
       ON  p.UserId   = u.Id
WHERE c.DocId = @DocId
  AND c.IsDeleted = 0
ORDER BY c.ThreadRootId, c.Depth, c.CreatedAt;";

            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

            var currentUserName = User?.Identity?.Name ?? "";
            var currentUserId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            var isAdmin = User?.IsInRole("Admin") ?? false;

            await using var rd = await cmd.ExecuteReaderAsync();
            while (await rd.ReadAsync())
            {
                var commentId = Convert.ToInt64(rd["CommentId"]);
                var parent = rd["ParentCommentId"] == DBNull.Value
                    ? (long?)null
                    : Convert.ToInt64(rd["ParentCommentId"]);
                var depth = Convert.ToInt32(rd["Depth"]);
                var hasAttachment = Convert.ToBoolean(rd["HasAttachment"]);

                var createdAtUtc = (DateTime)rd["CreatedAt"];
                var createdAtLocal = ToLocalStringFromUtc(
                    DateTime.SpecifyKind(createdAtUtc, DateTimeKind.Utc));

                var createdByRaw = rd["CreatedBy"]?.ToString() ?? "";

                // 2025.12.08 Changed: 예전 데이터(사용자 Id 저장)와 신규 데이터(UserName 저장)를 모두 허용
                var canModify =
                    (!string.IsNullOrWhiteSpace(createdByRaw) &&
                     (string.Equals(createdByRaw, currentUserName, StringComparison.OrdinalIgnoreCase) ||
                      string.Equals(createdByRaw, currentUserId, StringComparison.OrdinalIgnoreCase)))
                    || isAdmin;

                var authorName = rd["AuthorName"]?.ToString() ?? createdByRaw;
                var hasChildren = childMap.ContainsKey(commentId);

                items.Add(new
                {
                    commentId,
                    docId = (string)rd["DocId"],
                    parentCommentId = parent,
                    depth,
                    body = rd["Body"]?.ToString() ?? "",
                    hasAttachment,
                    createdBy = authorName,
                    createdAt = createdAtLocal,
                    canEdit = canModify && !hasChildren,
                    canDelete = canModify && !hasChildren,
                    hasChildren
                });
            }

            return Json(new { items });
        }

        [HttpPost("Comments")]
        [ValidateAntiForgeryToken] // JSON 요청도 EB-CSRF 헤더로 보호
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> PostComment([FromBody] DocCommentPostDto dto)
        {
            try
            {
                if (dto == null || string.IsNullOrWhiteSpace(dto.docId))
                    return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" }, detail = "docId missing" });

                var body = (dto.body ?? "").Trim();
                if (string.IsNullOrWhiteSpace(body))
                    return BadRequest(new { messages = new[] { "DOC_Val_Required" }, detail = "body empty" });

                var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
                var userId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
                var userName = User?.Identity?.Name ?? userId;

                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                int? parentId = dto.parentCommentId;
                int threadRootId;
                int depth;

                if (parentId.HasValue && parentId.Value > 0)
                {
                    var q = conn.CreateCommand();
                    q.CommandText = @"
SELECT ISNULL(ThreadRootId, CommentId) AS RootId, Depth
FROM dbo.DocumentComments
WHERE CommentId = @pid AND DocId = @docId;";
                    q.Parameters.Add(new SqlParameter("@pid", SqlDbType.Int) { Value = parentId.Value });
                    q.Parameters.Add(new SqlParameter("@docId", SqlDbType.NVarChar, 40) { Value = dto.docId });

                    await using var r = await q.ExecuteReaderAsync();
                    if (await r.ReadAsync())
                    {
                        // 2025.12.04 Changed: bigint tinyint 을 object 에서 바로 (int) 캐스팅하던 부분을 Convert 사용으로 변경 InvalidCastException 방지
                        threadRootId = Convert.ToInt32(r["RootId"]);
                        depth = Convert.ToInt32(r["Depth"]) + 1;
                    }
                    else
                    {
                        threadRootId = parentId.Value;
                        depth = 1;
                    }
                }
                else
                {
                    threadRootId = 0;
                    depth = 0;
                    parentId = null;
                }

                bool hasAttachment = dto.files != null && dto.files.Count > 0;

                var cmd = conn.CreateCommand();
                cmd.CommandText = @"
INSERT INTO dbo.DocumentComments
    (DocId, ParentCommentId, ThreadRootId, Depth,
     Body, HasAttachment, IsDeleted,
     CreatedBy, CreatedAt, UpdatedAt)
OUTPUT INSERTED.CommentId
VALUES
    (@DocId, @ParentCommentId, @ThreadRootId, @Depth,
     @Body, @HasAttachment, 0,
     @CreatedBy, SYSUTCDATETIME(), SYSUTCDATETIME());";

                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                cmd.Parameters.Add(new SqlParameter("@ParentCommentId", SqlDbType.Int)
                { Value = (object?)parentId ?? DBNull.Value });
                cmd.Parameters.Add(new SqlParameter("@ThreadRootId", SqlDbType.Int) { Value = threadRootId });
                cmd.Parameters.Add(new SqlParameter("@Depth", SqlDbType.Int) { Value = depth });
                cmd.Parameters.Add(new SqlParameter("@Body", SqlDbType.NVarChar, -1) { Value = body });
                cmd.Parameters.Add(new SqlParameter("@HasAttachment", SqlDbType.Bit) { Value = hasAttachment });
                cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 200) { Value = userName ?? "" });

                var newIdObj = await cmd.ExecuteScalarAsync();
                var newId = (newIdObj is int i) ? i : Convert.ToInt32(newIdObj);

                if (threadRootId == 0 && newId > 0)
                {
                    var u = conn.CreateCommand();
                    u.CommandText = @"UPDATE dbo.DocumentComments SET ThreadRootId = @id WHERE CommentId = @id;";
                    u.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = newId });
                    await u.ExecuteNonQueryAsync();
                }

                return Json(new { ok = true, id = newId });
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "PostComment failed");
                return StatusCode(500, new
                {
                    messages = new[] { "DOC_Err_RequestFailed" },
                    detail = ex.Message
                });
            }
        }

        [HttpDelete("Comments/{id}")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> DeleteComment(int id, string docId)
        {
            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var userName = User?.Identity?.Name ?? "";
            var isAdmin = User?.IsInRole("Admin") ?? false;

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            // 2025.12.04 Added 자식 댓글이 있는지 먼저 확인
            var checkChild = conn.CreateCommand();
            checkChild.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentComments
WHERE ParentCommentId = @id AND DocId = @docId AND IsDeleted = 0;";
            checkChild.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = id });
            checkChild.Parameters.Add(new SqlParameter("@docId", SqlDbType.NVarChar, 40) { Value = docId });

            var childCountObj = await checkChild.ExecuteScalarAsync();
            var childCount = (childCountObj is int ci) ? ci : Convert.ToInt32(childCountObj);
            if (childCount > 0)
            {
                // 하위 답글이 있으면 삭제 불가
                return StatusCode(409, new
                {
                    ok = false,
                    messages = new[] { "DOC_Err_DeleteFailed" },
                    detail = "comment has child replies"
                });
            }

            var cmd = conn.CreateCommand();
            cmd.CommandText = @"
UPDATE dbo.DocumentComments
SET IsDeleted = 1, UpdatedAt = SYSUTCDATETIME()
WHERE CommentId = @id
  AND DocId = @docId
  AND IsDeleted = 0
  AND (@IsAdmin = 1 OR CreatedBy = @UserName);";
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.Int) { Value = id });
            cmd.Parameters.Add(new SqlParameter("@docId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@IsAdmin", SqlDbType.Bit) { Value = isAdmin ? 1 : 0 });
            cmd.Parameters.Add(new SqlParameter("@UserName", SqlDbType.NVarChar, 200) { Value = userName ?? "" });

            var rows = await cmd.ExecuteNonQueryAsync();
            if (rows == 0)
                return Forbid();

            return Json(new { ok = true });
        }

        [HttpPut("Comments/{id}")]
        [ValidateAntiForgeryToken]
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> UpdateComment(int id, [FromBody] DocCommentPostDto dto)
        {
            if (dto == null || string.IsNullOrWhiteSpace(dto.body))
                return BadRequest(new { ok = false, message = "DOC_Val_Required" });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            var userName = User?.Identity?.Name ?? "";

            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            // 2025.12.04 Added 자식 댓글이 있는 경우 수정 차단
            var checkChild = conn.CreateCommand();
            checkChild.CommandText = @"
SELECT COUNT(1)
FROM dbo.DocumentComments
WHERE ParentCommentId = @Id AND DocId = @DocId AND IsDeleted = 0;";
            checkChild.Parameters.Add(new SqlParameter("@Id", SqlDbType.Int) { Value = id });
            checkChild.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId ?? "" });

            var childCountObj = await checkChild.ExecuteScalarAsync();
            var childCount = (childCountObj is int ci) ? ci : Convert.ToInt32(childCountObj);
            if (childCount > 0)
            {
                return StatusCode(409, new
                {
                    ok = false,
                    messages = new[] { "DOC_Err_RequestFailed" },
                    detail = "comment has child replies"
                });
            }

            var cmd = conn.CreateCommand();
            cmd.CommandText = @"
UPDATE dbo.DocumentComments
SET Body = @Body,
    UpdatedAt = SYSUTCDATETIME()
WHERE CommentId = @Id
  AND DocId = @DocId
  AND IsDeleted = 0
  AND CreatedBy = @UserName;";
            cmd.Parameters.Add(new SqlParameter("@Body", SqlDbType.NVarChar, -1) { Value = dto.body.Trim() });
            cmd.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = id });
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId ?? "" });
            cmd.Parameters.Add(new SqlParameter("@UserName", SqlDbType.NVarChar, 200) { Value = userName });

            var rows = await cmd.ExecuteNonQueryAsync();
            if (rows == 0)
                return Forbid();

            return Json(new { ok = true });
        }

        /// <summary>
        /// Detail 화면 진입 시 DocumentViewLogs + DocumentAuditLogs 에 조회 이력 기록
        /// </summary>
        private async Task InsertDocumentViewAndAuditAsync(
            SqlConnection conn,
            string docId,
            string? userId,
            string? userName)
        {
            if (string.IsNullOrWhiteSpace(docId))
                return;

            // 익명/시스템 계정까지 모두 남길지 여부는 필요 시 조건 추가
            var uid = userId ?? string.Empty;
            var uname = userName ?? string.Empty;

            try
            {
                // 1) DocumentViewLogs: 누가 언제 봤는지
                await using (var cmd = conn.CreateCommand())
                {
                    cmd.CommandText = @"
INSERT INTO dbo.DocumentViewLogs (DocId, ViewedAtUtc, UserId, UserName)
VALUES (@DocId, SYSUTCDATETIME(), @UserId, @UserName);";
                    cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450)
                    {
                        Value = string.IsNullOrEmpty(uid) ? DBNull.Value : uid
                    });
                    cmd.Parameters.Add(new SqlParameter("@UserName", SqlDbType.NVarChar, 200)
                    {
                        Value = string.IsNullOrEmpty(uname) ? DBNull.Value : uname
                    });

                    await cmd.ExecuteNonQueryAsync();
                }

                // 2) DocumentAuditLogs: 액션 코드만 간단히 남김 (기존 테이블 구조 기준)
                //    - 컬럼: DocId, UserId, ActionCode, CreatedAtUtc(기본값) 정도만 있다고 가정
                await using (var ac = conn.CreateCommand())
                {
                    ac.CommandText = @"
INSERT INTO dbo.DocumentAuditLogs (DocId, UserId, ActionCode)
VALUES (@DocId, @UserId, @ActionCode);";
                    ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    ac.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450)
                    {
                        Value = string.IsNullOrEmpty(uid) ? DBNull.Value : uid
                    });
                    ac.Parameters.Add(new SqlParameter("@ActionCode", SqlDbType.NVarChar, 50) { Value = "View" });

                    await ac.ExecuteNonQueryAsync();
                }
            }
            catch (Exception ex)
            {
                // 조회 로그 기록 실패해도 화면은 계속 보여야 하므로 Warning 으로만 남김
                _log.LogWarning(ex, "Detail: InsertDocumentViewAndAuditAsync failed for {DocId}", docId);
            }
        }

        // ---------- Helper: DocumentViewLogs INSERT ----------
        private async Task WriteDocumentViewLogAsync(
            SqlConnection conn,
            string docId,
            string viewerId,
            string viewerRole)
        {
            if (string.IsNullOrWhiteSpace(docId) || string.IsNullOrWhiteSpace(viewerId))
                return;

            var ip = HttpContext?.Connection?.RemoteIpAddress?.ToString() ?? string.Empty;
            var ua = HttpContext?.Request?.Headers["User-Agent"].ToString() ?? string.Empty;

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
INSERT INTO dbo.DocumentViewLogs
    (DocId, ViewerId, ViewerRole, ViewedAt, ClientIp, UserAgent)
VALUES
    (@DocId, @ViewerId, @ViewerRole, SYSUTCDATETIME(), @ClientIp, @UserAgent);";

            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@ViewerId", SqlDbType.NVarChar, 450) { Value = viewerId });
            cmd.Parameters.Add(new SqlParameter("@ViewerRole", SqlDbType.NVarChar, 50)
            {
                Value = string.IsNullOrWhiteSpace(viewerRole) ? (object)DBNull.Value : viewerRole
            });
            cmd.Parameters.Add(new SqlParameter("@ClientIp", SqlDbType.NVarChar, 64)
            {
                Value = string.IsNullOrWhiteSpace(ip) ? (object)DBNull.Value : ip
            });
            cmd.Parameters.Add(new SqlParameter("@UserAgent", SqlDbType.NVarChar, 512)
            {
                Value = string.IsNullOrWhiteSpace(ua) ? (object)DBNull.Value : ua
            });

            await cmd.ExecuteNonQueryAsync();
        }

        /// 현재 단계 이후(nextStep)의 승인자를 DocumentApprovals 에서 읽어서
        /// UserId 를 보정하고 메일 주소 목록을 반환합니다.
        private static async Task<List<string>> GetNextApproverEmailsFromDbAsync(SqlConnection conn, string docId, int stepOrder)
        {
            var result = new List<string>();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT ApproverValue
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND StepOrder = @StepOrder
  AND ISNULL(Status, N'Pending') LIKE N'Pending%'
  AND ApproverValue IS NOT NULL
  AND LTRIM(RTRIM(ApproverValue)) <> N'';";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = stepOrder });

            await using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                var v = r[0] as string;
                if (!string.IsNullOrWhiteSpace(v))
                    result.Add(v.Trim());
            }

            return result
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        //   승인란(H4 / I4 / J4 등) 좌표를 찾아오는 헬퍼입니다.
        //   dto.slot(=StepOrder) 값을 넘겨서 호출하게 됩니다.
        private static async Task<(int sheet, int row, int col)?> GetApprovalCellAsync(SqlConnection conn, SqlTransaction? tx, string docId, int stepOrder)
        {
            await using var cmd = conn.CreateCommand();
            if (tx != null) cmd.Transaction = tx;

            cmd.CommandText = @"
SELECT TOP (1) ta.Sheet,
               ta.CellRow,
               ta.CellColumn
FROM dbo.Documents          AS d
JOIN dbo.DocTemplateVersion AS tv ON tv.Id = d.TemplateVersionId
JOIN dbo.DocTemplateApproval AS ta ON ta.VersionId = tv.Id
WHERE d.DocId = @DocId
  AND ta.Slot = @StepOrder;";   // Slot = 현재 승인 단계(1차/2차/3차)

            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = stepOrder });

            await using var r = await cmd.ExecuteReaderAsync();
            if (await r.ReadAsync())
            {
                var sheet = Convert.ToInt32(r["Sheet"]);
                var row = Convert.ToInt32(r["CellRow"]);
                var col = Convert.ToInt32(r["CellColumn"]);
                return (sheet, row, col);
            }

            return null;
        }

        GetApprovalCapabilitiesAsync(string docId)

        {
            var userId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? "";
            if (string.IsNullOrWhiteSpace(docId) || string.IsNullOrWhiteSpace(userId))
                return (false, false, false, false);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? "";
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            string? status = null;
            string? createdBy = null;

            // 1) 문서 상태 / 작성자 조회
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

            // 회수: 작성자 && 아직 Pending 상태일 때만 허용
            var canRecall = isCreator && st.StartsWith("Pending", StringComparison.OrdinalIgnoreCase);

            // 2) 현재 Pending 단계 찾기
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

            // 3) 해당 단계 승인자(UserId, Status) 조회
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

        public sealed class ApproveDto
        {
            public string? docId { get; set; }
            public string? action { get; set; }
            public int slot { get; set; } = 1;
            public string? comment { get; set; }
        }

        public sealed class DocCommentPostDto
        {
            public string? docId { get; set; }
            public int? parentCommentId { get; set; }
            public string? body { get; set; }

            // 업로드 메타 (JS에서 files: [...] 로 보내는 구조와 맞추기)
            public List<CommentFileDto>? files { get; set; }
        }

        public sealed class CommentFileDto
        {
            public string? fileKey { get; set; }
            public string? originalName { get; set; }
            public string? contentType { get; set; }
            public long byteSize { get; set; }
        }

        private async Task<(bool canRecall, bool canApprove, bool canHold, bool canReject)>

            public sealed class UpdateSharesDto
        {
            public string? DocId { get; set; }
            public List<string>? SelectedRecipientUserIds { get; set; }
        }

    }
}
