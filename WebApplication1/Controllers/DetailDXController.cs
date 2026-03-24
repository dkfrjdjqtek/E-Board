// 2026.03.17 Added: DetailDX 구성
// 2026.03.18 Changed: DxTemp 임시파일 생성 제거 — ReadOnly 조회이므로 원본 파일 직접 사용
// 2026.03.23 Added: Cooperate 기능 추가, DetailController 전체 기능 마이그레이션
// 2026.03.24 Added: 협조 스탬프 지원 — DescriptorJson.Cooperations에서 CellA1 파싱, UserProfiles.SignatureRelativePath JOIN
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
        private readonly IWebPushNotifier _webPushNotifier;
        private readonly DocControllerHelper _helper;

        public DetailDXController(
            IStringLocalizer<SharedResource> S,
            IConfiguration cfg,
            IWebHostEnvironment env,
            ILogger<DetailDXController> log,
            IWebPushNotifier webPushNotifier)
        {
            _S = S;
            _cfg = cfg;
            _env = env;
            _log = log;
            _webPushNotifier = webPushNotifier;
            _helper = new DocControllerHelper(_cfg, _env, () => User);
        }

        // ========== DetailDX (GET) ==========
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
LEFT JOIN dbo.UserProfiles up ON up.UserId = a.UserId
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
SELECT Slot, ISNULL(CellA1, A1) AS A1
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
SELECT FileKey, OriginalName, ContentType, ByteSize, UploadedAt
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

            // ========== 협조 셀 맵: DescriptorJson.Cooperations 파싱 ==========
            // ApproverValue(UserId) → A1 주소 매핑
            var cooperationCellMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (!string.IsNullOrWhiteSpace(descriptorJson) && descriptorJson != "{}")
                {
                    using var descDoc = JsonDocument.Parse(descriptorJson);
                    if (descDoc.RootElement.TryGetProperty("Cooperations", out var coopsEl)
                        && coopsEl.ValueKind == JsonValueKind.Array)
                    {
                        foreach (var coop in coopsEl.EnumerateArray())
                        {
                            var av = coop.TryGetProperty("ApproverValue", out var avEl)
                                ? avEl.GetString() ?? string.Empty : string.Empty;
                            var a1 = string.Empty;
                            if (coop.TryGetProperty("Cell", out var cellEl)
                                && cellEl.TryGetProperty("A1", out var a1El))
                                a1 = a1El.GetString() ?? string.Empty;
                            if (!string.IsNullOrWhiteSpace(av) && !string.IsNullOrWhiteSpace(a1))
                                cooperationCellMap[av] = a1;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: DescriptorJson Cooperations parse failed for {DocId}", id);
            }

            // ========== 협조 데이터 로드 ==========
            var cooperations = new List<object>();
            try
            {
                await using var cc = conn.CreateCommand();
                cc.CommandText = @"
SELECT dc.Id,
       dc.UserId,
       dc.ApproverValue,
       dc.Status,
       dc.Action,
       dc.ActedAt,
       dc.ActorName,
       COALESCE(dc.ApproverDisplayText, up.DisplayName, dc.ActorName, dc.ApproverValue) AS ApproverDisplayText,
       up.SignatureRelativePath AS SignaturePath
FROM dbo.DocumentCooperations dc
LEFT JOIN dbo.UserProfiles up ON up.UserId = dc.UserId
WHERE dc.DocId = @DocId
ORDER BY dc.Id;";
                cc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id! });

                await using var cr = await cc.ExecuteReaderAsync();
                while (await cr.ReadAsync())
                {
                    DateTime? acted = cr["ActedAt"] is DateTime dt ? dt : (DateTime?)null;
                    var uid = cr["UserId"]?.ToString() ?? string.Empty;
                    var av = cr["ApproverValue"]?.ToString() ?? string.Empty;

                    // 셀 위치: UserId 우선, 없으면 ApproverValue로 매핑
                    var cellA1 = string.Empty;
                    if (!string.IsNullOrWhiteSpace(uid) && cooperationCellMap.TryGetValue(uid, out var ca1))
                        cellA1 = ca1;
                    else if (!string.IsNullOrWhiteSpace(av) && cooperationCellMap.TryGetValue(av, out var ca2))
                        cellA1 = ca2;

                    cooperations.Add(new
                    {
                        userId = uid,
                        approverValue = av,
                        status = cr["Status"]?.ToString(),
                        action = cr["Action"]?.ToString(),
                        actedAtText = acted.HasValue
                            ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc))
                            : string.Empty,
                        actorName = cr["ActorName"]?.ToString(),
                        approverDisplayText = cr["ApproverDisplayText"]?.ToString() ?? string.Empty,
                        cellA1,
                        signaturePath = cr["SignaturePath"]?.ToString() ?? string.Empty
                    });
                }
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "DetailDX: load DocumentCooperations failed for {DocId}", id);
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
                    if (seen.Add(uid)) sharedRecipientUserIds.Add(uid);
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
       v.ViewedAt, v.ViewerId, v.ViewerRole, v.ClientIp, v.UserAgent,
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
            ViewBag.CooperationsJson = JsonSerializer.Serialize(cooperations);

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
            ViewBag.CanCooperate = caps.canCooperate;

            return View("~/Views/Doc/DetailDX.cshtml");
        }

        // ========== ApproveOrHold ==========
        [HttpPost("ApproveOrHoldDX")]
        [ValidateAntiForgeryToken]
        [Produces("application/json")]
        public async Task<IActionResult> ApproveOrHold([FromBody] ApproveDto dto)
        {
            if (string.IsNullOrWhiteSpace(dto?.docId) || string.IsNullOrWhiteSpace(dto?.action))
                return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" } });

            var actionLower = dto.action!.ToLowerInvariant();

            static bool IsHoldLikeStatus(string? st)
            {
                if (string.IsNullOrWhiteSpace(st)) return false;
                st = st.Trim();
                if (st.StartsWith("PendingHold", StringComparison.OrdinalIgnoreCase)) return true;
                if (st.Equals("OnHold", StringComparison.OrdinalIgnoreCase)) return true;
                if (st.Equals("Hold", StringComparison.OrdinalIgnoreCase)) return true;
                return false;
            }

            // ── 회수 ──────────────────────────────────────────────────────────
            if (actionLower == "recall")
            {
                var csR = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(csR);
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
                    ac.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = _helper.GetCurrentUserDisplayNameStrict() ?? string.Empty });
                    await ac.ExecuteNonQueryAsync();
                }

                return Json(new { ok = true, docId = dto.docId, status = "Recalled" });
            }

            // ── 협조 확인 ─────────────────────────────────────────────────────
            if (actionLower == "cooperate")
            {
                var approverId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
                var approverName = _helper.GetCurrentUserDisplayNameStrict();

                // ── 프로필 스냅샷 (결재 POST와 완전히 동일) ──────────────────
                string? coopDisplayText = null;
                string? coopSignaturePath = null;
                try
                {
                    var csP = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                    await using var connP = new SqlConnection(csP);
                    await connP.OpenAsync();
                    await using var up = connP.CreateCommand();
                    up.CommandText = @"
SELECT TOP (1)
       u.DisplayName,
       COALESCE(pl.Name, pm.Name, N'') AS PositionName,
       COALESCE(dl.Name, dm.Name, N'') AS DepartmentName,
       u.SignatureRelativePath         AS SignaturePath
FROM dbo.UserProfiles AS u
LEFT JOIN dbo.DepartmentMasters   AS dm ON dm.CompCd = u.CompCd AND dm.Id       = u.DepartmentId
LEFT JOIN dbo.DepartmentMasterLoc AS dl ON dl.DepartmentId = dm.Id AND dl.LangCode = N'ko'
LEFT JOIN dbo.PositionMasters     AS pm ON pm.CompCd = u.CompCd AND pm.Id       = u.PositionId
LEFT JOIN dbo.PositionMasterLoc   AS pl ON pl.PositionId   = pm.Id AND pl.LangCode = N'ko'
WHERE u.UserId = @uid;";
                    up.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId });
                    await using var r = await up.ExecuteReaderAsync();
                    if (await r.ReadAsync())
                    {
                        var disp = r["DisplayName"] as string ?? string.Empty;
                        var pos = r["PositionName"] as string ?? string.Empty;
                        var dept = r["DepartmentName"] as string ?? string.Empty;
                        var sig = r["SignaturePath"] as string ?? string.Empty;
                        coopDisplayText = string.Join(" ",
                            new[] { dept, pos, disp }.Where(s => !string.IsNullOrWhiteSpace(s)));
                        coopSignaturePath = string.IsNullOrWhiteSpace(sig) ? null : sig;
                    }
                }
                catch (Exception ex)
                {
                    _log.LogError(ex, "cooperate: profile snapshot failed docId={docId}", dto.docId);
                }

                var csC = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(csC);
                await conn.OpenAsync();

                // ── Pending 협조자인지 확인 ───────────────────────────────────
                int? coopId = null;
                await using (var ck = conn.CreateCommand())
                {
                    ck.CommandText = @"
SELECT TOP 1 Id
FROM dbo.DocumentCooperations
WHERE DocId  = @DocId
  AND UserId = @UserId
  AND ISNULL(Status, N'Pending') = N'Pending';";
                    ck.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    ck.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = approverId });
                    var o = await ck.ExecuteScalarAsync();
                    if (o == null || o == DBNull.Value)
                        return Forbid();
                    coopId = Convert.ToInt32(o);
                }

                // ── 협조 완료 UPDATE (결재 approve와 완전히 동일한 컬럼 기준) ─
                await using (var upC = conn.CreateCommand())
                {
                    upC.CommandText = @"
UPDATE dbo.DocumentCooperations
SET Status              = N'Cooperated',
    Action              = N'Cooperate',
    ActedAt             = SYSUTCDATETIME(),
    ActorName           = @actor,
    ApproverDisplayText = COALESCE(@displayText, ApproverDisplayText)
WHERE Id = @Id;";
                    upC.Parameters.Add(new SqlParameter("@Id", SqlDbType.Int) { Value = coopId!.Value });
                    upC.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = approverName ?? string.Empty });
                    upC.Parameters.Add(new SqlParameter("@displayText", SqlDbType.NVarChar, 400) { Value = (object?)coopDisplayText ?? DBNull.Value });
                    await upC.ExecuteNonQueryAsync();
                }

                return Json(new { ok = true, docId = dto.docId, status = "Cooperated" });
            }

            // ── 승인 / 보류 / 반려 ────────────────────────────────────────────
            string outXlsx = string.Empty;
            try
            {
                var csOut = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var connOut = new SqlConnection(csOut);
                await connOut.OpenAsync();
                await using var cmdOut = connOut.CreateCommand();
                cmdOut.CommandText = @"SELECT TOP (1) OutputPath FROM dbo.Documents WHERE DocId = @DocId;";
                cmdOut.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                var obj = await cmdOut.ExecuteScalarAsync();
                outXlsx = (obj as string) ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(outXlsx) && !System.IO.Path.IsPathRooted(outXlsx))
                    outXlsx = System.IO.Path.Combine(_env.ContentRootPath, outXlsx.TrimStart('\\', '/'));
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ApproveOrHold: output path resolve failed docId={docId}", dto.docId);
            }

            if (string.IsNullOrWhiteSpace(outXlsx) || !System.IO.File.Exists(outXlsx))
                return NotFound(new { messages = new[] { "DOC_Err_DocumentNotFound" } });

            var approverId2 = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
            var approverName2 = _helper.GetCurrentUserDisplayNameStrict();

            string? approverDisplayText = null;
            string? signatureRelativePath = null;

            try
            {
                var csP = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var connP = new SqlConnection(csP);
                await connP.OpenAsync();
                await using var up = connP.CreateCommand();
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
                up.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId2 });

                await using var r = await up.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    var disp = r["DisplayName"] as string ?? string.Empty;
                    var pos = r["PositionName"] as string ?? string.Empty;
                    var dept = r["DepartmentName"] as string ?? string.Empty;
                    var sig = r["SignaturePath"] as string ?? string.Empty;
                    approverDisplayText = string.Join(" ", new[] { dept, pos, disp }.Where(s => !string.IsNullOrWhiteSpace(s)));
                    signatureRelativePath = string.IsNullOrWhiteSpace(sig) ? null : sig;
                }
            }
            catch (Exception ex)
            {
                _log.LogError(ex, "ApproveOrHold: profile snapshot failed docId={docId}", dto.docId);
            }

            string newStatus = "Updated";
            int? actedStep = null;
            int? nextStepForNotify = null;

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();

                string docStatus = string.Empty;
                await using (var ds = conn.CreateCommand())
                {
                    ds.CommandText = @"SELECT TOP (1) Status FROM dbo.Documents WHERE DocId = @DocId;";
                    ds.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    var o = await ds.ExecuteScalarAsync();
                    docStatus = (o as string) ?? string.Empty;
                }

                int? currentStep = null;
                await using (var findStep = conn.CreateCommand())
                {
                    findStep.CommandText = @"
SELECT TOP (1) StepOrder
FROM dbo.DocumentApprovals
WHERE DocId = @DocId
  AND ISNULL(Status,N'Pending') LIKE N'Pending%'
  AND (UserId = @UserId OR ApproverValue = @UserId)
ORDER BY StepOrder;";
                    findStep.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    findStep.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = approverId2 });
                    var stepObj = await findStep.ExecuteScalarAsync();
                    if (stepObj != null && stepObj != DBNull.Value)
                        currentStep = Convert.ToInt32(stepObj);
                }

                if (currentStep == null)
                    return Forbid();

                var step = currentStep.Value;
                actedStep = step;

                if (actionLower == "hold")
                {
                    string apprStatus = string.Empty;
                    await using (var asel = conn.CreateCommand())
                    {
                        asel.CommandText = @"
SELECT TOP (1) Status
FROM dbo.DocumentApprovals
WHERE DocId = @DocId AND StepOrder = @Step;";
                        asel.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        asel.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = step });
                        var ao = await asel.ExecuteScalarAsync();
                        apprStatus = (ao as string) ?? string.Empty;
                    }
                    if (IsHoldLikeStatus(apprStatus) || IsHoldLikeStatus(docStatus))
                        return Conflict(new { messages = new[] { "DOC_Err_BadRequest" } });
                }

                await using (var u = conn.CreateCommand())
                {
                    u.CommandText = @"
UPDATE dbo.DocumentApprovals
SET Action              = @act,
    ActedAt             = SYSUTCDATETIME(),
    ActorName           = @actor,
    Status              = CASE
                             WHEN @act = N'approve' THEN N'Approved'
                             WHEN @act = N'hold'    THEN N'PendingHold'
                             WHEN @act = N'reject'  THEN N'Rejected'
                             ELSE ISNULL(Status, N'Updated')
                          END,
    UserId              = COALESCE(UserId, @uid),
    ApproverDisplayText = COALESCE(@displayText, ApproverDisplayText),
    SignaturePath       = COALESCE(@sigPath, SignaturePath)
WHERE DocId = @id AND StepOrder = @step;";
                    u.Parameters.Add(new SqlParameter("@act", SqlDbType.NVarChar, 20) { Value = actionLower });
                    u.Parameters.Add(new SqlParameter("@actor", SqlDbType.NVarChar, 200) { Value = approverName2 ?? string.Empty });
                    u.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = approverId2 });
                    u.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                    u.Parameters.Add(new SqlParameter("@step", SqlDbType.Int) { Value = step });
                    u.Parameters.Add(new SqlParameter("@displayText", SqlDbType.NVarChar, 400) { Value = (object?)approverDisplayText ?? DBNull.Value });
                    u.Parameters.Add(new SqlParameter("@sigPath", SqlDbType.NVarChar, 400) { Value = (object?)signatureRelativePath ?? DBNull.Value });
                    await u.ExecuteNonQueryAsync();
                }

                newStatus = actionLower switch
                {
                    "approve" => $"ApprovedA{step}",
                    "hold" => $"PendingHoldA{step}",
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

                if (actionLower == "approve")
                {
                    var next = step + 1;
                    nextStepForNotify = next;
                    var toList = await GetNextApproverEmailsFromDbAsync(conn, dto.docId!, next);

                    if (toList.Count == 0)
                    {
                        await using var up2a = conn.CreateCommand();
                        up2a.CommandText = @"UPDATE dbo.Documents SET Status = N'Approved' WHERE DocId = @id;";
                        up2a.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = dto.docId });
                        await up2a.ExecuteNonQueryAsync();
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

                        try
                        {
                            var nextIds = await GetNextApproverUserIdsFromDbAsync(conn, dto.docId!, next);
                            await DocControllerHelper.SendApprovalPendingBadgeAsync(
                                notifier: _webPushNotifier, S: _S, conn: conn,
                                targetUserIds: nextIds ?? new List<string>(),
                                url: "/", tag: "badge-approval-pending");
                        }
                        catch (Exception exPush)
                        {
                            _log.LogWarning(exPush, "ApproveOrHold: push notify failed docId={docId}", dto.docId);
                        }
                    }
                }
                else if (actionLower == "hold" || actionLower == "reject")
                {
                    try
                    {
                        var authorId = await GetDocumentAuthorUserIdAsync(conn, dto.docId!);
                        if (!string.IsNullOrWhiteSpace(authorId)
                            && !string.Equals(authorId, approverId2, StringComparison.OrdinalIgnoreCase))
                        {
                            var titleText = (_S?["PUSH_SummaryTitle"] ?? "PUSH_SummaryTitle").ToString();
                            var bodyText = actionLower == "hold"
                                ? (_S?["PUSH_ApprovalHold"] ?? "PUSH_ApprovalHold").ToString()
                                : (_S?["PUSH_ApprovalReject"] ?? "PUSH_ApprovalReject").ToString();
                            var tag = actionLower == "hold" ? "approval-author-hold" : "approval-author-reject";
                            await _webPushNotifier.SendToUserIdAsync(
                                userId: authorId.Trim(), title: titleText, body: bodyText,
                                url: "/Doc/DetailDX?id=" + Uri.EscapeDataString(dto.docId!), tag: tag);
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

            var previewJson2 = BuildPreviewJsonFromExcel(outXlsx);
            return Json(new { ok = true, docId = dto.docId, status = newStatus, previewJson = previewJson2 });
        }

        // ========== UpdateShares ==========
        [HttpPost("UpdateShares")]
        [ValidateAntiForgeryToken]
        [Consumes("application/json")]
        [Produces("application/json")]
        public async Task<IActionResult> UpdateShares([FromBody] UpdateSharesDto dto)
        {
            if (dto is null || string.IsNullOrWhiteSpace(dto.DocId))
                return BadRequest(new { ok = false, messages = new[] { "DOC_Err_SaveFailed" }, stage = "arg" });

            var docId = dto.DocId!.Trim();
            var actorId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;

            var selected = (dto.SelectedRecipientUserIds ?? new List<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Where(x => !string.Equals(x, actorId, StringComparison.OrdinalIgnoreCase))
                .ToList();

            var newlyAdded = new List<string>();

            try
            {
                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
                await using var conn = new SqlConnection(cs);
                await conn.OpenAsync();
                await using var tx = await conn.BeginTransactionAsync();

                try
                {
                    var currentActive = new List<string>();
                    await using (var sel = conn.CreateCommand())
                    {
                        sel.Transaction = (SqlTransaction)tx;
                        sel.CommandText = @"SELECT UserId FROM dbo.DocumentShares WHERE DocId = @DocId AND ISNULL(IsRevoked,0) = 0;";
                        sel.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                        await using var r = await sel.ExecuteReaderAsync();
                        while (await r.ReadAsync())
                        {
                            var uid = (r["UserId"]?.ToString() ?? string.Empty).Trim();
                            if (!string.IsNullOrWhiteSpace(uid)) currentActive.Add(uid);
                        }
                    }

                    newlyAdded = selected
                        .Where(uid => !currentActive.Any(x => string.Equals(x, uid, StringComparison.OrdinalIgnoreCase)))
                        .Distinct(StringComparer.OrdinalIgnoreCase).ToList();

                    var toRevoke = currentActive
                        .Where(uid => !selected.Any(x => string.Equals(x, uid, StringComparison.OrdinalIgnoreCase)))
                        .ToList();

                    foreach (var targetUserId in toRevoke)
                    {
                        await using (var cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = (SqlTransaction)tx;
                            cmd.CommandText = @"UPDATE dbo.DocumentShares SET IsRevoked = 1, ExpireAt = SYSUTCDATETIME() WHERE DocId = @DocId AND UserId = @UserId;";
                            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                            await cmd.ExecuteNonQueryAsync();
                        }
                        await using (var logCmd = conn.CreateCommand())
                        {
                            logCmd.Transaction = (SqlTransaction)tx;
                            logCmd.CommandText = @"INSERT INTO dbo.DocumentShareLogs (DocId,ActorId,ChangeCode,TargetUserId,BeforeJson,AfterJson,ChangedAt) VALUES (@DocId,@ActorId,@ChangeCode,@TargetUserId,NULL,@AfterJson,SYSUTCDATETIME());";
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
    UPDATE dbo.DocumentShares SET IsRevoked = 0, ExpireAt = NULL WHERE DocId=@DocId AND UserId=@UserId;
END
ELSE
BEGIN
    INSERT INTO dbo.DocumentShares (DocId,UserId,AccessRole,ExpireAt,IsRevoked,CreatedBy,CreatedAt)
    VALUES (@DocId,@UserId,'Commenter',NULL,0,@CreatedBy,SYSUTCDATETIME());
END";
                            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.VarChar, 40) { Value = docId });
                            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = targetUserId });
                            cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 64) { Value = actorId });
                            await cmd.ExecuteNonQueryAsync();
                        }
                        await using (var logCmd = conn.CreateCommand())
                        {
                            logCmd.Transaction = (SqlTransaction)tx;
                            logCmd.CommandText = @"INSERT INTO dbo.DocumentShareLogs (DocId,ActorId,ChangeCode,TargetUserId,BeforeJson,AfterJson,ChangedAt) VALUES (@DocId,@ActorId,@ChangeCode,@TargetUserId,NULL,@AfterJson,SYSUTCDATETIME());";
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

                try
                {
                    var ids = newlyAdded.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                    if (ids.Count > 0)
                    {
                        await using var connPush = new SqlConnection(_cfg.GetConnectionString("DefaultConnection") ?? string.Empty);
                        await connPush.OpenAsync();
                        await DocControllerHelper.SendSharedUnreadBadgeAsync(
                            notifier: _webPushNotifier, S: _S, conn: connPush,
                            targetUserIds: ids, url: "/", tag: "badge-shared");
                    }
                }
                catch (Exception exPush)
                {
                    _log.LogWarning(exPush, "UpdateShares: push notify failed docId={docId}", docId);
                }

                return Json(new { ok = true, docId, selectedCount = selected.Count, addedNotifyCount = newlyAdded.Count });
            }
            catch (Exception ex)
            {
                _log.LogWarning(ex, "UpdateShares failed docId={docId}", docId);
                return BadRequest(new { ok = false, messages = new[] { "DOC_Err_SaveFailed" }, stage = "db", detail = ex.Message });
            }
        }
        // ========== Download ==========
        //        [HttpGet("Download/{FileKey}")]
        //        public async Task<IActionResult> Download([FromRoute] string FileKey)
        //        {
        //            if (string.IsNullOrWhiteSpace(FileKey))
        //                return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

        //            try
        //            {
        //                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
        //                await using var conn = new SqlConnection(cs);
        //                await conn.OpenAsync();
        //                await using var cmd = conn.CreateCommand();
        //                cmd.CommandText = @"
        //SELECT TOP (1) DocId, FileKey, OriginalName, StoragePath, ContentType, ByteSize
        //FROM dbo.DocumentFiles
        //WHERE FileKey = @FileKey
        //ORDER BY UploadedAt DESC, FileKey DESC;";
        //                cmd.Parameters.Add(new SqlParameter("@FileKey", SqlDbType.NVarChar, 200) { Value = FileKey });

        //                await using var r = await cmd.ExecuteReaderAsync(CommandBehavior.SingleRow);
        //                if (!await r.ReadAsync())
        //                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

        //                var storagePath = (r["StoragePath"] as string) ?? string.Empty;
        //                var originalName = (r["OriginalName"] as string) ?? FileKey;
        //                var contentType = (r["ContentType"] as string) ?? "application/octet-stream";

        //                var rel = storagePath.Replace('\\', '/').Trim();
        //                if (string.IsNullOrWhiteSpace(rel) || rel.StartsWith("/") || rel.Contains(":") || rel.Contains(".."))
        //                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

        //                var full = Path.GetFullPath(Path.Combine(_env.ContentRootPath, rel));
        //                var root = Path.GetFullPath(_env.ContentRootPath);
        //                if (!full.StartsWith(root, StringComparison.OrdinalIgnoreCase) || !System.IO.File.Exists(full))
        //                    return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });

        //                var stream = new FileStream(full, FileMode.Open, FileAccess.Read, FileShare.Read);
        //                return File(stream, contentType, originalName, enableRangeProcessing: true);
        //            }
        //            catch
        //            {
        //                return NotFound(new { messages = new[] { "DOC_File_Err_NotFound" } });
        //            }
        //        }

        // ========== Comments ==========
        //        [HttpGet("Comments")]
        //        [Produces("application/json")]
        //        public async Task<IActionResult> GetComments([FromQuery] string? docId, [FromQuery] string? id, [FromQuery] string? documentId)
        //        {
        //            docId = !string.IsNullOrWhiteSpace(docId) ? docId
        //                  : !string.IsNullOrWhiteSpace(id) ? id
        //                  : documentId;

        //            if (string.IsNullOrWhiteSpace(docId))
        //                return Json(new { items = Array.Empty<object>() });

        //            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
        //            var items = new List<object>();

        //            await using var conn = new SqlConnection(cs);
        //            await conn.OpenAsync();

        //            var childMap = new Dictionary<long, int>();
        //            await using (var childCmd = conn.CreateCommand())
        //            {
        //                childCmd.CommandText = @"
        //SELECT ParentCommentId, COUNT(*) AS ChildCount
        //FROM dbo.DocumentComments
        //WHERE DocId = @DocId AND IsDeleted = 0 AND ParentCommentId IS NOT NULL
        //GROUP BY ParentCommentId;";
        //                childCmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
        //                await using var childRd = await childCmd.ExecuteReaderAsync();
        //                while (await childRd.ReadAsync())
        //                    childMap[Convert.ToInt64(childRd["ParentCommentId"])] = Convert.ToInt32(childRd["ChildCount"]);
        //            }

        //            await using var cmd = conn.CreateCommand();
        //            cmd.CommandText = @"
        //;WITH Tree AS
        //(
        //    SELECT c.CommentId, c.DocId, c.ParentCommentId, c.ThreadRootId, c.Depth,
        //           c.Body, c.HasAttachment, c.IsDeleted, c.CreatedBy, c.CreatedAt, c.UpdatedAt,
        //           CAST(RIGHT(REPLICATE('0',20)+CONVERT(varchar(20),c.CommentId),20) AS nvarchar(max)) AS SortPath
        //    FROM dbo.DocumentComments c
        //    WHERE c.DocId = @DocId AND c.IsDeleted = 0 AND c.ParentCommentId IS NULL
        //    UNION ALL
        //    SELECT ch.CommentId, ch.DocId, ch.ParentCommentId, ch.ThreadRootId, ch.Depth,
        //           ch.Body, ch.HasAttachment, ch.IsDeleted, ch.CreatedBy, ch.CreatedAt, ch.UpdatedAt,
        //           CAST(t.SortPath+N'/'+RIGHT(REPLICATE('0',20)+CONVERT(varchar(20),ch.CommentId),20) AS nvarchar(max))
        //    FROM dbo.DocumentComments ch
        //    INNER JOIN Tree t ON t.CommentId = ch.ParentCommentId
        //    WHERE ch.DocId = @DocId AND ch.IsDeleted = 0
        //)
        //SELECT t.CommentId, t.DocId, t.ParentCommentId, t.ThreadRootId, t.Depth,
        //       t.Body, t.HasAttachment, t.IsDeleted, t.CreatedBy, t.CreatedAt, t.UpdatedAt,
        //       COALESCE(p.DisplayName, u.UserName, u.Email, t.CreatedBy) AS AuthorName
        //FROM Tree t
        //LEFT JOIN dbo.AspNetUsers  u ON u.Id = t.CreatedBy OR u.UserName = t.CreatedBy OR u.Email = t.CreatedBy
        //LEFT JOIN dbo.UserProfiles p ON p.UserId = u.Id
        //ORDER BY t.SortPath
        //OPTION (MAXRECURSION 100);";
        //            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });

        //            var currentUserId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
        //            var currentUserName = User?.Identity?.Name ?? string.Empty;
        //            var isAdmin = User?.IsInRole("Admin") ?? false;

        //            await using var rd = await cmd.ExecuteReaderAsync();
        //            while (await rd.ReadAsync())
        //            {
        //                var commentId = Convert.ToInt64(rd["CommentId"]);
        //                var parent = rd["ParentCommentId"] == DBNull.Value ? (long?)null : Convert.ToInt64(rd["ParentCommentId"]);
        //                var depth = Convert.ToInt32(rd["Depth"]);
        //                var hasAttachment = Convert.ToBoolean(rd["HasAttachment"]);
        //                var createdAtUtc = (DateTime)rd["CreatedAt"];
        //                var createdAtLocal = _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(createdAtUtc, DateTimeKind.Utc));
        //                var createdByRaw = rd["CreatedBy"]?.ToString() ?? string.Empty;
        //                var canModify = (!string.IsNullOrWhiteSpace(createdByRaw) &&
        //                                    (string.Equals(createdByRaw, currentUserName, StringComparison.OrdinalIgnoreCase) ||
        //                                     string.Equals(createdByRaw, currentUserId, StringComparison.OrdinalIgnoreCase)))
        //                                   || isAdmin;
        //                var authorName = rd["AuthorName"]?.ToString() ?? createdByRaw;
        //                var hasChildren = childMap.ContainsKey(commentId);

        //                items.Add(new
        //                {
        //                    commentId,
        //                    docId = (string)rd["DocId"],
        //                    parentCommentId = parent,
        //                    depth,
        //                    body = rd["Body"]?.ToString() ?? string.Empty,
        //                    hasAttachment,
        //                    createdBy = authorName,
        //                    createdAt = createdAtLocal,
        //                    canEdit = canModify && !hasChildren,
        //                    canDelete = canModify && !hasChildren,
        //                    hasChildren
        //                });
        //            }

        //            return Json(new { items });
        //        }

        //        [HttpPost("Comments")]
        //        [ValidateAntiForgeryToken]
        //        [Consumes("application/json")]
        //        [Produces("application/json")]
        //        public async Task<IActionResult> PostComment([FromBody] DocCommentPostDto dto)
        //        {
        //            try
        //            {
        //                if (dto == null || string.IsNullOrWhiteSpace(dto.docId))
        //                    return BadRequest(new { messages = new[] { "DOC_Err_BadRequest" }, detail = "docId missing" });

        //                var body = (dto.body ?? string.Empty).Trim();
        //                if (string.IsNullOrWhiteSpace(body))
        //                    return BadRequest(new { messages = new[] { "DOC_Val_Required" }, detail = "body empty" });

        //                var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
        //                var currentUserId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
        //                var currentUserName = User?.Identity?.Name ?? string.Empty;
        //                var createdByStore = !string.IsNullOrWhiteSpace(currentUserId) ? currentUserId : currentUserName;

        //                await using var conn = new SqlConnection(cs);
        //                await conn.OpenAsync();

        //                long? parentId = dto.parentCommentId.HasValue ? (long?)dto.parentCommentId.Value : null;
        //                long threadRootId;
        //                int depth;

        //                if (parentId.HasValue && parentId.Value > 0)
        //                {
        //                    var q = conn.CreateCommand();
        //                    q.CommandText = @"SELECT ISNULL(ThreadRootId,CommentId) AS RootId, Depth FROM dbo.DocumentComments WHERE CommentId = @pid AND DocId = @docId;";
        //                    q.Parameters.Add(new SqlParameter("@pid", SqlDbType.BigInt) { Value = parentId.Value });
        //                    q.Parameters.Add(new SqlParameter("@docId", SqlDbType.NVarChar, 40) { Value = dto.docId });
        //                    await using var r = await q.ExecuteReaderAsync();
        //                    if (await r.ReadAsync())
        //                    {
        //                        threadRootId = Convert.ToInt64(r["RootId"]);
        //                        depth = Convert.ToInt32(r["Depth"]) + 1;
        //                    }
        //                    else { threadRootId = parentId.Value; depth = 1; }
        //                }
        //                else { threadRootId = 0; depth = 0; parentId = null; }

        //                bool hasAttachment = dto.files != null && dto.files.Count > 0;

        //                var cmd = conn.CreateCommand();
        //                cmd.CommandText = @"
        //INSERT INTO dbo.DocumentComments
        //    (DocId,ParentCommentId,ThreadRootId,Depth,Body,HasAttachment,IsDeleted,CreatedBy,CreatedAt,UpdatedAt)
        //OUTPUT INSERTED.CommentId
        //VALUES
        //    (@DocId,@ParentCommentId,@ThreadRootId,@Depth,@Body,@HasAttachment,0,@CreatedBy,SYSUTCDATETIME(),SYSUTCDATETIME());";
        //                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId });
        //                cmd.Parameters.Add(new SqlParameter("@ParentCommentId", SqlDbType.BigInt) { Value = (object?)parentId ?? DBNull.Value });
        //                cmd.Parameters.Add(new SqlParameter("@ThreadRootId", SqlDbType.BigInt) { Value = threadRootId });
        //                cmd.Parameters.Add(new SqlParameter("@Depth", SqlDbType.Int) { Value = depth });
        //                cmd.Parameters.Add(new SqlParameter("@Body", SqlDbType.NVarChar, -1) { Value = body });
        //                cmd.Parameters.Add(new SqlParameter("@HasAttachment", SqlDbType.Bit) { Value = hasAttachment });
        //                cmd.Parameters.Add(new SqlParameter("@CreatedBy", SqlDbType.NVarChar, 200) { Value = createdByStore });

        //                var newIdObj = await cmd.ExecuteScalarAsync();
        //                var newId = Convert.ToInt64(newIdObj);

        //                if (threadRootId == 0 && newId > 0)
        //                {
        //                    var u = conn.CreateCommand();
        //                    u.CommandText = @"UPDATE dbo.DocumentComments SET ThreadRootId = @id WHERE CommentId = @id;";
        //                    u.Parameters.Add(new SqlParameter("@id", SqlDbType.BigInt) { Value = newId });
        //                    await u.ExecuteNonQueryAsync();
        //                }

        //                return Json(new { ok = true, id = newId });
        //            }
        //            catch (Exception ex)
        //            {
        //                _log.LogError(ex, "PostComment failed");
        //                return StatusCode(500, new { messages = new[] { "DOC_Err_RequestFailed" }, detail = ex.Message });
        //            }
        //        }

        //        [HttpDelete("Comments/{id}")]
        //        [ValidateAntiForgeryToken]
        //        [Produces("application/json")]
        //        public async Task<IActionResult> DeleteComment(long id, [FromQuery] string docId)
        //        {
        //            if (string.IsNullOrWhiteSpace(docId))
        //                return BadRequest(new { ok = false, messages = new[] { "DOC_Err_BadRequest" }, detail = "docId missing" });

        //            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
        //            var userName = User?.Identity?.Name ?? string.Empty;
        //            var userId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
        //            var isAdmin = User?.IsInRole("Admin") ?? false;

        //            await using var conn = new SqlConnection(cs);
        //            await conn.OpenAsync();

        //            var checkChild = conn.CreateCommand();
        //            checkChild.CommandText = @"SELECT COUNT(1) FROM dbo.DocumentComments WHERE ParentCommentId = @id AND DocId = @docId AND IsDeleted = 0;";
        //            checkChild.Parameters.Add(new SqlParameter("@id", SqlDbType.BigInt) { Value = id });
        //            checkChild.Parameters.Add(new SqlParameter("@docId", SqlDbType.NVarChar, 40) { Value = docId });
        //            var childCount = Convert.ToInt32(await checkChild.ExecuteScalarAsync());
        //            if (childCount > 0)
        //                return StatusCode(409, new { ok = false, messages = new[] { "DOC_Err_DeleteFailed" }, detail = "comment has child replies" });

        //            var cmd = conn.CreateCommand();
        //            cmd.CommandText = @"
        //UPDATE dbo.DocumentComments
        //SET IsDeleted = 1, UpdatedAt = SYSUTCDATETIME()
        //WHERE CommentId = @id AND DocId = @docId AND IsDeleted = 0
        //  AND (@IsAdmin = 1 OR CreatedBy = @UserName OR CreatedBy = @UserId);";
        //            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.BigInt) { Value = id });
        //            cmd.Parameters.Add(new SqlParameter("@docId", SqlDbType.NVarChar, 40) { Value = docId });
        //            cmd.Parameters.Add(new SqlParameter("@IsAdmin", SqlDbType.Bit) { Value = isAdmin ? 1 : 0 });
        //            cmd.Parameters.Add(new SqlParameter("@UserName", SqlDbType.NVarChar, 200) { Value = userName });
        //            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });

        //            var rows = await cmd.ExecuteNonQueryAsync();
        //            if (rows == 0) return Forbid();
        //            return Json(new { ok = true });
        //        }

        //        [HttpPut("Comments/{id}")]
        //        [ValidateAntiForgeryToken]
        //        [Consumes("application/json")]
        //        [Produces("application/json")]
        //        public async Task<IActionResult> UpdateComment(long id, [FromBody] DocCommentPostDto dto)
        //        {
        //            if (dto == null || string.IsNullOrWhiteSpace(dto.docId) || string.IsNullOrWhiteSpace(dto.body))
        //                return BadRequest(new { ok = false, message = "DOC_Val_Required" });

        //            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
        //            var userName = User?.Identity?.Name ?? string.Empty;
        //            var userId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
        //            var isAdmin = User?.IsInRole("Admin") ?? false;

        //            await using var conn = new SqlConnection(cs);
        //            await conn.OpenAsync();

        //            var checkChild = conn.CreateCommand();
        //            checkChild.CommandText = @"SELECT COUNT(1) FROM dbo.DocumentComments WHERE ParentCommentId = @Id AND DocId = @DocId AND IsDeleted = 0;";
        //            checkChild.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = id });
        //            checkChild.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId ?? string.Empty });
        //            var childCount = Convert.ToInt32(await checkChild.ExecuteScalarAsync());
        //            if (childCount > 0)
        //                return StatusCode(409, new { ok = false, messages = new[] { "DOC_Err_RequestFailed" }, detail = "comment has child replies" });

        //            var cmd = conn.CreateCommand();
        //            cmd.CommandText = @"
        //UPDATE dbo.DocumentComments
        //SET Body = @Body, UpdatedAt = SYSUTCDATETIME()
        //WHERE CommentId = @Id AND DocId = @DocId AND IsDeleted = 0
        //  AND (
        //        @IsAdmin = 1
        //        OR CreatedBy = @UserId
        //        OR CreatedBy = @UserName
        //        OR EXISTS (
        //            SELECT 1 FROM dbo.AspNetUsers u
        //            WHERE (u.Id = CreatedBy OR u.UserName = CreatedBy OR u.Email = CreatedBy)
        //              AND u.Id = @UserId
        //        )
        //      );";
        //            cmd.Parameters.Add(new SqlParameter("@Body", SqlDbType.NVarChar, -1) { Value = dto.body!.Trim() });
        //            cmd.Parameters.Add(new SqlParameter("@Id", SqlDbType.BigInt) { Value = id });
        //            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = dto.docId ?? string.Empty });
        //            cmd.Parameters.Add(new SqlParameter("@IsAdmin", SqlDbType.Bit) { Value = isAdmin ? 1 : 0 });
        //            cmd.Parameters.Add(new SqlParameter("@UserName", SqlDbType.NVarChar, 200) { Value = userName });
        //            cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = userId });

        //            var rows = await cmd.ExecuteNonQueryAsync();
        //            if (rows == 0) return Forbid();
        //            return Json(new { ok = true });
        //        }

        // ========== DetailsData ==========
        [HttpGet("DetailsData")]
        [Produces("application/json")]
        public async Task<IActionResult> DetailsData(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
                return BadRequest(new { message = "DOC_Err_BadRequest" });

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT d.DocId, d.TemplateCode, d.TemplateTitle, d.Status, d.DescriptorJson, d.OutputPath, d.CreatedAt
FROM dbo.Documents d WHERE d.DocId = @id;";
            cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = id });

            string? descriptor = null, output = null, title = null, code = null, status = null;
            DateTime? createdAt = null;

            await using (var rd = await cmd.ExecuteReaderAsync())
            {
                if (await rd.ReadAsync())
                {
                    code = rd["TemplateCode"] as string ?? string.Empty;
                    title = rd["TemplateTitle"] as string ?? string.Empty;
                    status = rd["Status"] as string ?? string.Empty;
                    descriptor = rd["DescriptorJson"] as string ?? "{}";
                    output = rd["OutputPath"] as string ?? string.Empty;
                    if (!rd.IsDBNull(rd.GetOrdinal("CreatedAt")))
                        createdAt = (DateTime)rd["CreatedAt"];
                }
                else return NotFound(new { message = "DOC_Err_DocumentNotFound" });
            }

            var resolvedOutput = output ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(resolvedOutput) && !System.IO.Path.IsPathRooted(resolvedOutput))
                resolvedOutput = System.IO.Path.Combine(_env.ContentRootPath, resolvedOutput.TrimStart('\\', '/'));

            var preview = string.IsNullOrWhiteSpace(resolvedOutput) || !System.IO.File.Exists(resolvedOutput)
                ? "{}" : BuildPreviewJsonFromExcel(resolvedOutput);

            var approvals = new List<object>();
            await using (var ac = conn.CreateCommand())
            {
                ac.CommandText = @"
SELECT a.StepOrder, a.RoleKey, a.ApproverValue, a.UserId, a.Status, a.Action, a.ActedAt, a.ActorName,
       COALESCE(a.ApproverDisplayText, up.DisplayName, a.ActorName, a.ApproverValue) AS ApproverDisplayText
FROM dbo.DocumentApprovals a
LEFT JOIN dbo.UserProfiles up ON up.UserId = a.UserId
WHERE a.DocId = @DocId ORDER BY a.StepOrder;";
                ac.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = id });
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
                        actedAtText = acted.HasValue ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(acted.Value, DateTimeKind.Utc)) : null,
                        actorName = r["ActorName"]?.ToString(),
                        ApproverDisplayText = r["ApproverDisplayText"]?.ToString() ?? string.Empty
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
                createdAt = createdAt.HasValue ? _helper.ToLocalStringFromUtc(DateTime.SpecifyKind(createdAt.Value, DateTimeKind.Utc)) : null,
                descriptorJson = descriptor,
                previewJson = preview,
                approvals
            });
        }

        // ========== Private helpers ==========
        private string ResolveContentRootRelativePath(string? path)
        {
            var raw = (path ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            if (System.IO.Path.IsPathRooted(raw)) return raw;
            raw = raw.TrimStart('\\', '/').Replace('/', System.IO.Path.DirectorySeparatorChar).Replace('\\', System.IO.Path.DirectorySeparatorChar);
            return System.IO.Path.Combine(_env.ContentRootPath, raw);
        }

        private string MakeRelativeToContentRoot(string? absPath)
        {
            var raw = (absPath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
            try { if (System.IO.Path.IsPathRooted(raw)) return System.IO.Path.GetRelativePath(_env.ContentRootPath, raw); }
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
                sc.CommandText = @"SELECT TOP 1 1 FROM dbo.DocumentShares s WITH (NOLOCK) WHERE s.DocId=@DocId AND s.UserId=@UserId AND ISNULL(s.IsRevoked,0)=0 AND (s.ExpireAt IS NULL OR s.ExpireAt>SYSUTCDATETIME());";
                sc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                sc.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });
                var o = await sc.ExecuteScalarAsync();
                isSharedRecipient = (o != null && o != DBNull.Value);
            }
            catch (Exception ex) { _log.LogWarning(ex, "DetailDX: shared recipient check failed. DocId={DocId}", docId); }

            string viewerRole = "Other";
            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"
SELECT CreatedBy,
       CASE WHEN EXISTS (SELECT 1 FROM dbo.DocumentApprovals da WHERE da.DocId=d.DocId AND da.UserId=@UserId) THEN 1 ELSE 0 END AS IsApprover
FROM dbo.Documents d WHERE d.DocId=@DocId;";
                cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                cmd.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });
                await using var r = await cmd.ExecuteReaderAsync();
                if (await r.ReadAsync())
                {
                    var createdBy = r["CreatedBy"] as string;
                    var isApprover = Convert.ToInt32(r["IsApprover"]) == 1;
                    var isCreator = !string.IsNullOrWhiteSpace(createdBy) && string.Equals(createdBy, userId, StringComparison.OrdinalIgnoreCase);
                    if (isSharedTab) viewerRole = "Shared";
                    else if (isCreator && isApprover) viewerRole = "Creator+Approver";
                    else if (isCreator) viewerRole = "Creator";
                    else if (isApprover) viewerRole = "Approver";
                }
            }

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"INSERT INTO dbo.DocumentViewLogs (DocId,ViewerId,ViewerRole,ViewedAt,ClientIp,UserAgent) VALUES (@DocId,@ViewerId,@ViewerRole,SYSUTCDATETIME(),@ClientIp,@UserAgent);";
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
                    uc.CommandText = @"UPDATE dbo.DocumentShares SET IsRead=1 WHERE DocId=@DocId AND UserId=@UserId AND ISNULL(IsRevoked,0)=0 AND (ExpireAt IS NULL OR ExpireAt>SYSUTCDATETIME()) AND ISNULL(IsRead,0)=0;";
                    uc.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                    uc.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 450) { Value = string.IsNullOrWhiteSpace(userId) ? DBNull.Value : userId });
                    await uc.ExecuteNonQueryAsync();
                }
                catch (Exception ex) { _log.LogWarning(ex, "DetailDX: shared read update failed. DocId={DocId}", docId); }
            }
        }

        private async Task<(bool canRecall, bool canApprove, bool canHold, bool canReject, bool canCooperate)> GetApprovalCapabilitiesAsync(string docId)
        {
            var userId = User?.FindFirstValue(ClaimTypes.NameIdentifier) ?? string.Empty;
            if (string.IsNullOrWhiteSpace(docId) || string.IsNullOrWhiteSpace(userId))
                return (false, false, false, false, false);

            var cs = _cfg.GetConnectionString("DefaultConnection") ?? string.Empty;
            await using var conn = new SqlConnection(cs);
            await conn.OpenAsync();

            string? status = null;
            string? createdBy = null;

            await using (var cmd = conn.CreateCommand())
            {
                cmd.CommandText = @"SELECT TOP 1 Status, CreatedBy FROM dbo.Documents WHERE DocId = @id;";
                cmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                await using var r = await cmd.ExecuteReaderAsync();
                if (await r.ReadAsync()) { status = r["Status"] as string ?? string.Empty; createdBy = r["CreatedBy"] as string ?? string.Empty; }
            }

            if (status == null) return (false, false, false, false, false);

            var st = status ?? string.Empty;
            var isCreator = string.Equals(createdBy ?? string.Empty, userId, StringComparison.OrdinalIgnoreCase);
            var isRecalled = st.Equals("Recalled", StringComparison.OrdinalIgnoreCase);
            var isFinal = st.StartsWith("Approved", StringComparison.OrdinalIgnoreCase) || st.StartsWith("Rejected", StringComparison.OrdinalIgnoreCase);
            var canRecall = isCreator && st.StartsWith("Pending", StringComparison.OrdinalIgnoreCase);

            bool canCooperate = false;
            await using (var coopCmd = conn.CreateCommand())
            {
                coopCmd.CommandText = @"SELECT TOP 1 1 FROM dbo.DocumentCooperations WHERE DocId=@id AND UserId=@uid AND ISNULL(Status,N'Pending')=N'Pending';";
                coopCmd.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                coopCmd.Parameters.Add(new SqlParameter("@uid", SqlDbType.NVarChar, 64) { Value = userId });
                var coopObj = await coopCmd.ExecuteScalarAsync();
                canCooperate = (coopObj != null && coopObj != DBNull.Value);
            }

            int currentStep = 0;
            await using (var cmd2 = conn.CreateCommand())
            {
                cmd2.CommandText = @"SELECT MIN(StepOrder) FROM dbo.DocumentApprovals WHERE DocId=@id AND ISNULL(Status,N'Pending') LIKE N'Pending%';";
                cmd2.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                var obj = await cmd2.ExecuteScalarAsync();
                if (obj != null && obj != DBNull.Value) currentStep = Convert.ToInt32(obj);
            }

            if (currentStep == 0 || isFinal || isRecalled)
                return (canRecall, false, false, false, canCooperate);

            string? stepUserId = null;
            string stepStatus = string.Empty;

            await using (var cmd3 = conn.CreateCommand())
            {
                cmd3.CommandText = @"SELECT TOP 1 UserId, ISNULL(Status,N'Pending') AS Status FROM dbo.DocumentApprovals WHERE DocId=@id AND StepOrder=@step;";
                cmd3.Parameters.Add(new SqlParameter("@id", SqlDbType.NVarChar, 40) { Value = docId });
                cmd3.Parameters.Add(new SqlParameter("@step", SqlDbType.Int) { Value = currentStep });
                await using var r = await cmd3.ExecuteReaderAsync();
                if (await r.ReadAsync()) { stepUserId = r["UserId"] as string ?? string.Empty; stepStatus = r["Status"] as string ?? string.Empty; }
            }

            if (string.IsNullOrWhiteSpace(stepUserId))
                return (canRecall, false, false, false, canCooperate);

            var isMyStep = string.Equals(stepUserId, userId, StringComparison.OrdinalIgnoreCase);
            var isPending = stepStatus.StartsWith("Pending", StringComparison.OrdinalIgnoreCase);
            var canDo = isMyStep && isPending && !isRecalled && !isFinal;

            return (canRecall, canDo, canDo, canDo, canCooperate);
        }

        private async Task<string?> GetDocumentAuthorUserIdAsync(SqlConnection conn, string docId)
        {
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"SELECT TOP (1) NULLIF(LTRIM(RTRIM(CreatedBy)),N'') AS CreatedBy FROM dbo.Documents WHERE DocId=@DocId;";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            var obj = await cmd.ExecuteScalarAsync();
            var s = (obj == null || obj == DBNull.Value) ? null : (obj.ToString() ?? string.Empty).Trim();
            return string.IsNullOrWhiteSpace(s) ? null : s;
        }

        private async Task<List<string>> GetNextApproverUserIdsFromDbAsync(SqlConnection conn, string docId, int nextStep)
        {
            var list = new List<string>();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT DISTINCT
       NULLIF(LTRIM(RTRIM(a.UserId)),N'')        AS UserId,
       NULLIF(LTRIM(RTRIM(a.ApproverValue)),N'') AS ApproverValue
FROM dbo.DocumentApprovals a
WHERE a.DocId=@DocId AND a.StepOrder=@Step AND ISNULL(a.Status,N'Pending') LIKE N'Pending%';";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = nextStep });

            var approverValuesToResolve = new List<string>();
            await using (var r = await cmd.ExecuteReaderAsync())
            {
                while (await r.ReadAsync())
                {
                    var uid = r["UserId"] as string;
                    var av = r["ApproverValue"] as string;
                    if (!string.IsNullOrWhiteSpace(uid)) { list.Add(uid.Trim()); continue; }
                    if (!string.IsNullOrWhiteSpace(av)) approverValuesToResolve.Add(av.Trim());
                }
            }

            var toBackfill = new List<(string approverValue, string resolvedUserId)>();
            foreach (var av in approverValuesToResolve.Where(x => !string.IsNullOrWhiteSpace(x)).Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var resolved = await TryResolveUserIdAsync(conn, tx: null, value: av);
                if (!string.IsNullOrWhiteSpace(resolved)) { list.Add(resolved.Trim()); toBackfill.Add((av, resolved.Trim())); }
            }

            foreach (var x in toBackfill
                .Select(t => new { approverValue = t.approverValue.Trim(), resolvedUserId = t.resolvedUserId.Trim(), key = t.approverValue.Trim() + "\u001F" + t.resolvedUserId.Trim() })
                .Where(x => x.approverValue.Length > 0 && x.resolvedUserId.Length > 0)
                .DistinctBy(x => x.key, StringComparer.OrdinalIgnoreCase))
            {
                await using var up = conn.CreateCommand();
                up.CommandText = @"UPDATE dbo.DocumentApprovals SET UserId=@UserId WHERE DocId=@DocId AND StepOrder=@Step AND (UserId IS NULL OR LTRIM(RTRIM(UserId))=N'') AND ApproverValue=@ApproverValue;";
                up.Parameters.Add(new SqlParameter("@UserId", SqlDbType.NVarChar, 64) { Value = x.resolvedUserId });
                up.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
                up.Parameters.Add(new SqlParameter("@Step", SqlDbType.Int) { Value = nextStep });
                up.Parameters.Add(new SqlParameter("@ApproverValue", SqlDbType.NVarChar, 256) { Value = x.approverValue });
                await up.ExecuteNonQueryAsync();
            }

            return list.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static async Task<string?> TryResolveUserIdAsync(SqlConnection conn, SqlTransaction? tx, string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;
            await using var cmd = conn.CreateCommand();
            if (tx != null) cmd.Transaction = tx;
            cmd.CommandText = @"SELECT TOP 1 Id FROM dbo.AspNetUsers WHERE Id=@v OR UserName=@v OR NormalizedUserName=UPPER(@v) OR Email=@v OR NormalizedEmail=UPPER(@v);";
            cmd.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            var id = (string?)await cmd.ExecuteScalarAsync();
            if (!string.IsNullOrWhiteSpace(id)) return id;

            await using var cmd2 = conn.CreateCommand();
            if (tx != null) cmd2.Transaction = tx;
            cmd2.CommandText = @"SELECT TOP 1 COALESCE(p.UserId,u.Id) FROM dbo.UserProfiles p LEFT JOIN dbo.AspNetUsers u ON u.Id=p.UserId WHERE p.UserId=@v OR p.Email=@v OR p.DisplayName=@v OR p.Name=@v;";
            cmd2.Parameters.Add(new SqlParameter("@v", SqlDbType.NVarChar, 256) { Value = value });
            id = (string?)await cmd2.ExecuteScalarAsync();
            return string.IsNullOrWhiteSpace(id) ? null : id;
        }

        private static async Task<List<string>> GetNextApproverEmailsFromDbAsync(SqlConnection conn, string docId, int stepOrder)
        {
            var result = new List<string>();
            await using var cmd = conn.CreateCommand();
            cmd.CommandText = @"
SELECT ApproverValue FROM dbo.DocumentApprovals
WHERE DocId=@DocId AND StepOrder=@StepOrder AND ISNULL(Status,N'Pending') LIKE N'Pending%'
  AND ApproverValue IS NOT NULL AND LTRIM(RTRIM(ApproverValue))<>N'';";
            cmd.Parameters.Add(new SqlParameter("@DocId", SqlDbType.NVarChar, 40) { Value = docId });
            cmd.Parameters.Add(new SqlParameter("@StepOrder", SqlDbType.Int) { Value = stepOrder });
            await using var r = await cmd.ExecuteReaderAsync();
            while (await r.ReadAsync())
            {
                var v = r[0] as string;
                if (!string.IsNullOrWhiteSpace(v)) result.Add(v.Trim());
            }
            return result.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        // ========== DTOs ==========
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
            public List<CommentFileDto>? files { get; set; }
        }

        public sealed class CommentFileDto
        {
            public string? FileKey { get; set; }
            public string? OriginalName { get; set; }
            public string? ContentType { get; set; }
            public long ByteSize { get; set; }
        }

        public sealed class UpdateSharesDto
        {
            public string? DocId { get; set; }
            public List<string>? SelectedRecipientUserIds { get; set; }
        }
    }
}
